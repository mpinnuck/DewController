import sys
import ctypes
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import time
import serial
import serial.tools.list_ports
import requests
import os
import json

# ------------ Commands to build executable ------------
# .venv\Scripts\activate.ps1
# pyinstaller --onefile --noconsole --name usbdewcontroller --add-data "config.json;." usbdewcontroller.py
# copy dist/usbdewcontroller.exe d:/astro/apps/UsbDewController
# -----------------------------------------------------

# ---------------- CONFIGURATION ----------------
VERSION = "Version: 1.8"
CONFIG_FILE = "config.json"
DEFAULT_DEWSPREAD_THRESHOLD = 3.0  # °C
HYSTERESIS_DEW = 1.0  # °C for heater off
WEATHER_API_URL = "https://api.weather.com/v2/pws/observations/current?stationId=ISYDNEY478&format=json&units=m&apiKey=5356e369de454c6f96e369de450c6f22"
REFRESH_INTERVAL = 5       # seconds for AUTO heater check
HUMIDITY_POLL_INTERVAL = 60  # seconds for fetching current weather
HEATER_ON_COLOR = "#90EE90"       # Light green for heater ON
HEATER_OFF_COLOR = "#FFB6C1"      # Light red/pink for heater OFF

# ---------------- GUI COLORS ----------------
MODE_AUTO_COLOR = "#90EE90"  # Light green for AUTO mode
MODE_MANUAL_COLOR = "SystemButtonFace"  # Default button color

# ---------------- SINGLE INSTANCE (Windows only) ----------------
def check_single_instance():
    mutex_name = "USBDEWCONTROLLER_MUTEX"
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, ctypes.c_bool(False), mutex_name)
    if kernel32.GetLastError() == 183:
        sys.exit(0)

check_single_instance()

# ---------------- GUI APP ----------------
class DewHeaterController(tk.Tk):
    def __init__(self):
        super().__init__()

        # Hide window initially
        self.withdraw()

        # Load persisted config
        self.config_data = self.load_config()

        # State variables
        self.serial_port = None
        self.mode = tk.StringVar(value=self.config_data.get("mode", "AUTO"))
        self.dewspread_threshold = self.config_data.get("dewspread_threshold", DEFAULT_DEWSPREAD_THRESHOLD)
        self.current_dewpoint = tk.DoubleVar(value=0.0)
        self.current_temp = tk.DoubleVar(value=0.0)
        self.current_dewspread = tk.DoubleVar(value=0.0)
        self.current_rh = tk.StringVar(value="0")
        self.heater_on = False
        self.running = True

        # Build the GUI
        self.build_gui()
        self.update_idletasks()
        self.center_window()

        #  Fetch initial weather to populate GUI before threads start
        self.fetch_weather()  

        # Set initial mode button color
        if self.mode.get() == "AUTO":
            self.btn_mode.config(bg=MODE_AUTO_COLOR)
        else:
            self.btn_mode.config(bg=MODE_MANUAL_COLOR)

        # Show window now that everything is ready
        self.deiconify()


        # Threads
        threading.Thread(target=self.auto_monitor, daemon=True).start()
        threading.Thread(target=self.poll_current_weather, daemon=True).start()
        self.refresh_serial_ports()

        # Auto-connect previous port
        self.auto_connect_previous_port()

        # Close protocol
        self.protocol("WM_DELETE_WINDOW", self.on_close)


    def center_window(self):
        """Center the main window on the screen."""
        self.update_idletasks()  # ensure geometry info is up-to-date
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # ---------------- GUI BUILD ----------------
    def build_gui(self):
        self.title("Dew Heater Controller")
        self.geometry("600x450")
        self.minsize(600, 300)

        # ---------------- Row 0: COM Port ----------------
        tk.Label(self, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combobox_ports = ttk.Combobox(self, values=self.get_serial_ports(), state="readonly")
        self.combobox_ports.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        if self.config_data.get("com_port") in self.combobox_ports['values']:
            self.combobox_ports.set(self.config_data["com_port"])

        self.btn_connect = tk.Button(self, text="Connect", command=self.toggle_connection)
        self.btn_connect.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        self.lbl_version = tk.Label(self, text=VERSION, fg="blue")
        self.lbl_version.grid(row=0, column=3, padx=5, pady=5, sticky="e")

        # ---------------- Row 1: Mode + Manual ----------------
        self.btn_mode = tk.Button(self, text=f"Mode: {self.mode.get()}", width=15, command=self.toggle_mode)
        self.btn_mode.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.btn_manual = tk.Button(
            self, text="Turn ON", width=15,
            state="normal" if self.mode.get() == "MANUAL" else "disabled",
            command=self.toggle_manual
        )
        self.btn_manual.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.lbl_heater_status = tk.Label(
            self, text="OFF", width=15, bg="red", fg="black", relief="sunken"
        )
        self.lbl_heater_status.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        # Initialize heater label color
        self.lbl_heater_status.config(
            bg=HEATER_ON_COLOR if self.heater_on else HEATER_OFF_COLOR
        )

        tk.Label(self, text="Current RH %:").grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.label_current_rh = tk.Label(self, textvariable=self.current_rh, bg="#d3d3d3", padx=5, pady=2, relief="sunken")
        self.label_current_rh.grid(row=1, column=4, padx=5, pady=5, sticky="w")

        # ---------------- Row 2: Dew Spread + Temp + Dew Point ----------------
        row2_frame = tk.Frame(self)
        row2_frame.grid(row=2, column=0, columnspan=8, sticky="ew", padx=5, pady=5)

        tk.Label(row2_frame, text="Dew Spread Trigger °C:").pack(side=tk.LEFT, padx=5)
        self.entry_dewspread = tk.Entry(row2_frame, width=5)
        self.entry_dewspread.insert(0, str(self.dewspread_threshold))
        self.entry_dewspread.pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(row2_frame, text="Dew Spread °C:").pack(side=tk.LEFT, padx=5)
        tk.Label(row2_frame, textvariable=self.current_dewspread, bg="#d3d3d3", padx=5, pady=2, relief="sunken").pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(row2_frame, text="Temp °C:").pack(side=tk.LEFT, padx=5)
        tk.Label(row2_frame, textvariable=self.current_temp, bg="#d3d3d3", padx=5, pady=2, relief="sunken").pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(row2_frame, text="Dew Point °C:").pack(side=tk.LEFT, padx=5)
        tk.Label(row2_frame, textvariable=self.current_dewpoint, bg="#d3d3d3", padx=5, pady=2, relief="sunken").pack(side=tk.LEFT, padx=5)

        row2_frame.columnconfigure(0, weight=1)

        # ---------------- Row 3: Log ----------------
        self.log_text = scrolledtext.ScrolledText(self, wrap=tk.WORD)
        self.log_text.grid(row=3, column=0, columnspan=8, padx=5, pady=5, sticky="nsew")
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(7, weight=1)

    # ---------------- Auto-connect previous COM ----------------
    def auto_connect_previous_port(self):
        saved_port = self.config_data.get("com_port")
        if saved_port and saved_port in self.combobox_ports['values']:
            self.combobox_ports.set(saved_port)
            self.log(f"Attempting auto-connect to {saved_port}...")
            try:
                self.serial_port = serial.Serial(saved_port, 9600, timeout=1)
                self.btn_connect.config(text="Disconnect")
                self.log(f"Auto-connected to {saved_port}")
            except Exception as e:
                self.serial_port = None
                self.log(f"Auto-connect failed: {e}")
                
    # ---------------- Helper Methods ----------------
    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def get_serial_ports(self):
        return [port.device for port in serial.tools.list_ports.comports()]

    def refresh_serial_ports(self):
        current_ports = set(self.combobox_ports['values'])
        detected_ports = set(self.get_serial_ports())
        if current_ports != detected_ports:
            selected = self.combobox_ports.get()
            self.combobox_ports['values'] = list(detected_ports)
            if selected in detected_ports:
                self.combobox_ports.set(selected)
            else:
                self.combobox_ports.set('')
                self.log("COM port list updated")
        self.after(3000, self.refresh_serial_ports)

    # ---------------- Serial Connection ----------------
    def toggle_connection(self):
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            self.serial_port = None
            self.btn_connect.config(text="Connect")
            self.log("Disconnected")
        else:
            port = self.combobox_ports.get()
            if not port:
                self.log("No COM port selected")
                return
            try:
                self.serial_port = serial.Serial(port, 9600, timeout=1)
                self.btn_connect.config(text="Disconnect")
                self.log(f"Connected to {port}")
                self.save_config()
            except Exception as e:
                self.serial_port = None
                self.log(f"Connection failed: {e}")

    # ---------------- Mode and Manual ----------------
    def toggle_mode(self):
        if self.mode.get() == "AUTO":
            self.mode.set("MANUAL")
            self.btn_manual.config(state="normal")
            self.btn_mode.config(bg=MODE_MANUAL_COLOR)  # Use constant

            # Update manual heater button text
            if self.heater_on:
                self.btn_manual.config(text="TURN OFF")
            else:
                self.btn_manual.config(text="TURN ON")
        else:
            self.mode.set("AUTO")
            self.btn_manual.config(state="disabled")
            self.btn_mode.config(bg=MODE_AUTO_COLOR)  # Use constant

        self.btn_mode.config(text=f"Mode: {self.mode.get()}")
        self.save_config()
        self.log(f"Mode changed to {self.mode.get()}")

    def toggle_manual(self):
        if self.heater_on:
            self.send_relay_command(False)
            self.btn_manual.config(text="TURN ON")   # update button text
        else:
            self.send_relay_command(True)
            self.btn_manual.config(text="TURN OFF")  # update button text

    # ---------------- Send Relay Command ----------------
    def send_relay_command(self, turn_on):
        if self.serial_port and self.serial_port.is_open:
            try:
                start_id = 0xA0
                switch_addr = 0x01
                op_data = 0x01 if turn_on else 0x00
                checksum = (start_id + switch_addr + op_data) & 0xFF  # simple sum & mask
                cmd = bytes([start_id, switch_addr, op_data, checksum])

                self.serial_port.write(cmd)
                self.heater_on = turn_on

                # Update GUI
                self.lbl_heater_status.config(
                    text="ON" if turn_on else "OFF",
                    bg=HEATER_ON_COLOR if turn_on else HEATER_OFF_COLOR,
                    fg="black"
                )
                self.btn_manual.config(
                    text="TURN OFF" if turn_on else "TURN ON"
                )

                self.log(f"Heater turned {'ON' if turn_on else 'OFF'}")
            except Exception as e:
                self.log(f"Serial write failed: {e}")
        else:
            self.log("Serial port not connected")

    # ---------------- AUTO Monitoring ----------------
    def auto_monitor(self):
        while self.running:
            try:
                if self.mode.get() == "AUTO":
                    try:
                        self.dewspread_threshold = float(self.entry_dewspread.get())
                    except ValueError:
                        self.log("Invalid Dew Spread threshold input")
                        time.sleep(REFRESH_INTERVAL)
                        continue

                    dewspread = self.current_dewspread.get()
                    if not self.heater_on and dewspread <= self.dewspread_threshold:
                        self.send_relay_command(True)
                    elif self.heater_on and dewspread >= (self.dewspread_threshold + HYSTERESIS_DEW):
                        self.send_relay_command(False)

            except Exception as e:
                self.log(f"Auto-monitoring error: {e}")

            time.sleep(REFRESH_INTERVAL)

    
    def fetch_weather(self):
        try:
            response = requests.get(WEATHER_API_URL, timeout=10)
            data = response.json()['observations'][0]
            metric = data['metric']
            
            temp = metric['temp']
            dewpt = metric['dewpt']
            rh = data.get('humidity')

            # Instead of updating Tk variables directly here in the thread:
            self.after(0, self.update_weather_gui, temp, dewpt, rh)

            self.log(f"Weather update: Temp={temp}°C, Dew={dewpt}°C, DewSpread={temp - dewpt}°C, RH={rh}%")
            return temp, dewpt, rh

        except Exception as e:
            self.log(f"Failed to fetch weather: {e}")
            return 0.0, 0.0, 0

    def update_weather_gui(self, temp, dewpt, rh):
        """This runs safely in the GUI thread."""
        self.current_temp.set(round(temp, 1))
        self.current_dewpoint.set(round(dewpt, 1))
        self.current_dewspread.set(round(temp - dewpt, 2))
        if rh is not None:
            self.current_rh.set(round(rh, 1))
    
    # ---------------- Poll Weather ----------------
    def poll_current_weather(self):
        while self.running:
            self.fetch_weather()  # call the same function here
            time.sleep(HUMIDITY_POLL_INTERVAL)

    # ---------------- Config Persistence ----------------
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_config(self):
        try:
            self.config_data["mode"] = self.mode.get()
            self.config_data["dewspread_threshold"] = self.dewspread_threshold
            self.config_data["com_port"] = self.combobox_ports.get()
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f)
        except Exception as e:
            self.log(f"Failed to save config: {e}")

    # ---------------- Clean Exit ----------------
    def on_close(self):
        self.running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.destroy()

# ---------------- MAIN ----------------
if __name__ == "__main__":
    app = DewHeaterController()
    app.mainloop()
