import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import time
import serial
import serial.tools.list_ports
import requests
import os
import json

# ---------------- CONFIGURATION ----------------
CONFIG_FILE = "config.json"
DEFAULT_RH_THRESHOLD = 80
WEATHER_API_URL = "https://api.weather.com/v2/pws/observations/current?stationId=ISYDNEY478&format=json&units=m&apiKey=5356e369de454c6f96e369de450c6f22"
REFRESH_INTERVAL = 5  # seconds for RH check

# ---------------- GUI APP ----------------
class DewHeaterController(tk.Tk):
    def __init__(self):
        super().__init__()

        # Load persisted config
        self.config_data = self.load_config()

        self.title("Dew Heater Controller")
        self.geometry("700x500")
        self.minsize(500, 300)

        # ---------------- State Variables ----------------
        self.serial_port = None
        self.mode = tk.StringVar(value=self.config_data.get("mode", "AUTO"))
        self.rh_threshold = self.config_data.get("rh_threshold", DEFAULT_RH_THRESHOLD)
        self.heater_on = False
        self.running = True

        # ---------------- GUI Layout ----------------
        # Row 0: COM Port selection
        tk.Label(self, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combobox_ports = ttk.Combobox(self, values=self.get_serial_ports(), state="readonly")
        self.combobox_ports.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Restore previous COM port if available
        if self.config_data.get("com_port") in self.combobox_ports['values']:
            self.combobox_ports.set(self.config_data["com_port"])

        self.btn_connect = tk.Button(self, text="Connect", command=self.toggle_connection)
        self.btn_connect.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Row 1: Mode and Manual On/Off
        self.btn_mode = tk.Button(self, text=f"Mode: {self.mode.get()}", width=15, command=self.toggle_mode)
        self.btn_mode.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.btn_manual = tk.Button(self, text="Turn ON", width=15,
                                    state="normal" if self.mode.get() == "MANUAL" else "disabled",
                                    command=self.toggle_manual)
        self.btn_manual.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Row 2: RH Threshold
        tk.Label(self, text="RH Threshold %:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entry_rh = tk.Entry(self, width=5)
        self.entry_rh.insert(0, str(self.rh_threshold))
        self.entry_rh.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Row 3: Log box
        self.log_text = scrolledtext.ScrolledText(self, wrap=tk.WORD)
        self.log_text.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        # Configure resizing
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(2, weight=1)

        # Start threads
        threading.Thread(target=self.auto_monitor, daemon=True).start()
        self.refresh_serial_ports()  # Dynamic COM port updates

        # Bind close event
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- Helper Methods ----------------
    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def get_serial_ports(self):
        return [port.device for port in serial.tools.list_ports.comports()]

    def refresh_serial_ports(self):
        """Update COM port list dynamically"""
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

        self.after(3000, self.refresh_serial_ports)  # refresh every 3 sec

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
                self.log("Select a COM port first")
                return
            try:
                self.serial_port = serial.Serial(port, 9600, timeout=1)
                self.btn_connect.config(text="Disconnect")
                self.log(f"Connected to {port}")
            except Exception as e:
                self.log(f"Failed to connect: {e}")
                self.serial_port = None

    # ---------------- Mode / Manual ----------------
    def toggle_mode(self):
        if self.mode.get() == "AUTO":
            self.mode.set("MANUAL")
            self.btn_manual.config(state="normal")
        else:
            self.mode.set("AUTO")
            self.btn_manual.config(state="disabled")
        self.btn_mode.config(text=f"Mode: {self.mode.get()}")
        self.log(f"Mode changed to {self.mode.get()}")

    def toggle_manual(self):
        if not self.serial_port or not self.serial_port.is_open:
            self.log("Connect to COM port first")
            return
        if self.heater_on:
            self.send_relay_command(False)
        else:
            self.send_relay_command(True)

    def send_relay_command(self, turn_on: bool):
        """Send command to USB relay"""
        try:
            if turn_on:
                self.serial_port.write(b'\xA0\x01\x01\xA2')  # Example ON command
                self.log("Heater turned ON")
            else:
                self.serial_port.write(b'\xA0\x01\x00\xA1')  # Example OFF command
                self.log("Heater turned OFF")
            self.heater_on = turn_on
        except Exception as e:
            self.log(f"Error sending relay command: {e}")

    # ---------------- Humidity Monitoring ----------------
    def auto_monitor(self):
        while self.running:
            try:
                if self.mode.get() == "AUTO":
                    rh_threshold = float(self.entry_rh.get())
                    self.rh_threshold = rh_threshold
                    response = requests.get(WEATHER_API_URL, timeout=5)
                    data = response.json()
                    rh = data['observations'][0]['humidity']
                    self.log(f"Current RH: {rh}%")
                    if rh >= self.rh_threshold and not self.heater_on:
                        self.send_relay_command(True)
                    elif rh < self.rh_threshold and self.heater_on:
                        self.send_relay_command(False)
            except Exception as e:
                self.log(f"Monitoring error: {e}")
            time.sleep(REFRESH_INTERVAL)

    # ---------------- Config Persistence ----------------
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f)
            except Exception as e:
                self.log(f"Error loading config: {e}")
        return {}

    def save_config(self):
        try:
            self.config_data["com_port"] = self.combobox_ports.get()
            self.config_data["mode"] = self.mode.get()
            try:
                self.config_data["rh_threshold"] = float(self.entry_rh.get())
            except ValueError:
                pass
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f, indent=2)
            self.log("Settings saved")
        except Exception as e:
            self.log(f"Error saving config: {e}")

    # ---------------- Close ----------------
    def on_close(self):
        self.running = False
        self.save_config()
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.destroy()

# ---------------- Main ----------------
if __name__ == "__main__":
    app = DewHeaterController()
    app.mainloop()

