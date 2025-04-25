import os
import tkinter as tk
from tkinter import Menu, messagebox
import pyudev
import subprocess
import netifaces
from openpyxl import Workbook, load_workbook
from datetime import datetime

REPORT_FILE = "wifi_test_results.xlsx"

class HardwareTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WiFi Test Application")
        self.root.minsize(800, 600)

        # Configurable parameters
        self.ssid = "Your-SSID"
        self.password = "Your-Password"

        # Test data
        self.interfaces = []      # list of iface names
        self.test_results = {}    # iface -> "PASS"/"FAIL"
        self.buttons = {}         # iface -> Button widget

        self.create_menu()
        self.create_widgets()
        self.refresh_interfaces()

    def create_menu(self):
        """Creates a File menu with options to open/save results, configure test parameters, and access help."""
        menu_bar = Menu(self.root)
        
        # File menu
        file_menu = Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Save Results", command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label="Reset", command=self.reset_tests)
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Configure menu
        menu_bar.add_command(label="Configure", command=self.configure_parameters)

        # Help Menu
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Help", command=self.show_help)
        # help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_separator()
        # help_menu.add_command(label="Contact Support", command=self.contact_support)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menu_bar)

    def create_widgets(self):
        main = tk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True)

        # Left pane for interface buttons
        left = tk.Frame(main, bd=2, relief=tk.SUNKEN)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        tk.Label(left, text="Detected Adapters").pack()
        self.left_frame = left

        # Right pane for logs
        right = tk.Frame(main, bd=2, relief=tk.SUNKEN)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        tk.Label(right, text="Log").pack(anchor="w")
        self.log = tk.Text(right, wrap=tk.WORD, height=15, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True)

        # Run Tests button
        self.run_button = tk.Button(self.root, text="Run Tests", command=self.run_all_tests)
        self.run_button.pack(pady=5)

    def log_message(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def refresh_interfaces(self):
        # 1) announce start
        self.log_message("Refreshing interfaces…")

        # 2) enumerate only 'net' devices that already carry the VID/PID properties
        ctx = pyudev.Context()
        # list_devices(**kwargs) forwards to match_subsystem() & match_property() :contentReference[oaicite:4]{index=4}
        matches = list(ctx.list_devices(
            subsystem="net",
            ID_VENDOR_ID="148f",
            ID_MODEL_ID="7601"
        ))

        # 3) collect interface names
        self.interfaces = [dev.sys_name for dev in matches]

        # 4) debug-log exactly what we found
        if not self.interfaces:
            self.log_message("No MT7601U interfaces found.")
        else:
            self.log_message(f"Found MT7601U interfaces: {self.interfaces}")

    def reset_tests(self):
        """Reset all tests for the next device."""

    def configure_parameters(self):
        """Popup a window to allow user to edit MAC address, serial port, and model number simultaneously."""
        
        # Create a new Toplevel window (acts as a modal dialog)
        window = tk.Toplevel(self.root)
        window.title("Configure Test Parameters")
        
        # --- SSID field ---
        tk.Label(window, text="WiFi SSID:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        ssid_entry = tk.Entry(window, width=30)
        ssid_entry.insert(0, self.ssid)
        ssid_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # --- Password field ---
        tk.Label(window, text="Password:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        password_entry = tk.Entry(window, width=30)
        password_entry.insert(0, self.password)
        password_entry.grid(row=2, column=1, padx=5, pady=5)

        # --- Serial Port field ---
        tk.Label(window, text="Serial Port:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        serial_entry = tk.Entry(window, width=30)
        serial_entry.insert(0, self.serial_port)
        serial_entry.grid(row=3, column=1, padx=5, pady=5)
        
        # --- OK Button to save changes ---
        def on_ok():
            self.ssid = ssid_entry.get()
            self.password = password_entry.get()
            self.serial_port = serial_entry.get()
            # Optionally, you can show a message box or log the changes
            # For example:
            # messagebox.showinfo("Parameters Set",
            #                     f"MAC Address: {self.mac_addr}\n"
            #                     f"IP Address: {self.server_ip}\n"
            #                     f"Serial Port: {self.serial_port}\n"
            #                     f"Model Number: {self.model_number}")
            window.destroy()

        tk.Button(window, text="OK", command=on_ok).grid(row=4, column=0, columnspan=2, pady=10)
        
        # Update window to ensure its size is computed.
        window.update_idletasks()

        # Calculate position: top center relative to self.root
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        win_width = window.winfo_reqwidth()
        # Place the popup centered horizontally relative to the main window,
        # and at the top (or a few pixels offset from the top).
        x = root_x + (root_width // 2) - (win_width // 2)
        y = root_y + 10  # 10 pixels below the top edge of the main window
        window.geometry(f"+{x}+{y}")

        # Make the window modal
        window.grab_set()
    
    def show_help(self):
        """Display help information."""
        messagebox.showinfo("Help", "This application conducts hardware tests via serial port.\n\n"
                             "• Use the File menu to open/save results and configure parameters.\n"
                             "• Use the buttons on the left to run tests.\n"
                             "• For tests requiring manual input, use the Pass/Fail buttons on the right.\n"
                             "• Use the Reconnect button to manually recheck the serial connection.")

    def test_interface(self, iface):
        self.log_message(f"Testing {iface}...")
        # Connect
        ok = self.connect_wifi(iface, self.ssid, self.password)
        # Verify IP
        passed = ok and self.check_ip(iface)
        status = "PASS" if passed else "FAIL"
        self.test_results[iface] = status
        self.log_message(f"{iface}: {status}")
        append_result(self.get_mac(iface), status)
        # Update button color
        color = "green" if passed else "red"
        self.buttons[iface].config(bg=color)

    def run_all_tests(self):
        if not self.interfaces:
            messagebox.showwarning("No Adapters", "No MT7601U adapters found.")
            return
        for iface in self.interfaces:
            self.test_interface(iface)

    def save_results(self):
        messagebox.showinfo("Saved", f"Results saved to {REPORT_FILE}")

    @staticmethod
    def connect_wifi(iface, ssid, password):
        try:
            subprocess.run(["nmcli", "radio", "wifi", "on"], check=True)
            subprocess.run(["nmcli", "device", "wifi", "rescan"], check=True)
            subprocess.run(
                ["nmcli", "device", "wifi", "connect", ssid, "password", password, "ifname", iface],
                check=True, timeout=20
            )
            return True
        except subprocess.SubprocessError:
            return False

    @staticmethod
    def check_ip(iface):
        addrs = netifaces.ifaddresses(iface)
        return netifaces.AF_INET in addrs and bool(addrs[netifaces.AF_INET])

    @staticmethod
    def get_mac(iface):
        addrs = netifaces.ifaddresses(iface)
        link = addrs.get(netifaces.AF_LINK)
        return link[0]['addr'] if link else "Unknown"

def initialize_workbook():
    if not os.path.exists(REPORT_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "TestResults"
        ws.append(["Timestamp", "MAC Address", "Status"])
        wb.save(REPORT_FILE)

def append_result(mac, status):
    initialize_workbook()
    wb = load_workbook(REPORT_FILE)
    ws = wb.active
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([ts, mac, status])
    wb.save(REPORT_FILE)

if __name__ == "__main__":
    initialize_workbook()
    root = tk.Tk()
    app = HardwareTestApp(root)
    root.mainloop()
