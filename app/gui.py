"""Tkinter GUI for orchestrating Wi-Fi adapter tests and displaying results.

This application:
  * Scans for MT7601U mPCIe Wi-Fi adapters.
  * Allows per-adapter or batch testing (connect + IP check).
  * Displays real-time logs, progress bar, and LED status.
  * Saves results to an Excel workbook via test_engine API.
"""

import tkinter as tk
from tkinter import Menu, messagebox, ttk
import tkinter.font as tkFont
import test_engine        # engine module
from test_engine import scan_interfaces, connect_wifi, check_ip, get_mac, random_mac
from excel_writer import append_result

import configparser
from appdirs import user_config_dir
import os

TESTING_ENABLED = True   # ← flip to False to disable result simulation
if TESTING_ENABLED:
    import time # For testing purposes, simulate connection time
# Load configuration
cfg = configparser.ConfigParser()
path = os.path.join(user_config_dir("WiFiTestApp","ECSI"), "settings.ini")

class HardwareTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WiFi Test Application")
        self.root.minsize(800, 600)

        # Configurable parameters
        self.ssid = ""
        self.password = ""
        self.show_progress_bar = False    # whether to display the progress bar
        self.auto_save_enabled = False   # whether to save results without prompting
        self.simulate_result = False # whether to enable result simulation 
        
        # Load configuration
        self.load_config()

        # Test data
        self.interfaces = []      # list of interface names
        self.test_results = {}    # iface -> "PASS"/"FAIL"
        self.buttons = {}         # iface -> Button widget
        self.iface_buttons = {}    # iface -> Interface button widget

        self.create_menu()
        self.create_widgets()
        self.refresh_interfaces()

    def create_menu(self):
        """Creates a File menu with options to open/save results, configure test parameters, and access help."""
        menu_bar = Menu(self.root)
        
        # File menu
        file_menu = Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Save Results", command=self.prompt_save_results)
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
        heading_font = tkFont.Font(family="Helvetica", size=11, weight="normal")

        # Left pane for interface buttons
        left = tk.Frame(main, bd=2, relief=tk.SUNKEN)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        self.left_frame = left

        # -- Heading + LED indicator canvas ------------------
        header_frame = tk.Frame(self.left_frame)
        header_frame.pack(fill=tk.X, pady=(0,10), padx=0)

        lbl = tk.Label(
            header_frame,
            text="Detected Devices",
            pady=5,
            font=heading_font,
            bg="#e0e0e0", bd=1, relief=tk.SOLID
        )
        lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # LED: tiny canvas for colored dot
        self.led = tk.Canvas(
            header_frame, width=25, height=25, highlightthickness=0, 
            bg="#e0e0e0", bd=1, relief=tk.SOLID)
        # draw circle centered in that 16×16
        self.led_dot = self.led.create_oval(6,6,20,20, fill="red")   # default red
        self.led.pack(side=tk.RIGHT)
        # Sub‐frame that will contain only the dynamic interface buttons
        self.iface_frame = tk.Frame(self.left_frame)
        self.iface_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0,10))

        # -- Rescan button ----------------
        self.rescan_btn = tk.Button(
            self.left_frame,
            text="Rescan",
            command=self.refresh_interfaces,
            width=15
        )
        # pack at bottom, with padding
        self.rescan_btn.pack(side=tk.BOTTOM, pady=10, padx=5)

        # Right pane for logs
        right = tk.Frame(main, bd=2, relief=tk.SUNKEN)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        tk.Label(right, text="Log").pack(anchor="w")
        self.log = tk.Text(right, wrap=tk.WORD, height=15, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True)

        # Progress bar (below the Run Tests button)
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            self.root,
            variable=self.progress_var,
            maximum=1,            # will reset later to number of interfaces
            mode='determinate',   # shows exact progress
        )
        if self.show_progress_bar:
            self.progress.pack(fill=tk.X, padx=5, pady=5)
        else:
            self.progress.pack_forget()

        # Run Tests button
        self.run_button = tk.Button(self.root, text="Test All", command=self.run_all_tests)
        self.run_button.pack(pady=5)

    def log_message(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
        # Force GUI to update
        self.root.update_idletasks()
    
    def clear_log(self):
        self.log.configure(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.configure(state=tk.DISABLED)
        self.progress_var.set(0)    # reset progress bar

    def reset_tests(self):
        """Reset all tests for the next device."""
        self.clear_log()

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

        # create a container frame in dialog
        opts_frame = tk.Frame(window)
        opts_frame.grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=5)

        # --- Progress Bar Checkbox ---
        show_progress_var = tk.BooleanVar(value=self.show_progress_bar)
        progress_cb = tk.Checkbutton(opts_frame, text="Progress Bar", variable=show_progress_var)
        progress_cb.grid(row=3, column=1, sticky="w", padx=5, pady=5)

        # --- Auto-Save Checkbox ---
        auto_save_var = tk.BooleanVar(value=self.auto_save_enabled)
        auto_save_cb = tk.Checkbutton(opts_frame, text="Auto-Save", variable=auto_save_var)
        auto_save_cb.grid(row=3, column=2, sticky="w", padx=5, pady=5)

        # --- Auto-Save Checkbox ---
        sim_rslt_var = tk.BooleanVar(value=self.simulate_result)
        sim_rslt_cb = tk.Checkbutton(opts_frame, text="Result Simulation", variable=sim_rslt_var)
        sim_rslt_cb.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        # --- OK Button to save changes ---
        def on_ok():
            cfg["network"]["ssid"] = ssid_entry.get()
            cfg["network"]["password"] = password_entry.get()
            cfg["ui"]["show_progress"] = str(show_progress_var.get())
            cfg["ui"]["auto_save"] = str(auto_save_var.get())
            cfg["ui"]["simulate_result"] = str(sim_rslt_var.get())
            with open(path, "w") as f:
                cfg.write(f)

            self.load_config()
            # Optionally, you can show a message box or log the changes
            # For example:
            # messagebox.showinfo("Parameters Set",
            #                     f"MAC Address: {self.mac_addr}\n"
            #                     f"IP Address: {self.server_ip}\n"
            #                     f"Serial Port: {self.serial_port}\n"
            #                     f"Model Number: {self.model_number}")
            self.update_widgets()  # Update the main window with new settings
            window.destroy()

        tk.Button(window, text="OK", command=on_ok).grid(row=5, column=0, columnspan=2, pady=10)
        
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
        messagebox.showinfo(
            "Help",  # Title in title case, no end punctuation
            "This application tests mPCIe-connected MT7601U Wi-Fi adapters by connecting each to a configured network and verifying IP assignment.\n\n"
            "• File → Save Results: write current test outcomes to the report file.\n" 
            "• Configure: set Wi-Fi SSID and password\n"     
            "• Run Tests: detect all connected adapters, attempt connections, and check for IPs.\n"  
            "• Progress bar: shows live count of completed tests; updates in real time.\n"      
            "• Logs: view step-by-step status in the text area for troubleshooting.\n"           
            "• Interface Buttons: click an individual interface to retest a single dongle.\n"         
            "• Rescan: re-scan USB and serial/Wi-Fi interfaces on demand.\n"                    
            "• Color codes: green=PASS, red=FAIL; results saved automatically to Excel.\n"      
            "• Close this dialog by clicking OK or the × in the corner.\n"                    
        )

    def rescan_interfaces(self):
        self.log_message("Rescanning interfaces…")
        self.reset_tests()
        self.refresh_interfaces()
        self.log_message("Rescan complete.")

    def refresh_interfaces(self):
        self.clear_log()
        self.log_message("Refreshing interfaces…")

        # collect interface names
        if self.simulate_result:
            self.interfaces = ['wlan0', 'wlan1']  # For testing purposes, hardcoded to wlan0 and wlan1
        else:
            self.interfaces = scan_interfaces()

        # debug-log exactly what we found
        if not self.interfaces:
            self.log_message("No MT7601U interfaces found.")
            self.led.itemconfig(self.led_dot, fill="red")  # red
        else:
            self.log_message(f"Found {len(self.interfaces)} WiFi devices: {self.interfaces}")
            self.led.itemconfig(self.led_dot, fill="green")  # green

        # Clear old buttons
        for widget in self.iface_frame.winfo_children():
            widget.destroy()

        # Create one button per interface
        for iface in self.interfaces:
            btn = tk.Button(
                self.iface_frame,
                text=iface,
                width=15,
                command=lambda i=iface: self.test_interface(i)  # bind current iface
            )
            btn.pack(pady=2)
            self.iface_buttons[iface] = btn            
            self.test_results[iface] = "Pending" # initialize status

    def test_interface(self, iface):
        self.log_message(f"Testing {iface}...")
        # Connect
        if not self.simulate_result:
            ok = connect_wifi(iface, self.ssid, self.password)
            self.log_message(f"{iface} WiFi connection {'succeeded' if ok else 'failed'}")
            # Verify IP
            passed = ok and check_ip(iface)
        else:
            time.sleep(2)  # Simulate connection time 
            passed = True
        status = "PASS" if passed else "FAIL"
        self.test_results[iface] = status
        self.log_message(f"{iface}: {status}")
        
        # Update button color
        color = "lightgreen" if passed else "red"
        self.iface_buttons[iface].config(bg=color)

        # If all tests have finished, prompt to save results.        
        if all(status in ("PASS","FAIL") for status in self.test_results.values()):
            if self.auto_save_enabled:
                self.save_results()
            else:
                self.prompt_save_results()

    def run_all_tests(self):
        self.refresh_interfaces()
        if not self.interfaces:
            messagebox.showwarning("No Adapters", "No MT7601U adapters found.")
            return
        # reset progress bar
        total = len(self.interfaces)
        self.progress.config(maximum=total)
        self.progress_var.set(0)

        for iface in self.interfaces:
            self.test_interface(iface)
            if self.simulate_result:
                time.sleep(1)  # Simulate time taken for each test
            # advance progress bar immediately
            self.progress_var.set(self.progress_var.get() + 1)
            self.root.update_idletasks()        # ensure bar moves in real time            

    def update_widgets(self):
        # Update the interface buttons 
        self.refresh_interfaces()
        # Show or hide progress bar based on the setting
        if self.show_progress_bar:
            self.progress.pack(fill=tk.X, padx=5, pady=5)
        else:
            self.progress.pack_forget()
    
    def save_results(self):
        for i in self.interfaces:
            status = self.test_results.get(i, "Unknown")
            mac = get_mac(i) if not self.simulate_result else random_mac()
            append_result(mac, status)
        
        if(self.auto_save_enabled):
            self.log_message("Results saved\n")

    def prompt_save_results(self):
        """Prompt the user to save test results once all tests are completed."""
        if messagebox.askyesno("Save Results", "All devices tested. Do you want to save the results?"):
            self.save_results()

    def load_config(self):
        # read settings.ini
        # Load or create defaults
        if os.path.exists(path):
            cfg.read(path)
        else:
            cfg["network"] = {"ssid": "ssid", "password": "psswd"}
            cfg["ui"] = {"show_progress": "no", "auto_save": "no", "simulate_result": "no"}
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w") as f:
                cfg.write(f)
        self.ssid = cfg["network"]["ssid"]
        self.password = cfg["network"]["password"]
        self.show_progress_bar = cfg.getboolean("ui","show_progress")
        self.auto_save_enabled = cfg.getboolean("ui","auto_save")
        self.simulate_result = cfg.getboolean("ui","simulate_result")
            

if __name__ == "__main__":
    root = tk.Tk()
    app = HardwareTestApp(root)
    root.mainloop()