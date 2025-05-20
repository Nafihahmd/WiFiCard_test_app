import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import re

bullet = "\u2022"  # Unicode for bullet character
diamond = "\u25C6"  # Unicode for diamond character
arrow = "\u2192"  # Unicode for arrow character
star = "\u2605"  # Unicode for star character
check = "\u2713"  # Unicode for check mark character
cross = "\u274C"  # Unicode for cross mark character

class HelpCenter(tk.Toplevel):
    """A Help Center window with a tree-style navigation pane with styled multi-line content."""
    BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Help Center")
        self.geometry("600x400")
        self.resizable(False, False)

        # Grid configuration
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        self.rowconfigure(0, weight=1)

        # Navigation tree
        self.tree = ttk.Treeview(self, show="tree")
        topics = [
            ("running", "Running Test"),
            ("rescan", "Rescan"),
            ("saving", "Saving Results"),
            ("logging", "Logging"),
            ("config", "Configure"),
            ("simulating", "Simulating Test")
        ]
        for key, label in topics:
            self.tree.insert("", "end", key, text=label)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10,5), pady=10)

        # ScrolledText for rich styled content
        self.text = ScrolledText(self, wrap="word", state="disabled")
        self.text.grid(row=0, column=1, sticky="nsew", padx=(5,10), pady=10)

        # Style tags
        self.text.tag_configure("heading", font=(None, 14, "bold"), foreground="#003366", justify="center")
        self.text.tag_configure("subheading", font=(None, 11, "bold"), foreground="darkblue", spacing1=5)
        self.text.tag_configure("body", font=(None, 11), foreground="#333333")
        self.text.tag_configure("highlight", font=(None, 11, "italic"), foreground="darkgreen")
        self.text.tag_configure("warning", font=(None, 11, "normal"), foreground="red", spacing3=5)
        self.text.tag_configure("bold", font=(None, 11, "bold"), foreground="black")

        # Content dict: lists of (tag, text) tuples, with \n for line breaks
        self.content = {
            "running": [
                ("heading", "Running Tests\n"),
                ("body", (
                    "When the application starts, it automatically scans for MT7601U adapters.\n"
                )),
                ("subheading", "No Adapters Found\n"),
                ("body", "If no interfaces are detected, you’ll see:\n"),
                ("warning", "No MT7601U interfaces found.\n"),
                ("subheading", "Adapters Detected\n"),
                ("body", (
                    "A green indicator next to “Detected Devices” confirms that at least one adapter is available. "
                    "Each detected interface appears as a button under “Detected Devices.”\n"
                )),
                ("subheading", "Selecting Devices\n"),
                ("body", (
                    "To test a single adapter, click its corresponding button. "
                    "To test all available adapters simultaneously, click “Test All.”\n"
                )),
                ("subheading", "Test Results\n"),
                ("highlight", "Success: The button for a passing test turns green.\n"),
                ("warning", "Failure: The button for a failed test turns red.\n"),
            ],
            "rescan": [
                ("heading", "Refreshing Device List\n"),
                ("body", (
                    "Press the **Rescan** button to update the list whenever you've added or removed an adapter.\n"
                    "The list also refreshes automatically when you click **Test All** or change any configuration settings.\n\n"
                )),
                ("highlight", "Tip: Verify all adapters appear before clicking “Test All” to avoid missed devices.\n"),
            ],
            "saving": [
                ("heading", "Saving Test Results\n"),
                ("body", (
                    "After all tests finish, you’ll be prompted to save your results as “wifi_test_results.xlsx” in the current folder.\n"
                )),
                ("subheading", "Manual Save\n"),
                ("body", (
                    "At any time, choose **Save** from the File menu to write the latest results to disk.\n"
                )),
                ("subheading", "Automatic Save\n"),
                ("body", (
                    "Enable **Auto Save** in the Configuration panel to have each test’s outcome saved immediately.\n"
                )),
            ],
            "logging": [
                ("heading", "Logging\n"),
                ("subheading", "Log Types\n"),
                ("body", (
                    f"{diamond} **Application Log:** Displays high‑level test status in the log pane.\n"
                    f"{diamond} **Terminal Log:** Streams detailed, step‑by‑step messages to the console.\n"
                    f"{diamond} **Saved Log:** Archives all entries to files for post‑test troubleshooting.\n"
                )),
                ("subheading", "Log Files\n"),
                ("body", (
                    "Inside the log folder you’ll find:\n"
                    f"  {bullet} runtime.log\n"
                    f"  {bullet} error.log\n"
                    f"  {bullet} crash.log\n"
                    "Refer to these files when diagnosing bugs or unexpected crashes.\n"
                )),
            ],
            "config": [
                ("heading", "Configuring Test Parameters\n"),
                ("body", (
                    "Click the **Configure** button to adjust your test settings.\n"
                )),
                ("subheading", "Available Settings\n"),
                ("body", (
                    "1. **SSID:** Network name to test.\n"
                    "2. **Password:** Network password.\n"
                    "3. **Auto Save:** Save results automatically after tests complete.\n"
                    "4. **Progress Bar:** Display a progress bar during testing.\n"
                    "5. **Result Simulation:** Run a dry‑run without actual hardware tests.\n"
                )),
                ("warning", "Note: Any change to these settings triggers an automatic rescan of adapters.\n"),
            ],
            "simulating": [
                ("heading", "Simulating Test\n"),
                ("body", (
                    "Use **Simulation Mode** to run tests without connecting to actual hardware. "
                    "The application generates logs and progress feedback, but no real network connections are made.\n"
                )),
                ("warning", "Warning: Simulation Mode does not change or report real adapter status.\n"),
                ("subheading", "When to Use Simulation Mode\n"),
                ("body", (
                    "- Safe debugging of test sequences without risking hardware.  \n"
                    "- Demonstrations and training sessions where real adapters aren’t required.\n"
                )),
            ]
        }

        # Event binding
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        # Select default
        self.tree.selection_set("running")
        self.on_select(None)

    def on_select(self, event):
        key = self.tree.selection()[0]
        lines = self.content.get(key, [])
        # Update text widget
        self.text.config(state="normal")
        self.text.delete("1.0", tk.END)
        for tag, text in lines:
            self._insert_markdown_bold(text, tag)
            # self.text.insert(tk.END, text, tag)
        self.text.config(state="disabled")

    def _insert_markdown_bold(self, full_text, tag="body"):
        """Parse full_text for **…** spans and insert into Text with bold tags."""
        idx = 0
        for m in self.BOLD_RE.finditer(full_text):
            # Plain text before **
            pre = full_text[idx:m.start()]
            if pre:
                self.text.insert(tk.END, pre, tag)
            # Text inside ** **
            bold_chunk = m.group(1)
            self.text.insert(tk.END, bold_chunk, "bold")
            idx = m.end()
        # Any remaining text after last **
        tail = full_text[idx:]
        if tail:
            self.text.insert(tk.END, tail, tag)

# class App(tk.Tk):
#     def __init__(self):
#         super().__init__()
#         self.title("Wi‑Fi Adapter Tester")
#         self.geometry("300x200")
#         ttk.Button(self, text="Help Center", command=lambda: HelpCenter(self)).pack(pady=20)

# if __name__ == '__main__':
#     App().mainloop()
