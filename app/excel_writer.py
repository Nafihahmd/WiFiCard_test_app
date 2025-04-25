import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

# Define the XLSX file path.
REPORT_FILE = "wifi_test_results.xlsx"

# Define fill colors for pass and fail
GREEN_FILL = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
# Define a grey fill to mark a MAC address as used.
GREY_FILL = PatternFill(start_color="A9A9A9", end_color="A9A9A9", fill_type="solid")

def initialize_workbook():
    """Creates a new workbook with headers if it doesn't exist."""
    if not os.path.exists(REPORT_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "TestResults"
        # Write headers: first column is a timestamp, then one per test.
        headers = ["Timestamp", "MAC Addr", "Status"]
        ws.append(headers)
        wb.save(TEST_REPORT)
        print("Created new workbook with headers.")
    else:
        print("Workbook already exists.")

def append_result(mac, status):
    initialize_workbook()
    wb = load_workbook(REPORT_FILE)
    ws = wb.active
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([ts, mac, status])
    wb.save(REPORT_FILE)