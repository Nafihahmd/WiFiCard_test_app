"""Code for logging results.

This module provides functions to:
  * Initialize or load the Excel report workbook.
  * Append timestamped test results (MAC, status) to the workbook.

Public API:
    initialize_workbook
    append_result
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
import os

# Define the XLSX file path.
REPORT_FILE = "wifi_test_results.xlsx"
# Define fill colors for pass and fail
GREEN_FILL = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

def initialize_workbook():    
    """Create the Excel report file with headers if it does not exist.

    The workbook will be named by REPORT_FILE and contain a "TestResults" sheet
    with columns: Timestamp, MAC, Status.
    """
    if not os.path.exists(REPORT_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "TestResults"
        ws.append(["Timestamp", "MAC", "Status"])
        wb.save(REPORT_FILE)
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