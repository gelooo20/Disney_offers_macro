import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
from datetime import datetime, timedelta


# ============================================================
# CONFIGURATION (FILES THAT NEVER CHANGE)
# ============================================================

SCHEDULING_XLSX = r"Disney Creative Scheduling.xlsx"
MACRO_XLSM = r"Disney_CreativeQA_Macro[1].xlsm"


# ============================================================
# HELPER: AUTO-DETECT FY SHEET NAME
# ============================================================

def get_disney_schedule_sheet_name():
    """
    Jan 1, 2025 -> FY25_Disney_Creative
    Jan 1, 2026 -> FY26_Disney_Creative
    """
    year = datetime.now().year
    fy = str(year)[-2:]
    return f"FY{fy}_Disney_Creative"


# ============================================================
# MAIN LOGIC
# ============================================================

def run_report(reports_csv):
    import xlwings as xw

    print("Start")

    current_directory = os.getcwd()

    first_book = xw.Book(reports_csv)
    second_book = xw.Book(SCHEDULING_XLSX)
    third_book = xw.Book(MACRO_XLSM)

    # ---- Auto-detect FY sheet ----
    scheduling_sheet_name = get_disney_schedule_sheet_name()

    if scheduling_sheet_name not in [s.name for s in second_book.sheets]:
        raise Exception(
            f"Sheet '{scheduling_sheet_name}' not found in Disney Scheduling document."
        )

    scheduling_doc = second_book.sheets[scheduling_sheet_name]
    reports_sheet = first_book.sheets['reports']
    macro_scheduling_doc = third_book.sheets['GOOGLE DOCS HERE']
    macro_reports = third_book.sheets['CREATIVE CHECK DAILY RPT HERE']
    macro_page = third_book.sheets['GENERATE REPORTS']

    # Scheduling data
    campaign_range = scheduling_doc.range('A:A').value
    adConcept_range = scheduling_doc.range('B:B').value
    creative_doc_range = scheduling_doc.range('E:E').value
    start_range = scheduling_doc.range('G:G').value
    end_range = scheduling_doc.range('H:H').value

    # Reports data
    date_range = reports_sheet.range('A:A').value
    creative_range = reports_sheet.range('B:B').value
    creative_id_range = reports_sheet.range('C:C').value
    totalads_range = reports_sheet.range('D:D').value

    # ========================================================
    # Paste values (PRESERVE 0 VALUES)
    # ========================================================

    macro_scheduling_doc.range('A1').value = [[i] for i in campaign_range if i is not None]
    macro_scheduling_doc.range('B1').value = [[i] for i in adConcept_range if i is not None]
    macro_scheduling_doc.range('E1').value = [[i] for i in creative_doc_range if i is not None]
    macro_scheduling_doc.range('F1').value = [[i] for i in start_range if i is not None]
    macro_scheduling_doc.range('G1').value = [[i] for i in end_range if i is not None]

    macro_reports.range('A1').value = [[i] for i in date_range if i is not None]
    macro_reports.range('B1').value = [[i] for i in creative_range if i is not None]
    macro_reports.range('C1').value = [[i] for i in creative_id_range if i is not None]
    macro_reports.range('D1').value = [[i] for i in totalads_range if i is not None]

    macro_page.range('B10').value = current_directory
    macro_page.range('I10').value = "Pass 1"

    # First macro
    third_book.macro('Module1.GenerateDisneyReports')()

    # ========================================================
    # Replace MAUI codes (PRESERVE "0")
    # ========================================================

    search_items = [
        'WDWDOM', 'WDWEPCOT', 'WDWFLRES', 'WDWRSTS',
        'WDWDHS', 'DSPNGS', 'WDWRES', 'WDWEEC',
        'WDWLATAM', 'CONSUMER', '320x50', '0', 'DIQF'
    ]

    for i, val in enumerate(macro_scheduling_doc.range('C:C').value):
        if val in search_items:
            col_e = macro_scheduling_doc.range(f'E{i+1}').value
            if col_e is not None:
                parts = col_e.split('_')
                if len(parts) >= 4:
                    macro_scheduling_doc.range(f'C{i+1}').value = parts[3]
                else:
                    macro_scheduling_doc.range(f'C{i+1}').value = val  # keep original

    # Second macro
    formatted_date = datetime.now().strftime('%m.%d.%Y')
    macro_page.range('I10').value = f"{formatted_date}_Creative_QA_Report"
    third_book.macro('Module1.GenerateDisneyReports')()

    # ========================================================
    # Clean output
    # ========================================================

    file_name = f"{formatted_date}_Creative_QA_Report.xlsx"
    output_report = xw.Book(file_name)

    RemoveRotation = output_report.sheets['Remove From Rotation']
    ManualChecking = output_report.sheets['Manual Checking Needed']

    RemoveRotation.range('A1').value = (
        'There are no creatives (having a minimum of 150 impressions) '
        'in rotation past their end dates.'
    )
    ManualChecking.range('A1').value = (
        'The below creatives have multiple end dates associated with '
        'the given MAUI code and job number.'
    )

    def rgb(r, g, b):
        return (r << 16) + (g << 8) + b

    yesterday = (datetime.now() - timedelta(days=1)).date()
    used_range = RemoveRotation.range('G3').current_region

    for cell in used_range:
        if isinstance(cell.value, (int, float)) and cell.value <= 150:
            cell.api.Interior.Color = rgb(255, 0, 0)
        elif hasattr(cell.value, "date") and cell.value.date() == yesterday:
            cell.offset(0, -1).api.Interior.Color = rgb(255, 255, 0)

    output_report.save()
    output_report.close()
    first_book.close()
    second_book.close()

    print("Program finished.")


# ============================================================
# UI – ONE BUTTON
# ============================================================

def run_clicked():
    reports_csv = filedialog.askopenfilename(
        title="Select reports.csv",
        filetypes=[("CSV Files", "*.csv")]
    )

    if not reports_csv:
        return

    status_var.set("Running...")
    run_btn.config(state="disabled")

    def task():
        try:
            run_report(reports_csv)
            status_var.set("Completed successfully ✅")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            status_var.set("Failed ❌")
        finally:
            run_btn.config(state="normal")

    threading.Thread(target=task).start()


root = tk.Tk()
root.title("Disney Creative QA Automation")
root.geometry("420x160")
root.resizable(False, False)

tk.Label(
    root,
    text="Select reports.csv and run the QA process",
    font=("Segoe UI", 10)
).pack(pady=15)

run_btn = tk.Button(
    root,
    text="Run Creative QA Report",
    height=2,
    width=30,
    command=run_clicked
)
run_btn.pack()

status_var = tk.StringVar(value="Idle")
tk.Label(root, textvariable=status_var, fg="blue").pack(pady=10)

root.mainloop()
