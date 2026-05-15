import os
import shutil
import sys
import tempfile
import time
from datetime import datetime, timedelta
import pandas as pd
import win32com.client as win32
from win32com.client import dynamic

# FIX 4: Use dynamic path based on script location instead of hardcoded D:\
BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zonal Zodiac Automation")

# Dependency files (MUST open first)
DEPENDENCY_FILES = [
    "Final Zodiac Automate.xlsx",
    "Key Surgery.xlsx",
    "SLR_Automate.xlsx"
]

# Main dashboard file
WORKBOOK = "Zonal Dashboard Format Automation.xlsx"
SHEET = "Monday_Dashboard"
RANGE = "C6:AO50"

TO = ""
CC = ""
CONTROL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "control.xlsx")
RESET_UNITS = "MHW,MVB,MSB,MHS,MBB,MBW,MIW,MSC"


def report_date_str():
    return (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")


def is_win32com_cache_error(exc):
    msg = str(exc)
    return ("CLSIDToClassMap" in msg) or ("CLSIDToPackageMap" in msg)


def clear_win32com_cache():
    candidates = set()

    try:
        import win32com  # Local import keeps startup side effects minimal
        candidates.add(win32com.__gen_path__)
    except Exception:
        pass

    candidates.add(os.path.join(tempfile.gettempdir(), "gen_py"))

    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}"

    for base in candidates:
        if not base:
            continue

        # base can be ...\gen_py or ...\gen_py\3.xx depending on pywin32 state.
        if os.path.basename(base) == py_ver:
            targets = [base]
        else:
            targets = [os.path.join(base, py_ver), base]

        for target in targets:
            if os.path.exists(target):
                shutil.rmtree(target, ignore_errors=True)


def create_excel_app():
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except AttributeError as exc:
        if is_win32com_cache_error(exc):
            print("[WARN] win32com cache is corrupted. Rebuilding cache and retrying Excel COM.")
            clear_win32com_cache()
            try:
                return win32.gencache.EnsureDispatch("Excel.Application")
            except AttributeError as retry_exc:
                if is_win32com_cache_error(retry_exc):
                    print("[WARN] Cache rebuild did not fully recover. Falling back to dynamic Excel COM dispatch.")
                    return dynamic.DumbDispatch("Excel.Application")
                raise
        raise


def export_range_as_image(xl_app, wb, sheet_name, rng_addr, out_path):
    ws = wb.Worksheets(sheet_name)

    wb.Activate()
    ws.Activate()

    xl_app.Visible = True
    xl_app.ScreenUpdating = True
    xl_app.CutCopyMode = False

    time.sleep(2)

    rng = ws.Range(rng_addr)

    rng.CopyPicture(Appearance=1, Format=2)
    time.sleep(1)

    chart_obj = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
    chart = chart_obj.Chart

    # Retry paste
    for _ in range(5):
        try:
            chart.Paste()
            break
        except:
            time.sleep(1)
    else:
        chart_obj.Delete()
        raise Exception("Failed to paste image into chart.")

    time.sleep(1)

    chart.Export(out_path)
    chart_obj.Delete()


def send_email_with_inline_image(image_path, subject, body_text):
    outlook = dynamic.DumbDispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = TO
    mail.CC = CC
    mail.Subject = subject

    attachment = mail.Attachments.Add(image_path)
    cid = "dashboard_image"

    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        cid
    )

    html_body = (
        body_text.replace("\n", "<br>")
        + f"<br><br><img src='cid:{cid}'><br>"
    )

    mail.HTMLBody = html_body
    mail.Send()


def reset_control_workbook():
    today_text = datetime.now().strftime("%d-%m-%Y  00:00:00")

    main_df = pd.read_excel(CONTROL_PATH, sheet_name="Main", header=None, index_col=0)
    report_df = pd.read_excel(CONTROL_PATH, sheet_name="Report")

    main_df.loc["Units", 1] = RESET_UNITS

    if "Download" not in report_df.columns:
        report_df["Download"] = "TRUE"
    else:
        report_df["Download"] = "TRUE"

    if "From Date" not in report_df.columns:
        report_df["From Date"] = today_text
    else:
        report_df["From Date"] = today_text

    if "To Date" not in report_df.columns:
        report_df["To Date"] = today_text
    else:
        report_df["To Date"] = today_text

    with pd.ExcelWriter(CONTROL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        main_df.to_excel(writer, sheet_name="Main", header=False)
        report_df.to_excel(writer, sheet_name="Report", index=False)


def main():
    xl = create_excel_app()

    xl.Visible = True
    xl.DisplayAlerts = False
    xl.ScreenUpdating = True
    xl.AskToUpdateLinks = False

    opened_workbooks = []

    try:
        # 🔥 Open dependency files first
        for file_name in DEPENDENCY_FILES:
            file_path = os.path.join(BASE_DIR, file_name)

            if os.path.isfile(file_path):
                wb_dep = xl.Workbooks.Open(file_path, UpdateLinks=0)
                opened_workbooks.append(wb_dep)
                print(f"[OK] Opened dependency: {file_name}")
            else:
                print(f"[WARN] Missing dependency file: {file_name}")

        # Give Excel time to calculate links
        xl.Calculate()
        time.sleep(3)

        # Open main dashboard
        dashboard_path = os.path.join(BASE_DIR, WORKBOOK)

        if not os.path.isfile(dashboard_path):
            print(f"[ERROR] Dashboard file not found.")
            return

        wb_dashboard = xl.Workbooks.Open(dashboard_path, UpdateLinks=1)
        opened_workbooks.append(wb_dashboard)

        time.sleep(3)

        image_path = os.path.join(
            BASE_DIR,
            f"zonal_dashboard_{report_date_str()}.png"
        )

        export_range_as_image(
            xl,
            wb_dashboard,
            SHEET,
            RANGE,
            image_path
        )

    finally:
        # Close all workbooks safely
        for wb in opened_workbooks:
            wb.Close(SaveChanges=False)

        xl.Quit()

    subject = f"Zonal Daily Dashboard - {report_date_str()}"
    body = (
        "Dear All,<br><br>"
        f"Please find below the Zonal Daily Dashboard for {report_date_str()}."
    )

    send_email_with_inline_image(image_path, subject, body)
    reset_control_workbook()

    print("[OK] Email sent.")
    print(f"[OK] control.xlsx reset: Units={RESET_UNITS}, Download=TRUE, From/To Date={datetime.now().strftime('%d-%m-%Y  00:00:00')}")


if __name__ == "__main__":
    main()
