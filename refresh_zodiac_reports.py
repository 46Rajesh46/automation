import os
import time
import ctypes
import win32com.client as win32
from datetime import datetime

# FIX 4: Use dynamic path based on script location instead of hardcoded D:\
BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zonal Zodiac Automation")
# FIX 4: Added Zonal Dashboard Format Automation.xlsx — was missing from refresh list
FILES = [
    "Final Zodiac Automate.xlsx",
    "SLR_Automate.xlsx",
    "Key Surgery.xlsx",
    "Zonal Dashboard Format Automation.xlsx",
]


def append_run_log(message):
    log_path = os.path.join(os.getcwd(), f"run_log_{datetime.now().strftime('%Y-%m-%d')}.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def refresh_workbook(excel_app, path, refresh_cycles=2):
    wb = excel_app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
    try:
        for cycle in range(1, refresh_cycles + 1):
            # Refresh all data connections
            wb.RefreshAll()
            # Wait for async queries to finish
            try:
                excel_app.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            # Small buffer for refresh completion
            time.sleep(2)
            wb.Save()
            append_run_log(f"[REFRESH] Saved after refresh cycle {cycle}: {path}")
    finally:
        wb.Close(SaveChanges=True)


def main():
    # Prevent sleep while refresh is running
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
    except Exception:
        pass

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    failed = []
    append_run_log(f"[REFRESH] Starting workbook refresh for {len(FILES)} file(s).")
    try:
        for name in FILES:
            path = os.path.join(BASE_DIR, name)
            if not os.path.isfile(path):
                print(f"[WARN] File not found: {path}")
                failed.append(path)
                append_run_log(f"[REFRESH] File not found: {path}")
                continue
            print(f"[RUN] Refreshing: {path}")
            try:
                refresh_workbook(excel, path, refresh_cycles=2)
                print(f"[OK] Refreshed: {path}")
                append_run_log(f"[REFRESH] Success: {path} (2 refresh cycles completed)")
            except Exception as e:
                failed.append(path)
                print(f"[WARN] Refresh failed for {path}: {e}")
                append_run_log(f"[REFRESH] Failed: {path} | error={e}")
    finally:
        excel.Quit()
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        except Exception:
            pass

    if failed:
        append_run_log(f"[REFRESH] Completed with failures: {len(failed)} file(s).")
        raise SystemExit(1)
    append_run_log("[REFRESH] Completed successfully for all files.")


if __name__ == "__main__":
    main()
