import os
import time
import glob
import shutil
import random
import threading
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
import subprocess
import sys
import re
import ctypes
from pywinauto import Desktop, keyboard
from datetime import datetime


def append_run_log(message):
    log_path = os.path.join(os.getcwd(), f"run_log_{datetime.now().strftime('%Y-%m-%d')}.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")

def wake_and_prevent_sleep():
    disable_windows_screensaver()
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    MOUSEEVENTF_MOVE = 0x0001
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    VK_SHIFT = 0x10
    KEYEVENTF_KEYUP = 0x0002
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
        ctypes.windll.user32.keybd_event(VK_SHIFT, 0, 0, 0)
        ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, 100, 0, 0, 0)
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, -100, 0, 0, 0)
        append_run_log("[WAKE] Keep-awake enabled with keypress + mouse click wake signals before core run.")
        print("[OK] Wake/keep-awake signal sent (keypress + mouse click).")
    except Exception as e:
        append_run_log(f"[WARN] Could not send wake signal before core run: {e}")
        print(f"[WARN] Could not send wake signal: {e}")

def disable_windows_screensaver():
    SPI_SETSCREENSAVEACTIVE = 0x0012
    SPI_SETSCREENSAVETIMEOUT = 0x000F
    SPI_SETSCREENSAVESECURE = 0x0061
    SPIF_UPDATEINIFILE = 0x01
    SPIF_SENDCHANGE = 0x02
    flags = SPIF_UPDATEINIFILE | SPIF_SENDCHANGE
    try:
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETSCREENSAVEACTIVE, 0, None, flags)
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETSCREENSAVETIMEOUT, 0, None, flags)
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETSCREENSAVESECURE, 0, None, flags)
        append_run_log("[WAKE] Screensaver disabled via SystemParametersInfoW before automation run.")
        print("[OK] Screensaver disabled for this session.")
    except Exception as e:
        append_run_log(f"[WARN] Could not disable screensaver: {e}")
        print(f"[WARN] Could not disable screensaver: {e}")


DATE_WISE_REPORT_NAMES = {
    'patient_payable_date_wise',
}

CLINIC_ONLY_REPORT_NAMES = {
    'revenue_service_level',
    'sales_day_book_excel',
    'opd_foot_falls_new',
}


def parse_report_date(value, default=None):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        value = default
    text = str(value).strip() if value is not None else ""
    if not text:
        value = default
        text = str(value).strip() if value is not None else ""
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}($|\s)", text):
        return pd.to_datetime(text, yearfirst=True)
    if any(sep in text for sep in ["/", ".", "-"]):
        return pd.to_datetime(text, dayfirst=True)
    return pd.to_datetime(value, dayfirst=True)


def normalize_report_name(report_name):
    normalized = re.sub(r'[^a-z0-9]+', '_', str(report_name).strip().lower())
    return normalized.strip('_')

def normalize_col_name(col_name):
    return re.sub(r'[^a-z0-9]+', '_', str(col_name).strip().lower()).strip('_')

def bool_from_cell(value):
    if isinstance(value, bool):
        return value
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    text = str(value).strip().lower()
    return text in {"true", "1", "yes", "y", "t"}


def is_date_wise_report(report_name):
    return normalize_report_name(report_name) in DATE_WISE_REPORT_NAMES

def get_units_for_report(report_name):
    normalized = normalize_report_name(report_name)
    all_units = set(unit_codes)
    if normalized in CLINIC_ONLY_REPORT_NAMES:
        return all_units
    return all_units - set(clinic_units_in_scope)

def report_applies_to_unit(report_name, unit_code):
    return unit_code in get_units_for_report(report_name)


def iter_report_date_ranges(rep):
    rep_from = parse_report_date(rep.get('From Date', from_date), default=from_date)
    rep_to = parse_report_date(rep.get('To Date', to_date), default=to_date)
    if is_date_wise_report(rep.get('Report Name', '')):
        total_days = (rep_to - rep_from).days + 1
        print(f"[INFO] Date-wise extraction for {rep.get('Report Name', '')}: {total_days} daily reports ({rep_from.strftime('%d/%m/%Y')} to {rep_to.strftime('%d/%m/%Y')})")
        for day in pd.date_range(rep_from, rep_to, freq='D'):
            yield day, day
    else:
        yield rep_from, rep_to

wake_and_prevent_sleep()

# === Load Main sheet raw ===
main_raw_df = pd.read_excel("control.xlsx", sheet_name='Main', header=None)

report_header_idx = None
for i in range(len(main_raw_df)):
    first_cell = normalize_col_name(main_raw_df.iloc[i, 0])
    if first_cell == "report_name":
        report_header_idx = i
        break

# === Load config (key/value area above report table) ===
config_rows = []
config_scan_limit = report_header_idx if report_header_idx is not None else len(main_raw_df)
for i in range(config_scan_limit):
    if pd.isna(main_raw_df.iloc[i, 0]) and pd.isna(main_raw_df.iloc[i, 1]):
        break
    config_rows.append((str(main_raw_df.iloc[i, 0]).strip().lower(), main_raw_df.iloc[i, 1]))
config = dict(config_rows)
print("[OK] Config loaded:", config)

url = config.get('url')
username = config.get('username')
password = config.get('password')
output_folder = str(config.get('output folder', r'C:\laragon\www\automation-project - Main\download'))
os.makedirs(output_folder, exist_ok=True)

from_date = parse_report_date(config.get('from date', pd.Timestamp.today())).strftime('%d/%m/%Y')
to_date = parse_report_date(config.get('to date', pd.Timestamp.today())).strftime('%d/%m/%Y')

configured_unit_codes = [u.strip() for u in str(config.get('units', '')).split(',') if u.strip()]

# === Load Reports and Units sheets ===
reports_df = pd.read_excel("control.xlsx", sheet_name='Report')
reports_df.columns = [c.strip() for c in reports_df.columns]
units_df = pd.read_excel("control.xlsx", sheet_name='Unit')
units_df.columns = [c.strip() for c in units_df.columns]
units_df["Units"] = units_df["Units"].astype(str).str.strip()

# === Load report control table from Main sheet (if present) ===
main_reports_df = pd.DataFrame()
main_unit_columns = []
main_control_mode = False
if report_header_idx is not None:
    report_header = [str(c).strip() for c in main_raw_df.iloc[report_header_idx].tolist()]
    tmp_df = main_raw_df.iloc[report_header_idx + 1:].copy()
    tmp_df.columns = report_header
    tmp_df = tmp_df.dropna(how="all")
    tmp_df = tmp_df.loc[tmp_df["Report Name"].astype(str).str.strip().ne("")]
    tmp_df["Report Name"] = tmp_df["Report Name"].astype(str).str.strip()
    main_reports_df = tmp_df.reset_index(drop=True)
    known_non_unit_cols = {"Report Name", "From Date", "To Date", "Download", "Unit"}
    unit_sheet_codes = set(units_df["Units"].tolist())
    main_unit_columns = [
        c for c in main_reports_df.columns
        if c not in known_non_unit_cols and str(c).strip() in unit_sheet_codes
    ]
    main_control_mode = len(main_reports_df) > 0 and len(main_unit_columns) > 0
    if main_control_mode:
        print(f"[OK] Main-sheet report control detected ({len(main_reports_df)} rows). Unit columns: {', '.join(main_unit_columns)}")
    else:
        print("[WARN] Main-sheet report control table not detected or has no valid unit columns. Falling back to legacy Report sheet flow.")

# Final unit scope for this run
if main_control_mode:
    enabled_units = []
    for unit_code in main_unit_columns:
        if main_reports_df[unit_code].apply(bool_from_cell).any():
            enabled_units.append(unit_code)
    if configured_unit_codes:
        unit_codes = [u for u in enabled_units if u in configured_unit_codes]
    else:
        unit_codes = enabled_units
else:
    unit_codes = configured_unit_codes

clinic_units_in_scope = []
for _, unit_row in units_df.iterrows():
    unit_code = str(unit_row.get('Units', '')).strip()
    if unit_code not in unit_codes:
        continue
    hospital_text = str(unit_row.get('Hospital', '')).strip().lower()
    if hospital_text and "clinic" in hospital_text:
        clinic_units_in_scope.append(unit_code)

# === Edge driver setup ===
def get_driver():
    options = EdgeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-insecure-localhost')
    options.add_argument('--safebrowsing-disable-download-protection')
    _ext_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fix2_edge_extension")
    if os.path.isdir(_ext_path):
        options.add_argument(f"--load-extension={_ext_path}")
    options.set_capability("acceptInsecureCerts", True)
    options.add_experimental_option("prefs", {
        "download.default_directory": output_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
    })
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 30)
    return driver, wait

# === Move downloaded file into folder structure ===
def move_downloaded_file(report_name, unit_code, from_date_disp, to_date_disp, output_folder, downloaded_file=None):
    if downloaded_file and os.path.isfile(downloaded_file) and not downloaded_file.endswith(".crdownload"):
        latest_file = downloaded_file
    else:
        files = [
            f for f in glob.glob(os.path.join(output_folder, '*'))
            if os.path.isfile(f) and not f.endswith(".crdownload")
        ]
        if not files:
            print("[WARN] No completed downloaded file found to move.")
            return
        latest_file = max(files, key=os.path.getctime)

    ext = os.path.splitext(latest_file)[1]
    safe_report_name = report_name.replace(" ", "_")
    safe_from = from_date_disp.replace("/", "-")
    safe_to = to_date_disp.replace("/", "-")

    dest_dir = os.path.join(output_folder, safe_report_name)
    os.makedirs(dest_dir, exist_ok=True)
    new_name = f"{unit_code}_{safe_from}_{safe_to}_{safe_report_name}{ext}"
    dest_path = os.path.join(dest_dir, new_name)

    try:
        shutil.move(latest_file, dest_path)
        print(f"[OK] File moved to: {dest_path}")
    except Exception as e:
        print(f"[WARN] Could not move file: {e}")

def has_expected_report_file(report_name, unit_code, from_date_disp, to_date_disp, output_folder):
    safe_report_name = report_name.replace(" ", "_")
    safe_from = from_date_disp.replace("/", "-")
    safe_to = to_date_disp.replace("/", "-")
    base_name = f"{unit_code}_{safe_from}_{safe_to}_{safe_report_name}"
    report_dir = os.path.join(output_folder, safe_report_name)
    if not os.path.isdir(report_dir):
        return False
    pattern = os.path.join(report_dir, f"{base_name}.*")
    files = [f for f in glob.glob(pattern) if os.path.isfile(f)]
    return len(files) > 0

def report_instance_key(report_name, rep_from_dt, rep_to_dt):
    from_disp = pd.to_datetime(rep_from_dt).strftime('%d/%m/%Y')
    to_disp = pd.to_datetime(rep_to_dt).strftime('%d/%m/%Y')
    return f"{report_name} ({from_disp} to {to_disp})"

def get_expected_report_runs_for_unit(unit_code):
    expected = []
    if main_control_mode:
        for _, main_row in main_reports_df.iterrows():
            if unit_code not in main_unit_columns:
                continue
            if not bool_from_cell(main_row.get(unit_code)):
                continue

            report_name = str(main_row.get("Report Name", "")).strip()
            if not report_name:
                continue

            normalized_target = normalize_report_name(report_name)
            candidates = reports_df[
                reports_df["Report Name"].apply(normalize_report_name) == normalized_target
            ]
            if candidates.empty:
                print(f"[WARN] Report metadata not found in 'Report' sheet for Main row: {report_name}")
                continue

            rep = candidates.iloc[0].copy()
            rep["From Date"] = main_row.get("From Date", from_date)
            rep["To Date"] = main_row.get("To Date", to_date)

            for rep_from_dt, rep_to_dt in iter_report_date_ranges(rep):
                expected.append((rep, rep_from_dt, rep_to_dt))
    else:
        for _, rep in reports_df.iterrows():
            download_flag = str(rep.get('Download', 'FALSE')).strip().upper() == 'TRUE'
            if not download_flag:
                continue
            if not report_applies_to_unit(rep.get('Report Name', ''), unit_code):
                continue
            for rep_from_dt, rep_to_dt in iter_report_date_ranges(rep):
                expected.append((rep, rep_from_dt, rep_to_dt))
    return expected

def find_missing_reports_for_unit(unit_code):
    missing_reports = []
    for rep, rep_from_dt, rep_to_dt in get_expected_report_runs_for_unit(unit_code):
        rep_from_date_disp = pd.to_datetime(rep_from_dt).strftime('%d/%m/%Y')
        rep_to_date_disp = pd.to_datetime(rep_to_dt).strftime('%d/%m/%Y')
        if not has_expected_report_file(rep['Report Name'], unit_code, rep_from_date_disp, rep_to_date_disp, output_folder):
            missing_reports.append((rep, rep_from_dt, rep_to_dt))
    return missing_reports

def wait_for_download_start(output_folder, timeout_s=30, known_files=None):
    known_files = known_files or set()
    start = time.time()
    while time.time() - start < timeout_s:
        files = [
            f for f in glob.glob(os.path.join(output_folder, '*'))
            if os.path.isfile(f)
        ]
        new_files = [f for f in files if f not in known_files]
        if new_files:
            return True
        time.sleep(0.5)
    return False

# FIX 3: Snapshot-based detection — finds any NEW completed file not in known_files.
# Old version used mtime comparison which failed when Edge renamed .crdownload files.
def wait_for_download_complete(output_folder, timeout_s=120, known_files=None):
    known_files = known_files or set()
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        files = [
            f for f in glob.glob(os.path.join(output_folder, '*'))
            if os.path.isfile(f) and not f.endswith(".crdownload")
        ]
        new_files = [f for f in files if f not in known_files]
        if new_files:
            return max(new_files, key=os.path.getmtime)
        time.sleep(0.5)
    return None

def handle_edge_download_prompt_ui(attempts=8, pause_s=1.0):
    """
    Use Windows UI Automation to click Keep -> Keep anyway in Edge downloads flyout.
    """
    def _click_control(ctrl):
        try:
            if hasattr(ctrl, "click_input"):
                ctrl.click_input()
                return True
        except Exception:
            pass
        try:
            ctrl.invoke()
            return True
        except Exception:
            return False

    def _click_all(edge_win, label_regex):
        clicked_any = False
        try:
            for ctrl in edge_win.descendants():
                try:
                    name = (ctrl.window_text() or "").strip()
                    if not name:
                        continue
                    # FIX 6: corrected regex — raw strings in Python don't need double backslash
                    if re.search(label_regex, name, re.IGNORECASE):
                        if _click_control(ctrl):
                            clicked_any = True
                except Exception:
                    continue
        except Exception:
            pass
        return clicked_any

    for _ in range(attempts):
        try:
            edge_windows = Desktop(backend="uia").windows(title_re=".*Edge.*")
            if not edge_windows:
                time.sleep(pause_s)
                continue
            edge = edge_windows[0]
            edge.set_focus()
            time.sleep(0.2)

            keyboard.send_keys("^j")
            time.sleep(0.6)

            # FIX 6: regex strings fixed — were r"^retry$|\\bretry\\b" (broken double-escape)
            if _click_all(edge, r"^retry$|\bretry\b"):
                print("[OK] Retry clicked via UI automation.")
                return True

            if _click_all(edge, r"^keep$|\bkeep\b"):
                time.sleep(0.5)
                _click_all(edge, r"^keep anyway$|keep\s+anyway")
                time.sleep(5)
                print("[OK] Keep/Keep anyway clicked via UI automation.")
                return True
        except Exception:
            pass
        time.sleep(pause_s)
    print("[WARN] UI automation could not click Retry/Keep/Keep anyway.")
    return False

# === Export button helper ===
def click_export_button(driver, wait, export_btn):
    if pd.isna(export_btn) or str(export_btn).strip().lower() == 'none':
        return False

    xpath = str(export_btn).strip()
    driver.switch_to.default_content()
    all_frames = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")

    try:
        exp_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        exp_btn.click()
        print(f"[OK] Export clicked (top-level) using XPath: {xpath}")
        time.sleep(10)
        return True
    except:
        pass

    def search_frames(frames):
        for f in frames:
            try:
                driver.switch_to.frame(f)
                try:
                    btn = driver.find_element(By.XPATH, xpath)
                    btn.click()
                    print(f"[OK] Export clicked inside frame using XPath: {xpath}")
                    time.sleep(10)
                    return True
                except:
                    nested_frames = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")
                    if nested_frames:
                        found = search_frames(nested_frames)
                        if found:
                            return True
            finally:
                driver.switch_to.parent_frame()
        return False

    clicked = search_frames(all_frames)
    if not clicked:
        print(f"[WARN] Export button '{xpath}' not found in any frame")
    driver.switch_to.default_content()
    return clicked

# === Handle Edge download security prompt ===
def handle_edge_download_prompt(driver, attempts=6, pause_s=1.0, target_name=None):
    """
    Use Selenium on edge://downloads to click Keep / Keep anyway via Shadow DOM.
    """
    js_click_keep = """
    const callback = arguments[arguments.length - 1];
    const targetName = (arguments.length > 1 && arguments[0]) ? String(arguments[0]).toLowerCase() : '';
    (async () => {
        const mgr = document.querySelector('downloads-manager');
        if (!mgr || !mgr.shadowRoot) { callback(false); return; }
        const mgrRoot = mgr.shadowRoot;
        const items = mgrRoot.querySelectorAll('downloads-item');
        if (!items || items.length === 0) { callback(false); return; }

        const sleep = (ms) => new Promise(r => setTimeout(r, ms));

        function getTexts(root) {
            const nameEl = root.querySelector('#file-link') || root.querySelector('#fileName') || root.querySelector('#name');
            const statusEl = root.querySelector('#dangerousDescription') || root.querySelector('#tag') || root.querySelector('#status');
            return {
                nameText: (nameEl ? nameEl.textContent : '').toLowerCase(),
                statusText: (statusEl ? statusEl.textContent : '').toLowerCase(),
            };
        }

        function isTarget(root) {
            const { nameText, statusText } = getTexts(root);
            const isDanger = statusText.includes("can't be downloaded securely") ||
                             statusText.includes('downloaded securely') ||
                             statusText.includes('blocked') ||
                             statusText.includes('virus scan failed') ||
                             statusText.includes("couldn't download") ||
                             statusText.includes('could not download');
            const nameMatch = targetName ? nameText.includes(targetName) : true;
            return { isDanger, nameMatch, nameText, statusText };
        }

        function hardClick(el) {
            if (!el) return;
            el.dispatchEvent(new MouseEvent('mouseover', {bubbles: true, view: window}));
            el.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, view: window}));
            el.dispatchEvent(new MouseEvent('mouseup', {bubbles: true, view: window}));
            el.dispatchEvent(new MouseEvent('click', {bubbles: true, view: window}));
            if (el.click) el.click();
        }

        async function openMenu(root) {
            const row = root.querySelector('#content') || root;
            if (row) {
                row.scrollIntoView({block: 'center'});
                row.dispatchEvent(new MouseEvent('mouseover', {bubbles: true, view: window}));
                row.dispatchEvent(new MouseEvent('mouseenter', {bubbles: true, view: window}));
                row.dispatchEvent(new MouseEvent('mousemove', {bubbles: true, view: window}));
                row.focus && row.focus();
                hardClick(row);
                await sleep(350);
            }

            const inlineButtons = root.querySelectorAll('cr-button, button');
            for (const b of inlineButtons) {
                const text = (b.textContent || '').trim().toLowerCase();
                if (text === 'retry' || text.includes('retry')) {
                    hardClick(b);
                    await sleep(400);
                    return true;
                }
                if (text === 'keep' || text.includes('keep')) {
                    hardClick(b);
                    await sleep(250);
                    await clickKeepAnyway();
                    return true;
                }
            }

            const moreBtn = root.querySelector('cr-icon-button#moreActions') ||
                            root.querySelector('cr-icon-button[aria-label*="More"]');
            if (moreBtn) {
                moreBtn.scrollIntoView({block: 'center'});
                moreBtn.dispatchEvent(new MouseEvent('mouseover', {bubbles: true, view: window}));
                hardClick(moreBtn);
                await sleep(700);
                return true;
            }
            return false;
        }

        async function clickKeepAnyway() {
            const scopes = [mgrRoot, document];
            for (const scope of scopes) {
                const buttons = scope.querySelectorAll('button, cr-button, cr-menu-item');
                for (const b of buttons) {
                    const text = (b.textContent || '').trim().toLowerCase();
                    if (text === 'keep anyway' || text.includes('keep anyway')) {
                        hardClick(b);
                        return true;
                    }
                }
            }
            return false;
        }

        async function clickKeepFromMenu() {
            const menu = mgrRoot.querySelector('cr-action-menu') ||
                         mgrRoot.querySelector('cr-menu') ||
                         document.querySelector('cr-action-menu') ||
                         document.querySelector('cr-menu');
            if (!menu) return false;
            const buttons = menu.querySelectorAll('button, cr-button, cr-menu-item');
            for (const b of buttons) {
                const text = (b.textContent || '').trim().toLowerCase();
                if (text === 'retry' || text.includes('retry')) {
                    hardClick(b);
                    await sleep(400);
                    return true;
                }
                if (text === 'keep' || text.includes('keep')) {
                    hardClick(b);
                    await sleep(300);
                    await clickKeepAnyway();
                    return true;
                }
            }
            return false;
        }

        for (const item of items) {
            const root = item.shadowRoot;
            if (!root) continue;
            const { isDanger, nameMatch } = isTarget(root);
            if (isDanger && nameMatch) {
                if (await openMenu(root)) {
                    if (await clickKeepFromMenu()) { callback(true); return; }
                }
            }
        }

        for (const item of items) {
            const root = item.shadowRoot;
            if (!root) continue;
            const { isDanger, nameText } = isTarget(root);
            if (isDanger && (nameText.includes('.xls') || nameText.includes('export'))) {
                if (await openMenu(root)) {
                    if (await clickKeepFromMenu()) { callback(true); return; }
                }
            }
        }

        callback(false);
    })().catch(() => callback(false));
    """
    original_handle = None
    try:
        try:
            original_handle = driver.current_window_handle
        except Exception:
            original_handle = None
        driver.get("edge://downloads/")
        time.sleep(pause_s)

        items_ready = False
        for _ in range(12):
            try:
                items = driver.execute_script(
                    """
                    const mgr = document.querySelector('downloads-manager');
                    if (!mgr || !mgr.shadowRoot) return [];
                    const items = mgr.shadowRoot.querySelectorAll('downloads-item');
                    return Array.from(items).length;
                    """
                )
                if items and int(items) > 0:
                    items_ready = True
                    break
            except Exception:
                pass
            try:
                driver.refresh()
            except Exception:
                pass
            time.sleep(1.0)

        try:
            row_elem = driver.execute_script(
                """
                const targetName = arguments[0] ? String(arguments[0]).toLowerCase() : '';
                const mgr = document.querySelector('downloads-manager');
                if (!mgr || !mgr.shadowRoot) return null;
                const items = mgr.shadowRoot.querySelectorAll('downloads-item');
                if (!items || items.length === 0) return null;
                for (const item of items) {
                    const root = item.shadowRoot;
                    if (!root) continue;
                    const nameEl = root.querySelector('#file-link') || root.querySelector('#fileName') || root.querySelector('#name');
                    const statusEl = root.querySelector('#dangerousDescription') || root.querySelector('#tag') || root.querySelector('#status');
                    const nameText = (nameEl ? nameEl.textContent : '').toLowerCase();
                    const statusText = (statusEl ? statusEl.textContent : '').toLowerCase();
                    const isDanger = statusText.includes("can't be downloaded securely") ||
                                     statusText.includes('downloaded securely') ||
                                     statusText.includes('blocked') ||
                                     statusText.includes('virus scan failed') ||
                                     statusText.includes("couldn't download") ||
                                     statusText.includes('could not download');
                    const nameMatch = targetName ? nameText.includes(targetName) : true;
                    if (isDanger && nameMatch) {
                        return root.querySelector('#content') || root;
                    }
                }
                const root0 = items[0].shadowRoot;
                if (!root0) return null;
                return root0.querySelector('#content') || root0;
                """,
                target_name,
            )
            if row_elem:
                ActionChains(driver).move_to_element(row_elem).perform()
                time.sleep(0.35)
        except Exception:
            pass

        clicked = False
        for _ in range(attempts):
            try:
                clicked = bool(
                    driver.execute_script(
                        """
                        function* allNodes(root) {
                            const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT, null);
                            let node = walker.currentNode;
                            while (node) {
                                yield node;
                                if (node.shadowRoot) {
                                    yield* allNodes(node.shadowRoot);
                                }
                                node = walker.nextNode();
                            }
                        }
                        function clickByText(text) {
                            const t = text.toLowerCase();
                            for (const el of allNodes(document)) {
                                const label = (el.textContent || '').trim().toLowerCase();
                                if (label === t || label.includes(t)) {
                                    el.dispatchEvent(new MouseEvent('mouseover', {bubbles:true, view:window}));
                                    el.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, view:window}));
                                    el.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, view:window}));
                                    el.dispatchEvent(new MouseEvent('click', {bubbles:true, view:window}));
                                    if (el.click) el.click();
                                    return true;
                                }
                            }
                            return false;
                        }
                        const retryClicked = clickByText('retry');
                        if (retryClicked) { return true; }
                        const keepClicked = clickByText('keep');
                        if (keepClicked) {
                            setTimeout(() => clickByText('keep anyway'), 300);
                            return true;
                        }
                        return false;
                        """
                    )
                )
            except Exception:
                clicked = False

            if not clicked:
                clicked = bool(driver.execute_async_script(js_click_keep, target_name))
            if clicked:
                time.sleep(pause_s)
                driver.execute_async_script(js_click_keep, target_name)
                time.sleep(5)
                print("[OK] Edge download security prompt handled via downloads page.")
                break
            try:
                driver.refresh()
                time.sleep(pause_s)
            except Exception:
                pass
            time.sleep(pause_s)

        if not clicked:
            try:
                items_dump = driver.execute_script(
                    """
                    const mgr = document.querySelector('downloads-manager');
                    if (!mgr || !mgr.shadowRoot) return [];
                    const items = mgr.shadowRoot.querySelectorAll('downloads-item');
                    const out = [];
                    for (const item of items) {
                        const root = item.shadowRoot;
                        if (!root) continue;
                        const nameEl = root.querySelector('#file-link') || root.querySelector('#fileName') || root.querySelector('#name');
                        const statusEl = root.querySelector('#dangerousDescription') || root.querySelector('#tag') || root.querySelector('#status');
                        out.push({
                            name: nameEl ? nameEl.textContent : '',
                            status: statusEl ? statusEl.textContent : ''
                        });
                    }
                    return out;
                    """
                )
                import json
                print("LOG: Download items: " + json.dumps(items_dump, ensure_ascii=False))
            except Exception:
                print("[WARN] No Keep/Keep anyway button found in downloads.")
    except Exception as e:
        print(f"[WARN] Could not handle Edge download prompt via downloads: {e}")
    finally:
        try:
            if original_handle and original_handle in driver.window_handles:
                if driver.current_window_handle != original_handle:
                    driver.close()
                    driver.switch_to.window(original_handle)
        except Exception:
            pass

# === FIX 6: Session keepalive — pings the session bar every 4 min to prevent logout ===
# Selenium is not thread-safe, but this only reads (find_element without interaction) and
# is wrapped in try/except, so a collision produces at worst a harmless exception.
def _keepalive_worker(driver_ref, stop_event):
    while not stop_event.wait(timeout=240):
        try:
            driver_ref[0].switch_to.default_content()
            driver_ref[0].switch_to.frame("eprmenu")
            driver_ref[0].find_element(By.ID, "MainMenuSessionBarLocation")
            driver_ref[0].switch_to.default_content()
            print("[KEEPALIVE] Session ping sent.")
        except Exception:
            try:
                driver_ref[0].switch_to.default_content()
            except Exception:
                pass

def start_keepalive(driver):
    stop_event = threading.Event()
    t = threading.Thread(target=_keepalive_worker, args=([driver], stop_event), daemon=True)
    t.start()
    return stop_event

def stop_keepalive(stop_event):
    if stop_event:
        stop_event.set()

# === Process single report ===
def process_report(driver, wait, rep, last_main_menu, unit_code, report_from_date=None, report_to_date=None):
    main_menu = rep['Main Menu']
    sub_menu = rep['Sub Menu']
    from_field = rep.get('From Field')
    to_field = rep.get('To Field')
    checkbox_id = rep.get('Checkbox (if any)')
    export_btn = rep.get('Export Button')

    base_from_date = report_from_date if report_from_date is not None else rep.get('From Date', from_date)
    base_to_date = report_to_date if report_to_date is not None else rep.get('To Date', to_date)

    parsed_from = parse_report_date(base_from_date, default=from_date)
    parsed_to = parse_report_date(base_to_date, default=to_date)
    rep_from_date_iso = parsed_from.strftime('%Y-%m-%d')
    rep_to_date_iso   = parsed_to.strftime('%Y-%m-%d')
    rep_from_date_disp = parsed_from.strftime('%d/%m/%Y')
    rep_to_date_disp   = parsed_to.strftime('%d/%m/%Y')

    download_flag = str(rep.get('Download', 'FALSE')).strip().upper() == 'TRUE'
    if not download_flag:
        print(f"[SKIP] Skipping report: {rep['Report Name']} (Download=FALSE)")
        return last_main_menu

    print(f"[RUN] Processing report: {rep['Report Name']} ({rep_from_date_disp} to {rep_to_date_disp})")

    driver.switch_to.default_content()
    driver.switch_to.frame("eprmenu")
    try:
        tools_link = wait.until(EC.element_to_be_clickable((By.ID, "MainMenuLinksTools")))
        tools_link.click()
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "TRAK_tools")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "TRAKTools_menu")))
        print("[OK] Tools clicked")
    except:
        driver.switch_to.default_content()
        print("[WARN] Could not click Tools")

    try:
        if main_menu and main_menu != last_main_menu:
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{main_menu}']"))).click()
            print(f"[OK] Main menu clicked: {main_menu}")
            last_main_menu = main_menu
            # FIX 6: human-like random delay after menu click
            time.sleep(random.uniform(1.2, 2.5))
        else:
            print(f"[SKIP] Skipping Main menu click (same as previous): {main_menu}")
        if sub_menu:
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{sub_menu}']"))).click()
            print(f"[OK] Sub menu clicked: {sub_menu}")
            time.sleep(random.uniform(1.0, 2.0))
    except:
        print(f"[WARN] Could not click Main/Sub menu for {rep['Report Name']}")
        return last_main_menu

    time.sleep(random.uniform(1.5, 3.0))
    handles = driver.window_handles
    report_window = handles[-1]
    driver.switch_to.window(report_window)
    driver.switch_to.default_content()

    try:
        if pd.notna(from_field) and str(from_field).strip().lower() != 'none':
            attr, val = from_field.split("=")
            val = val.strip('"')
            from_elem = wait.until(EC.presence_of_element_located((By.XPATH, f"//*[@{attr}='{val}']")))
            from_elem.clear()
            if from_elem.get_attribute("type") == "date":
                from_elem.send_keys(rep_from_date_iso)
            else:
                from_elem.send_keys(rep_from_date_disp)
        if pd.notna(to_field) and str(to_field).strip().lower() != 'none':
            attr, val = to_field.split("=")
            val = val.strip('"')
            to_elem = wait.until(EC.presence_of_element_located((By.XPATH, f"//*[@{attr}='{val}']")))
            to_elem.clear()
            if to_elem.get_attribute("type") == "date":
                to_elem.send_keys(rep_to_date_iso)
            else:
                to_elem.send_keys(rep_to_date_disp)
        print("[OK] Dates entered")
    except Exception as e:
        print(f"[WARN] Could not enter dates: {e}")

    if pd.notna(checkbox_id) and str(checkbox_id).strip().lower() != 'none':
        try:
            attr, val = checkbox_id.split("=")
            val = val.strip('"')
            cb = driver.find_element(By.XPATH, f"//*[@{attr}='{val}']")
            if not cb.is_selected():
                cb.click()
            print(f"[OK] Checkbox '{checkbox_id}' clicked")
        except:
            print(f"[WARN] Checkbox '{checkbox_id}' not found")

    # FIX 3: Snapshot of files BEFORE export click — used by wait_for_download_complete
    existing_files = {
        f for f in glob.glob(os.path.join(output_folder, '*'))
        if os.path.isfile(f)
    }

    click_export_button(driver, wait, export_btn)

    download_started = wait_for_download_start(output_folder, timeout_s=25, known_files=existing_files)
    if download_started:
        # FIX 3: Poll every 2 seconds for up to 120 seconds total.
        # Handle Keep prompt every 5 seconds during the wait.
        completed_file = None
        deadline = time.time() + 120
        next_keep_check = time.time() + 5
        while time.time() < deadline:
            completed_file = wait_for_download_complete(output_folder, timeout_s=2, known_files=existing_files)
            if completed_file:
                break
            if time.time() >= next_keep_check:
                handle_edge_download_prompt(driver, attempts=1, target_name=rep.get('Report Name'))
                handle_edge_download_prompt_ui(attempts=1)
                next_keep_check = time.time() + 5

        if completed_file:
            move_downloaded_file(rep['Report Name'], unit_code, rep_from_date_disp, rep_to_date_disp, output_folder, downloaded_file=completed_file)
        else:
            print("[WARN] Download did not complete within 120 seconds.")
    else:
        print("[WARN] Download did not start (no new file detected within 25 seconds).")

    try:
        driver.close()
    except Exception as e:
        print(f"[WARN] Could not close report window cleanly: {e}")
    if handles:
        try:
            driver.switch_to.window(handles[0])
        except Exception as e:
            print(f"[WARN] Could not switch back to main window: {e}")
    driver.switch_to.default_content()

    try:
        driver.switch_to.frame("eprmenu")
        tools_link = wait.until(EC.element_to_be_clickable((By.ID, "MainMenuLinksTools")))
        tools_link.click()
        time.sleep(random.uniform(1.5, 2.5))
        tools_link = wait.until(EC.element_to_be_clickable((By.ID, "MainMenuLinksTools")))
        tools_link.click()
        time.sleep(random.uniform(1.5, 2.5))
        tools_link = wait.until(EC.element_to_be_clickable((By.ID, "MainMenuLinksTools")))
        tools_link.click()
        time.sleep(random.uniform(1.5, 2.5))
        driver.switch_to.default_content()
        print("[OK] Tools toggled after report download")
    except:
        driver.switch_to.default_content()
        print("[WARN] Could not toggle Tools after report")

    return last_main_menu

# === FIX 6: Process single unit — driver is shared across all units (login once) ===
# Old version called get_driver() + login inside this function for every unit.
def process_unit(driver, wait, unit_code):
    last_main_menu_clicked = None
    keepalive_stop = start_keepalive(driver)
    try:
        row = units_df.loc[units_df['Units'] == unit_code].iloc[0]
        location = row['Location'].strip()
        sec_group = row['Security Group'].strip()
        hospital = row['Hospital'].strip()
        print(f"[RUN] Processing unit: {unit_code} ({location})")

        try:
            driver.switch_to.default_content()
            driver.switch_to.frame("eprmenu")
            loc_link = wait.until(EC.element_to_be_clickable((By.ID, "MainMenuSessionBarLocation")))
            loc_link.click()
            driver.switch_to.default_content()
        except:
            driver.switch_to.default_content()
            print("[WARN] Could not open unit selector")
            try:
                _shot = os.path.join(os.getcwd(), f"debug_screenshot_{unit_code}.png")
                driver.save_screenshot(_shot)
                print(f"[DEBUG] Screenshot saved: {_shot}")
            except Exception as _e:
                print(f"[DEBUG] Screenshot failed: {_e}")

        unit_clicked = False

        # Wait for unit selector page to load before searching rows
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Choose Logon Location')]")))
        except:
            pass

        for attempt in range(8):
            try:
                rows = driver.find_elements(By.TAG_NAME, "tr")
                for tr in rows:
                    tds = tr.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 2:
                        continue
                    loc_text = tds[1].text.strip()
                    if unit_code in loc_text:
                        tds[1].click()
                        unit_clicked = True
                        print(f"[OK] Selected unit {unit_code} (matched '{loc_text}') on attempt {attempt+1}")
                        break

                if unit_clicked:
                    break

                try:
                    next_btn = driver.find_element(By.CSS_SELECTOR, "img.clsComponentNextPageImage")
                    driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(random.uniform(1.5, 2.5))
                    print(f"[RUN] Checked next page (attempt {attempt+1}) for unit selection...")
                except:
                    print(f"[FAIL] No next page found on attempt {attempt+1}, stopping.")
                    break
            except:
                time.sleep(1)

        if not unit_clicked:
            print(f"[FAIL] Could not select unit {unit_code}, skipping.")
            append_run_log(f"[UNIT {unit_code}] Skipped: could not select unit on portal.")
            return

        time.sleep(random.uniform(1.5, 3.0))
        try:
            driver.switch_to.frame("eprmenu")
            dashboard_loc = driver.find_element(By.ID, "MainMenuSessionBarLocation").text.strip()
            driver.switch_to.default_content()
            if dashboard_loc == location:
                print(f"[OK] Dashboard location verified: {dashboard_loc}")
        except:
            print("[WARN] Dashboard location verification failed")

        expected_reports = get_expected_report_runs_for_unit(unit_code)
        if not expected_reports:
            print(f"[INFO] No reports configured for unit {unit_code} in current control mode.")
        for rep, rep_from_dt, rep_to_dt in expected_reports:
            last_main_menu_clicked = process_report(
                driver,
                wait,
                rep,
                last_main_menu_clicked,
                unit_code,
                report_from_date=rep_from_dt,
                report_to_date=rep_to_dt,
            )

        missing_reports = find_missing_reports_for_unit(unit_code)
        if missing_reports:
            print(f"[WARN] Missing reports after first pass for unit {unit_code}: {len(missing_reports)}")
            for rep, rep_from_dt, rep_to_dt in missing_reports:
                rep_from_disp = pd.to_datetime(rep_from_dt).strftime('%d/%m/%Y')
                rep_to_disp = pd.to_datetime(rep_to_dt).strftime('%d/%m/%Y')
                print(f"[RETRY] Re-downloading: {rep['Report Name']} ({rep_from_disp} to {rep_to_disp})")
                last_main_menu_clicked = process_report(
                    driver,
                    wait,
                    rep,
                    last_main_menu_clicked,
                    unit_code,
                    report_from_date=rep_from_dt,
                    report_to_date=rep_to_dt,
                )

            still_missing = find_missing_reports_for_unit(unit_code)
            if still_missing:
                print(f"[WARN] Still missing after retry for unit {unit_code}:")
                for rep, rep_from_dt, rep_to_dt in still_missing:
                    rep_from_disp = pd.to_datetime(rep_from_dt).strftime('%d/%m/%Y')
                    rep_to_disp = pd.to_datetime(rep_to_dt).strftime('%d/%m/%Y')
                    print(f"       - {rep['Report Name']} ({rep_from_disp} to {rep_to_disp})")
            else:
                print(f"[OK] Missing reports recovered after retry for unit {unit_code}.")
        else:
            print(f"[OK] No missing reports for unit {unit_code}.")

        still_missing = find_missing_reports_for_unit(unit_code)
        missing_keys = {
            report_instance_key(rep['Report Name'], rep_from_dt, rep_to_dt)
            for rep, rep_from_dt, rep_to_dt in still_missing
        }
        downloaded_keys = [
            report_instance_key(rep['Report Name'], rep_from_dt, rep_to_dt)
            for rep, rep_from_dt, rep_to_dt in expected_reports
            if report_instance_key(rep['Report Name'], rep_from_dt, rep_to_dt) not in missing_keys
        ]

        if downloaded_keys:
            append_run_log(f"[UNIT {unit_code}] Downloaded reports: " + "; ".join(downloaded_keys))
        if still_missing:
            append_run_log(
                f"[UNIT {unit_code}] Missing even after re-download: "
                + "; ".join(sorted(missing_keys))
            )
        else:
            append_run_log(f"[UNIT {unit_code}] All expected reports downloaded.")

    except Exception as e:
        print(f"[FAIL] Unit {unit_code} error: {e}")
        append_run_log(f"[UNIT {unit_code}] Error: {e}")
    finally:
        stop_keepalive(keepalive_stop)

def collect_missing_report_map(units):
    missing_map = {}
    for unit_code in units:
        missing = find_missing_reports_for_unit(unit_code)
        if missing:
            missing_map[unit_code] = missing
    return missing_map

def list_files_recursive(base_dir):
    files = []
    if not os.path.isdir(base_dir):
        return files
    for root, _, names in os.walk(base_dir):
        for name in names:
            files.append(os.path.join(root, name))
    return files

def list_spreadsheets_recursive(base_dir):
    return [
        f for f in list_files_recursive(base_dir)
        if os.path.splitext(f)[1].lower() in {".xls", ".xlsx", ".xlsm"}
    ]

def run_subprocess_step(script_name, step_label, verify_fn=None, max_attempts=4, retry_sleep_s=5):
    for attempt in range(1, max_attempts + 1):
        print(f"[RUN] {step_label} (attempt {attempt}/{max_attempts})")
        try:
            subprocess.run([sys.executable, script_name], check=True)
        except Exception as e:
            append_run_log(f"[ERROR] {script_name} attempt {attempt} failed: {e}")
            if attempt < max_attempts:
                time.sleep(retry_sleep_s)
                continue
            return False

        if verify_fn is None:
            append_run_log(f"[OK] {step_label} succeeded on attempt {attempt}.")
            return True

        verified, details = verify_fn()
        if verified:
            append_run_log(f"[OK] {step_label} verified on attempt {attempt}.")
            return True

        append_run_log(f"[WARN] {step_label} verification failed on attempt {attempt}: {details}")
        if attempt < max_attempts:
            time.sleep(retry_sleep_s)
    return False

# === FIX 6: Login ONCE, change location per unit, logout at the very end ===
# Old version: called process_unit(u) which created a new browser + login per unit.
# New version: one shared browser, one login, switch unit via location bar.
MAX_DOWNLOAD_ROUNDS = 4
target_units = list(unit_codes)
remaining_missing = {}

if not target_units:
    print("[WARN] No units selected for processing.")
else:
    driver, wait = get_driver()
    try:
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.ID, "USERNAME"))).send_keys(username)
        driver.find_element(By.ID, "PASSWORD").send_keys(password)
        driver.find_element(By.ID, "Logon").click()
        print("[OK] Logged in once for entire run")
        # FIX 6: random delay after login (human-like, avoids fixed 3s fingerprint)
        time.sleep(random.uniform(2.5, 4.5))

        for round_no in range(1, MAX_DOWNLOAD_ROUNDS + 1):
            print(f"[RUN] Download round {round_no}/{MAX_DOWNLOAD_ROUNDS} for units: {', '.join(target_units)}")
            for u in target_units:
                process_unit(driver, wait, u)

            remaining_missing = collect_missing_report_map(unit_codes)
            if not remaining_missing:
                append_run_log(f"[DOWNLOAD] All expected reports downloaded by round {round_no}.")
                break

            missing_summary = ", ".join(
                f"{unit}:{len(items)}" for unit, items in sorted(remaining_missing.items())
            )
            append_run_log(
                f"[DOWNLOAD] Missing reports remain after round {round_no}: {missing_summary}"
            )
            target_units = sorted(remaining_missing.keys())
        else:
            append_run_log(
                "[WARN] Download rounds exhausted; some reports still missing after maximum retries."
            )
    finally:
        driver.quit()
        print("[OK] Browser closed after all units processed")

# === Trigger post-processing script ===
print("[RUN] Triggering post-processing script...")
def post_run_checks_and_email():
    control_path = "control.xlsx"
    ok_downloaded = False
    ok_converted = False

    try:
        cfg_df = pd.read_excel(control_path, sheet_name="Main", header=None, index_col=0)
        units_val = str(cfg_df.loc["Units", 1]).strip() if "Units" in cfg_df.index else ""
        units_empty = (units_val == "" or units_val.lower() == "nan")

        main_raw = pd.read_excel(control_path, sheet_name="Main", header=None)
        header_idx = None
        for i in range(len(main_raw)):
            if normalize_col_name(main_raw.iloc[i, 0]) == "report_name":
                header_idx = i
                break

        all_download_false = False
        if header_idx is not None:
            header = [str(c).strip() for c in main_raw.iloc[header_idx].tolist()]
            mdf = main_raw.iloc[header_idx + 1:].copy()
            mdf.columns = header
            mdf = mdf.dropna(how="all")
            if "Report Name" in mdf.columns:
                mdf = mdf.loc[mdf["Report Name"].astype(str).str.strip().ne("")]
            unit_codes_in_unit_sheet = set(
                pd.read_excel(control_path, sheet_name="Unit")["Units"].astype(str).str.strip().tolist()
            )
            non_unit_cols = {"Report Name", "From Date", "To Date", "Download", "Unit"}
            unit_cols = [c for c in mdf.columns if c not in non_unit_cols and str(c).strip() in unit_codes_in_unit_sheet]
            if unit_cols:
                any_enabled = False
                for col in unit_cols:
                    if mdf[col].apply(bool_from_cell).any():
                        any_enabled = True
                        break
                all_download_false = not any_enabled

        if units_empty and all_download_false:
            ok_downloaded = True
            append_run_log("Today all reports downloaded.")
    except Exception as e:
        append_run_log(f"[WARN] Post-run check failed reading control.xlsx: {e}")

    try:
        dashboard_folder = os.path.join(os.path.dirname(output_folder), "dashboard")
        download_xlsx = list_spreadsheets_recursive(output_folder)
        dashboard_xlsx = list_spreadsheets_recursive(dashboard_folder)
        if len(download_xlsx) == 0 and len(dashboard_xlsx) == 0:
            ok_converted = True
            append_run_log("Today all reports converted successfully.")
        else:
            append_run_log(
                f"[WARN] Pending files remain. output={len(download_xlsx)}, dashboard={len(dashboard_xlsx)}"
            )
    except Exception as e:
        append_run_log(f"[WARN] Post-run check failed for folders: {e}")

    if ok_downloaded and ok_converted:
        append_run_log("All scripts run successfully.")
    else:
        append_run_log("[WARN] Conditions not met; email not sent.")

dashboard_folder = os.path.join(os.path.dirname(output_folder), "dashboard")

def verify_post_process():
    pending = list_spreadsheets_recursive(output_folder)
    return (len(pending) == 0, f"pending spreadsheets in output folder={len(pending)}")

def verify_move():
    pending = list_files_recursive(dashboard_folder)
    return (len(pending) == 0, f"pending files in dashboard folder={len(pending)}")

all_ok = True
if not run_subprocess_step(
    "post_process_downloads.py",
    "Post-process convert/validate/move-to-dashboard",
    verify_fn=verify_post_process,
    max_attempts=4,
    retry_sleep_s=6,
):
    all_ok = False

if all_ok and not run_subprocess_step(
    "move_dashboard_files.py",
    "Move dashboard files to Input Data",
    verify_fn=verify_move,
    max_attempts=4,
    retry_sleep_s=4,
):
    all_ok = False

if all_ok and not run_subprocess_step(
    "refresh_zodiac_reports.py",
    "Refresh zodiac workbooks",
    verify_fn=None,
    max_attempts=4,
    retry_sleep_s=8,
):
    all_ok = False

post_run_checks_and_email() if all_ok else append_run_log("[WARN] Skipping post-run checks due to step failure after retries.")

# FIX 5: Generate Plotly HTML report and send via email after all steps complete
if all_ok and not run_subprocess_step(
    "generate_report.py",
    "Generate Plotly HTML dashboard and send email",
    verify_fn=None,
    max_attempts=2,
    retry_sleep_s=5,
):
    append_run_log("[WARN] generate_report.py failed — dashboard email not sent.")

print("\n[OK] All units processed.")
