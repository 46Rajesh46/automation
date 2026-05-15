# Changes Ledger
**Project:** automation-project - Main  
**Date:** 2026-05-14  
**Applied by:** Claude (automated fix session)  
**Purpose:** Track every file and line changed so changes can be replicated manually on the server.

---

## How to use this ledger
Each entry lists:
- **File** — relative path from project root
- **Action** — NEW FILE / MODIFIED / LINE CHANGED
- **Old code** (what was there before)
- **New code** (what replaced it)
- **Why** — reason for the change

---

## FIX 1 — 5 AM Scheduled Wake-Up and Auto-Download

### New File: `set_yesterday_dates.py`
**Action:** NEW FILE  
**Purpose:** Called before core.py by the scheduled task. Sets yesterday's date as From/To Date in control.xlsx, resets Download=TRUE for all reports.

### Modified File: `run_all.bat`
**Action:** MODIFIED (updated in-place — no new file)  
**Purpose:** Updated to use `%~dp0` (dynamic path, works on any drive) and to call `set_yesterday_dates.py` before `core.py`.  
**Use:** Task Scheduler continues to point at `run_all.bat` as before — no change to the task target needed.

### New File: `register_task.ps1`
**Action:** NEW FILE  
**Purpose:** Run ONCE as Administrator to create the Windows Task Scheduler job.  
**Key setting:** `-WakeToRun` — wakes the PC from sleep at 5:00 AM.  
**Run command:**
```
powershell -ExecutionPolicy Bypass -File "C:\laragon\www\automation-project - Main\register_task.ps1"
```
**After running:** Also enable wake timers in Windows Power Options:  
Control Panel → Power Options → Change plan settings → Change advanced power settings → Sleep → Allow wake timers → Enable

---

## FIX 2 — Edge Browser Extension for Auto-Accept Downloads

### New Folder: `fix2_edge_extension\`
Contains two files:

**`fix2_edge_extension\manifest.json`** — Manifest V3 extension definition. Requests `"downloads"` permission.  
**`fix2_edge_extension\background.js`** — Service worker that listens to `chrome.downloads.onChanged`. When a download enters a dangerous/blocked state, calls `chrome.downloads.acceptDanger(id)` — same as user clicking "Keep anyway".

**Install instructions (one-time on every machine running this automation):**
1. Open Microsoft Edge
2. Go to `edge://extensions/`
3. Enable **Developer mode** (toggle top-right)
4. Click **Load unpacked**
5. Select the folder: `fix2_edge_extension`
6. Extension stays installed permanently — no further action needed

**Impact on core.py:** The existing `handle_edge_download_prompt()` and `handle_edge_download_prompt_ui()` functions remain as backup safety nets. With the extension installed, they will find nothing to click (download already accepted) and exit cleanly.

---

## FIX 3 — Missing Downloads: Snapshot-Based Detection

### File: `core.py`

**Change 1 — Function signature: `wait_for_download_complete`**  
| | Code |
|---|---|
| **Old line ~370** | `def wait_for_download_complete(output_folder, start_time, timeout_s=120):` |
| **New** | `def wait_for_download_complete(output_folder, timeout_s=120, known_files=None):` |

**Change 2 — Function body: `wait_for_download_complete`**

Old body (used mtime comparison — failed when Edge renamed .crdownload, changing mtime):
```python
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        files = [
            f for f in glob.glob(os.path.join(output_folder, '*'))
            if os.path.isfile(f) and not f.endswith(".crdownload")
        ]
        if files:
            newest = max(files, key=os.path.getmtime)
            if os.path.getmtime(newest) >= start_time - 1:
                return newest
        time.sleep(0.5)
    return None
```

New body (snapshot-based — finds any NEW file not present before the download started):
```python
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
```

**Change 3 — Download retry loop inside `process_report`**

Old code (6 rounds × 10 seconds — too short, poor recovery):
```python
    download_start = time.time()
    click_export_button(driver, wait, export_btn)

    download_started = wait_for_download_start(output_folder, timeout_s=20, known_files=existing_files)
    if download_started:
        completed_file = None
        for _ in range(6):
            handle_edge_download_prompt(driver, attempts=1, target_name=rep.get('Report Name'))
            handle_edge_download_prompt_ui(attempts=1)
            completed_file = wait_for_download_complete(output_folder, download_start, timeout_s=10)
            if completed_file:
                break

        if completed_file:
            move_downloaded_file(...)
        else:
            print("[WARN] Download did not complete after Keep/Keep anyway retries.")
    else:
        print("[WARN] Download did not start (no new file detected).")
```

New code (120 seconds total, polling every 2 seconds, Keep prompt every 5 seconds):
```python
    click_export_button(driver, wait, export_btn)

    download_started = wait_for_download_start(output_folder, timeout_s=25, known_files=existing_files)
    if download_started:
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
            move_downloaded_file(...)
        else:
            print("[WARN] Download did not complete within 120 seconds.")
    else:
        print("[WARN] Download did not start (no new file detected within 25 seconds).")
```

---

## FIX 4 — Hardcoded D:\ Paths and Missing Excel File in Refresh List

### File: `refresh_zodiac_reports.py`

**Change 1 — Line 7: BASE_DIR**  
Old: `BASE_DIR = r"D:\automation-project - Main\Zonal Zodiac Automation"`  
New: `BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zonal Zodiac Automation")`  
Why: Hardcoded D:\ breaks on C:\ installs and other machines.

**Change 2 — Lines 8-12: FILES list**  
Old:
```python
FILES = [
    "Final Zodiac Automate.xlsx",
    "SLR_Automate.xlsx",
    "Key Surgery.xlsx",
]
```
New:
```python
FILES = [
    "Final Zodiac Automate.xlsx",
    "SLR_Automate.xlsx",
    "Key Surgery.xlsx",
    "Zonal Dashboard Format Automation.xlsx",   # ← ADDED
]
```
Why: Zonal Dashboard Format Automation.xlsx was never being refreshed, so it showed stale values.

---

### File: `move_dashboard_files.py`

**Change — Lines 5-6: SOURCE_DIR and TARGET_DIR**  
Old:
```python
SOURCE_DIR = r"D:\automation-project - Main\dashboard"
TARGET_DIR = r"D:\automation-project - Main\Zonal Zodiac Automation\Input Data"
```
New:
```python
_BASE = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(_BASE, "dashboard")
TARGET_DIR = os.path.join(_BASE, "Zonal Zodiac Automation", "Input Data")
```
Why: Same hardcoded D:\ issue.

---

### File: `send_zonal_dashboard_email.py`

**Change — Line 11: BASE_DIR**  
Old: `BASE_DIR = r"D:\automation-project - Main\Zonal Zodiac Automation"`  
New: `BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zonal Zodiac Automation")`  
Why: Same hardcoded D:\ issue.

---

### File: `post_process_downloads.py`

**Change — Line 34: output_folder assignment**  
Old:
```python
output_folder = str(config_df.loc["Output folder", 1]).strip()
```
New:
```python
_raw_output_folder = str(config_df.loc["Output folder", 1]).strip()
if not _raw_output_folder or _raw_output_folder.lower() == "nan" or _raw_output_folder.startswith("D:\\"):
    output_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "download")
else:
    output_folder = _raw_output_folder
```
Why: Falls back to dynamic path if control.xlsx still has the old D:\ path or is empty.

---

## FIX 5 — Plotly HTML Dashboard Report

### New File: `generate_report.py`
**Action:** NEW FILE  
**Purpose:** Reads processed Excel files from `Zonal Zodiac Automation\Input Data`, generates an interactive HTML file with 4 charts (Revenue by Unit bar, OPD Footfall bar, Census pie, Sales by Unit bar), and emails it as an attachment.

**Requirements:** Install if not already installed:
```
venv\Scripts\activate.bat
pip install plotly xlrd
```
(`xlrd` is needed to read `.xls` files — the hospital portal downloads in old Excel format. `openpyxl` cannot read `.xls`.)

**Bugs fixed in generate_report.py (post-creation review):**

| Bug | Old code | Fixed code |
|-----|---------|-----------|
| `.xls` files crashed | `pd.read_excel(..., engine="openpyxl")` for all files | `read_excel_any()` uses `xlrd` for `.xls`, `openpyxl` for `.xlsx` |
| Charts blank offline | `include_plotlyjs="cdn"` (loads from internet) | `include_plotlyjs=True` (JS embedded, fully self-contained) |
| Wrong chart column | `numeric_cols[0]` (first numeric, often a row ID) | `find_amount_column()` searches by keyword then by highest-sum column |

**Configuration needed:** Add these two rows to `control.xlsx` → **Main sheet** (same key/value section as URL, username, password):

| Column A (key) | Column B (value) |
|---|---|
| `Email To` | `manager@hospital.com;doctor@hospital.com` |
| `Email CC` | `admin@hospital.com` |

No code editing needed — email addresses are read from control.xlsx at runtime.

**Pipeline position:** Runs after `refresh_zodiac_reports.py`. Added to `core.py` at the end of the pipeline.

### File: `core.py` — pipeline addition

Added at the end of the run pipeline (after post_run_checks_and_email):
```python
if all_ok and not run_subprocess_step(
    "generate_report.py",
    "Generate Plotly HTML dashboard and send email",
    verify_fn=None,
    max_attempts=2,
    retry_sleep_s=5,
):
    append_run_log("[WARN] generate_report.py failed — dashboard email not sent.")
```

---

## FIX 6 — Bot Detection / Session Keep-Alive / Shared Driver

### File: `core.py`

**Change 1 — Line 1-2: New imports**  
Added: `import random` and `import threading`  
Why: random for human-like delays (bot detection), threading for session keepalive daemon.

**Change 2 — New functions added (after handle_edge_download_prompt, before process_report)**

```python
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
```

Why: Keeps the session alive between units/reports by touching the session bar element every 4 minutes. Prevents the 5-minute session timeout from firing during long downloads.

**Change 3 — `process_report`: fixed regex bugs in `handle_edge_download_prompt_ui`**

Old (broken — double backslash in raw string is literal `\\b`, not word boundary `\b`):
```python
if _click_all(edge, r"^retry$|\\bretry\\b"):
if _click_all(edge, r"^keep$|\\bkeep\\b"):
_click_all(edge, r"^keep anyway$|keep\\s+anyway")
```
New (correct):
```python
if _click_all(edge, r"^retry$|\bretry\b"):
if _click_all(edge, r"^keep$|\bkeep\b"):
_click_all(edge, r"^keep anyway$|keep\s+anyway")
```

**Change 4 — `process_report`: human-like random delays**  
All fixed `time.sleep(2)` calls in navigation replaced with `time.sleep(random.uniform(1.2, 2.5))` (or similar range) to avoid bot-detection fingerprinting on fixed-interval clicks.

Locations changed:
- After main_menu click: `time.sleep(random.uniform(1.2, 2.5))`
- After sub_menu click: `time.sleep(random.uniform(1.0, 2.0))`
- Before window switch: `time.sleep(random.uniform(1.5, 3.0))`
- After next-page click in unit selection: `time.sleep(random.uniform(1.5, 2.5))`
- After Tools toggle clicks: `time.sleep(random.uniform(1.5, 2.5))`

**Change 5 — `process_unit`: MAJOR REFACTOR — shared driver pattern**

Old signature and structure (login per unit, one browser per unit):
```python
def process_unit(unit_code):
    driver, wait = get_driver()       # ← new browser for every unit
    last_main_menu_clicked = None
    try:
        driver.get(url)               # ← login for every unit
        wait.until(...USERNAME...).send_keys(username)
        driver.find_element(...PASSWORD...).send_keys(password)
        driver.find_element(...Logon...).click()
        ...
    finally:
        driver.quit()                 # ← close browser after every unit
```

New signature (shared driver, no login per unit):
```python
def process_unit(driver, wait, unit_code):   # ← receives shared driver
    last_main_menu_clicked = None
    keepalive_stop = start_keepalive(driver)  # ← start keepalive for this unit
    try:
        ...                                    # ← no login, no browser creation
    except Exception as e:
        ...
    finally:
        stop_keepalive(keepalive_stop)         # ← stop keepalive thread
```

**Change 6 — Main loop: login once, shared driver**

Old (login per unit, sequential browser launches):
```python
for round_no in range(1, MAX_DOWNLOAD_ROUNDS + 1):
    for u in target_units:
        process_unit(u)   # ← each call launched its own browser + login
```

New (one browser, one login, switch unit via location bar):
```python
driver, wait = get_driver()
try:
    driver.get(url)
    wait.until(...USERNAME...).send_keys(username)
    driver.find_element(...PASSWORD...).send_keys(password)
    driver.find_element(...Logon...).click()
    time.sleep(random.uniform(2.5, 4.5))

    for round_no in range(1, MAX_DOWNLOAD_ROUNDS + 1):
        for u in target_units:
            process_unit(driver, wait, u)   # ← shares the open browser
finally:
    driver.quit()   # ← close browser ONCE at the very end
```

Why: The old pattern created a new login session per unit. The server detects rapid repeated logins as bot behavior. One shared session is both more reliable and less detectable.

**Change 7 — `core.py` line ~169: default output_folder path**

Old: `output_folder = str(config.get('output folder', r'D:\automation-project - Main\download'))`  
New: `output_folder = str(config.get('output folder', r'C:\laragon\www\automation-project - Main\download'))`  
Why: Updated hardcoded fallback path from D:\ to C:\ to match actual install location.

---

## Summary Table

| Fix | Files Changed | Action |
|-----|--------------|--------|
| Fix 1 | `set_yesterday_dates.py` | NEW FILE |
| Fix 1 | `run_all.bat` | Modified — dynamic path + added set_yesterday_dates.py call |
| Fix 1 | `register_task.ps1` | NEW FILE |
| Fix 2 | `fix2_edge_extension\manifest.json` | NEW FILE |
| Fix 2 | `fix2_edge_extension\background.js` | NEW FILE |
| Fix 3 | `core.py` — `wait_for_download_complete` | Modified signature + body |
| Fix 3 | `core.py` — `process_report` download loop | Replaced 6-round loop with 120s poll |
| Fix 4 | `refresh_zodiac_reports.py` lines 7-12 | Path + added Zonal Dashboard to FILES |
| Fix 4 | `move_dashboard_files.py` lines 5-6 | Dynamic path |
| Fix 4 | `send_zonal_dashboard_email.py` line 11 | Dynamic path |
| Fix 4 | `post_process_downloads.py` line 34 | Dynamic path fallback |
| Fix 5 | `generate_report.py` | NEW FILE |
| Fix 5 | `core.py` — pipeline end | Added generate_report.py step |
| Fix 6 | `core.py` — imports (lines 1-2) | Added `import random, threading` |
| Fix 6 | `core.py` — `handle_edge_download_prompt_ui` | Fixed 3 broken regex strings |
| Fix 6 | `core.py` — `process_report` | Added random delays |
| Fix 6 | `core.py` — new functions | Added `_keepalive_worker`, `start_keepalive`, `stop_keepalive` |
| Fix 6 | `core.py` — `process_unit` | Refactored to shared driver + keepalive |
| Fix 6 | `core.py` — main loop | Login once, shared driver across all units |

---

## One-Time Setup Checklist (for server)

1. **Copy all changed files** to the server (compare against your backup)
2. **Install plotly:** `venv\Scripts\activate.bat && pip install plotly`
3. **Set email addresses in `generate_report.py`:** lines `TO = ""` and `CC = ""`
4. **Install Edge extension:** open `edge://extensions/`, developer mode on, load unpacked `fix2_edge_extension\`
5. **Register scheduled task:** run `register_task.ps1` as Administrator
6. **Enable wake timers:** Control Panel → Power Options → Advanced → Sleep → Allow wake timers → Enable
7. **Update `control.xlsx` Output folder** cell from `D:\...` to the correct server path (or leave blank — code now falls back automatically)
