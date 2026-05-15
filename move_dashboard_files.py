import os
import shutil
from datetime import datetime

# FIX 4: Use dynamic path based on script location instead of hardcoded D:\
_BASE = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(_BASE, "dashboard")
TARGET_DIR = os.path.join(_BASE, "Zonal Zodiac Automation", "Input Data")


def append_run_log(message):
    log_path = os.path.join(os.getcwd(), f"run_log_{datetime.now().strftime('%Y-%m-%d')}.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def move_files(src_dir=SOURCE_DIR, dst_dir=TARGET_DIR):
    os.makedirs(dst_dir, exist_ok=True)
    if not os.path.isdir(src_dir):
        print(f"[WARN] Source folder not found: {src_dir}")
        append_run_log(f"[MOVE] Source folder not found: {src_dir}")
        return

    required_rel_files = []
    for root, _, files in os.walk(src_dir):
        for name in files:
            src_path = os.path.join(root, name)
            rel_path = os.path.relpath(src_path, src_dir)
            required_rel_files.append(rel_path)

    append_run_log(
        f"[MOVE] Starting move: source={src_dir}, target={dst_dir}, required_files={len(required_rel_files)}"
    )

    moved = 0
    failed_src_files = []
    for root, _, files in os.walk(src_dir):
        for name in files:
            src_path = os.path.join(root, name)
            rel_dir = os.path.relpath(root, src_dir)
            dst_folder = os.path.join(dst_dir, rel_dir) if rel_dir != "." else dst_dir
            os.makedirs(dst_folder, exist_ok=True)
            dst_path = os.path.join(dst_folder, name)
            try:
                if os.path.exists(dst_path):
                    os.remove(dst_path)
                shutil.move(src_path, dst_path)
                moved += 1
            except Exception as e:
                print(f"[WARN] Could not move {src_path}: {e}")
                failed_src_files.append(src_path)
                append_run_log(f"[MOVE] Failed to move {src_path}: {e}")

    missing_in_target = []
    for rel_path in required_rel_files:
        expected_path = os.path.join(dst_dir, rel_path)
        if not os.path.exists(expected_path):
            missing_in_target.append(rel_path)

    print(f"[OK] Moved {moved} file(s) to {dst_dir}")
    append_run_log(f"[MOVE] Moved files count={moved}")
    if failed_src_files:
        append_run_log(f"[MOVE] Move failures count={len(failed_src_files)}")
    if missing_in_target:
        append_run_log(f"[MOVE] Missing in target after move count={len(missing_in_target)}")
        for rel_path in missing_in_target:
            append_run_log(f"[MOVE] Missing required file: {rel_path}")
        print(f"[WARN] {len(missing_in_target)} required file(s) missing in target.")
    else:
        append_run_log("[MOVE] All required files are present in target after move.")
        print("[OK] All required files verified in target.")


if __name__ == "__main__":
    move_files()
