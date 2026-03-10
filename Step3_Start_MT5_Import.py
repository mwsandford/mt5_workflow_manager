"""
MT5 Tick Import Launcher.

Ensures MT5 is running so the TickImportService can process any *_QDM.csv
files in the MQL5\\Files folder.

How it works:
    1. Launches MT5 (minimized) if not already running.
    2. The TickImportService (running in MT5) automatically:
       - Scans for *_QDM.csv files every 5 seconds
       - Imports ticks to matching custom symbols
       - Deletes CSV files after import
    3. Optionally closes MT5 when all imports are complete.

Usage:
    python Step3_Start_MT5_Import.py                                  # Just launch MT5 if needed
    python Step3_Start_MT5_Import.py --wait                           # Wait until all CSVs processed
    python Step3_Start_MT5_Import.py --wait --close-mt5               # Wait, then close MT5 when done
    python Step3_Start_MT5_Import.py --mt5-path "C:\\...\\terminal64.exe"  # Custom MT5 path

Arguments:
    --wait                      Wait until all CSV files have been processed.
    --close-mt5                 Close MT5 after imports complete (implies --wait).
    --wait-timeout              Timeout in seconds for --wait (default: 600).
    --mt5-path                  Path to terminal64.exe (default: C:\\Program Files\\Pepperstone_MT5_01\\terminal64.exe).
    --mt5-startup-wait          Seconds to wait for MT5 to fully initialise (default: 20).
    --mt5-data-folder           Path to the MetaQuotes Terminal data folder
                                (default: %APPDATA%\\MetaQuotes\\Terminal).
"""

import argparse
import glob
import os
import re
import subprocess
import sys
import time
from datetime import datetime


# ---------------------------------------------------------------------------
# ANSI color codes for terminal output
# ---------------------------------------------------------------------------
class Colors:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GRAY = '\033[90m'
    RESET = '\033[0m'


def init_colors():
    """Enable ANSI colors on Windows."""
    if sys.platform == 'win32':
        os.system('')


def format_duration(seconds: float) -> str:
    """Format duration as hours, minutes, seconds."""
    total_hours = int(seconds // 3600)
    total_minutes = int((seconds % 3600) // 60)
    total_seconds = int(seconds % 60)
    
    if total_hours > 0:
        return f"{total_hours}h {total_minutes}m {total_seconds}s"
    elif total_minutes > 0:
        return f"{total_minutes}m {total_seconds}s"
    else:
        return f"{total_seconds}s"


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
SW_SHOWMINNOACTIVE = 7  # Win32: show window minimized without activating


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def is_mt5_running() -> bool:
    """Check if terminal64.exe is currently running."""
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq terminal64.exe", "/NH"],
        capture_output=True, text=True,
    )
    return "terminal64.exe" in result.stdout.lower()


def find_mt5_data_folder(base_folder: str) -> str | None:
    """Find the most recently modified MT5 data folder containing an MQL5 subfolder."""
    if not os.path.isdir(base_folder):
        return None

    candidates = []
    for entry in os.listdir(base_folder):
        full_path = os.path.join(base_folder, entry)
        mql5_path = os.path.join(full_path, "MQL5")
        if os.path.isdir(full_path) and os.path.isdir(mql5_path):
            candidates.append((full_path, os.path.getmtime(full_path)))

    if not candidates:
        return None

    # Return the most recently modified
    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0]


def get_pending_csv_files(files_folder: str) -> list[str]:
    """Return list of *_QDM.csv files in the given folder."""
    if not os.path.isdir(files_folder):
        return []
    return glob.glob(os.path.join(files_folder, "*_QDM.csv"))


def launch_mt5_minimized(mt5_path: str) -> bool:
    """Launch MT5 in a minimized window. Returns True on success."""
    if not os.path.isfile(mt5_path):
        print(f"ERROR: MT5 not found at: {mt5_path}")
        print("Please set the correct path using --mt5-path")
        return False

    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = SW_SHOWMINNOACTIVE

        subprocess.Popen([mt5_path], startupinfo=startupinfo)
        print("  MT5 launched (minimized)")
        return True
    except OSError as e:
        print(f"ERROR: Failed to launch MT5: {e}")
        return False


def start_mt5_and_wait(mt5_path: str, startup_wait: int) -> bool:
    """Launch MT5 minimized, wait for process to appear, then wait for initialisation."""
    print("MT5 is not running. Launching...")

    if not launch_mt5_minimized(mt5_path):
        return False

    # Wait for MT5 process to appear
    print("  Waiting for MT5 to start...")
    timeout = 30
    for _ in range(timeout):
        if is_mt5_running():
            break
        time.sleep(1)
    else:
        print(f"ERROR: MT5 did not start within {timeout} seconds")
        return False

    # Wait for MT5 to fully initialise
    print(f"  Waiting {startup_wait} seconds for MT5 and services to initialise...")
    for i in range(startup_wait, 0, -1):
        print(f"\r  {i} seconds remaining...    ", end="", flush=True)
        time.sleep(1)
    print("\r  MT5 should be ready now.      ")

    return True


def stop_mt5() -> None:
    """Gracefully close MT5, falling back to force kill if needed."""
    print("Closing MT5...")

    if not is_mt5_running():
        print("  MT5 is not running")
        return

    # Try graceful close first
    subprocess.run(
        ["taskkill", "/IM", "terminal64.exe"],
        capture_output=True,
    )

    # Wait up to 10 seconds for graceful close
    for _ in range(10):
        if not is_mt5_running():
            break
        time.sleep(1)

    # Force kill if still running
    if is_mt5_running():
        print("  Force closing MT5...")
        subprocess.run(
            ["taskkill", "/F", "/IM", "terminal64.exe"],
            capture_output=True,
        )
        time.sleep(2)

    if not is_mt5_running():
        print("  MT5 closed")
    else:
        print("  WARNING: Could not close MT5")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    init_colors()
    
    default_mt5_path = r"C:\Program Files\Pepperstone_MT5_01\terminal64.exe"
    default_data_folder = os.path.join(os.environ.get("APPDATA", ""), "MetaQuotes", "Terminal")

    parser = argparse.ArgumentParser(description="MT5 Tick Import Launcher")
    parser.add_argument("--wait", action="store_true",
                        help="Wait until all CSV files have been processed")
    parser.add_argument("--close-mt5", action="store_true",
                        help="Close MT5 after imports complete (implies --wait)")
    parser.add_argument("--wait-timeout", type=int, default=600,
                        help="Timeout in seconds for --wait (default: 600)")
    parser.add_argument("--mt5-path", default=default_mt5_path,
                        help=f"Path to terminal64.exe (default: {default_mt5_path})")
    parser.add_argument("--mt5-startup-wait", type=int, default=20,
                        help="Seconds to wait for MT5 to fully initialise (default: 20)")
    parser.add_argument("--mt5-data-folder", default=default_data_folder,
                        help=f"MetaQuotes Terminal data folder (default: {default_data_folder})")
    args = parser.parse_args()

    # Record start time
    script_start_time = time.time()
    start_datetime = datetime.now()

    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print(f"{Colors.CYAN} MT5 Tick Import Launcher{Colors.RESET}")
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Started:{Colors.RESET} {Colors.GREEN}{start_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print()

    # --close-mt5 implies --wait
    if args.close_mt5 and not args.wait:
        print("NOTE: --close-mt5 requires --wait. Enabling --wait automatically.")
        args.wait = True
        print()

    # Find MT5 data folder
    data_folder = find_mt5_data_folder(args.mt5_data_folder)
    if not data_folder:
        print("ERROR: Could not find MT5 data folder.")
        sys.exit(1)

    mql5_files_folder = os.path.join(data_folder, "MQL5", "Files")
    print(f"MQL5\\Files: {mql5_files_folder}")
    print()

    # Check for pending CSV files
    pending_files = get_pending_csv_files(mql5_files_folder)

    if not pending_files:
        print("No *_QDM.csv files found in MQL5\\Files folder.")
        print("Export your tick data CSVs there first.")
        print()
        sys.exit(0)

    print(f"Found {len(pending_files)} CSV file(s) to import:")
    for filepath in pending_files:
        filename = os.path.basename(filepath)
        symbol_name = re.sub(r"_([^_]+)$", r".\1", os.path.splitext(filename)[0])
        print(f"  - {filename} -> {symbol_name}")
    print()

    # Check if MT5 is running, launch if not
    if not is_mt5_running():
        if not start_mt5_and_wait(args.mt5_path, args.mt5_startup_wait):
            sys.exit(1)
        print()
    else:
        print("MT5 is already running.")
        print()

    print("TickImportService will process the files automatically.")
    print("Check MT5 'Experts' tab for progress.")
    print()

    if args.wait:
        print("Waiting for all imports to complete...")
        print()

        start_time = time.time()
        last_dot = time.time()
        remaining = get_pending_csv_files(mql5_files_folder)

        while remaining:
            elapsed = time.time() - start_time

            if elapsed > args.wait_timeout:
                print()
                print(f"TIMEOUT: Import did not complete within {args.wait_timeout} seconds.")
                print("Remaining files:")
                for f in remaining:
                    print(f"  - {os.path.basename(f)}")
                sys.exit(1)

            # Print a dot every 2 seconds
            if time.time() - last_dot >= 2:
                print(".", end="", flush=True)
                last_dot = time.time()

            time.sleep(0.5)
            remaining = get_pending_csv_files(mql5_files_folder)

        print()
        print()
        elapsed = time.time() - start_time
        print(f"{Colors.GREEN}All imports complete!{Colors.RESET}")
        print(f"{Colors.CYAN}Import Time:{Colors.RESET} {Colors.GREEN}{elapsed:.1f} seconds{Colors.RESET}")
        print()

        # Close MT5 if requested
        if args.close_mt5:
            stop_mt5()

    # Record end time
    end_datetime = datetime.now()
    total_duration = time.time() - script_start_time

    print()
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print(f"{Colors.GREEN} Done{Colors.RESET}")
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Finished:{Colors.RESET}       {Colors.GREEN}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print(f"{Colors.CYAN}Total Duration:{Colors.RESET} {Colors.GREEN}{format_duration(total_duration)}{Colors.RESET}")


if __name__ == "__main__":
    main()
