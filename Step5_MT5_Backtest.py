"""
Automated MT5 Backtesting Script for Multiple EAs
==================================================

Reads all EAs from a specified folder, dynamically creates ini files based on
EA naming conventions, executes backtests sequentially via MT5 command line,
and copies reports to a destination folder.

EA Naming Convention: SQ SYMBOL TIMEFRAME VERSION.ex5
Example: SQ USDJPY H1 2.1.191.ex5

Symbol Convention in MT5: SymbolName.QDM (e.g., USDJPY.QDM)

Parameters
----------
--mt5-terminal-path : Path to the MT5 terminal64.exe executable
--mt5-data-folder   : MT5 data folder (found via File -> Open Data Folder in MT5)
--mt5-ea-folder     : Folder containing compiled .ex5 EA files
--report-dest-folder: Destination folder for completed HTML reports
--model             : Tick model for backtesting (0-4). Model: 0 = Every tick, 1 = 1 minute OHLC, 2 = Open price only, 3 = Math calculations, 4 = Every tick based on real ticks
--from-date         : Backtest start date in YYYY.MM.DD format
--to-date           : Backtest end date in YYYY.MM.DD format
--timeout           : Hard backstop in seconds before force-terminating a backtest (default: 3600)
--idle-threshold    : CPU% below which MT5 is considered idle (default: 0.1)
--idle-minutes      : Consecutive minutes of idle CPU before treating as hung (default: 3)

Usage Examples
--------------

1. Run with all default values:
   python step6_mt5_backtest.py

2. Override the date range only:
   python step6_mt5_backtest.py --from-date 2020.01.01 --to-date 2024.12.31

3. Use a different tick model (1 minute OHLC for faster testing):
   python step6_mt5_backtest.py --model 1

4. Specify a different EA folder and report destination:
   python step6_mt5_backtest.py --mt5-ea-folder "D:\\MyEAs" --report-dest-folder "D:\\Reports"

5. Override all parameters:
   python step6_mt5_backtest.py \\
       --mt5-terminal-path "C:\\Program Files\\Pepperstone_MT5_01\\terminal64.exe" \\
       --mt5-data-folder "C:\\Users\\msand\\AppData\\Roaming\\MetaQuotes\\Terminal\\ABCDEF123456" \\
       --mt5-ea-folder "C:\\Users\\msand\\AppData\\Roaming\\MetaQuotes\\Terminal\\ABCDEF123456\\MQL5\\Experts\\Advisors\\PineappleStrats" \\
       --report-dest-folder "E:\\Trading\\Tools\\MT5_Scripted_Backtests" \\
       --model 4 \\
       --from-date 2010.01.01 \\
       --to-date 2025.12.31

6. Quick test with a short date range and fast model:
   python step6_mt5_backtest.py --model 1 --from-date 2025.01.01 --to-date 2025.03.31

7. Set a custom timeout (20 minutes):
   python step5_mt5_backtest.py --timeout 1200

Tick Model Reference
--------------------
0 = Every tick
1 = 1 minute OHLC
2 = Open price only
3 = Math calculations
4 = Every tick based on real ticks

Known Issues
------------
MT5 occasionally hangs during or after a backtest (e.g. on the "saving report"
dialog, or on an EA-raised dialog). This script detects completion and hangs by:
1. Polling the reports folder every minute — when the expected report appears
   and its size is stable, MT5 is force-closed and the run treated as successful.
2. Measuring combined CPU% of terminal64.exe + metatester64.exe every minute.
   If CPU stays below --idle-threshold for --idle-minutes consecutive minutes
   without a report appearing, the run is treated as hung and force-terminated.
3. Using --timeout as an ultimate backstop so a single EA cannot run forever.
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import psutil
except ImportError:
    print("ERROR: psutil is required. Install with: pip install psutil", file=sys.stderr)
    sys.exit(1)

# =============================================================================
# ANSI colour codes for terminal output
# =============================================================================
class Colours:
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    GRAY = "\033[90m"
    WHITE = "\033[97m"
    RESET = "\033[0m"

MODEL_DESCRIPTIONS = {
    0: "Every tick",
    1: "1 minute OHLC",
    2: "Open price only",
    3: "Math calculations",
    4: "Every tick based on real ticks",
}

# =============================================================================
# DEFAULT VALUES
# =============================================================================
DEFAULT_MT5_TERMINAL_PATH = r"C:\Program Files\Pepperstone_MT5_01\terminal64.exe"
DEFAULT_MT5_DATA_FOLDER = r"C:\Users\msand\AppData\Roaming\MetaQuotes\Terminal\98196E2B1CEDEE516442D255B458C6C2"
DEFAULT_REPORT_DEST_FOLDER = r"E:\Trading\Analysis_Ouput"
DEFAULT_MODEL = 4
DEFAULT_FROM_DATE = "2025.01.01"
DEFAULT_TO_DATE = "2025.12.31"
DEFAULT_TIMEOUT = 3600  # 1 hour backstop — normally we exit far earlier via report-on-disk or CPU-idle detection
DEFAULT_IDLE_THRESHOLD = 0.1  # CPU% below this is considered idle (backtesting sits ~0.8%)
DEFAULT_IDLE_MINUTES = 3  # Consecutive idle minutes with no report → treat as hung


# =============================================================================
# FUNCTIONS
# =============================================================================

def print_cyan(text: str) -> None:
    print(f"{Colours.CYAN}{text}{Colours.RESET}")


def print_green(text: str) -> None:
    print(f"{Colours.GREEN}{text}{Colours.RESET}")


def print_yellow(text: str) -> None:
    print(f"{Colours.YELLOW}{text}{Colours.RESET}")


def print_gray(text: str) -> None:
    print(f"{Colours.GRAY}{text}{Colours.RESET}")


def print_red(text: str) -> None:
    print(f"{Colours.RED}{text}{Colours.RESET}")


def parse_ea_name(filename: str) -> dict | None:
    """
    Parses EA filename to extract symbol, timeframe, and base name.

    Expected format: SQ SYMBOL TIMEFRAME VERSION.ex5
    Example: SQ USDJPY H1 2.1.191.ex5

    Returns:
        dict with keys: symbol, timeframe, base_name
        None if filename does not match expected format
    """
    base_name = Path(filename).stem  # Remove .ex5 extension
    parts = base_name.split()

    if len(parts) < 4:
        print_yellow(f"  WARNING: EA filename '{filename}' does not match expected format: SQ SYMBOL TIMEFRAME VERSION")
        return None

    # parts[0] = "SQ"
    # parts[1] = Symbol (e.g., "USDJPY")
    # parts[2] = Timeframe (e.g., "H1")
    # parts[3+] = Version (e.g., "2.1.191")
    return {
        "symbol": parts[1],
        "timeframe": parts[2],
        "base_name": base_name,
    }


def create_ini_file(
    ea_base_name: str,
    symbol: str,
    timeframe: str,
    ini_path: str,
    model: int,
    from_date: str,
    to_date: str,
) -> None:
    """
    Creates the backtest.ini file for MT5 Strategy Tester.

    Uses ASCII encoding (no BOM) as MT5 can have issues with UTF-8 BOM.
    """
    ini_lines = [
        "[Tester]",
        f"Expert=Advisors\\PineappleStrats\\{ea_base_name}",
        f"ExpertParameters={ea_base_name}.set",
        f"Symbol={symbol}.QDM",
        f"Period={timeframe}",
        f"Model={model}",
        f"FromDate={from_date}",
        f"ToDate={to_date}",
        f"Report=reports\\{ea_base_name} MT5",
        "ReplaceReport=1",
        "ShutdownTerminal=1",
    ]

    # Write with ASCII encoding (no BOM) - MT5 prefers this
    with open(ini_path, "w", encoding="ascii") as f:
        f.write("\n".join(ini_lines))

    print_gray(f"  Created ini file: {ini_path}")


def _collect_mt5_processes(terminal_pid: int) -> list:
    """Collect psutil.Process objects for the terminal, its descendants, and any
    stray metatester64.exe processes that may not be parented to our terminal."""
    procs: list = []
    seen_pids: set[int] = set()

    try:
        parent = psutil.Process(terminal_pid)
        procs.append(parent)
        seen_pids.add(parent.pid)
        for child in parent.children(recursive=True):
            if child.pid not in seen_pids:
                procs.append(child)
                seen_pids.add(child.pid)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

    for p in psutil.process_iter(['name', 'pid']):
        try:
            name = p.info.get('name') or ''
            if name.lower() == 'metatester64.exe' and p.pid not in seen_pids:
                procs.append(p)
                seen_pids.add(p.pid)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    return procs


def sample_mt5_cpu_percent(terminal_pid: int, sample_seconds: float = 1.0) -> float:
    """Return combined CPU% of MT5 terminal + tester processes over sample_seconds.

    Values are per-core-normalised (can exceed 100% on multi-core workloads), but
    the tester typically reports well under 1% even when active.
    """
    procs = _collect_mt5_processes(terminal_pid)
    if not procs:
        return 0.0

    # Prime CPU readings — first call always returns 0.0
    for p in procs:
        try:
            p.cpu_percent(interval=None)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    time.sleep(sample_seconds)

    total = 0.0
    for p in procs:
        try:
            total += p.cpu_percent(interval=None)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return total


def _force_close_mt5(process) -> None:
    """Terminate then kill the MT5 process if still running."""
    try:
        process.terminate()
        time.sleep(2)
        if process.poll() is None:
            process.kill()
        print_gray("  MT5 process terminated")
    except Exception as e:
        print_yellow(f"  WARNING: Error terminating process: {e}")


def run_backtest(
    mt5_terminal_path: str,
    ini_path: str,
    report_path: str,
    timeout_seconds: int = DEFAULT_TIMEOUT,
    check_interval: int = 5,
    idle_threshold: float = DEFAULT_IDLE_THRESHOLD,
    idle_minutes: int = DEFAULT_IDLE_MINUTES,
) -> tuple[int, bool]:
    """
    Executes MT5 with the specified ini file and waits for completion.

    Exits early in three ways:
      1. MT5 exits cleanly — return exit code.
      2. Report file appears on disk and its size is stable — force-close MT5
         and return success. This handles the common "save dialog hang".
      3. Combined CPU% of MT5 processes stays below idle_threshold for
         idle_minutes consecutive minutes without a report — treat as hung,
         force-close and return failure.

    timeout_seconds is retained as an ultimate backstop.

    Returns:
        (exit_code, was_killed): exit_code 0 = success, 1 = failure; was_killed
        indicates whether we force-terminated MT5 rather than letting it exit.
    """
    print_gray("  Launching MT5 backtest...")

    report_mtime_before = None
    if os.path.exists(report_path):
        report_mtime_before = os.path.getmtime(report_path)

    try:
        process = subprocess.Popen(
            [mt5_terminal_path, f'/config:{ini_path}'],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        print_red(f"  ERROR: Failed to start MT5: {e}")
        return -1, False

    elapsed = 0
    idle_streak = 0  # consecutive minutes MT5 has been CPU-idle

    while elapsed < timeout_seconds:
        try:
            return_code = process.poll()
            if return_code is not None:
                print_gray(f"  Backtest completed (Exit code: {return_code})")
                return return_code, False
        except Exception as e:
            print_yellow(f"  WARNING: Error checking process: {e}")

        time.sleep(check_interval)
        elapsed += check_interval

        if elapsed % 60 != 0:
            continue

        print_gray(f"  Still running... ({format_duration(elapsed)} elapsed)")

        # --- Check 1: report file appeared on disk ---
        if os.path.exists(report_path):
            current_mtime = os.path.getmtime(report_path)
            if report_mtime_before is None or current_mtime > report_mtime_before:
                print_yellow("  Report file detected on disk — verifying write is complete...")
                time.sleep(5)
                size_first = os.path.getsize(report_path)
                time.sleep(3)
                size_second = os.path.getsize(report_path)
                if size_second == size_first and size_second > 0:
                    print_yellow("  Report file confirmed stable — force-closing MT5")
                    _force_close_mt5(process)
                    return 0, True
                print_gray("  Report file still being written, continuing to wait...")

        # --- Check 2: CPU-idle watchdog ---
        cpu = sample_mt5_cpu_percent(process.pid)
        if cpu < idle_threshold:
            idle_streak += 1
            print_yellow(f"  MT5 CPU idle ({cpu:.2f}%) — idle streak {idle_streak}/{idle_minutes}")
            if idle_streak >= idle_minutes:
                print_red(
                    f"  ERROR: MT5 appears hung ({idle_minutes} min idle, no report) — force-closing"
                )
                _force_close_mt5(process)
                return 1, True
        else:
            if idle_streak > 0:
                print_gray(f"  MT5 CPU active ({cpu:.2f}%) — idle streak reset")
            idle_streak = 0

    # Backstop timeout reached
    print_yellow(f"  WARNING: MT5 did not exit within backstop timeout {format_duration(timeout_seconds)}")

    report_saved = False
    if os.path.exists(report_path):
        report_mtime_after = os.path.getmtime(report_path)
        if report_mtime_before is None or report_mtime_after > report_mtime_before:
            report_saved = True
            print_yellow("  Report file was saved successfully despite hang")

    print_yellow("  Force-terminating MT5 process...")
    _force_close_mt5(process)

    if report_saved:
        return 0, True
    print_red("  ERROR: Report was not saved before backstop timeout")
    return 1, True


def format_duration(seconds: float) -> str:
    """Formats seconds into HH:MM:SS string."""
    hours, remainder = divmod(int(seconds), 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def format_duration_friendly(seconds: float) -> str:
    """Formats seconds into a friendly string like '1h 23m 45s'."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)

    if hours > 0:
        return f"{hours}h {minutes}m {secs}s"
    elif minutes > 0:
        return f"{minutes}m {secs}s"
    else:
        return f"{secs}s"


def load_top_strategies(json_path: str, max_strategies: int = 10) -> list[str]:
    """
    Load top ranked strategies from strategies_data.json.

    Args:
        json_path: Path to strategies_data.json
        max_strategies: Maximum number of strategies to return

    Returns:
        List of strategy names that should be backtested (top N by rank)
    """
    if not os.path.exists(json_path):
        print_yellow(f"WARNING: Strategies JSON not found: {json_path}")
        return []

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print_yellow(f"WARNING: Could not load strategies JSON: {e}")
        return []

    # Get ranking data (strategies sorted by composite score)
    ranking = data.get('ranking', [])

    # Return top N ranked strategies regardless of KEEP/ABANDON decision
    top_strategies = []
    for strat in ranking:
        name = strat.get('name', '')
        if name:
            top_strategies.append(name)
            if len(top_strategies) >= max_strategies:
                break

    return top_strategies


def normalize_strategy_name(name: str) -> str:
    """Normalize a strategy name for comparison.
    
    Handles variations like:
    - "SQ NATGAS H1 1.107" (from Dashboard)
    - "SQ NATGAS H1 1.1.107" (from EA filename)
    """
    # Remove common prefixes/suffixes and normalize
    normalized = name.strip().upper()
    return normalized


def match_ea_to_strategy(ea_name: str, strategy_names: list[str]) -> bool:
    """Check if an EA filename matches any of the strategy names.
    
    Args:
        ea_name: EA filename (e.g., "SQ NATGAS H1 1.1.107.ex5")
        strategy_names: List of strategy names from Dashboard
        
    Returns:
        True if EA matches a strategy, False otherwise
    """
    # Get base name without extension
    ea_base = Path(ea_name).stem.upper()
    
    for strat in strategy_names:
        strat_upper = strat.upper()
        # Exact match
        if ea_base == strat_upper:
            return True
        # Fuzzy match - EA name contains strategy name or vice versa
        if strat_upper in ea_base or ea_base in strat_upper:
            return True
        # Handle version number variations (1.107 vs 1.1.107)
        # Remove all dots and compare
        ea_nodots = ea_base.replace('.', ' ').replace('  ', ' ')
        strat_nodots = strat_upper.replace('.', ' ').replace('  ', ' ')
        if ea_nodots == strat_nodots:
            return True
    
    return False


def parse_arguments() -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Automated MT5 Backtesting Script for Multiple EAs",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Tick Model Reference:
  0 = Every tick
  1 = 1 minute OHLC
  2 = Open price only
  3 = Math calculations
  4 = Every tick based on real ticks

Examples:
  python step6_mt5_backtest.py
  python step6_mt5_backtest.py --from-date 2020.01.01 --to-date 2024.12.31
  python step6_mt5_backtest.py --model 1 --from-date 2025.01.01 --to-date 2025.03.31
  python step6_mt5_backtest.py --timeout 1200
        """,
    )

    parser.add_argument(
        "--mt5-terminal-path",
        default=DEFAULT_MT5_TERMINAL_PATH,
        help=f"Path to MT5 terminal64.exe (default: {DEFAULT_MT5_TERMINAL_PATH})",
    )
    parser.add_argument(
        "--mt5-data-folder",
        default=DEFAULT_MT5_DATA_FOLDER,
        help=f"MT5 data folder, found via File -> Open Data Folder in MT5 (default: {DEFAULT_MT5_DATA_FOLDER})",
    )
    parser.add_argument(
        "--mt5-ea-folder",
        default=None,
        help="EA folder containing .ex5 files (default: <mt5-data-folder>\\MQL5\\Experts\\Advisors\\PineappleStrats)",
    )
    parser.add_argument(
        "--report-dest-folder",
        default=DEFAULT_REPORT_DEST_FOLDER,
        help=f"Destination folder for completed reports (default: {DEFAULT_REPORT_DEST_FOLDER})",
    )
    parser.add_argument(
        "--model",
        type=int,
        default=DEFAULT_MODEL,
        choices=[0, 1, 2, 3, 4],
        help=f"Tick model for backtesting, 0-4 (default: {DEFAULT_MODEL})",
    )
    parser.add_argument(
        "--from-date",
        default=DEFAULT_FROM_DATE,
        help=f"Backtest start date in YYYY.MM.DD format (default: {DEFAULT_FROM_DATE})",
    )
    parser.add_argument(
        "--to-date",
        default=DEFAULT_TO_DATE,
        help=f"Backtest end date in YYYY.MM.DD format (default: {DEFAULT_TO_DATE})",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f"Hard backstop in seconds before force-terminating a backtest (default: {DEFAULT_TIMEOUT})",
    )
    parser.add_argument(
        "--idle-threshold",
        type=float,
        default=DEFAULT_IDLE_THRESHOLD,
        help=f"CPU%% below which MT5 is considered idle (default: {DEFAULT_IDLE_THRESHOLD})",
    )
    parser.add_argument(
        "--idle-minutes",
        type=int,
        default=DEFAULT_IDLE_MINUTES,
        help=f"Consecutive idle minutes before treating MT5 as hung (default: {DEFAULT_IDLE_MINUTES})",
    )
    parser.add_argument(
        "--strategies-json",
        default=None,
        help="Path to strategies_data.json to filter EAs to only top ranked strategies (optional)",
    )
    parser.add_argument(
        "--max-strategies",
        type=int,
        default=10,
        help="Maximum number of strategies to backtest when using --strategies-json (default: 10)",
    )

    args = parser.parse_args()

    # Default EA folder is derived from data folder if not explicitly provided
    if args.mt5_ea_folder is None:
        args.mt5_ea_folder = os.path.join(args.mt5_data_folder, "MQL5", "Experts", "Advisors", "PineappleStrats")

    return args


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main() -> None:
    args = parse_arguments()

    # Record start time
    start_time = time.time()
    start_datetime = datetime.now()

    print_cyan("============================================================")
    print_cyan("MT5 Batch Backtesting Script")
    print_cyan("============================================================")
    print()
    print(f"{Colours.CYAN}Started:{Colours.RESET} {Colours.GREEN}{start_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colours.RESET}")
    print()

    # Validate paths
    if not os.path.exists(args.mt5_terminal_path):
        print_red(f"ERROR: MT5 Terminal not found at: {args.mt5_terminal_path}")
        sys.exit(1)

    if not os.path.isdir(args.mt5_ea_folder):
        print_red(f"ERROR: EA Folder not found at: {args.mt5_ea_folder}")
        sys.exit(1)

    # Delete and recreate reports folder to avoid duplicates from previous runs
    reports_folder = os.path.join(args.mt5_data_folder, "reports")
    if os.path.exists(reports_folder):
        shutil.rmtree(reports_folder)
        print_yellow("Cleared existing reports folder")
    os.makedirs(reports_folder, exist_ok=True)
    print_yellow(f"Created reports folder: {reports_folder}")

    # Create destination folder if it doesn't exist
    os.makedirs(args.report_dest_folder, exist_ok=True)

    # Get all EA files
    ea_files = sorted([f for f in os.listdir(args.mt5_ea_folder) if f.lower().endswith(".ex5")])

    if not ea_files:
        print_yellow(f"WARNING: No .ex5 files found in: {args.mt5_ea_folder}")
        sys.exit(0)

    # Filter EAs based on strategies_json if provided
    if args.strategies_json:
        print_gray(f"Loading strategies from: {args.strategies_json}")
        top_strategies = load_top_strategies(args.strategies_json, args.max_strategies)

        if not top_strategies:
            print_yellow("WARNING: No strategies found in JSON. No EAs to backtest.")
            sys.exit(0)

        print_green(f"Found {len(top_strategies)} top ranked strategies to backtest:")
        for strat in top_strategies:
            print_gray(f"  - {strat}")
        print()

        # Filter EA files to only those matching top ranked strategies
        filtered_eas = []
        for ea in ea_files:
            if match_ea_to_strategy(ea, top_strategies):
                filtered_eas.append(ea)

        if not filtered_eas:
            print_yellow("WARNING: No EA files match the top ranked strategies.")
            print_gray("Available EAs:")
            for ea in ea_files[:10]:
                print_gray(f"  - {ea}")
            if len(ea_files) > 10:
                print_gray(f"  ... and {len(ea_files) - 10} more")
            sys.exit(0)
        
        print_green(f"Matched {len(filtered_eas)} EA(s) to backtest:")
        for ea in filtered_eas:
            print_gray(f"  - {ea}")
        print()
        
        ea_files = filtered_eas

    model_desc = MODEL_DESCRIPTIONS.get(args.model, "Unknown")
    print_green(f"Found {len(ea_files)} EA(s) to backtest")
    print_gray(f"Model: {args.model} ({model_desc})")
    print_gray(f"Date Range: {args.from_date} to {args.to_date}")
    print_gray(f"Timeout: {format_duration_friendly(args.timeout)} per backtest (backstop)")
    print_gray(f"Idle watchdog: kill after {args.idle_minutes} min below {args.idle_threshold}% CPU")
    print()

    # Path for the ini file - stored in MT5 data folder for reliability
    ini_path = os.path.join(args.mt5_data_folder, "backtest.ini")

    # Process each EA
    counter = 0
    successful = 0
    failed = 0
    force_killed = 0
    total_eas = len(ea_files)

    for ea_file in ea_files:
        counter += 1

        print_gray("------------------------------------------------------------")
        print_yellow(f"[{counter}/{total_eas}] Processing: {ea_file}")

        # Parse the EA filename
        parsed = parse_ea_name(ea_file)

        if parsed is None:
            print_yellow(f"  Skipping EA due to naming issue: {ea_file}")
            failed += 1
            continue

        print_gray(f"  Symbol: {parsed['symbol']}.QDM | Timeframe: {parsed['timeframe']}")

        # Create the ini file
        create_ini_file(
            ea_base_name=parsed["base_name"],
            symbol=parsed["symbol"],
            timeframe=parsed["timeframe"],
            ini_path=ini_path,
            model=args.model,
            from_date=args.from_date,
            to_date=args.to_date,
        )

        # Expected report path (for timeout detection)
        report_filename = f"{parsed['base_name']} MT5.htm"
        expected_report_path = os.path.join(reports_folder, report_filename)

        # Run the backtest
        bt_start = time.time()
        result, was_killed = run_backtest(
            mt5_terminal_path=args.mt5_terminal_path,
            ini_path=ini_path,
            report_path=expected_report_path,
            timeout_seconds=args.timeout,
            idle_threshold=args.idle_threshold,
            idle_minutes=args.idle_minutes,
        )
        bt_duration = time.time() - bt_start

        if result == 0:
            successful += 1
            if was_killed:
                force_killed += 1
        else:
            failed += 1

        print_gray(f"  Duration: {format_duration(bt_duration)}")

    print()
    print_cyan("============================================================")
    print_green("All backtests completed!")
    print_cyan("============================================================")

    # Copy reports to destination folder
    print()
    print_yellow(f"Copying reports to: {args.report_dest_folder}")

    report_files = []
    if os.path.exists(reports_folder):
        report_files = [f for f in os.listdir(reports_folder) if os.path.isfile(os.path.join(reports_folder, f))]

    if report_files:
        for report in report_files:
            src = os.path.join(reports_folder, report)
            dst = os.path.join(args.report_dest_folder, report)
            shutil.copy2(src, dst)
            print_gray(f"  Copied: {report}")
        print_green(f"Copied {len(report_files)} report(s)")
    else:
        print_yellow("WARNING: No report files found in: " + reports_folder)

    # Summary
    end_datetime = datetime.now()
    total_duration = time.time() - start_time
    duration_str = format_duration_friendly(total_duration)

    print()
    print_cyan("============================================================")
    print_cyan("SUMMARY")
    print_cyan("============================================================")
    print(f"{Colours.CYAN}EAs Processed:{Colours.RESET}    {Colours.GREEN}{counter}{Colours.RESET}")
    print(f"{Colours.CYAN}Successful:{Colours.RESET}       {Colours.GREEN}{successful}{Colours.RESET}")
    if force_killed > 0:
        print(f"{Colours.CYAN}Force-killed:{Colours.RESET}     {Colours.YELLOW}{force_killed}{Colours.RESET} (report saved despite hang)")
    if failed > 0:
        print(f"{Colours.CYAN}Failed/Skipped:{Colours.RESET}  {Colours.RED}{failed}{Colours.RESET}")
    print(f"{Colours.CYAN}Reports Location:{Colours.RESET} {Colours.GREEN}{args.report_dest_folder}{Colours.RESET}")
    print()
    print(f"{Colours.CYAN}Finished:{Colours.RESET}        {Colours.GREEN}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colours.RESET}")
    print(f"{Colours.CYAN}Total Duration:{Colours.RESET}  {Colours.GREEN}{duration_str}{Colours.RESET}")
    print()

    # Cleanup - remove temporary reports folder
    if os.path.exists(reports_folder):
        shutil.rmtree(reports_folder)
        print_gray("Cleaned up temporary reports folder")


if __name__ == "__main__":
    main()
