"""
QuantDataManager Batch Export Script.

Exports tick data for all symbols found in QuantDataManager's history folder
to MT5 format, then renames the exported files from SYMBOL.QDM.csv to SYMBOL_QDM.csv.

Symbols are discovered automatically by reading folder names from:
    <qdm_path>\\user\\data\\History

Usage:
    python Step2_Export_Data_From_QDM.py --date-from <date> --date-to <date> --qdm-path <path> --export-path <path>

Example:
    python Step2_Export_Data_From_QDM.py --date-from 2025.01.01 --date-to 2025.12.31 --qdm-path "C:\\QuantDataManager125" --export-path "C:\\Users\\msand\\...\\MQL5\\Files"

Arguments:
    --date-from       Start date for export (e.g. 2025.01.01).
    --date-to         End date for export (e.g. 2025.12.31).
    --qdm-path        Path to the QuantDataManager folder (qdmcli.exe is appended automatically).
    --export-path     Output directory for exported CSV files.
"""

import argparse
import os
import re
import subprocess
import sys
import time
from datetime import datetime


# ANSI color codes for terminal output
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


def get_symbols(qdm_path: str) -> list[str]:
    """Discover symbols by reading folder names from the QDM history directory."""
    history_path = os.path.join(qdm_path, "user", "data", "History")

    if not os.path.isdir(history_path):
        print(f"Error: History folder not found at {history_path}", file=sys.stderr)
        sys.exit(1)

    symbols = [
        entry for entry in os.listdir(history_path)
        if os.path.isdir(os.path.join(history_path, entry))
    ]

    if not symbols:
        print(f"Error: No symbol folders found in {history_path}", file=sys.stderr)
        sys.exit(1)

    symbols.sort()
    return symbols


def export_symbols(qdm_cli: str, export_path: str, date_from: str, date_to: str, symbols: list[str]) -> None:
    """Export tick data for each symbol using qdmcli."""
    print(f"Date range: {date_from} to {date_to}")
    print(f"Symbols found: {len(symbols)}")
    print()

    for symbol in symbols:
        print(f"Exporting {symbol}...")

        result = subprocess.run(
            [
                qdm_cli,
                "-data",
                f"action=exportToMT5",
                f"symbol={symbol}",
                f"datefrom={date_from}",
                f"dateto={date_to}",
                f"timeframe=TICK",
                f"outputdir={export_path}\\",
            ],
            check=False,
            stderr=subprocess.DEVNULL,
        )

        if result.returncode != 0:
            print(f"  Warning: {symbol} export returned exit code {result.returncode}", file=sys.stderr)
        else:
            print(f"Completed: {symbol}")


def rename_exported_files(export_path: str) -> None:
    """Rename exported files from SYMBOL.QDM.csv to SYMBOL_QDM.csv."""
    print()
    print("Renaming exported files...")

    for filename in os.listdir(export_path):
        if re.search(r"\.QDM\.csv$", filename):
            new_name = re.sub(r"\.QDM\.csv$", "_QDM.csv", filename)
            old_path = os.path.join(export_path, filename)
            new_path = os.path.join(export_path, new_name)

            # Remove existing file with new name if it exists (to allow overwrite)
            if os.path.exists(new_path):
                os.remove(new_path)

            os.rename(old_path, new_path)
            print(f"  Renamed: {filename} -> {new_name}")


def main():
    init_colors()
    
    parser = argparse.ArgumentParser(description="Export tick data from QuantDataManager to MT5 format.")
    parser.add_argument("--date-from", required=True, help="Start date for export (e.g. 2025.01.01)")
    parser.add_argument("--date-to", required=True, help="End date for export (e.g. 2025.12.31)")
    parser.add_argument("--qdm-path", required=True, help="Path to the QuantDataManager folder")
    parser.add_argument("--export-path", required=True, help="Output directory for exported CSV files")
    args = parser.parse_args()

    # Record start time
    start_time = time.time()
    start_datetime = datetime.now()

    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print(f"{Colors.CYAN} QuantDataManager Batch Export{Colors.RESET}")
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Started:{Colors.RESET} {Colors.GREEN}{start_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print()

    qdm_cli = os.path.join(args.qdm_path, "qdmcli.exe")

    if not os.path.isfile(qdm_cli):
        print(f"{Colors.RED}Error: qdmcli.exe not found at {qdm_cli}{Colors.RESET}", file=sys.stderr)
        sys.exit(1)

    symbols = get_symbols(args.qdm_path)
    export_symbols(qdm_cli, args.export_path, args.date_from, args.date_to, symbols)
    rename_exported_files(args.export_path)

    # Record end time
    end_datetime = datetime.now()
    total_duration = time.time() - start_time

    print()
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print(f"{Colors.GREEN}All exports completed!{Colors.RESET}")
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Finished:{Colors.RESET}       {Colors.GREEN}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print(f"{Colors.CYAN}Total Duration:{Colors.RESET} {Colors.GREEN}{format_duration(total_duration)}{Colors.RESET}")


if __name__ == "__main__":
    main()
