"""
QuantDataManager Data Refresh Script.

Refreshes all data in QuantDataManager using the CLI tool.

Usage:
    python Step1_Refresh_QDM_Data.py <path_to_qdmcli>

Example:
    python Step1_Refresh_QDM_Data.py "C:\\QuantDataManager125\\qdmcli.exe"

Arguments:
    qdm_path    Full path to the qdmcli.exe executable.

"""

import argparse
import os
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


def main():
    init_colors()
    
    parser = argparse.ArgumentParser(description="Refresh QuantDataManager data.")
    parser.add_argument(
        "qdm_path",
        help="Path to qdmcli.exe (e.g. C:\\QuantDataManager125\\qdmcli.exe)",
    )
    args = parser.parse_args()

    # Record start time
    start_time = time.time()
    start_datetime = datetime.now()

    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print(f"{Colors.CYAN} QuantDataManager Data Refresh{Colors.RESET}")
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Started:{Colors.RESET} {Colors.GREEN}{start_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print()

    # Update all data
    result = subprocess.run(
        [args.qdm_path, "-data", "action=update"],
        check=False,
    )

    # Record end time
    end_datetime = datetime.now()
    total_duration = time.time() - start_time

    if result.returncode != 0:
        print(f"{Colors.RED}Data refresh failed with exit code {result.returncode}{Colors.RESET}", file=sys.stderr)
        print()
        print(f"{Colors.CYAN}Finished:{Colors.RESET}       {Colors.RED}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
        print(f"{Colors.CYAN}Total Duration:{Colors.RESET} {Colors.RED}{format_duration(total_duration)}{Colors.RESET}")
        sys.exit(result.returncode)

    print()
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print(f"{Colors.GREEN}Data Refresh Complete!{Colors.RESET}")
    print(f"{Colors.CYAN}============================================{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Finished:{Colors.RESET}       {Colors.GREEN}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print(f"{Colors.CYAN}Total Duration:{Colors.RESET} {Colors.GREEN}{format_duration(total_duration)}{Colors.RESET}")


if __name__ == "__main__":
    main()
