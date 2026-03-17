#!/usr/bin/env python3
"""
Batch compile MQL5 files using MetaEditor and deploy compiled EAs

Compiles all .mq5 files and deploys compiled .ex5 EAs to a target MT5 folder.
When an EA folder (-e) is specified:
  1. Existing .ex5 files in the EA folder are deleted first
  2. .mq5 source files are copied there for compilation
  3. After compilation, the .mq5 copies are cleaned up

This ensures a clean deployment with only the newly compiled EAs.

Note: MetaEditor will NOT output .ex5 files to network/mapped drives.
      Using -e with a local MT5 Experts path ensures reliable .ex5 generation.

Error log (compile_errors.log) is only created if there are failures.
Any existing error log is deleted at the start of each run.

Usage Examples:
---------------
    # Compile .mq5 files from source folder and deploy to EA folder (recommended)
    python step5_compile_mt5_eas.py -s "D:\\Strategies" -m "D:\\MT5" -i "C:\\Users\\Mark\\MT5\\MQL5" -e "C:\\Users\\Mark\\MT5\\MQL5\\Experts\\MyStrats"

    # Compile all .mq5 files in current directory (no EA deployment)
    python step5_compile_mt5_eas.py

    # Compile files in a specific local folder (no EA deployment)
    python step5_compile_mt5_eas.py -s "C:\\Users\\Mark\\MT5\\MQL5\\Experts\\MyStrategies"

    # Specify custom MetaEditor path (folder or exe)
    python step5_compile_mt5_eas.py -s "D:\\Strategies" -m "D:\\MT5\\metaeditor64.exe"

    # Include custom include path (point to MQL5 folder, NOT MQL5\\Include - MetaEditor adds \\Include automatically)
    python step5_compile_mt5_eas.py -s "D:\\Strategies" -i "C:\\Users\\Mark\\MT5\\MQL5"

    # Clear stale Tester .set cache before compiling (fixes wrong parameter values in backtests)
    python step5_compile_mt5_eas.py -s "D:\\Strategies" -e "C:\\Users\\Mark\\MT5\\MQL5\\Experts\\MyStrats" -t "C:\\Users\\Mark\\AppData\\Roaming\\MetaQuotes\\Terminal\\<instance_id>\\MQL5\\Profiles\\Tester"

    # Show help
    python step5_compile_mt5_eas.py --help

Arguments:
----------
    -s, --source-folder     Path to folder containing .mq5 files (default: current directory)
    -m, --metaeditor-path   Path to metaeditor64.exe or its parent folder (auto-detected if not specified)
    -i, --include-path      Additional include path for compilation
    -e, --ea-folder         Local path to deploy compiled .ex5 files (e.g. MT5 Experts folder).
                            Existing .ex5 files are deleted, then source .mq5 files are copied
                            here for compilation and cleaned up afterward.
    -t, --tester-profile    Path to MT5 Tester profiles folder whose cached .set files should be
                            deleted before compilation. Stale .set files cause the Strategy Tester
                            to use old parameter values (e.g. wrong lot size) instead of EA defaults.
                            Typically: %AppData%\\MetaQuotes\\Terminal\\<instance_id>\\MQL5\\Profiles\\Tester
"""

import argparse
import shutil
import subprocess
import sys
import os
import re
import time
from pathlib import Path
from datetime import datetime


# ANSI color codes for terminal output
class Colors:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GRAY = '\033[90m'
    RESET = '\033[0m'


def print_color(message: str, color: str = Colors.RESET, end: str = '\n'):
    """Print colored message to console."""
    # Enable ANSI colors on Windows
    if sys.platform == 'win32':
        os.system('')
    print(f"{color}{message}{Colors.RESET}", end=end)


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


def find_metaeditor() -> Path | None:
    """Auto-detect MetaEditor installation path."""
    common_paths = [
        Path(r"C:\Program Files\MetaTrader 5\metaeditor64.exe"),
        Path(r"C:\Program Files (x86)\MetaTrader 5\metaeditor64.exe"),
        Path(r"C:\Program Files\Pepperstone MetaTrader 5\metaeditor64.exe"),
        Path(os.environ.get('APPDATA', ''), '..', 'Local', 'Programs', 'MetaTrader 5', 'metaeditor64.exe'),
    ]
    
    for path in common_paths:
        if path.exists():
            return path.resolve()
    
    return None


def compile_mq5_file(
    metaeditor_path: Path,
    source_file: Path,
    temp_log_file: Path,
    include_path: Path | None = None
) -> tuple[bool, str, int]:
    """
    Compile a single MQ5 file.
    
    Returns:
        tuple: (success: bool, log_content: str, exit_code: int)
    """
    # Build command string with explicit quoting
    # MetaEditor requires paths quoted after the colon: /compile:"path with spaces"
    cmd = f'"{metaeditor_path}" /compile:"{source_file}" /log:"{temp_log_file}"'
    
    if include_path:
        cmd += f' /include:"{include_path}"'
    
    # Run MetaEditor
    result = subprocess.run(cmd, capture_output=True, shell=True)
    
    # Read log file
    log_content = ""
    has_errors = False
    
    if temp_log_file.exists():
        try:
            # MetaEditor writes logs in UTF-16 LE (Unicode)
            log_content = temp_log_file.read_text(encoding='utf-16-le')
        except UnicodeDecodeError:
            try:
                log_content = temp_log_file.read_text(encoding='utf-8')
            except UnicodeDecodeError:
                log_content = "Unable to read log file"
        
        # Parse the Result line for error count (handles both "error(s)" and "errors")
        match = re.search(r'Result:\s*(\d+)\s+error', log_content)
        if match:
            error_count = int(match.group(1))
            has_errors = error_count > 0
        else:
            # No Result line found - check for error indicators in log
            log_lower = log_content.lower()
            if 'error' in log_lower and 'information' not in log_lower.split('error')[0][-20:]:
                has_errors = True
        
        # Clean up temp log
        temp_log_file.unlink(missing_ok=True)
    
    # MetaEditor exit codes are unreliable (returns 1 even on success)
    # So we rely primarily on log file parsing for success/failure
    success = not has_errors
    return success, log_content, result.returncode


def main():
    parser = argparse.ArgumentParser(
        description='Batch compile MQL5 files using MetaEditor and deploy compiled EAs'
    )
    parser.add_argument(
        '-s', '--source-folder',
        type=Path,
        default=Path.cwd(),
        help='Path to folder containing .mq5 files (default: current directory)'
    )
    parser.add_argument(
        '-m', '--metaeditor-path',
        type=Path,
        default=None,
        help='Path to metaeditor64.exe or its parent folder (auto-detected if not specified)'
    )
    parser.add_argument(
        '-i', '--include-path',
        type=Path,
        default=None,
        help='Additional include path for compilation'
    )
    parser.add_argument(
        '-e', '--ea-folder',
        type=Path,
        default=None,
        help='Local path to deploy compiled .ex5 files (e.g. MT5 Experts folder)'
    )
    parser.add_argument(
        '-t', '--tester-profile',
        type=Path,
        default=None,
        help=(
            'Path to MT5 Tester profiles folder to clear cached .set files before compilation. '
            'Typically: %%AppData%%\\MetaQuotes\\Terminal\\<instance_id>\\MQL5\\Profiles\\Tester'
        )
    )
    
    args = parser.parse_args()
    
    # Record start time
    start_time = time.time()
    start_datetime = datetime.now()

    print_color("============================================================", Colors.CYAN)
    print_color(" MQL5 Batch Compiler", Colors.CYAN)
    print_color("============================================================", Colors.CYAN)
    print()
    print(f"{Colors.CYAN}Started:{Colors.RESET} {Colors.GREEN}{start_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print()

    # Resolve source folder
    source_folder = args.source_folder.resolve()
    if not source_folder.exists():
        print_color(f"Source folder not found: {source_folder}", Colors.RED)
        sys.exit(1)
    
    # Find MetaEditor
    metaeditor_path = args.metaeditor_path
    if metaeditor_path is None:
        metaeditor_path = find_metaeditor()
        if metaeditor_path is None:
            print_color(
                "MetaEditor not found. Please specify path with --metaeditor-path",
                Colors.RED
            )
            sys.exit(1)
    else:
        metaeditor_path = metaeditor_path.resolve()
        # If a directory was passed, look for metaeditor64.exe inside it
        if metaeditor_path.is_dir():
            exe_path = metaeditor_path / "metaeditor64.exe"
            if exe_path.exists():
                metaeditor_path = exe_path
            else:
                print_color(f"metaeditor64.exe not found in: {metaeditor_path}", Colors.RED)
                sys.exit(1)
        elif not metaeditor_path.exists():
            print_color(f"MetaEditor not found at: {metaeditor_path}", Colors.RED)
            sys.exit(1)
    
    # Resolve EA folder if provided
    ea_folder = None
    if args.ea_folder:
        ea_folder = args.ea_folder.resolve()
    
    # Fixed log filename - define early so we can clear it (always in source folder)
    error_log_path = source_folder / "compile_errors.log"
    
    # Delete existing log file if present
    if error_log_path.exists():
        error_log_path.unlink()
        print_color("Removed existing log file", Colors.GRAY)

    # Clear cached .set files from the MT5 Tester profiles folder.
    # Stale .set files cause the Strategy Tester to use old parameter values
    # (e.g. mmLots = 1.0) instead of the EA's compiled defaults.
    # Only .set files are deleted — .ini files (e.g. Custom_Strategy_Tester_Settings.ini)
    # and subfolders (e.g. Groups) are intentionally preserved.
    if args.tester_profile:
        tester_folder = Path(os.path.normpath(args.tester_profile))
        if tester_folder.exists() and tester_folder.is_dir():
            # Delete all files starting with "SQ " (covers both .set and .ini SQ strategy files).
            # Preserves manually created files (e.g. CommercialsIndex.set, Moving Average.set,
            # Custom_Strategy_Tester_Settings.ini) and subfolders (e.g. Groups).
            sq_files = [f for f in tester_folder.iterdir() if f.is_file() and f.name.startswith("SQ ")]
            if sq_files:
                print_color(
                    f"Clearing {len(sq_files)} cached SQ file(s) from Tester profiles folder...",
                    Colors.CYAN
                )
                cleared = 0
                skipped = 0
                for f in sq_files:
                    try:
                        f.unlink()
                        cleared += 1
                    except Exception as e:
                        print_color(f"  Skipped {f.name}: {e}", Colors.YELLOW)
                        skipped += 1
                msg = f"  Cleared {cleared} SQ file(s)"
                if skipped:
                    msg += f", skipped {skipped} (in use?)"
                print_color(msg, Colors.GREEN)
            else:
                print_color("No SQ files found in Tester profiles folder", Colors.GRAY)
        else:
            print_color(
                f"Tester profiles folder not found, skipping: {tester_folder}",
                Colors.YELLOW
            )

    print_color(f"Using MetaEditor: {metaeditor_path}", Colors.CYAN)
    print_color(f"Source folder: {source_folder}", Colors.CYAN)
    if ea_folder:
        print_color(f"EA output folder: {ea_folder}", Colors.CYAN)
    print()
    
    # Get all .mq5 files from source
    mq5_files = list(source_folder.glob("*.mq5"))
    
    if not mq5_files:
        print_color(f"No .mq5 files found in {source_folder}", Colors.YELLOW)
        sys.exit(0)
    
    print_color(f"Found {len(mq5_files)} .mq5 file(s) to compile", Colors.GREEN)
    
    # If EA folder specified, copy .mq5 files there first and compile from there.
    # MetaEditor only outputs .ex5 to local paths, so compiling inside the EA folder
    # (which should be under the MT5 data directory on C:) ensures .ex5 files are created.
    copied_mq5_files = []
    if ea_folder:
        ea_folder.mkdir(parents=True, exist_ok=True)
        
        # Delete existing .ex5 files in EA folder before compiling new ones
        existing_ex5 = list(ea_folder.glob("*.ex5"))
        if existing_ex5:
            print()
            print_color(f"Removing {len(existing_ex5)} existing .ex5 file(s) from EA folder...", Colors.CYAN)
            removed = 0
            for ex5_file in existing_ex5:
                try:
                    ex5_file.unlink()
                    removed += 1
                except Exception as e:
                    print_color(f"  Failed to remove {ex5_file.name}: {e}", Colors.YELLOW)
            print_color(f"  Removed {removed} .ex5 file(s)", Colors.GREEN)
        print()
        print_color("Staging .mq5 files to EA folder for compilation...", Colors.CYAN)
        for mq5_file in mq5_files:
            dest = ea_folder / mq5_file.name
            try:
                shutil.copy2(mq5_file, dest)
                copied_mq5_files.append(dest)
            except Exception as e:
                print_color(f"  Failed to copy {mq5_file.name}: {e}", Colors.RED)
        print_color(f"  Staged {len(copied_mq5_files)} .mq5 file(s) to {ea_folder}", Colors.GREEN)
    
    print()
    print("=" * 60)
    
    # Track results
    failed_compiles = []
    success_count = 0
    temp_log_file = Path(os.environ.get('TEMP', '/tmp')) / "mql5_compile_temp.log"
    
    # Compile from EA folder if we staged there, otherwise from source
    compile_files = copied_mq5_files if copied_mq5_files else mq5_files
    
    for mq5_file in compile_files:
        print_color(f"Compiling: {mq5_file.name}... ", Colors.RESET, end='')
        
        success, log_content, exit_code = compile_mq5_file(
            metaeditor_path,
            mq5_file,
            temp_log_file,
            args.include_path
        )
        
        if success:
            # Verify .ex5 was actually created
            ex5_file = mq5_file.with_suffix('.ex5')
            if ex5_file.exists():
                print_color("OK", Colors.GREEN)
            else:
                print_color("OK (compiled, .ex5 not found)", Colors.YELLOW)
            success_count += 1
        else:
            print_color("FAILED", Colors.RED)
            failed_compiles.append({
                'filename': mq5_file.name,
                'full_path': str(mq5_file),
                'exit_code': exit_code,
                'log_output': log_content
            })
    
    print()
    print("=" * 60)
    print_color(
        f"Compilation complete: {success_count} succeeded, {len(failed_compiles)} failed",
        Colors.CYAN
    )
    
    # Write failure log only if there were failures
    if failed_compiles:
        log_lines = [
            "MQL5 Batch Compilation Error Log",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Source Folder: {source_folder}",
            "=" * 60,
            ""
        ]
        
        for fail in failed_compiles:
            log_lines.extend([
                f"FILE: {fail['filename']}",
                f"PATH: {fail['full_path']}",
                f"EXIT CODE: {fail['exit_code']}",
                "LOG OUTPUT:",
                fail['log_output'],
                "-" * 60,
                ""
            ])
        
        error_log_path.write_text('\n'.join(log_lines), encoding='utf-8')
        
        print()
        print_color(f"Error log written to: {error_log_path}", Colors.YELLOW)
    else:
        print_color("No errors - no log file created", Colors.GREEN)
    
    # Clean up: remove staged .mq5 copies from EA folder (keep only .ex5)
    if copied_mq5_files:
        print()
        print_color("Cleaning up staged .mq5 files from EA folder...", Colors.CYAN)
        cleaned = 0
        for mq5_copy in copied_mq5_files:
            try:
                mq5_copy.unlink(missing_ok=True)
                cleaned += 1
            except Exception as e:
                print_color(f"  Failed to remove {mq5_copy.name}: {e}", Colors.YELLOW)
        print_color(f"  Removed {cleaned} .mq5 file(s)", Colors.GREEN)
    
    # Report final .ex5 count in EA folder
    if ea_folder and ea_folder.exists():
        ex5_count = len(list(ea_folder.glob("*.ex5")))
        print()
        print("=" * 60)
        print_color(f"Deployed {ex5_count} .ex5 EA(s) to: {ea_folder}", Colors.CYAN)

    # Record end time
    end_datetime = datetime.now()
    total_duration = time.time() - start_time

    print()
    print_color("============================================================", Colors.CYAN)
    print(f"{Colors.CYAN}Finished:{Colors.RESET}       {Colors.GREEN}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print(f"{Colors.CYAN}Total Duration:{Colors.RESET} {Colors.GREEN}{format_duration(total_duration)}{Colors.RESET}")


if __name__ == "__main__":
    main()
