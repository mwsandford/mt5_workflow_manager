#!/usr/bin/env python3
"""
Step 10: Update Tick Survival Dashboard
========================================

Parses the latest run's Dashboard folder, persists strategy characteristics
and backtest results into the SQLite store under tick_survival/, regenerates
the Tick Survival Analysis HTML dashboard, and opens it in the default browser.

Usage:
    python Step10_Update_Tick_Survival.py --dashboard-folder <path>
        [--output-folder <path>]       # defaults to parent of --dashboard-folder
        [--db-path <path>]             # defaults to tick_survival/tick_survival.db
        [--report-path <path>]         # defaults to tick_survival/reports/tick_survival_dashboard.html
        [--tick-threshold <float>]     # default 2.0
        [--batch-id <str>]             # override auto-derived batch id
        [--no-open]                    # skip opening in browser
"""

from __future__ import annotations

import argparse
import os
import sys
import webbrowser
from pathlib import Path


# ─────────────────────────────────────────────
# ANSI colour codes (mirrors other Step scripts)
# ─────────────────────────────────────────────
class Colors:
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    GRAY = "\033[90m"
    RESET = "\033[0m"


def print_cyan(msg):   print(f"{Colors.CYAN}{msg}{Colors.RESET}")
def print_green(msg):  print(f"{Colors.GREEN}{msg}{Colors.RESET}")
def print_yellow(msg): print(f"{Colors.YELLOW}{msg}{Colors.RESET}")
def print_red(msg):    print(f"{Colors.RED}{msg}{Colors.RESET}")
def print_gray(msg):   print(f"{Colors.GRAY}{msg}{Colors.RESET}")


def _script_dir() -> Path:
    """Folder containing this script (works frozen or not)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main() -> int:
    # Make the local tick_survival package importable even when run from a
    # different working directory.
    script_dir = _script_dir()
    if str(script_dir) not in sys.path:
        sys.path.insert(0, str(script_dir))

    from tick_survival import db as ts_db
    from tick_survival import parser as ts_parser
    from tick_survival import dashboard as ts_dashboard

    ap = argparse.ArgumentParser(
        description="Update Tick Survival Analysis DB + Dashboard",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--dashboard-folder", required=True,
                    help="Path to the run's Dashboard folder (contains strategies_data.json + index.html)")
    ap.add_argument("--output-folder", default=None,
                    help="Run output folder (parent of Dashboard). Used as .txt fallback source. Defaults to parent of --dashboard-folder.")
    ap.add_argument("--db-path", default=str(script_dir / "tick_survival" / "tick_survival.db"),
                    help="SQLite DB path (persists across runs).")
    ap.add_argument("--report-path", default=str(script_dir / "tick_survival" / "reports" / "tick_survival_dashboard.html"),
                    help="Output HTML dashboard path.")
    ap.add_argument("--tick-threshold", type=float, default=2.0,
                    help="MC95 Tick value at or above which a strategy is considered to have survived (default: 2.0).")
    ap.add_argument("--batch-id", default=None,
                    help="Explicit batch id. If omitted, derived from strategies_data.json's generated_at or file mtime.")
    ap.add_argument("--no-open", action="store_true",
                    help="Do not open the generated dashboard in the browser.")
    args = ap.parse_args()

    dashboard_folder = Path(args.dashboard_folder)
    if not dashboard_folder.is_dir():
        print_red(f"ERROR: Dashboard folder not found: {dashboard_folder}")
        return 1

    output_folder = Path(args.output_folder) if args.output_folder else dashboard_folder.parent

    print_cyan("=" * 60)
    print_cyan("Update Tick Survival Analysis")
    print_cyan("=" * 60)
    print_gray(f"Dashboard folder : {dashboard_folder}")
    print_gray(f"Output folder    : {output_folder}")
    print_gray(f"DB path          : {args.db_path}")
    print_gray(f"Report path      : {args.report_path}")
    print_gray(f"Tick threshold   : {args.tick_threshold}")
    print()

    # 1. Parse the current run
    print_gray("Parsing run data...")
    try:
        run = ts_parser.load_run(
            dashboard_folder=dashboard_folder,
            output_folder=output_folder,
            tick_threshold=args.tick_threshold,
            batch_id_override=args.batch_id,
        )
    except FileNotFoundError as e:
        print_red(f"ERROR: {e}")
        return 1

    records = run["records"]
    print_green(f"  Batch id      : {run['batch_id']}")
    print_green(f"  Run timestamp : {run['run_timestamp']}")
    print_green(f"  Records       : {len(records)} top-10 strategies with tick MC values")

    if not records:
        print_yellow("WARNING: no tick-tested strategies found in this run. DB unchanged.")

    # 2. Upsert into the DB
    print()
    print_gray("Writing to SQLite store...")
    conn = ts_db.init_db(args.db_path)
    try:
        with conn:  # implicit transaction
            for rec in records:
                strategy_id = ts_db.upsert_strategy(conn, rec["strategy"])
                ts_db.upsert_backtest_result(conn, strategy_id, rec["result"])
        strat_count = ts_db.row_count(conn, "Strategies")
        result_count = ts_db.row_count(conn, "BacktestResults")
        print_green(f"  Strategies     : {strat_count}")
        print_green(f"  BacktestResults: {result_count}")

        # 3. Build the dashboard from the full DB history
        print()
        print_gray("Generating HTML dashboard...")
        all_records = ts_db.fetch_joined_records(conn)
    finally:
        conn.close()

    out_path = ts_dashboard.write_html(
        records=all_records,
        out_path=args.report_path,
        default_threshold=args.tick_threshold,
    )
    print_green(f"  Wrote {out_path}")

    # 4. Open in default browser
    if not args.no_open:
        print()
        print_gray("Opening dashboard in default browser...")
        try:
            webbrowser.open(out_path.resolve().as_uri())
        except Exception as e:
            print_yellow(f"WARNING: could not open browser: {e}")

    print()
    print_cyan("=" * 60)
    print_green("Tick Survival Analysis update complete!")
    print_cyan("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
