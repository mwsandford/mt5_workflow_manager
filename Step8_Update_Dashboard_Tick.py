#!/usr/bin/env python3
"""
Step 9: Update Dashboard with Tick Monte Carlo Results
=======================================================

This script updates the Performance Dashboard with MC95 Tick values
after running tick-based backtests and Monte Carlo analysis.

It reads:
  - BatchMC_Results.csv from the ticks subfolder
  - strategies_data.json from the Dashboard folder

It updates:
  - index.html - the MC95 TICK column
  - strategies_data.json - adds mc95_ret_dd_tick values

Usage:
    python Step9_Update_Dashboard_Tick.py <dashboard_folder> --tick-mc-results <path_to_tick_mc_results.csv>
"""

import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path


# =============================================================================
# ANSI colour codes for terminal output
# =============================================================================
class Colors:
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    GRAY = "\033[90m"
    RESET = "\033[0m"


def print_cyan(msg): print(f"{Colors.CYAN}{msg}{Colors.RESET}")
def print_green(msg): print(f"{Colors.GREEN}{msg}{Colors.RESET}")
def print_yellow(msg): print(f"{Colors.YELLOW}{msg}{Colors.RESET}")
def print_red(msg): print(f"{Colors.RED}{msg}{Colors.RESET}")
def print_gray(msg): print(f"{Colors.GRAY}{msg}{Colors.RESET}")


def load_mc_results(csv_path: str) -> dict:
    """
    Load Monte Carlo results from BatchMC_Results.csv.
    
    Returns dict mapping strategy name -> MC95 RetDD value
    """
    results = {}
    
    if not os.path.exists(csv_path):
        print_red(f"ERROR: MC results file not found: {csv_path}")
        return results
    
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            strategy = row.get('Strategy', '').strip()
            confidence = row.get('ConfidenceLevel', '').strip()
            ret_dd = row.get('RetDD', '').strip()
            
            # We want the 95% confidence level
            if confidence == '95' and strategy:
                try:
                    results[strategy] = float(ret_dd) if ret_dd else None
                except ValueError:
                    results[strategy] = None
    
    return results


def normalize_strategy_name(name: str) -> str:
    """Normalize strategy name for matching."""
    # Remove extra spaces, convert to uppercase
    return ' '.join(name.upper().split())


def match_strategy_names(mc_name: str, dashboard_name: str) -> bool:
    """Check if MC strategy name matches Dashboard strategy name."""
    mc_norm = normalize_strategy_name(mc_name)
    dash_norm = normalize_strategy_name(dashboard_name)
    
    # Exact match
    if mc_norm == dash_norm:
        return True
    
    # One contains the other
    if mc_norm in dash_norm or dash_norm in mc_norm:
        return True
    
    # Handle version variations (1.107 vs 1.1.107)
    mc_simple = mc_norm.replace('.', ' ').replace('  ', ' ')
    dash_simple = dash_norm.replace('.', ' ').replace('  ', ' ')
    if mc_simple == dash_simple:
        return True
    
    return False


def find_mc95_for_strategy(strategy_name: str, mc_results: dict) -> float | None:
    """Find MC95 RetDD value for a strategy, handling name variations."""
    # Try exact match first
    if strategy_name in mc_results:
        return mc_results[strategy_name]
    
    # Try fuzzy matching
    for mc_name, value in mc_results.items():
        if match_strategy_names(mc_name, strategy_name):
            return value
    
    return None


def update_strategies_json(json_path: str, mc_results: dict) -> int:
    """
    Update strategies_data.json with MC95 Tick values.
    
    Returns number of strategies updated.
    """
    if not os.path.exists(json_path):
        print_red(f"ERROR: strategies_data.json not found: {json_path}")
        return 0
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    updated_count = 0
    
    # Update ranking entries
    for strat in data.get('ranking', []):
        name = strat.get('name', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
            updated_count += 1
            print_gray(f"  Updated {name}: MC95 Tick = {mc95_tick:.2f}")
    
    # Update portfolio entries
    for strat in data.get('portfolio', []):
        name = strat.get('name', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
    
    # Save updated JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    
    return updated_count


def update_dashboard_html(html_path: str, mc_results: dict) -> int:
    """
    Update Dashboard index.html with MC95 Tick values.

    The dashboard generated by Step7 embeds all data as a single JS object:
        const DATA = { ranking: [...], portfolio: [...], ... };
    The table is rendered dynamically by JavaScript, so we parse and update
    the embedded DATA object directly.

    Returns number of strategies updated.
    """
    if not os.path.exists(html_path):
        print_red(f"ERROR: Dashboard HTML not found: {html_path}")
        return 0

    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    updated_count = 0

    # -------------------------------------------------------------------------
    # Parse the embedded "const DATA = {...};" object from the <script> block
    # -------------------------------------------------------------------------
    js_data_pattern = r'(const\s+DATA\s*=\s*)(.*?)(;\s*\n)'
    js_match = re.search(js_data_pattern, html_content, re.DOTALL)

    if not js_match:
        print_red("  ERROR: Could not find 'const DATA = ...' in dashboard HTML")
        return 0

    try:
        js_data = json.loads(js_match.group(2))
    except json.JSONDecodeError as e:
        print_red(f"  ERROR: Could not parse embedded DATA object: {e}")
        return 0

    # Update ranking entries
    for strat in js_data.get('ranking', []):
        name = strat.get('name', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            old_val = strat.get('mc95_ret_dd_tick')
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
            updated_count += 1
            print_gray(f"  Ranking updated: {name} -> {mc95_tick:.2f} (was: {old_val})")

    # Update portfolio entries
    for strat in js_data.get('portfolio', []):
        name = strat.get('name', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)

    if updated_count > 0:
        # Replace the DATA object in the HTML
        new_js_data = json.dumps(js_data)
        html_content = (
            html_content[:js_match.start()]
            + f'{js_match.group(1)}{new_js_data}{js_match.group(3)}'
            + html_content[js_match.end():]
        )
        print_green(f"  Updated {updated_count} strategies in embedded DATA object")

        # Save updated HTML
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    else:
        print_yellow("  No matching strategies found in embedded DATA object")

    return updated_count


def main():
    parser = argparse.ArgumentParser(
        description='Update Dashboard with Tick Monte Carlo Results',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    
    parser.add_argument(
        'dashboard_folder',
        help='Path to Dashboard folder containing index.html and strategies_data.json'
    )
    parser.add_argument(
        '--tick-mc-results',
        required=True,
        help='Path to BatchMC_Results.csv from tick Monte Carlo analysis'
    )
    
    args = parser.parse_args()
    
    print_cyan("=" * 60)
    print_cyan("Update Dashboard with Tick MC Results")
    print_cyan("=" * 60)
    print()
    
    # Validate paths
    dashboard_folder = args.dashboard_folder
    if not os.path.isdir(dashboard_folder):
        print_red(f"ERROR: Dashboard folder not found: {dashboard_folder}")
        sys.exit(1)
    
    html_path = os.path.join(dashboard_folder, 'index.html')
    json_path = os.path.join(dashboard_folder, 'strategies_data.json')
    
    if not os.path.exists(html_path):
        print_red(f"ERROR: Dashboard HTML not found: {html_path}")
        sys.exit(1)
    
    if not os.path.exists(json_path):
        print_red(f"ERROR: strategies_data.json not found: {json_path}")
        sys.exit(1)
    
    # Load tick MC results
    print_gray(f"Loading tick MC results from: {args.tick_mc_results}")
    mc_results = load_mc_results(args.tick_mc_results)
    
    if not mc_results:
        print_yellow("WARNING: No MC95 results found in tick MC results file")
        sys.exit(0)
    
    print_green(f"Found {len(mc_results)} strategy results:")
    for name, value in mc_results.items():
        formatted = f"{value:.2f}" if value else "N/A"
        print_gray(f"  {name}: {formatted}")
    print()
    
    # Update strategies_data.json
    print_gray("Updating strategies_data.json...")
    json_updated = update_strategies_json(json_path, mc_results)
    print_green(f"  Updated {json_updated} strategies in JSON")
    print()
    
    # Update Dashboard HTML
    print_gray("Updating Dashboard HTML...")
    html_updated = update_dashboard_html(html_path, mc_results)
    print_green(f"  Updated {html_updated} strategies in HTML")
    print()
    
    print_cyan("=" * 60)
    print_green("Dashboard update complete!")
    print_cyan("=" * 60)


if __name__ == '__main__':
    main()
