#!/usr/bin/env python3
"""
Step 8: Update Dashboard with Tick Backtest + Monte Carlo Results
=================================================================

This script builds the *Tick* grid on the Performance Dashboard after the
tick-based backtests and Monte Carlo analysis have run.

It reads:
  - BatchMC_Results.csv from the ticks subfolder (tick MC95 values)
  - The tick MT5 .htm Strategy Tester reports in the ticks subfolder
  - strategies_data.json from the Dashboard folder (for canonical M1 names)

It updates:
  - index.html          - injects a full `tick_ranking` array into the embedded
                          DATA object (rendered as the toggleable "Tick" grid)
                          and keeps ranking[].mc95_ret_dd_tick populated
  - strategies_data.json - same two updates, for downstream consumers (Step 10)

The tick grid is built with the exact same parser + scoring as the M1 grid by
reusing Step 7's parse_mt5_report() / compute_mt5_rankings(), so metrics line up.
ranking[].mc95_ret_dd_tick is still populated (Step 10 / Tick Survival depends on it).

Usage:
    python Step8_Update_Dashboard_Tick.py <dashboard_folder> \\
        --tick-mc-results <path_to_tick_mc_results.csv> \\
        [--tick-reports-folder <folder_with_tick_htm_reports>]
"""

import argparse
import csv
import io
import json
import math
import os
import re
import sys
from pathlib import Path

# -----------------------------------------------------------------------------
# Reuse Step 7's MT5 report parser + ranking logic so the tick grid is built
# with the exact same metrics/scoring as the M1 grid (see project convention:
# don't re-implement strategy/report parsing — Step 7 already does it).
# -----------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

try:
    import Step7_Strategy_Ranking as step7
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    _STEP7_AVAILABLE = True
    _STEP7_IMPORT_ERROR = None
except Exception as _exc:  # pragma: no cover - defensive
    step7 = None
    plt = None
    mdates = None
    _STEP7_AVAILABLE = False
    _STEP7_IMPORT_ERROR = _exc


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


# =============================================================================
# Tick grid construction (parse tick .htm reports -> full ranking rows)
# =============================================================================
def _rnd(value, ndigits: int = 2):
    """Round defensively, coercing None/NaN/non-numeric to 0."""
    try:
        if value is None:
            return 0
        f = float(value)
        if math.isnan(f):
            return 0
        return round(f, ndigits)
    except (TypeError, ValueError):
        return 0


def pd_notna(value) -> bool:
    """True if value is not None/NaN (avoids importing pandas directly here)."""
    if value is None:
        return False
    try:
        return not math.isnan(float(value))
    except (TypeError, ValueError):
        return True


def strategy_name_from_report(path: str) -> str:
    """Derive the strategy name from an MT5 .htm report filename.

    'SQ NAS100 M30 15.1.485 MT5.htm' -> 'SQ NAS100 M30 15.1.485'
    """
    base = os.path.splitext(os.path.basename(path))[0]
    for suffix in (' MT5', ' MT4', '_MT5', '_MT4'):
        if base.endswith(suffix):
            base = base[:-len(suffix)]
            break
    return base.strip()


def canonical_name(name: str, m1_names: list) -> str:
    """Reconcile a tick strategy name to the M1 ranking spelling so that the
    dashboard's Overview/Strategy/chart lookups (keyed by M1 name) still work."""
    for mn in m1_names:
        if match_strategy_names(name, mn):
            return mn
    return name


def load_m1_ranking_names(json_path: str) -> list:
    """Read the existing M1 ranking strategy names from strategies_data.json."""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return [r.get('name', '') for r in data.get('ranking', []) if r.get('name')]
    except Exception:
        return []


def render_tick_chart(report_path: str, name: str) -> str:
    """Render the tick backtest equity curve as a base64 PNG data URI.

    Uses Step 7's Deals-table parser; single MetaTrader-5 (tick) curve.
    Returns '' if the equity curve can't be built.
    """
    if not _STEP7_AVAILABLE or not report_path:
        return ''
    try:
        equity = step7.parse_mt5_deals_equity(report_path)
    except Exception:
        equity = None
    if not equity or len(equity) < 2:
        return ''

    dates = [d for d, _ in equity]
    vals = [v for _, v in equity]

    fig, ax = plt.subplots(figsize=(8, 3.5), dpi=130)
    fig.patch.set_facecolor('#12151e')
    ax.set_facecolor('#12151e')
    ax.plot(dates, vals, color='#fbbf24', linewidth=1.3,
            label='MetaTrader 5 (Tick)', alpha=0.9)
    ax.set_title(name, color='#e8eaf0', fontsize=10, fontweight='bold', pad=10)
    ax.legend(loc='upper left', fontsize=7.5, framealpha=0.3,
              facecolor='#1e2333', edgecolor='#2a2f42', labelcolor='#8b90a5')
    ax.tick_params(colors='#5c6178', labelsize=7)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#2a2f42')
    ax.spines['left'].set_color('#2a2f42')
    ax.grid(True, alpha=0.15, color='#353b52', linewidth=0.5)
    ax.set_ylabel('Cumulative P&L ($)', color='#8b90a5', fontsize=7.5)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    fig.autofmt_xdate(rotation=0)
    ax.axhline(y=0, color='#5c6178', linewidth=0.5, linestyle='--', alpha=0.5)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', facecolor='#12151e', edgecolor='none',
                bbox_inches='tight', pad_inches=0.1)
    plt.close(fig)
    import base64
    b64 = base64.b64encode(buf.getvalue()).decode('ascii')
    return f'data:image/png;base64,{b64}'


def build_tick_ranking(reports_folder: str, mc_results: dict, m1_names: list) -> list:
    """Parse the tick MT5 .htm reports and build a full ranking-grid array.

    Mirrors the row shape produced by Step 7's ranking_data so the dashboard's
    Tick grid renders identically to the M1 grid. The MC95 Ret/DD column carries
    the *tick* Monte Carlo value; Score is recomputed on the tick metrics; MC Rank
    has no tick equivalent (left as None -> shows '—').
    """
    if not _STEP7_AVAILABLE:
        print_yellow(
            f"  Step 7 module unavailable ({_STEP7_IMPORT_ERROR}); "
            "skipping full tick grid (MC95 values still updated)")
        return []

    if not os.path.isdir(reports_folder):
        print_red(f"  Tick reports folder not found: {reports_folder}")
        return []

    report_files = step7.find_mt5_reports(reports_folder)
    if not report_files:
        print_yellow(f"  No tick .htm reports found in: {reports_folder}")
        return []

    print_gray(f"  Parsing {len(report_files)} tick MT5 report(s)...")

    mt5_metrics = {}
    report_by_name = {}
    for path in report_files:
        metrics = step7.parse_mt5_report(path)
        if not metrics:
            print_yellow(f"    Skipped (unparseable): {os.path.basename(path)}")
            continue
        name = canonical_name(strategy_name_from_report(path), m1_names)
        mt5_metrics[name] = metrics
        report_by_name[name] = path

    if not mt5_metrics:
        print_yellow("  No parseable tick reports — tick grid will be empty")
        return []

    ranked_df = step7.compute_mt5_rankings(mt5_metrics)
    if ranked_df is None or ranked_df.empty:
        print_yellow("  Could not compute tick rankings")
        return []

    tick_ranking = []
    for _, row in ranked_df.iterrows():
        name = row.get('Strategy', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        chart = render_tick_chart(report_by_name.get(name, ''), name)
        tick_ranking.append({
            'rank': int(row.get('Rank', 0)) if pd_notna(row.get('Rank')) else 0,
            'name': name,
            'score': _rnd(row.get('Composite Score'), 3),
            'symbol': row.get('Symbol', '') or '',
            'net_profit': _rnd(row.get('Total Net Profit')),
            'ret_dd': _rnd(row.get('Ret/DD Ratio')),
            # MC95 Ret/DD column on the Tick grid == tick Monte Carlo value
            'mc95_ret_dd': _rnd(mc95_tick) if mc95_tick is not None else 0,
            'mc95_ret_dd_tick': _rnd(mc95_tick) if mc95_tick is not None else None,
            'mc_rank': None,  # no tick MC_Ranked.csv -> renders as '—'
            'wl_ratio': _rnd(row.get('Win/Loss Ratio')),
            'pf': _rnd(row.get('Profit Factor')),
            'sharpe': _rnd(row.get('Sharpe Ratio')),
            'recovery': _rnd(row.get('Recovery Factor')),
            'lr_corr': _rnd(row.get('LR Correlation')),
            'win_rate': _rnd(row.get('Win Rate %'), 1),
            'trades': int(row.get('Total Trades')) if pd_notna(row.get('Total Trades')) else 0,
            'exp_payoff': _rnd(row.get('Expected Payoff')),
            'dd_dollar': _rnd(row.get('Balance DD Max $', row.get('Equity DD Max $', 0))),
            'dd_pct': _rnd(row.get('Balance DD Rel %', row.get('Equity DD Max %', 0)), 1),
            'lr_stderr': _rnd(row.get('LR Standard Error')),
            'chart': chart,
        })
        mc_txt = f"{mc95_tick:.2f}" if mc95_tick is not None else "—"
        print_gray(f"    Tick row: {name} (score {tick_ranking[-1]['score']}, MC95 {mc_txt})")

    return tick_ranking


def update_strategies_json(json_path: str, mc_results: dict, tick_ranking: list) -> int:
    """
    Update strategies_data.json with MC95 Tick values and the tick_ranking grid.
    
    Returns number of strategies updated.
    """
    if not os.path.exists(json_path):
        print_red(f"ERROR: strategies_data.json not found: {json_path}")
        return 0
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    updated_count = 0
    
    # Update ranking entries (top 10 only)
    top10_names = set()
    for strat in data.get('ranking', []):
        rank = strat.get('rank', 999)
        if rank > 10:
            continue
        name = strat.get('name', '')
        top10_names.add(name)
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
            updated_count += 1
            print_gray(f"  Updated {name}: MC95 Tick = {mc95_tick:.2f}")

    # Update portfolio entries (only those in top 10)
    for strat in data.get('portfolio', []):
        name = strat.get('name', '')
        if name not in top10_names:
            continue
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)

    # Attach the full tick grid (rendered as the toggleable "Tick" grid)
    data['tick_ranking'] = tick_ranking

    # Save updated JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)

    return updated_count


def update_dashboard_html(html_path: str, mc_results: dict, tick_ranking: list) -> int:
    """
    Update Dashboard index.html with MC95 Tick values and the tick_ranking grid.

    The dashboard generated by Step7 embeds all data as a single JS object:
        const DATA = { ranking: [...], portfolio: [...], tick_ranking: [...], ... };
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

    # Update ranking entries (top 10 only)
    top10_names = set()
    for strat in js_data.get('ranking', []):
        rank = strat.get('rank', 999)
        if rank > 10:
            continue
        name = strat.get('name', '')
        top10_names.add(name)
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            old_val = strat.get('mc95_ret_dd_tick')
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
            updated_count += 1
            print_gray(f"  Ranking updated: {name} -> {mc95_tick:.2f} (was: {old_val})")

    # Update portfolio entries (only those in top 10)
    for strat in js_data.get('portfolio', []):
        name = strat.get('name', '')
        if name not in top10_names:
            continue
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)

    # Attach the full tick grid (rendered as the toggleable "Tick" grid)
    js_data['tick_ranking'] = tick_ranking

    if updated_count > 0 or tick_ranking:
        # Replace the DATA object in the HTML
        new_js_data = json.dumps(js_data)
        html_content = (
            html_content[:js_match.start()]
            + f'{js_match.group(1)}{new_js_data}{js_match.group(3)}'
            + html_content[js_match.end():]
        )
        print_green(
            f"  Updated {updated_count} MC95 value(s) and "
            f"{len(tick_ranking)} tick grid row(s) in embedded DATA object")

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
    parser.add_argument(
        '--tick-reports-folder',
        default=None,
        help='Folder containing the tick MT5 .htm reports '
             '(default: same folder as --tick-mc-results)'
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

    # Build the full tick ranking grid from the tick MT5 .htm reports
    tick_reports_folder = (
        args.tick_reports_folder
        or os.path.dirname(os.path.abspath(args.tick_mc_results))
    )
    print_gray(f"Building tick grid from reports in: {tick_reports_folder}")
    m1_names = load_m1_ranking_names(json_path)
    tick_ranking = build_tick_ranking(tick_reports_folder, mc_results, m1_names)
    print_green(f"  Built {len(tick_ranking)} tick grid row(s)")
    print()

    # Update strategies_data.json
    print_gray("Updating strategies_data.json...")
    json_updated = update_strategies_json(json_path, mc_results, tick_ranking)
    print_green(f"  Updated {json_updated} strategies in JSON")
    print()

    # Update Dashboard HTML
    print_gray("Updating Dashboard HTML...")
    html_updated = update_dashboard_html(html_path, mc_results, tick_ranking)
    print_green(f"  Updated {html_updated} strategies in HTML")
    print()
    
    print_cyan("=" * 60)
    print_green("Dashboard update complete!")
    print_cyan("=" * 60)


if __name__ == '__main__':
    main()
