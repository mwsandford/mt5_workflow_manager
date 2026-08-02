#!/usr/bin/env python3
"""
Step 8: Update Dashboard with Tick Backtest + Monte Carlo Results
=================================================================

This script builds the *Tick* grid on the Performance Dashboard after the
tick-based backtests and Monte Carlo analysis have run, and runs the strategy
correlation check over those tick results.

It reads:
  - BatchMC_Results.csv from the ticks subfolder (tick MC95 values)
  - The tick MT5 .htm Strategy Tester reports in the ticks subfolder
  - strategies_data.json from the Dashboard folder (for canonical M1 names)

It updates:
  - index.html          - injects a full `tick_ranking` array into the embedded
                          DATA object (rendered as the toggleable "Tick" grid),
                          keeps ranking[].mc95_ret_dd_tick populated, and fills
                          the correlation-driven sections (portfolio / corr_names
                          / correlations / clusters / best_pairs / cards)
  - strategies_data.json - the same updates, for downstream consumers

The tick grid is built with the exact same parser + scoring as the M1 grid by
reusing Step 7's parse_mt5_report() / compute_mt5_rankings(), so metrics line up.

Correlation check
-----------------
Per-trade P&L is reconstructed from the Deals table of each tick .htm report
(successive Balance deltas on 'out' deals), then Step 7's correlation primitives
(build_pnl_series / compute_pairwise_correlation / identify_clusters) group the
strategies into clusters at weekly |r| >= 0.5. The best tick composite score in
each cluster is KEEP; the rest are ABANDON and get flagged on the Tick grid.

Usage:
    python Step8_Update_Dashboard_Tick.py <dashboard_folder> \\
        --tick-mc-results <path_to_tick_mc_results.csv> \\
        [--tick-reports-folder <folder_with_tick_htm_reports>] \\
        [--top-n 20]
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
    import pandas as pd
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    _STEP7_AVAILABLE = True
    _STEP7_IMPORT_ERROR = None
except Exception as _exc:  # pragma: no cover - defensive
    step7 = None
    pd = None
    plt = None
    mdates = None
    _STEP7_AVAILABLE = False
    _STEP7_IMPORT_ERROR = _exc

# Weekly |correlation| at or above this groups two strategies into one cluster
# (same threshold the M1 pass used before the check moved to tick).
CORR_CLUSTER_THRESHOLD = 0.5

# Default number of top-ranked M1 strategies that get tick back-tested.
DEFAULT_TOP_N = 20


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


def build_tick_ranking(reports_folder: str, mc_results: dict, m1_names: list) -> tuple:
    """Parse the tick MT5 .htm reports and build a full ranking-grid array.

    Mirrors the row shape produced by Step 7's ranking_data so the dashboard's
    Tick grid renders identically to the M1 grid. The MC95 Ret/DD column carries
    the *tick* Monte Carlo value; Score is recomputed on the tick metrics.

    Returns (tick_ranking, report_by_name) — the report map feeds the correlation
    check, which needs the Deals table from each report.
    """
    if not _STEP7_AVAILABLE:
        print_yellow(
            f"  Step 7 module unavailable ({_STEP7_IMPORT_ERROR}); "
            "skipping full tick grid (MC95 values still updated)")
        return [], {}

    if not os.path.isdir(reports_folder):
        print_red(f"  Tick reports folder not found: {reports_folder}")
        return [], {}

    report_files = step7.find_mt5_reports(reports_folder)
    if not report_files:
        print_yellow(f"  No tick .htm reports found in: {reports_folder}")
        return [], {}

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
        return [], {}

    ranked_df = step7.compute_mt5_rankings(mt5_metrics)
    if ranked_df is None or ranked_df.empty:
        print_yellow("  Could not compute tick rankings")
        return [], {}

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

    return tick_ranking, report_by_name


# =============================================================================
# Tick correlation check (clusters -> KEEP / ABANDON)
# =============================================================================
def _parse_float(cell: str):
    """Parse an MT5 report number ('1 234.56', '-1\xa0234.56') -> float or None."""
    try:
        s = str(cell).replace(' ', '').replace('\xa0', '').replace(' ', '')
        if not s:
            return None
        return float(s)
    except (TypeError, ValueError):
        return None


def parse_tick_trades(filepath: str):
    """Reconstruct a per-trade P&L series from an MT5 report's Deals table.

    Only 'out' (position-closing) deals carry a realised result. Per-trade P&L is
    the Balance delta between successive 'out' rows, so commission and swap are
    included — the same Balance column Step 7's equity curve uses.

    Returns a DataFrame with 'Close time', 'Profit/Loss' and 'Type' (the entry
    side, taken from the preceding 'in' deal), or None if the table can't be read.
    """
    content = None
    for encoding in ['utf-16-le', 'utf-16', 'utf-8', 'latin-1']:
        try:
            with open(filepath, 'r', encoding=encoding, errors='replace') as f:
                content = f.read()
            if 'Strategy Tester' in content or 'Deals' in content:
                break
        except Exception:
            continue

    if content is None:
        return None

    parser = step7._MT5DealsParser()
    try:
        parser.feed(content)
    except Exception:
        return None

    if not parser.rows:
        return None

    # Locate the Deals header row and the columns we need
    deals_start = balance_col = direction_col = type_col = None
    for i, row in enumerate(parser.rows):
        if len(row) < 12:
            continue
        bc = dc = tc = None
        for j, cell in enumerate(row):
            s = cell.strip()
            if s == 'Balance':
                bc = j
            elif s == 'Direction':
                dc = j
            elif s == 'Type':
                tc = j
        if bc is not None and dc is not None:
            deals_start, balance_col, direction_col, type_col = i, bc, dc, tc
            break

    if deals_start is None:
        return None

    trades = []
    prev_balance = None
    pending_entry_side = ''

    for row in parser.rows[deals_start + 1:]:
        # Initial balance row (shorter, type="balance")
        if len(row) < 12:
            if 6 <= len(row) <= 8:
                type_cell = row[2].strip().lower() if len(row) > 2 else ''
                if type_cell == 'balance':
                    bal = _parse_float(row[-1])
                    if bal is not None:
                        prev_balance = bal
            continue

        direction = row[direction_col].strip().lower() if direction_col < len(row) else ''
        deal_type = row[type_col].strip() if type_col is not None and type_col < len(row) else ''

        if direction == 'in':
            # Remember the entry side so the closed trade can be labelled
            pending_entry_side = deal_type.capitalize()
            continue

        if direction != 'out':
            continue

        dt = None
        for fmt in ['%Y.%m.%d %H:%M:%S', '%Y.%m.%d %H:%M', '%Y-%m-%d %H:%M:%S']:
            try:
                dt = pd.to_datetime(row[0], format=fmt)
                break
            except (ValueError, TypeError):
                continue
        if dt is None:
            continue

        balance = _parse_float(row[balance_col]) if balance_col < len(row) else None
        if balance is None:
            continue

        if prev_balance is None:
            # No opening balance row — first close establishes the baseline
            prev_balance = balance
            continue

        trades.append({
            'Close time': dt,
            'Profit/Loss': balance - prev_balance,
            'Type': pending_entry_side or 'Buy',
        })
        prev_balance = balance
        pending_entry_side = ''

    if len(trades) < 2:
        return None

    return pd.DataFrame(trades)


def _direction_from_trades(df) -> str:
    """Long Only / Short Only / Both, from the entry side of each closed trade."""
    long_types = {'Buy', 'BuyStop', 'BuyLimit'}
    short_types = {'Sell', 'SellStop', 'SellLimit'}
    n_long = df['Type'].isin(long_types).sum()
    n_short = df['Type'].isin(short_types).sum()
    if n_short == 0:
        return 'Long Only'
    if n_long == 0:
        return 'Short Only'
    return 'Both'


def compute_tick_correlation(tick_ranking: list, report_by_name: dict) -> dict:
    """Run the correlation / cluster / KEEP-ABANDON check over the tick results.

    Returns a dict of the dashboard sections to inject: portfolio, corr_names,
    correlations, clusters, best_pairs and the tick_* summary cards. Returns an
    empty dict when there isn't enough tick trade data to correlate.
    """
    empty = {}
    if not _STEP7_AVAILABLE or not tick_ranking:
        return empty

    # Names in tick rank order — identify_clusters anchors each cluster on the
    # first (i.e. best-ranked) member it sees.
    ordered = sorted(tick_ranking, key=lambda r: r.get('rank', 9999))
    row_by_name = {r['name']: r for r in ordered}

    trades = {}
    for r in ordered:
        name = r['name']
        path = report_by_name.get(name)
        if not path:
            continue
        try:
            df = parse_tick_trades(path)
        except Exception as exc:
            print_yellow(f"    Could not read tick trades for {name}: {exc}")
            df = None
        if df is not None and len(df) >= 2:
            trades[name] = df
        else:
            print_yellow(f"    No usable tick trade data for {name} — excluded from correlation")

    names = [r['name'] for r in ordered if r['name'] in trades]
    if len(names) < 2:
        print_yellow("  Need at least 2 strategies with tick trade data — skipping correlation check")
        return empty

    strategies = {n: trades[n] for n in names}

    print_gray(f"  Computing tick P&L correlations across {len(names)} strategies...")
    daily_df = step7.build_pnl_series(strategies, 'D')
    weekly_df = step7.build_pnl_series(strategies, 'W')
    monthly_df = step7.build_pnl_series(strategies, 'M')
    corr_daily = step7.compute_pairwise_correlation(daily_df)
    corr_weekly = step7.compute_pairwise_correlation(weekly_df)
    corr_monthly = step7.compute_pairwise_correlation(monthly_df, min_observations=6)

    clusters = step7.identify_clusters(names, corr_weekly, threshold=CORR_CLUSTER_THRESHOLD)

    # Best tick composite score in each cluster survives
    def _score(n):
        return row_by_name.get(n, {}).get('score') or 0

    keep = {}
    abandon = {}
    for cid, members in enumerate(clusters, 1):
        if len(members) == 1:
            keep[members[0]] = cid
            continue
        best = max(members, key=_score)
        keep[best] = cid
        for m in members:
            if m != best:
                abandon[m] = (cid, best)

    def _portfolio_row(name, decision, cid, reason):
        r = row_by_name.get(name, {})
        return {
            'name': name,
            'decision': decision,
            'cluster': cid,
            'reason': reason,
            'direction': _direction_from_trades(trades[name]),
            'total_pnl': _rnd(r.get('net_profit')),
            'avg_trade': _rnd(r.get('exp_payoff')),
            'win_rate': _rnd(r.get('win_rate'), 1),
            'mc95_ret_dd': r.get('mc95_ret_dd', 0),
            'mc95_ret_dd_tick': r.get('mc95_ret_dd_tick'),
            'chart': r.get('chart', ''),
        }

    portfolio = []
    for name in sorted(keep.keys(), key=lambda s: keep[s]):
        cid = keep[name]
        size = len(clusters[cid - 1])
        if size == 1:
            reason = 'Only in cluster'
        else:
            reason = f'Best tick score ({_score(name):.3f}) in cluster of {size}'
        portfolio.append(_portfolio_row(name, 'KEEP', cid, reason))
    for name in sorted(abandon.keys(), key=lambda s: abandon[s][0]):
        cid, replaced_by = abandon[name]
        reason = (f'Correlated with {replaced_by} '
                  f'(score {_score(replaced_by):.3f} vs {_score(name):.3f})')
        portfolio.append(_portfolio_row(name, 'ABANDON', cid, reason))

    def corr_to_list(matrix):
        return [[_rnd(matrix.loc[a, b], 3) for b in names] for a in names]

    best_pairs = []
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            best_pairs.append({
                'pair': f'{a} vs {b}',
                'weekly': _rnd(corr_weekly.loc[a, b], 3),
                'daily': _rnd(corr_daily.loc[a, b], 3),
            })
    best_pairs.sort(key=lambda p: abs(p['weekly']))

    print_green(f"  {len(clusters)} cluster(s): {len(keep)} KEEP, {len(abandon)} ABANDON")
    for name in sorted(abandon.keys(), key=lambda s: abandon[s][0]):
        cid, replaced_by = abandon[name]
        print_gray(f"    x {name} — correlated with {replaced_by} (cluster {cid})")

    return {
        'portfolio': portfolio,
        'corr_names': names,
        'correlations': {
            'daily': corr_to_list(corr_daily),
            'weekly': corr_to_list(corr_weekly),
            'monthly': corr_to_list(corr_monthly),
        },
        'clusters': [{'id': i + 1, 'members': c, 'count': len(c)}
                     for i, c in enumerate(clusters)],
        'best_pairs': best_pairs,
        'cards': {
            'tick_clusters': len(clusters),
            'tick_keep': len(keep),
            'tick_abandon': len(abandon),
        },
    }


def build_tick_overviews(report_by_name: dict) -> dict:
    """Parse the full QA-style overview out of each *tick* .htm report.

    Step 7 only ever builds overviews from the M1 reports (DATA.overviews), so
    without this the dashboard's Overview panel and the Export payload showed M1
    numbers while the Tick grid was selected. Keyed by strategy name and shaped
    by Step 7's own shaper so both models share field names.
    """
    if not _STEP7_AVAILABLE or not report_by_name:
        return {}

    overviews = {}
    for name, path in report_by_name.items():
        try:
            ov = step7.parse_mt5_full_overview(path)
        except Exception as exc:
            print_yellow(f"    Could not parse tick overview for {name}: {exc}")
            continue
        if ov:
            overviews[name] = step7.shape_overview_for_js(ov)

    if overviews:
        print_green(f"  Built {len(overviews)} tick overview(s)")
    else:
        print_yellow("  No tick overviews could be built")
    return overviews


def apply_correlation(data: dict, corr: dict) -> None:
    """Merge the correlation sections into an embedded DATA / strategies_data dict."""
    if not corr:
        return
    for key in ('portfolio', 'corr_names', 'correlations', 'clusters', 'best_pairs'):
        data[key] = corr[key]
    cards = data.get('cards')
    if isinstance(cards, dict):
        cards.update(corr['cards'])


def update_strategies_json(json_path: str, mc_results: dict, tick_ranking: list,
                           corr: dict, top_n: int, tick_overviews: dict = None) -> int:
    """
    Update strategies_data.json with MC95 Tick values, the tick_ranking grid and
    the tick correlation sections.

    Returns number of strategies updated.
    """
    if not os.path.exists(json_path):
        print_red(f"ERROR: strategies_data.json not found: {json_path}")
        return 0

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    updated_count = 0

    # Update ranking entries (tick-tested top N only)
    for strat in data.get('ranking', []):
        rank = strat.get('rank', 999)
        if rank > top_n:
            continue
        name = strat.get('name', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
            updated_count += 1
            print_gray(f"  Updated {name}: MC95 Tick = {mc95_tick:.2f}")

    # Attach the full tick grid (rendered as the toggleable "Tick" grid) and the
    # correlation sections computed from the tick trades
    data['tick_ranking'] = tick_ranking
    if tick_overviews:
        data['tick_overviews'] = tick_overviews
    apply_correlation(data, corr)

    # Save updated JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)

    return updated_count


def update_dashboard_html(html_path: str, mc_results: dict, tick_ranking: list,
                          corr: dict, top_n: int, tick_overviews: dict = None) -> int:
    """
    Update Dashboard index.html with MC95 Tick values, the tick_ranking grid and
    the tick correlation sections.

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

    # Update ranking entries (tick-tested top N only)
    for strat in js_data.get('ranking', []):
        rank = strat.get('rank', 999)
        if rank > top_n:
            continue
        name = strat.get('name', '')
        mc95_tick = find_mc95_for_strategy(name, mc_results)
        if mc95_tick is not None:
            old_val = strat.get('mc95_ret_dd_tick')
            strat['mc95_ret_dd_tick'] = round(mc95_tick, 2)
            updated_count += 1
            print_gray(f"  Ranking updated: {name} -> {mc95_tick:.2f} (was: {old_val})")

    # Attach the full tick grid (rendered as the toggleable "Tick" grid) and the
    # correlation sections computed from the tick trades
    js_data['tick_ranking'] = tick_ranking
    if tick_overviews:
        js_data['tick_overviews'] = tick_overviews
    apply_correlation(js_data, corr)

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
    parser.add_argument(
        '--top-n',
        type=int,
        default=DEFAULT_TOP_N,
        help=f'Number of top-ranked M1 strategies that were tick back-tested '
             f'(default: {DEFAULT_TOP_N})'
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
    tick_ranking, report_by_name = build_tick_ranking(tick_reports_folder, mc_results, m1_names)
    print_green(f"  Built {len(tick_ranking)} tick grid row(s)")

    # Full overviews from the same tick reports — these back the Overview panel
    # and the Export payload whenever the Tick grid is the active view.
    tick_overviews = build_tick_overviews(report_by_name)
    print()

    # Correlation check over the tick results
    print_gray("Running correlation check on the tick backtest results...")
    try:
        corr = compute_tick_correlation(tick_ranking, report_by_name)
    except Exception as exc:
        import traceback
        print_yellow(f"  WARNING: tick correlation check failed: {exc}")
        traceback.print_exc()
        corr = {}
    if not corr:
        print_yellow("  No correlation results — Portfolio/Correlation/Clusters tabs left empty")
    print()

    # Update strategies_data.json
    print_gray("Updating strategies_data.json...")
    json_updated = update_strategies_json(json_path, mc_results, tick_ranking, corr,
                                          args.top_n, tick_overviews)
    print_green(f"  Updated {json_updated} strategies in JSON")
    print()

    # Update Dashboard HTML
    print_gray("Updating Dashboard HTML...")
    html_updated = update_dashboard_html(html_path, mc_results, tick_ranking, corr,
                                         args.top_n, tick_overviews)
    print_green(f"  Updated {html_updated} strategies in HTML")
    print()
    
    print_cyan("=" * 60)
    print_green("Dashboard update complete!")
    print_cyan("=" * 60)


if __name__ == '__main__':
    main()
