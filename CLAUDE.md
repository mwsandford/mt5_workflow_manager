# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MetaTrader 5 Workflow Manager — a PySide6 (Qt6) desktop GUI that orchestrates a multi-step automated trading workflow across three external tools: **QuantDataManager** (data acquisition), **MetaTrader 5** (backtesting), and **Quant Analyzer** (Monte Carlo analysis & strategy ranking).

## Build & Run

```bash
# Run from source
python mt5_workflow_manager.py

# Build standalone executable (output: dist/MT5 Workflow Manager.exe)
pyinstaller --onefile --windowed --name "MT5 Workflow Manager" mt5_workflow_manager.py

# Or via batch launcher
Run_MT5_WFM.cmd
```

**Deployment requirement**: All `Step*.py` scripts must be in the same directory as the compiled `.exe`.

## Dependencies

PySide6, pyautogui, pywinauto, psutil, opencv-python, pandas, openpyxl, matplotlib. Install via `pip install`.

## Architecture

**Execution model**: The main GUI (`mt5_workflow_manager.py`, ~1970 lines) launches each workflow step as a **subprocess** running individual `Step*.py` scripts. Output is captured in real-time via threads and rendered as color-coded HTML in the log panel.

**Key classes in `mt5_workflow_manager.py`**:
- `WorkflowWindow(QMainWindow)` — main application window, process management, and step orchestration
- `WorkflowStep` (dataclass) — step definition with `id`, `title`, `description`, `script_name`, `build_args` callback, `depends_on` chain, and optional confirm-dialog mode (`is_confirmation` + `confirmation_message`) for user-gated pseudo-steps that don't launch a subprocess
- `Settings` (dataclass) — all user configuration, persisted as JSON
- `StepCard` — individual step UI widget with status indicator and run/confirm button
- `WorkflowSection` — grouped section of related steps. Four sections are instantiated with these display titles: "Update MetaTrader Data", "Back Test MetaTrader Expert Advisors", "Monte Carlo Analysis - M1", and "Monte Carlo Analysis - Tick"
- `Theme` — dark theme color constants (GitHub dark mode inspired)
- `StepStatus` (enum) — IDLE, RUNNING, COMPLETE, FAILED

**Step factory functions** (`build_data_update_steps()`, `build_backtest_steps()`, etc.) construct the dependency graph. Each step's `build_args` callback receives the current `Settings` and returns the CLI argument list.

**Sequential execution mode**: When enabled, completing one step auto-triggers the next in the dependency chain. Stops on failure.

## Workflow Steps

| Step | Script | Section | Purpose |
|------|--------|---------|---------|
| 1 | Step1_Refresh_QDM_Data.py | Data Update | Refresh symbol data via `qdmcli.exe` |
| 2 | Step2_Export_Data_From_QDM.py | Data Update | Export tick data to CSV |
| 3 | Step3_Start_MT5_Import.py | Data Update | Import CSVs into MT5 custom symbols |
| 4 | Step4_Compile_MT5_EAs.py | Backtest | Batch compile `.mq5` → `.ex5` via `metaeditor64.exe` |
| 5 | Step5_MT5_Backtest.py | Backtest | Run backtests via MT5 terminal CLI with INI files |
| 6 | Step6_Run_QA_Script.py | Monte Carlo | Automate Quant Analyzer scripting (pywinauto + image recognition) |
| 7 | Step7_Strategy_Ranking.py | Monte Carlo | Composite-score ranking, Excel + HTML dashboard generation |
| 8 | Step8_Update_Dashboard_Tick.py | Monte Carlo (Tick) | Merge tick MC results into run-level Dashboard + run the tick correlation check + open the finished Dashboard in the browser |

Note: Step 8 is the last step — Steps 9 and 10 do not exist. Steps 5/6 have "tick" variants (5b/6b) configured in the GUI.

## Dashboard Column Provenance (MT5 Backtest Rankings table)

The run-level Dashboard runs **two separate MT5 backtests** against the same EAs using different tick models (`Step5_MT5_Backtest.py --model`):
- **M1 backtest** — `--model 1` (**1 minute OHLC**), output to `BacktestOutputFolder`. This is the primary/faster pass over all strategies.
- **Tick backtest** — `--model 4` (**Every tick based on real ticks**), output to `BacktestOutputFolder/ticks`, run only for the top-N strategies by M1 composite rank (`--max-strategies`, driven by `Settings.TickBacktestCount`, default **20**).

Every column in the "MT5 Backtest Rankings" table is sourced from the **1-minute-OHLC pipeline**, *except* `MC95 Tick`:
- **SCORE, NET PROFIT, RET/DD, W/L, PF, SHARPE, RECOVERY, LR CORR, WIN%, TRADES, DD($), DD%** — parsed directly from the M1 MT5 Strategy Tester `.htm` reports (`parse_mt5_report`) plus trade CSVs, in `Step7_Strategy_Ranking.py`.
- **MC95 RET/DD** — derived from the Monte Carlo pass (`BatchMC_Results.csv`) run *on the M1 backtest* (Step 6).
- **MC95 TICK** — the only column from the tick model. Step 7 emits it as a `None` placeholder (`ranking[].mc95_ret_dd_tick`); `Step8_Update_Dashboard_Tick.py` fills it from `ticks/BatchMC_Results.csv` (Monte Carlo on the tick backtest).

### Deliberate mismatches with StrategyQuant

The dashboard mirrors **MetaTrader 5's backtest report**, not StrategyQuant X. Two columns will therefore never agree with SQX, by design — do not "fix" them:

- **SHARPE** — read verbatim from the MT5 `.htm`. MT5's convention is roughly **4x** a conventional annualised Sharpe and **~40x** the per-trade Sharpe SQX reports (measured across 97 reports: MT5 median 3.66, annualised median 0.87, per-trade median 0.093). Note this is not a monotone transform of a correct Sharpe — rank correlation against annualised Sharpe is only ~0.25, and it tracks *per-trade* Sharpe (~0.87), so it systematically favours low-frequency strategies. Relevant because Sharpe carries 11% of the composite score.
- **RET/DD** — derived, since MT5 publishes no such field: `Total Net Profit / Balance Drawdown Maximal`, both straight from the MT5 report. Dividing by *Equity* Drawdown Maximal instead would reproduce MT5's own **Recovery Factor** exactly (verified identical across all 97 reports), collapsing two columns into one — hence Balance DD is used to keep the balance-DD and equity-DD views distinct.

## Strategy Correlation Check (Step 8, tick data only)

The correlation / cluster / KEEP-ABANDON check runs on the **tick** backtest results, not M1. Step 7 no longer computes correlation at all — it emits the correlation-driven dashboard sections empty (`portfolio`, `corr_names`, `correlations`, `clusters`, `best_pairs`, `cards.tick_*`) and `Step8_Update_Dashboard_Tick.py` fills them in.

**How it works** (`compute_tick_correlation` in Step 8):
1. `parse_tick_trades` reconstructs a per-trade P&L series from each tick `.htm` report's Deals table — successive **Balance** deltas on `out` deals, so commission and swap are included. The entry side comes from the preceding `in` deal (used for the Long Only / Short Only / Both label).
2. Step 7's primitives do the maths — `build_pnl_series` → `compute_pairwise_correlation` (daily/weekly/monthly) → `identify_clusters` at weekly |r| ≥ `CORR_CLUSTER_THRESHOLD` (0.5).
3. Highest **tick** composite score in each cluster is KEEP; the rest are ABANDON.

**Where it surfaces**: ✗ ABANDON badges + dimmed rows on the **Tick** grid only (the M1 grid never shows them), plus the Portfolio, Correlation and Clusters tabs and the Tick Clusters / Tick Keep / Tick Abandon summary cards. Before Step 8 runs, those tabs show an empty state.

**Step 7's correlation functions are still live** — `build_pnl_series`, `compute_pairwise_correlation` and `identify_clusters` remain in `Step7_Strategy_Ranking.py` because Step 8 imports them. Trade-overlap analysis was removed entirely (it needed trade open times, which the Deals table doesn't expose directly).

## Configuration

User settings stored at `%USERPROFILE%\.mt5_workflow\`:
- `mt5_workflow_config.json` — folder paths, date ranges, thresholds, automation flags
- `ui_state.json` — window geometry and panel state

## QA Automation (Step 6)

Uses two strategies for automating Quant Analyzer's UI:
1. **Image recognition** (preferred): OpenCV template matching against PNGs in `qa_templates/`
2. **Coordinate-based**: Fallback using pixel offsets via pyautogui

## Conventions

- Python 3.10+ (uses `X | Y` union type syntax)
- Dark theme with consistent color palette defined in `Theme` class
- Subprocess output captured with `PYTHONUNBUFFERED=1` and `CREATE_NO_WINDOW` flag on Windows
- Each `Step*.py` is independently runnable from CLI with `argparse` arguments
- No test suite exists; manual testing against live MT5/QDM/QA installations