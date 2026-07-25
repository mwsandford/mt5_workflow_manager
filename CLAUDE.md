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
- `WorkflowSection` — grouped section of related steps. Five sections are instantiated with these display titles: "Update MetaTrader Data", "Back Test MetaTrader Expert Advisors", "Monte Carlo Analysis - M1", "Monte Carlo Analysis - Tick", and "Tick Survival Analysis"
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
| 7 | Step7_Strategy_Ranking.py | Monte Carlo | Correlation analysis, Excel + HTML dashboard generation |
| 8 | Step8_Update_Dashboard_Tick.py | Monte Carlo (Tick) | Merge tick MC results into run-level Dashboard |
| 10 | Step10_Update_Tick_Survival.py | Tick Survival Analysis | Update persistent SQLite store + regenerate cross-run survival dashboard |

Note: Step 9 does not exist (number reserved). Steps 5/6 have "tick" variants (5b/6b) configured in the GUI.

## Dashboard Column Provenance (MT5 Backtest Rankings table)

The run-level Dashboard runs **two separate MT5 backtests** against the same EAs using different tick models (`Step5_MT5_Backtest.py --model`):
- **M1 backtest** — `--model 1` (**1 minute OHLC**), output to `BacktestOutputFolder`. This is the primary/faster pass over all strategies.
- **Tick backtest** — `--model 4` (**Every tick based on real ticks**), output to `BacktestOutputFolder/ticks`, run only for the top-10 KEEP strategies from the M1 analysis (`--max-strategies 10`).

Every column in the "MT5 Backtest Rankings" table is sourced from the **1-minute-OHLC pipeline**, *except* `MC95 Tick`:
- **SCORE, NET PROFIT, RET/DD, W/L, PF, SHARPE, RECOVERY, LR CORR, WIN%, TRADES, DD($), DD%** — parsed directly from the M1 MT5 Strategy Tester `.htm` reports (`parse_mt5_report`) plus trade CSVs, in `Step7_Strategy_Ranking.py`.
- **MC95 RET/DD** and **MC RANK** — derived from the Monte Carlo pass (`BatchMC_Results.csv` / `MC_Ranked.csv`) run *on the M1 backtest* (Step 6).
- **MC95 TICK** — the only column from the tick model. Step 7 emits it as a `None` placeholder (`ranking[].mc95_ret_dd_tick`); `Step8_Update_Dashboard_Tick.py` fills it from `ticks/BatchMC_Results.csv` (Monte Carlo on the tick backtest).

## Tick Survival Analysis (Step 10)

A persistent, cross-run analytics store that accumulates strategy characteristics joined to backtest results and renders an HTML survival-analysis dashboard.

**Why separate from the run-level Dashboard**: `BacktestOutputFolder` is wiped every run, so anything needing to build up history across runs must live outside it. The tick_survival store lives inside the Workflow Manager folder.

**Layout**:
- `tick_survival/db.py` — stdlib `sqlite3` schema + upsert helpers. Two tables: `Strategies` (natural key `StrategyName`), `BacktestResults` (`UNIQUE(StrategyID, BatchID)`)
- `tick_survival/parser.py` — reads `Dashboard/strategies_data.json` + embedded `const DATA` from `Dashboard/index.html`; falls back to re-running Step 7's `parse_strategy_pseudo_code` for missing entries
- `tick_survival/dashboard.py` — self-contained dark-theme HTML with filter bar, 12 survival-by-characteristic tables, Bayesian-shrunk predictive ranking, expandable per-strategy history
- `tick_survival/tick_survival.db` — persistent store (never checked into the run output folder)
- `tick_survival/reports/tick_survival_dashboard.html` — generated view, rebuilt from the DB every run
- `Step10_Update_Tick_Survival.py` — CLI orchestrator: parse → upsert → regenerate HTML → `webbrowser.open()`

**Scope**: only top-10 tick-tested strategies per batch are stored (records where `ranking.mc95_ret_dd_tick is not None`). Re-running Step 10 against the same folder upserts; new runs accumulate as history rows. Pass threshold defaults to MC95 Tick ≥ 2.0 (configurable via `Settings.TickSurvivalThreshold` in the GUI and `--tick-threshold` CLI arg).

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