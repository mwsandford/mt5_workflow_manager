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

PySide6, pyautogui, pywinauto, opencv-python, pandas, openpyxl, matplotlib. Install via `pip install`.

## Architecture

**Execution model**: The main GUI (`mt5_workflow_manager.py`, ~1900 lines) launches each workflow step as a **subprocess** running individual `Step*.py` scripts. Output is captured in real-time via threads and rendered as color-coded HTML in the log panel.

**Key classes in `mt5_workflow_manager.py`**:
- `WorkflowWindow(QMainWindow)` — main application window, process management, and step orchestration
- `WorkflowStep` (dataclass) — step definition with name, script path, `build_args` callback, and `depends_on` chain
- `Settings` (dataclass) — all user configuration, persisted as JSON
- `StepCard` — individual step UI widget with status indicator and run button
- `WorkflowSection` — grouped section of related steps (Data Update, Backtest, Monte Carlo)
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
| 9 | Step9_Update_Dashboard_Tick.py | Monte Carlo (Tick) | Merge tick MC results into dashboard |

Note: Step 8 does not exist. Steps 5/6 have "tick" variants (5b/6b) configured in the GUI.

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