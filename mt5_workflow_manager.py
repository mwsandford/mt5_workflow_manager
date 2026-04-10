"""
MetaTrader 5 Workflow Manager
A modern GUI for managing MT5 data updates and EA backtesting workflows.

Usage:
    python mt5_workflow_manager.py

    Or build as standalone exe:
    pyinstaller --onefile --windowed --name "MT5 Workflow Manager" mt5_workflow_manager.py

Requirements:
    pip install PySide6

Note: When running as a PyInstaller exe, place the exe in the same folder
as the Step*.py workflow scripts. Python must be installed and in PATH.
"""

import sys
import json
import os
import re
import subprocess
import threading
from pathlib import Path
from enum import Enum, auto
from dataclasses import dataclass
from typing import Optional, Callable

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QScrollArea, QFrame, QLineEdit,
    QSplitter, QGroupBox, QGridLayout, QFileDialog, QSizePolicy,
    QMessageBox, QDateEdit, QCheckBox
)
from PySide6.QtCore import Qt, Signal, QDate
from PySide6.QtGui import QFont, QTextCursor


# ─────────────────────────────────────────────
# Theme / Colour Constants (matching HTML mockup exactly)
# ─────────────────────────────────────────────
class Theme:
    # Backgrounds
    BG_DARKEST   = "#0d1117"   # Main window / deepest layer
    BG_DARK      = "#161b22"   # Panels, cards
    BG_MID       = "#1c2333"   # Elevated surfaces
    BG_LIGHT     = "#242d3d"   # Hover states
    BG_INPUT     = "#1a2030"   # Input fields

    # Borders
    BORDER       = "#2a3444"
    BORDER_LIGHT = "#3a4858"

    # Text
    TEXT_PRIMARY   = "#e6edf3"
    TEXT_SECONDARY = "#8b949e"
    TEXT_MUTED     = "#6e7681"

    # Accent – blue
    ACCENT         = "#58a6ff"
    ACCENT_HOVER   = "#79bfff"
    ACCENT_DIM     = "#1a3a5c"

    # Status colours
    STATUS_PENDING   = "#6e7681"
    STATUS_RUNNING   = "#58a6ff"
    STATUS_COMPLETE  = "#3fb950"
    STATUS_FAILED    = "#f85149"

    # Section colours (matching HTML mockup)
    SECTION_DATA           = "#8b5cf6"  # Purple for data update section
    SECTION_DATA_BG        = "rgba(139, 92, 246, 0.08)"  # Purple at 8% opacity
    SECTION_BACKTEST       = "#f59e0b"  # Amber for backtest section
    SECTION_BACKTEST_BG    = "rgba(245, 158, 11, 0.08)"  # Amber at 8% opacity
    SECTION_MONTECARLO     = "#10b981"  # Teal/emerald for Monte Carlo section
    SECTION_MONTECARLO_BG  = "rgba(16, 185, 129, 0.08)"  # Teal at 8% opacity

    # Button colours
    BTN_RUN          = "#238636"
    BTN_RUN_HOVER    = "#2ea043"


# ─────────────────────────────────────────────
# Step Status
# ─────────────────────────────────────────────
class StepStatus(Enum):
    IDLE      = auto()
    RUNNING   = auto()
    COMPLETE  = auto()
    FAILED    = auto()


# ─────────────────────────────────────────────
# Step Definition
# ─────────────────────────────────────────────
@dataclass
class WorkflowStep:
    id: str
    title: str
    description: str
    script_name: str = ""
    build_args: Callable[['Settings'], list[str]] = None
    is_confirmation: bool = False  # If True, shows a confirm dialog instead of running a script
    confirmation_message: str = ""  # Message to show in confirmation dialog
    depends_on: str = ""  # Step ID that must be completed before this step can run


# ─────────────────────────────────────────────
# Settings Data Class
# ─────────────────────────────────────────────
@dataclass
class Settings:
    # Common
    MT5Folder: str = r"C:\Program Files\MetaTrader 5"
    MT5MQL5Folder: str = r"C:\Users\<UserID>\AppData\Roaming\MetaQuotes\Terminal\<ID>\MQL5"

    # QuantDataManager
    QDMFolder: str = r"C:\QuantDataManager125"
    ExportDataFrom: str = "2026.01.01"
    ExportDataTo: str = "2026.01.31"

    # Back Test
    BacktestOutputFolder: str = r"C:\Trading\Analysis_Ouput"
    MT5BackTestFrom: str = "2010.01.01"
    MT5BackTestTo: str = "2025.12.31"

    # Monte Carlo
    MC95Threshold: str = "2.5"
    QAPath: str = r"C:\QuantAnalyzer4\QuantAnalyzer4.exe"
    QAUseImages: bool = True  # Use image recognition for QA automation (more reliable)

    # Workflow Options
    SequentialExecution: bool = False  # Run Back Test and Monte Carlo steps sequentially


# ─────────────────────────────────────────────
# Build workflow steps
# ─────────────────────────────────────────────
def build_data_update_steps() -> list[WorkflowStep]:
    return [
        WorkflowStep(
            id="refresh_qdm",
            title="Step 1 — Refresh Data in QuantDataManager",
            description="Update the data for all symbols in QuantDataManager",
            script_name="Step1_Refresh_QDM_Data.py",
            build_args=lambda s: [
                os.path.join(s.QDMFolder, "qdmcli.exe")
            ],
        ),
        WorkflowStep(
            id="export_qdm",
            title="Step 2 — Export Tick Data from QuantDataManager",
            description="Export all tick data for all symbols in QuantDataManager",
            script_name="Step2_Export_Data_From_QDM.py",
            build_args=lambda s: [
                "--date-from", s.ExportDataFrom,
                "--date-to", s.ExportDataTo,
                "--qdm-path", s.QDMFolder,
                "--export-path", os.path.join(s.MT5MQL5Folder, "Files"),
            ],
        ),
        WorkflowStep(
            id="import_mt5",
            title="Step 3 — Import Tick Data into MetaTrader",
            description="Opens MetaTrader 5 and import all tick data",
            script_name="Step3_Start_MT5_Import.py",
            build_args=lambda s: [
                "--mt5-path", os.path.join(s.MT5Folder, "terminal64.exe"),
                "--wait",
                "--close-mt5",
            ],
        ),
    ]


def build_backtest_steps() -> list[WorkflowStep]:
    return [
        WorkflowStep(
            id="compile_eas",
            title="Step 1 — Compile Expert Advisors",
            description="Compile all Expert Advisors and deploy to PineappleStrats folder",
            script_name="Step4_Compile_MT5_EAs.py",
            build_args=lambda s: [
                "-s", s.BacktestOutputFolder,
                "-m", s.MT5Folder,
                "-i", s.MT5MQL5Folder,
                "-e", os.path.join(s.MT5MQL5Folder, "Experts", "Advisors", "PineappleStrats"),
                "-t", os.path.join(s.MT5MQL5Folder, "Profiles", "Tester"),
            ],
        ),
        WorkflowStep(
            id="run_backtests",
            title="Step 2 — Back Test Expert Advisors",
            description="Back test each Expert Advisor and save HTML report output",
            script_name="Step5_MT5_Backtest.py",
            build_args=lambda s: [
                "--mt5-terminal-path", os.path.join(s.MT5Folder, "terminal64.exe"),
                "--report-dest-folder", s.BacktestOutputFolder,
                "--model", "1",
                "--from-date", s.MT5BackTestFrom,
                "--to-date", s.MT5BackTestTo,
                "--timeout", "900",
            ],
        ),
    ]


def build_montecarlo_steps() -> list[WorkflowStep]:
    """Build Monte Carlo Analysis - M1 steps."""
    def build_mc_args(s):
        args = [
            "BatchMC_Minute_Data.java",  # Hardcoded - M1 backtest analysis
            "--qa-path", s.QAPath,
            "--output-folder", s.BacktestOutputFolder,
        ]
        if s.QAUseImages:
            args.append("--use-images")
        return args

    return [
        WorkflowStep(
            id="run_mc_analysis",
            title="Step 1 — Run Monte Carlo Analysis",
            description="Run Batch Monte Carlo Analysis script in QuantAnalyzer",
            script_name="Step6_Run_QA_Script.py",
            build_args=build_mc_args,
        ),
        WorkflowStep(
            id="rank_strategies",
            title="Step 2 — Rank by Correlation and Performance",
            description="Discard correlated strategies and rank remaining strategies",
            script_name="Step7_Strategy_Ranking.py",
            build_args=lambda s: [
                s.BacktestOutputFolder,
                "--mc-results", os.path.join(s.BacktestOutputFolder, "BatchMC_Results.csv"),
                "--mc95-threshold", s.MC95Threshold,
                "--mt5-reports", s.BacktestOutputFolder,
            ],
            depends_on="run_mc_analysis",
        ),
    ]


def build_tick_montecarlo_steps() -> list[WorkflowStep]:
    """Build Monte Carlo Analysis - Tick steps."""
    def build_tick_backtest_args(s):
        # Tick backtests use model 4 (every tick based on real ticks)
        # Output goes to ticks subfolder
        # Only backtest top KEEP strategies from M1 analysis
        return [
            "--mt5-terminal-path", os.path.join(s.MT5Folder, "terminal64.exe"),
            "--report-dest-folder", os.path.join(s.BacktestOutputFolder, "ticks"),
            "--model", "4",  # Every tick based on real ticks
            "--from-date", s.MT5BackTestFrom,
            "--to-date", s.MT5BackTestTo,
            "--strategies-json", os.path.join(s.BacktestOutputFolder, "Dashboard", "strategies_data.json"),
            "--max-strategies", "10",
            "--timeout", "1800",
        ]
    
    def build_tick_mc_args(s):
        args = [
            "BatchMC_Tick_Data.java",  # Hardcoded - Tick backtest analysis
            "--qa-path", s.QAPath,
            "--output-folder", os.path.join(s.BacktestOutputFolder, "ticks"),
        ]
        if s.QAUseImages:
            args.append("--use-images")
        return args
    
    return [
        WorkflowStep(
            id="tick_backtest",
            title="Step 1 — Back Test Expert Advisors (Tick)",
            description="Run tick-based backtests for top-ranked strategies from M1 analysis",
            script_name="Step5_MT5_Backtest.py",
            build_args=build_tick_backtest_args,
        ),
        WorkflowStep(
            id="tick_mc_analysis",
            title="Step 2 — Run Monte Carlo Analysis",
            description="Run Batch Monte Carlo Analysis on tick backtest results",
            script_name="Step6_Run_QA_Script.py",
            build_args=build_tick_mc_args,
            depends_on="tick_backtest",
        ),
        WorkflowStep(
            id="tick_update_dashboard",
            title="Step 3 — Update Performance Dashboard",
            description="Update Dashboard with MC95 Tick values",
            script_name="Step8_Update_Dashboard_Tick.py",
            build_args=lambda s: [
                os.path.join(s.BacktestOutputFolder, "Dashboard"),
                "--tick-mc-results", os.path.join(s.BacktestOutputFolder, "ticks", "BatchMC_Results.csv"),
            ],
            depends_on="tick_mc_analysis",
        ),
    ]


# ─────────────────────────────────────────────
# Global Stylesheet (matching HTML mockup exactly)
# ─────────────────────────────────────────────
STYLESHEET = f"""
/* ── Base ── */
QMainWindow, QWidget {{
    background-color: {Theme.BG_DARKEST};
    color: {Theme.TEXT_PRIMARY};
    font-family: "Segoe UI", "SF Pro Text", "Helvetica Neue", sans-serif;
    font-size: 13px;
}}

/* ── Scroll areas ── */
QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollBar:vertical {{
    background: {Theme.BG_DARK};
    width: 8px;
    margin: 0;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {Theme.BORDER_LIGHT};
    min-height: 30px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical:hover {{
    background: {Theme.TEXT_MUTED};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    height: 0; background: none;
}}

/* ── Splitter ── */
QSplitter::handle {{
    background: {Theme.BORDER};
    width: 1px;
}}

/* ── Group boxes ── */
QGroupBox {{
    background-color: {Theme.BG_DARK};
    border: 1px solid {Theme.BORDER};
    border-radius: 8px;
    margin-top: 8px;
    padding: 16px 12px 12px 12px;
    font-weight: 700;
    font-size: 10px;
    color: {Theme.TEXT_SECONDARY};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 16px;
    padding: 0 6px;
    color: {Theme.TEXT_SECONDARY};
}}

/* ── Line edits ── */
QLineEdit {{
    background-color: {Theme.BG_INPUT};
    border: 1px solid {Theme.BORDER};
    border-radius: 6px;
    padding: 7px 10px;
    color: {Theme.TEXT_PRIMARY};
    font-size: 12px;
    selection-background-color: {Theme.ACCENT_DIM};
}}
QLineEdit:focus {{
    border-color: {Theme.ACCENT};
}}

/* ── Date edits with calendar popup ── */
QDateEdit {{
    background-color: {Theme.BG_INPUT};
    border: 1px solid {Theme.BORDER};
    border-radius: 6px;
    padding: 7px 10px;
    color: {Theme.TEXT_PRIMARY};
    font-size: 12px;
    selection-background-color: {Theme.ACCENT_DIM};
}}
QDateEdit:focus {{
    border-color: {Theme.ACCENT};
}}
QDateEdit::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 20px;
    border: none;
    padding-right: 5px;
}}
QDateEdit::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid {Theme.TEXT_SECONDARY};
    margin-right: 5px;
}}
QDateEdit::down-arrow:hover {{
    border-top-color: {Theme.ACCENT};
}}

/* ── Calendar popup styling ── */
QCalendarWidget {{
    background-color: {Theme.BG_DARK};
    border: 1px solid {Theme.BORDER};
    border-radius: 8px;
}}
QCalendarWidget QToolButton {{
    color: {Theme.TEXT_PRIMARY};
    background-color: {Theme.BG_MID};
    border: none;
    border-radius: 4px;
    padding: 5px;
    margin: 2px;
}}
QCalendarWidget QToolButton:hover {{
    background-color: {Theme.ACCENT};
    color: white;
}}
QCalendarWidget QMenu {{
    background-color: {Theme.BG_DARK};
    color: {Theme.TEXT_PRIMARY};
    border: 1px solid {Theme.BORDER};
}}
QCalendarWidget QMenu::item:selected {{
    background-color: {Theme.ACCENT};
    color: white;
}}
QCalendarWidget QSpinBox {{
    background-color: {Theme.BG_INPUT};
    color: {Theme.TEXT_PRIMARY};
    border: 1px solid {Theme.BORDER};
    border-radius: 4px;
    padding: 2px 5px;
}}
QCalendarWidget QWidget#qt_calendar_navigationbar {{
    background-color: {Theme.BG_MID};
    border-bottom: 1px solid {Theme.BORDER};
    padding: 4px;
}}
QCalendarWidget QAbstractItemView:enabled {{
    background-color: {Theme.BG_DARK};
    color: {Theme.TEXT_PRIMARY};
    selection-background-color: {Theme.ACCENT};
    selection-color: white;
}}
QCalendarWidget QAbstractItemView:disabled {{
    color: {Theme.TEXT_MUTED};
}}

/* ── Labels ── */
QLabel {{
    background: transparent;
    border: none;
}}

/* ── Buttons ── */
QPushButton {{
    background-color: {Theme.BG_MID};
    color: {Theme.TEXT_PRIMARY};
    border: 1px solid {Theme.BORDER};
    border-radius: 6px;
    padding: 6px 14px;
    font-weight: 600;
    font-size: 12px;
}}
QPushButton:hover {{
    background-color: {Theme.BG_LIGHT};
    border-color: {Theme.BORDER_LIGHT};
}}
QPushButton:pressed {{
    background-color: {Theme.ACCENT_DIM};
}}
QPushButton:disabled {{
    color: {Theme.TEXT_MUTED};
    background-color: {Theme.BG_DARK};
    border-color: {Theme.BORDER};
}}

/* ── Primary button (blue) ── */
QPushButton[cssClass="primary"] {{
    background-color: {Theme.ACCENT};
    color: #ffffff;
    border: none;
    font-weight: 700;
}}
QPushButton[cssClass="primary"]:hover {{
    background-color: {Theme.ACCENT_HOVER};
}}
QPushButton[cssClass="primary"]:disabled {{
    background-color: {Theme.ACCENT_DIM};
    color: {Theme.TEXT_MUTED};
}}

/* ── Text edits (log viewer) ── */
QTextEdit {{
    background-color: {Theme.BG_DARK};
    color: {Theme.TEXT_PRIMARY};
    border: 1px solid {Theme.BORDER};
    border-radius: 8px;
    padding: 10px;
    font-family: "Cascadia Code", "Fira Code", "Consolas", monospace;
    font-size: 12px;
    selection-background-color: {Theme.ACCENT_DIM};
}}

/* ── Frames ── */
QFrame[cssClass="separator"] {{
    background-color: {Theme.BORDER};
    max-height: 1px;
}}
"""


# ─────────────────────────────────────────────
# Step Card Widget (matching HTML mockup)
# ─────────────────────────────────────────────
class StepCard(QFrame):
    run_clicked = Signal(object)

    STATUS_COLOURS = {
        StepStatus.IDLE:     Theme.STATUS_PENDING,
        StepStatus.RUNNING:  Theme.STATUS_RUNNING,
        StepStatus.COMPLETE: Theme.STATUS_COMPLETE,
        StepStatus.FAILED:   Theme.STATUS_FAILED,
    }

    STATUS_ICONS = {
        StepStatus.IDLE:     "○",
        StepStatus.RUNNING:  "▶",
        StepStatus.COMPLETE: "✓",
        StepStatus.FAILED:   "✗",
    }

    STATUS_LABELS = {
        StepStatus.IDLE:     "READY",
        StepStatus.RUNNING:  "RUNNING...",
        StepStatus.COMPLETE: "COMPLETE",
        StepStatus.FAILED:   "FAILED",
    }

    def __init__(self, step: WorkflowStep, parent=None):
        super().__init__(parent)
        self.step = step
        self.status = StepStatus.IDLE
        self.dependency_met = True  # Track if dependencies are satisfied
        self.setObjectName("StepCard")
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        # Main horizontal layout
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(10)

        # Status indicator dot
        self.status_dot = QLabel()
        self.status_dot.setFixedSize(24, 24)
        self.status_dot.setAlignment(Qt.AlignCenter)
        self.status_dot.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(self.status_dot)

        # Text block (takes remaining space)
        text_widget = QWidget()
        text_layout = QVBoxLayout(text_widget)
        text_layout.setSpacing(2)
        text_layout.setContentsMargins(0, 0, 0, 0)

        self.title_label = QLabel(self.step.title)
        self.title_label.setFont(QFont("Segoe UI", 10, QFont.DemiBold))
        self.title_label.setStyleSheet(f"color: {Theme.TEXT_PRIMARY};")
        text_layout.addWidget(self.title_label)

        self.desc_label = QLabel(self.step.description)
        self.desc_label.setFont(QFont("Segoe UI", 8))
        self.desc_label.setStyleSheet(f"color: {Theme.TEXT_SECONDARY};")
        self.desc_label.setWordWrap(True)
        text_layout.addWidget(self.desc_label)

        text_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(text_widget, 1)

        # Status badge (fixed width - matching HTML mockup style)
        self.status_badge = QLabel()
        self.status_badge.setFont(QFont("Segoe UI", 7, QFont.Bold))
        self.status_badge.setAlignment(Qt.AlignCenter)
        self.status_badge.setFixedHeight(20)
        self.status_badge.setMinimumWidth(70)
        layout.addWidget(self.status_badge)

        # Run/Confirm button (gray background with green text, matching HTML mockup)
        # All buttons same width (70px) for consistent alignment
        button_text = "CONFIRM" if self.step.is_confirmation else "RUN"
        self.run_btn = QPushButton(button_text)
        self.run_btn.setFixedSize(70, 24)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(lambda: self.run_clicked.emit(self.step))
        layout.addWidget(self.run_btn)

    def refresh(self):
        colour = self.STATUS_COLOURS[self.status]
        icon = self.STATUS_ICONS[self.status]
        label = self.STATUS_LABELS[self.status]

        # Override for waiting state (dependency not met)
        if not self.dependency_met and self.status == StepStatus.IDLE:
            colour = Theme.TEXT_MUTED
            icon = "◷"  # Clock icon for waiting
            label = "WAITING"

        # Status dot
        self.status_dot.setText(icon)
        self.status_dot.setStyleSheet(
            f"color: {colour}; "
            f"background: {Theme.BG_DARKEST}; "
            f"border-radius: 12px; "
            f"border: 2px solid {colour};"
        )

        # Status badge - different background colors for each status (matching HTML mockup)
        badge_styles = {
            StepStatus.IDLE: f"color: #6e7681; background: rgba(110, 118, 129, 0.1);",
            StepStatus.RUNNING: f"color: #58a6ff; background: rgba(88, 166, 255, 0.1);",
            StepStatus.COMPLETE: f"color: #3fb950; background: rgba(63, 185, 80, 0.1);",
            StepStatus.FAILED: f"color: #f85149; background: rgba(248, 81, 73, 0.1);",
        }
        
        # Use waiting style if dependency not met
        if not self.dependency_met and self.status == StepStatus.IDLE:
            badge_style = f"color: {Theme.TEXT_MUTED}; background: rgba(110, 118, 129, 0.1);"
        else:
            badge_style = badge_styles[self.status]
        
        self.status_badge.setText(label)
        self.status_badge.setStyleSheet(
            badge_style + 
            "border-radius: 4px; "
            "padding: 2px 8px; "
            "font-weight: bold; "
            "font-size: 9px; "
            "letter-spacing: 0.5px;"
        )

        # Run button (gray background with green text, matching HTML mockup style)
        button_text = "CONFIRM" if self.step.is_confirmation else "RUN"
        if self.status == StepStatus.RUNNING:
            self.run_btn.setEnabled(False)
            self.run_btn.setText(button_text)
            self.run_btn.setStyleSheet(
                f"background-color: {Theme.BG_MID}; "
                f"color: {Theme.TEXT_MUTED}; "
                f"border: 1px solid {Theme.BORDER}; "
                f"border-radius: 5px; "
                f"font-weight: 700; "
                f"font-size: 9px;"
            )
        elif not self.dependency_met:
            # Dependency not met - show disabled with waiting state
            self.run_btn.setEnabled(False)
            self.run_btn.setText(button_text)
            self.run_btn.setStyleSheet(
                f"background-color: {Theme.BG_MID}; "
                f"color: {Theme.TEXT_MUTED}; "
                f"border: 1px solid {Theme.BORDER}; "
                f"border-radius: 5px; "
                f"font-weight: 700; "
                f"font-size: 9px;"
            )
        else:
            self.run_btn.setEnabled(True)
            self.run_btn.setText(button_text)
            self.run_btn.setStyleSheet(
                f"background-color: {Theme.BG_MID}; "
                f"color: {Theme.STATUS_COMPLETE}; "
                f"border: 1px solid {Theme.BORDER}; "
                f"border-radius: 5px; "
                f"font-weight: 700; "
                f"font-size: 9px;"
            )

        # Card border (highlight when running)
        if self.status == StepStatus.RUNNING:
            border_colour = Theme.STATUS_RUNNING
        else:
            border_colour = Theme.BORDER

        self.setStyleSheet(
            f"#StepCard {{ "
            f"background: {Theme.BG_DARK}; "
            f"border: 1px solid {border_colour}; "
            f"border-radius: 8px; "
            f"}}"
        )

    def set_status(self, status: StepStatus):
        self.status = status
        self.refresh()

    def set_dependency_met(self, met: bool):
        """Set whether this step's dependencies are met."""
        self.dependency_met = met
        self.refresh()

    def set_sequential_waiting(self, waiting: bool):
        """Set whether this step is waiting in sequential mode."""
        # Reuse dependency_met for visual state - False = waiting
        self.dependency_met = not waiting
        self.refresh()


# ─────────────────────────────────────────────
# Workflow Section Widget (matching HTML mockup)
# ─────────────────────────────────────────────
class WorkflowSection(QWidget):
    step_run_requested = Signal(object)

    def __init__(self, title: str, accent_colour: str, bg_colour: str, steps: list[WorkflowStep], parent=None):
        super().__init__(parent)
        self.title = title
        self.accent_colour = accent_colour
        self.bg_colour = bg_colour
        self.steps = steps
        self.cards: dict[str, StepCard] = {}
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # Section header (matching HTML: subtle tinted background with left border)
        header = QLabel(self.title.upper())
        header.setFont(QFont("Segoe UI", 8, QFont.Bold))
        header.setStyleSheet(
            f"color: {self.accent_colour}; "
            f"padding: 6px 10px; "
            f"letter-spacing: 1px; "
            f"border-left: 3px solid {self.accent_colour}; "
            f"background-color: {self.bg_colour}; "
            f"border-radius: 0px 6px 6px 0px;"
        )
        layout.addWidget(header)

        # Step cards
        for step in self.steps:
            card = StepCard(step)
            card.run_clicked.connect(self.step_run_requested.emit)
            self.cards[step.id] = card
            layout.addWidget(card)
        
        # Initialize dependencies
        self._update_dependencies()

    def _update_dependencies(self, enforce: bool = True):
        """Update dependency state for all cards in this section.
        
        Args:
            enforce: If False, mark all dependencies as met (allow manual running)
        """
        for step in self.steps:
            card = self.cards[step.id]
            if not enforce:
                # Don't enforce dependencies - allow manual running
                card.set_dependency_met(True)
            elif step.depends_on:
                # Check if the dependency step is complete
                dep_card = self.cards.get(step.depends_on)
                if dep_card:
                    card.set_dependency_met(dep_card.status == StepStatus.COMPLETE)
                else:
                    card.set_dependency_met(False)
            else:
                card.set_dependency_met(True)

    def on_step_completed(self, step_id: str):
        """Called when a step completes - update dependencies."""
        self._update_dependencies()

    def get_card(self, step_id: str) -> Optional[StepCard]:
        return self.cards.get(step_id)

    def set_all_buttons_enabled(self, enabled: bool):
        for card in self.cards.values():
            button_text = "CONFIRM" if card.step.is_confirmation else "RUN"
            # Check if dependencies are met
            dep_met = card.dependency_met
            
            if enabled and card.status != StepStatus.RUNNING and dep_met:
                card.run_btn.setEnabled(True)
                card.run_btn.setText(button_text)
                card.run_btn.setStyleSheet(
                    f"background-color: {Theme.BG_MID}; "
                    f"color: {Theme.STATUS_COMPLETE}; "
                    f"border: 1px solid {Theme.BORDER}; "
                    f"border-radius: 5px; "
                    f"font-weight: 700; "
                    f"font-size: 9px;"
                )
            else:
                card.run_btn.setEnabled(False)
                card.run_btn.setText(button_text)
                card.run_btn.setStyleSheet(
                    f"background-color: {Theme.BG_MID}; "
                    f"color: {Theme.TEXT_MUTED}; "
                    f"border: 1px solid {Theme.BORDER}; "
                    f"border-radius: 5px; "
                    f"font-weight: 700; "
                    f"font-size: 9px;"
                )


# ─────────────────────────────────────────────
# Settings Panel
# ─────────────────────────────────────────────
class SettingsPanel(QWidget):
    CONFIG_FILE = "mt5_workflow_config.json"
    
    # Signal emitted when settings are saved
    settings_saved = Signal()

    COMMON_FIELDS = [
        ("MT5Folder", "MetaTrader 5 Folder", True, False),
        ("MT5MQL5Folder", "MQL5 Data Folder", True, False),
    ]

    QDM_FIELDS = [
        ("QDMFolder", "QuantDataManager Folder", True, False),
        ("ExportDataFrom", "Export Data From", False, True),
        ("ExportDataTo", "Export Data To", False, True),
    ]

    BACKTEST_FIELDS = [
        ("BacktestOutputFolder", "Backtest Output Folder", True, False),
        ("MT5BackTestFrom", "Backtest From Date", False, True),
        ("MT5BackTestTo", "Backtest To Date", False, True),
    ]

    MONTECARLO_FIELDS = [
        ("QAPath", "Quant Analyzer Exe", False, False),
        ("MC95Threshold", "95% Confidence Level", False, False),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._fields: dict[str, QLineEdit] = {}
        self._build_ui()
        self._load_config()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 12, 16, 12)
        main_layout.setSpacing(12)

        # Three column layout
        sections_layout = QHBoxLayout()
        sections_layout.setSpacing(16)

        # MetaTrader Settings
        common_group = self._create_section_group(
            "METATRADER SETTINGS",
            "MetaTrader Folder Settings",
            self.COMMON_FIELDS
        )
        sections_layout.addWidget(common_group)

        # QuantDataManager Settings
        qdm_group = self._create_section_group(
            "QUANTDATAMANAGER SETTINGS",
            "Settings for data export",
            self.QDM_FIELDS
        )
        sections_layout.addWidget(qdm_group)

        # Back Test Settings
        backtest_group = self._create_section_group(
            "BACK TEST SETTINGS",
            "Settings for EA backtesting",
            self.BACKTEST_FIELDS
        )
        sections_layout.addWidget(backtest_group)

        # Monte Carlo Settings
        montecarlo_group = self._create_section_group(
            "MONTE CARLO SETTINGS",
            "Settings for Monte Carlo analysis",
            self.MONTECARLO_FIELDS
        )
        sections_layout.addWidget(montecarlo_group)

        # Workflow Options
        workflow_group = self._create_workflow_options_group()
        sections_layout.addWidget(workflow_group)

        main_layout.addLayout(sections_layout)

        # Save button row
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.save_btn = QPushButton("Save Settings")
        self.save_btn.setProperty("cssClass", "primary")
        self.save_btn.clicked.connect(self._save_config)
        btn_row.addWidget(self.save_btn)

        main_layout.addLayout(btn_row)

    def _create_section_group(self, title: str, description: str, fields: list) -> QGroupBox:
        group = QGroupBox(title)
        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        desc_label = QLabel(description)
        desc_label.setFont(QFont("Segoe UI", 10))
        desc_label.setStyleSheet(f"color: {Theme.TEXT_MUTED}; padding-bottom: 4px;")
        layout.addWidget(desc_label)

        grid = QGridLayout()
        grid.setSpacing(6)

        for row, (key, label, is_folder, is_date) in enumerate(fields):
            lbl = QLabel(label)
            lbl.setFont(QFont("Segoe UI", 11))
            lbl.setStyleSheet(f"color: {Theme.TEXT_SECONDARY};")
            grid.addWidget(lbl, row, 0)

            if is_date:
                # Use QDateEdit with calendar popup for date fields
                # Span across columns 1 and 2 to align with folder fields
                date_edit = QDateEdit()
                date_edit.setCalendarPopup(True)
                date_edit.setDisplayFormat("yyyy.MM.dd")
                date_edit.setFixedHeight(30)
                date_edit.setDate(QDate.currentDate())
                self._fields[key] = date_edit
                grid.addWidget(date_edit, row, 1, 1, 2)  # Span 2 columns
            else:
                # Use QLineEdit for folder paths
                line_edit = QLineEdit()
                line_edit.setPlaceholderText("Select folder…")
                line_edit.setAlignment(Qt.AlignLeft)
                line_edit.setCursorPosition(0)
                self._fields[key] = line_edit
                grid.addWidget(line_edit, row, 1)

                if is_folder:
                    browse_btn = QPushButton("…")
                    browse_btn.setFixedWidth(36)
                    browse_btn.setFixedHeight(30)
                    browse_btn.clicked.connect(lambda checked, k=key: self._browse(k))
                    grid.addWidget(browse_btn, row, 2)

        layout.addLayout(grid)
        layout.addStretch()
        return group

    def _create_workflow_options_group(self) -> QGroupBox:
        """Create the workflow options group with checkboxes."""
        group = QGroupBox("WORKFLOW OPTIONS")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        desc_label = QLabel("Automation settings")
        desc_label.setFont(QFont("Segoe UI", 10))
        desc_label.setStyleSheet(f"color: {Theme.TEXT_MUTED}; padding-bottom: 4px;")
        layout.addWidget(desc_label)

        # Sequential Execution checkbox
        self._sequential_checkbox = QCheckBox("Sequential Execution")
        self._sequential_checkbox.setFont(QFont("Segoe UI", 11))
        self._sequential_checkbox.setStyleSheet(f"color: {Theme.TEXT_SECONDARY};")
        self._sequential_checkbox.setToolTip(
            "When enabled, Back Test and Monte Carlo steps will run automatically in sequence.\n"
            "Only the first Back Test step will have an active RUN button."
        )
        self._fields["SequentialExecution"] = self._sequential_checkbox
        layout.addWidget(self._sequential_checkbox)

        # QA Use Images checkbox
        self._qa_use_images_checkbox = QCheckBox("Use Image Recognition for QA")
        self._qa_use_images_checkbox.setFont(QFont("Segoe UI", 11))
        self._qa_use_images_checkbox.setStyleSheet(f"color: {Theme.TEXT_SECONDARY};")
        self._qa_use_images_checkbox.setToolTip(
            "When enabled, Quant Analyzer automation will use image recognition\n"
            "instead of coordinate-based clicking. More reliable across different\n"
            "screen resolutions and RDP sessions.\n\n"
            "Requires running --capture-templates once to create template images."
        )
        self._fields["QAUseImages"] = self._qa_use_images_checkbox
        layout.addWidget(self._qa_use_images_checkbox)

        layout.addStretch()
        return group

    def _browse(self, key: str):
        folder = QFileDialog.getExistingDirectory(self, f"Select {key}")
        if folder:
            self._fields[key].setText(folder)
            self._fields[key].setCursorPosition(0)  # Show start of text

    def get_settings(self) -> Settings:
        values = {}
        for key, widget in self._fields.items():
            if isinstance(widget, QDateEdit):
                # Get date as string in YYYY.MM.DD format
                values[key] = widget.date().toString("yyyy.MM.dd")
            elif isinstance(widget, QCheckBox):
                # Get checkbox as boolean
                values[key] = widget.isChecked()
            else:
                values[key] = widget.text()
        return Settings(**values)

    def _config_path(self) -> Path:
        return Path.home() / ".mt5_workflow" / self.CONFIG_FILE

    def _set_field_value(self, key: str, value):
        """Set a field value, handling QLineEdit, QDateEdit, and QCheckBox."""
        widget = self._fields.get(key)
        if widget is None:
            return
        
        if isinstance(widget, QDateEdit):
            # Parse date string (YYYY.MM.DD format)
            date = QDate.fromString(str(value), "yyyy.MM.dd")
            if date.isValid():
                widget.setDate(date)
        elif isinstance(widget, QCheckBox):
            # Handle boolean or string "true"/"false"
            if isinstance(value, bool):
                widget.setChecked(value)
            else:
                widget.setChecked(str(value).lower() in ("true", "1", "yes"))
        else:
            widget.setText(str(value))
            widget.setCursorPosition(0)

    def _get_field_value(self, key: str):
        """Get a field value, handling QLineEdit, QDateEdit, and QCheckBox."""
        widget = self._fields.get(key)
        if widget is None:
            return ""
        
        if isinstance(widget, QDateEdit):
            return widget.date().toString("yyyy.MM.dd")
        elif isinstance(widget, QCheckBox):
            return widget.isChecked()
        else:
            return widget.text()

    def _load_config(self):
        defaults = Settings()
        for key in self._fields.keys():
            default_value = getattr(defaults, key, "")
            self._set_field_value(key, default_value)

        path = self._config_path()
        if path.exists():
            try:
                data = json.loads(path.read_text())
                for key, value in data.items():
                    if key in self._fields:
                        self._set_field_value(key, value)
            except Exception as e:
                print(f"Warning: Could not load config: {e}")

    def _save_config(self):
        path = self._config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        data = {k: self._get_field_value(k) for k in self._fields.keys()}
        path.write_text(json.dumps(data, indent=2))

        # Emit signal to notify that settings were saved (used to collapse panel)
        self.settings_saved.emit()


# ─────────────────────────────────────────────
# Log Panel with colored output
# ─────────────────────────────────────────────
class LogPanel(QWidget):
    # Colors for log output
    COLOR_TIMESTAMP = "#6e7681"
    COLOR_INFO = "#58a6ff"
    COLOR_SUCCESS = "#3fb950"
    COLOR_WARNING = "#f59e0b"
    COLOR_ERROR = "#f85149"
    COLOR_DEFAULT = "#e6edf3"
    COLOR_DIM = "#8b949e"

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(16, 12, 16, 12)

        self.title_label = QLabel("Output Log")
        self.title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.title_label.setStyleSheet(f"color: {Theme.TEXT_PRIMARY};")
        header_layout.addWidget(self.title_label)

        header_layout.addStretch()

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setFixedWidth(60)
        self.clear_btn.clicked.connect(self.clear)
        header_layout.addWidget(self.clear_btn)

        layout.addWidget(header)

        # Separator
        sep = QFrame()
        sep.setProperty("cssClass", "separator")
        sep.setFrameShape(QFrame.HLine)
        sep.setFixedHeight(1)
        layout.addWidget(sep)

        # Container for log viewer with padding so border-radius is visible
        log_container = QWidget()
        log_layout = QVBoxLayout(log_container)
        log_layout.setContentsMargins(16, 16, 16, 16)
        log_layout.setSpacing(0)

        # Log viewer with HTML support - matching HTML mockup CSS
        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setPlaceholderText("Click 'Run' on any step to see output here…")
        self.log_viewer.setStyleSheet(
            f"QTextEdit {{ "
            f"background: {Theme.BG_DARK}; "
            f"border: 1px solid {Theme.BORDER}; "
            f"border-radius: 8px; "
            f"padding: 14px 16px; "
            f"font-family: 'JetBrains Mono', 'Cascadia Code', 'Fira Code', 'Consolas', monospace; "
            f"font-size: 12px; "
            f"color: {Theme.TEXT_PRIMARY}; "
            f"}}"
        )
        # Set default document style for line-height
        self.log_viewer.document().setDefaultStyleSheet(
            "* { line-height: 1.7; }"
        )
        log_layout.addWidget(self.log_viewer, 1)
        
        layout.addWidget(log_container, 1)

    def _html_escape(self, text: str) -> str:
        """Escape HTML special characters."""
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _format_line(self, text: str) -> str:
        """Format a line of text with appropriate colors based on content."""
        escaped = self._html_escape(text)
        
        # Check for error patterns (case-insensitive)
        text_lower = text.lower()
        
        # Error patterns - color entire line red
        error_patterns = [
            'error:', 'error ', '[error]', 'failed', 'failure', 
            'exception:', 'traceback', 'errno', 'winerror',
            'not found', 'cannot find', 'could not', 'unable to'
        ]
        for pattern in error_patterns:
            if pattern in text_lower:
                return f'<span style="color: {self.COLOR_ERROR};">{escaped}</span>'
        
        # Warning patterns - color entire line orange/yellow
        warning_patterns = ['warning:', 'warning ', '[warning]', 'warn:', 'skipping']
        for pattern in warning_patterns:
            if pattern in text_lower:
                return f'<span style="color: {self.COLOR_WARNING};">{escaped}</span>'
        
        # Success patterns - color entire line green
        success_patterns = ['success', 'completed', 'complete!', 'done', 'ok', 'passed']
        for pattern in success_patterns:
            if pattern in text_lower and 'not' not in text_lower:
                return f'<span style="color: {self.COLOR_SUCCESS};">{escaped}</span>'
        
        # Info patterns - color entire line blue
        info_patterns = ['info:', '[info]', 'starting', 'processing', 'running']
        for pattern in info_patterns:
            if pattern in text_lower:
                return f'<span style="color: {self.COLOR_INFO};">{escaped}</span>'
        
        # Detect and color timestamps like [10:42:15] or 15:27:04.182
        timestamp_pattern = r'(\[?\d{1,2}:\d{2}:\d{2}(?:\.\d+)?\]?)'
        
        def color_timestamp(match):
            return f'<span style="color: {self.COLOR_TIMESTAMP};">{match.group(1)}</span>'
        
        escaped = re.sub(timestamp_pattern, color_timestamp, escaped)
        
        return escaped

    def append(self, text: str):
        """Append text, auto-detecting and coloring output."""
        # Split into lines and format each
        lines = text.split('\n')
        html_parts = []
        
        for line in lines:
            if not line:
                html_parts.append("<br>")
                continue
                
            formatted = self._format_line(line)
            html_parts.append(f'<div style="margin: 0; padding: 0; line-height: 1.7;">{formatted}</div>')
        
        html = "".join(html_parts)
        
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        self.log_viewer.setTextCursor(cursor)
        self.log_viewer.ensureCursorVisible()

    def append_line(self, text: str):
        self.append(text + "\n")

    def append_header(self, text: str):
        """Append a styled header."""
        separator = "═" * 60
        html = (
            f'<div style="line-height: 1.7;"><br></div>'
            f'<div style="color: {self.COLOR_INFO}; line-height: 1.7;">{separator}</div>'
            f'<div style="color: {self.COLOR_INFO}; font-weight: bold; line-height: 1.7;">  {self._html_escape(text)}</div>'
            f'<div style="color: {self.COLOR_INFO}; line-height: 1.7;">{separator}</div>'
            f'<div style="line-height: 1.7;"><br></div>'
        )
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        self.log_viewer.setTextCursor(cursor)
        self.log_viewer.ensureCursorVisible()

    def append_info(self, text: str):
        """Append info-styled text (blue)."""
        html = f'<div style="color: {self.COLOR_INFO}; line-height: 1.7;">{self._html_escape(text)}</div>'
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        self.log_viewer.setTextCursor(cursor)
        self.log_viewer.ensureCursorVisible()

    def append_success(self, text: str):
        """Append success-styled text (green)."""
        html = f'<div style="line-height: 1.7;"><br></div><div style="color: {self.COLOR_SUCCESS}; line-height: 1.7;">✓ {self._html_escape(text)}</div>'
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        self.log_viewer.setTextCursor(cursor)
        self.log_viewer.ensureCursorVisible()

    def append_error(self, text: str):
        """Append error-styled text (red)."""
        html = f'<div style="line-height: 1.7;"><br></div><div style="color: {self.COLOR_ERROR}; line-height: 1.7;">✗ ERROR: {self._html_escape(text)}</div>'
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        self.log_viewer.setTextCursor(cursor)
        self.log_viewer.ensureCursorVisible()

    def append_dim(self, text: str):
        """Append dimmed text."""
        html = f'<div style="color: {self.COLOR_DIM}; line-height: 1.7;">{self._html_escape(text)}</div>'
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        self.log_viewer.setTextCursor(cursor)
        self.log_viewer.ensureCursorVisible()

    def clear(self):
        self.log_viewer.clear()

    def set_title(self, title: str):
        self.title_label.setText(title)


# ─────────────────────────────────────────────
# Main Window
# ─────────────────────────────────────────────
class WorkflowWindow(QMainWindow):
    # Signal for thread-safe output updates
    _output_signal = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MetaTrader 5 Workflow Manager")
        self.resize(1400, 800)
        self.setMinimumSize(1150, 600)

        # Connect output signal
        self._output_signal.connect(self._append_output)

        # Build step lists
        self.data_steps = build_data_update_steps()
        self.backtest_steps = build_backtest_steps()
        self.montecarlo_steps = build_montecarlo_steps()
        self.tick_montecarlo_steps = build_tick_montecarlo_steps()

        # Current running process (subprocess.Popen)
        self.current_process = None
        self.current_step_id: Optional[str] = None
        
        # Sequential execution mode
        self._sequential_running = False  # True when running a sequential batch

        # Config file path for UI state
        self._ui_config_path = Path.home() / ".mt5_workflow" / "ui_state.json"

        self._build_ui()
        self._load_ui_state()
        
        # Apply initial sequential mode state and connect checkbox change
        self._update_sequential_mode()
        self.settings_panel._sequential_checkbox.stateChanged.connect(self._update_sequential_mode)

    def _save_ui_state(self):
        """Save window geometry and splitter position."""
        try:
            self._ui_config_path.parent.mkdir(parents=True, exist_ok=True)
            state = {
                "window_x": self.x(),
                "window_y": self.y(),
                "window_width": self.width(),
                "window_height": self.height(),
                "splitter_sizes": self.splitter.sizes(),
                "window_maximized": self.isMaximized(),
                "settings_visible": self.settings_panel.isVisible(),
            }
            self._ui_config_path.write_text(json.dumps(state, indent=2))
        except Exception as e:
            print(f"Warning: Could not save UI state: {e}")

    def _load_ui_state(self):
        """Load window geometry and splitter position."""
        if not self._ui_config_path.exists():
            return

        try:
            state = json.loads(self._ui_config_path.read_text())

            # Restore window geometry
            if state.get("window_maximized"):
                self.showMaximized()
            else:
                if "window_x" in state and "window_y" in state:
                    self.move(state["window_x"], state["window_y"])
                if "window_width" in state and "window_height" in state:
                    self.resize(state["window_width"], state["window_height"])

            # Restore splitter sizes
            if "splitter_sizes" in state:
                self.splitter.setSizes(state["splitter_sizes"])

            # Restore settings panel visibility
            if state.get("settings_visible"):
                self.settings_panel.setVisible(True)
                self.settings_toggle_btn.setChecked(True)

        except Exception as e:
            print(f"Warning: Could not load UI state: {e}")

    def closeEvent(self, event):
        """Save UI state when window is closed."""
        self._save_ui_state()
        super().closeEvent(event)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ── Title bar ──
        title_bar = QWidget()
        title_bar.setStyleSheet(
            f"background: {Theme.BG_DARK}; "
            f"border-bottom: 1px solid {Theme.BORDER};"
        )
        tb_layout = QHBoxLayout(title_bar)
        tb_layout.setContentsMargins(20, 12, 20, 12)

        app_title = QLabel("MetaTrader 5 Workflow Manager")
        app_title.setFont(QFont("Segoe UI", 15, QFont.Bold))
        app_title.setStyleSheet(f"color: {Theme.ACCENT};")
        tb_layout.addWidget(app_title)

        tb_layout.addStretch()

        # Settings button
        self.settings_toggle_btn = QPushButton("Settings")
        self.settings_toggle_btn.setCheckable(True)
        self.settings_toggle_btn.clicked.connect(self._toggle_settings)
        tb_layout.addWidget(self.settings_toggle_btn)

        main_layout.addWidget(title_bar)

        # ── Settings panel (collapsible) ──
        self.settings_panel = SettingsPanel()
        self.settings_panel.setVisible(False)
        self.settings_panel.setStyleSheet(
            f"background: {Theme.BG_DARK}; "
            f"border-bottom: 1px solid {Theme.BORDER};"
        )
        # Connect signal to collapse settings after saving
        self.settings_panel.settings_saved.connect(self._on_settings_saved)
        main_layout.addWidget(self.settings_panel)

        # ── Body: left steps + right log ──
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setHandleWidth(1)

        # Left panel — workflow sections
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(16)

        # Scrollable area for steps
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(0, 0, 12, 0)  # More right margin for scrollbar
        scroll_layout.setSpacing(20)

        # Data Update Section
        self.data_section = WorkflowSection(
            "Update MetaTrader Data",
            Theme.SECTION_DATA,
            Theme.SECTION_DATA_BG,
            self.data_steps
        )
        self.data_section.step_run_requested.connect(self._on_run_step)
        scroll_layout.addWidget(self.data_section)

        # Backtest Section
        self.backtest_section = WorkflowSection(
            "Back Test MetaTrader Expert Advisors",
            Theme.SECTION_BACKTEST,
            Theme.SECTION_BACKTEST_BG,
            self.backtest_steps
        )
        self.backtest_section.step_run_requested.connect(self._on_run_step)
        scroll_layout.addWidget(self.backtest_section)

        # Monte Carlo Analysis - M1 Section
        self.montecarlo_section = WorkflowSection(
            "Monte Carlo Analysis - M1",
            Theme.SECTION_MONTECARLO,
            Theme.SECTION_MONTECARLO_BG,
            self.montecarlo_steps
        )
        self.montecarlo_section.step_run_requested.connect(self._on_run_step)
        scroll_layout.addWidget(self.montecarlo_section)

        # Monte Carlo Analysis - Tick Section
        self.tick_montecarlo_section = WorkflowSection(
            "Monte Carlo Analysis - Tick",
            Theme.SECTION_MONTECARLO,
            Theme.SECTION_MONTECARLO_BG,
            self.tick_montecarlo_steps
        )
        self.tick_montecarlo_section.step_run_requested.connect(self._on_run_step)
        scroll_layout.addWidget(self.tick_montecarlo_section)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        left_layout.addWidget(scroll, 1)

        # Set minimum width for left panel to prevent button cutoff
        left_panel.setMinimumWidth(520)
        self.splitter.addWidget(left_panel)

        # Right panel — log output
        self.log_panel = LogPanel()
        self.splitter.addWidget(self.log_panel)

        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        self.splitter.setSizes([520, 880])  # Initial sizes for left and right panels

        main_layout.addWidget(self.splitter, 1)

        # ── Status bar ──
        status_bar = QWidget()
        status_bar.setFixedHeight(32)
        status_bar.setStyleSheet(
            f"background: {Theme.BG_DARK}; "
            f"border-top: 1px solid {Theme.BORDER};"
        )
        sb_layout = QHBoxLayout(status_bar)
        sb_layout.setContentsMargins(16, 0, 16, 0)

        self.status_label = QLabel("Ready")
        self.status_label.setFont(QFont("Segoe UI", 10))
        self.status_label.setStyleSheet(f"color: {Theme.TEXT_MUTED};")
        sb_layout.addWidget(self.status_label)

        sb_layout.addStretch()

        version_label = QLabel("v1.0.0")
        version_label.setFont(QFont("Segoe UI", 10))
        version_label.setStyleSheet(f"color: {Theme.TEXT_MUTED};")
        sb_layout.addWidget(version_label)

        main_layout.addWidget(status_bar)

    def _toggle_settings(self):
        visible = self.settings_toggle_btn.isChecked()
        self.settings_panel.setVisible(visible)

    def _on_settings_saved(self):
        """Called when settings are saved - collapse the settings panel."""
        self.settings_toggle_btn.setChecked(False)
        self.settings_panel.setVisible(False)
        # Update sequential mode in case checkbox changed
        self._update_sequential_mode()

    def _get_sequential_steps(self) -> list[WorkflowStep]:
        """Get the ordered list of steps for sequential execution."""
        # Backtest steps followed by Monte Carlo M1 steps, then Tick MC steps
        return self.backtest_steps + self.montecarlo_steps + self.tick_montecarlo_steps

    def _get_next_sequential_step(self, current_step_id: str) -> Optional[WorkflowStep]:
        """Get the next step in the sequential execution order."""
        steps = self._get_sequential_steps()
        for i, step in enumerate(steps):
            if step.id == current_step_id:
                if i + 1 < len(steps):
                    return steps[i + 1]
                return None
        return None

    def _is_sequential_mode(self) -> bool:
        """Check if sequential execution mode is enabled."""
        settings = self.settings_panel.get_settings()
        return settings.SequentialExecution

    def _update_sequential_mode(self):
        """Update UI based on sequential execution mode setting."""
        is_sequential = self._is_sequential_mode()
        
        if is_sequential and not self._sequential_running:
            # Disable all buttons except the first backtest step
            self._update_sequential_button_states()
        else:
            # Enable all buttons normally (respecting dependencies)
            self._restore_normal_button_states()

    def _update_sequential_button_states(self):
        """Update button states for sequential mode - only first step is enabled."""
        sequential_steps = self._get_sequential_steps()
        first_step_id = sequential_steps[0].id if sequential_steps else None
        
        # Update backtest section
        for step in self.backtest_steps:
            card = self.backtest_section.cards.get(step.id)
            if card:
                if step.id == first_step_id:
                    card.set_sequential_waiting(False)  # First step is active
                else:
                    card.set_sequential_waiting(True)   # Other steps show WAITING
        
        # Update monte carlo section
        for step in self.montecarlo_steps:
            card = self.montecarlo_section.cards.get(step.id)
            if card:
                card.set_sequential_waiting(True)  # All MC steps show WAITING initially

        # Update tick monte carlo section
        for step in self.tick_montecarlo_steps:
            card = self.tick_montecarlo_section.cards.get(step.id)
            if card:
                card.set_sequential_waiting(True)  # All Tick MC steps show WAITING initially

    def _restore_normal_button_states(self):
        """Restore normal button states (non-sequential mode)."""
        # Update backtest section - all ready
        for step in self.backtest_steps:
            card = self.backtest_section.cards.get(step.id)
            if card:
                card.set_sequential_waiting(False)

        # Update monte carlo section - clear waiting and don't enforce dependencies
        for step in self.montecarlo_steps:
            card = self.montecarlo_section.cards.get(step.id)
            if card:
                card.set_sequential_waiting(False)
        self.montecarlo_section._update_dependencies(enforce=False)

        # Update tick monte carlo section - clear waiting and don't enforce dependencies
        for step in self.tick_montecarlo_steps:
            card = self.tick_montecarlo_section.cards.get(step.id)
            if card:
                card.set_sequential_waiting(False)
        self.tick_montecarlo_section._update_dependencies(enforce=False)

    def _find_card(self, step_id: str) -> Optional[StepCard]:
        card = self.data_section.get_card(step_id)
        if card:
            return card
        card = self.backtest_section.get_card(step_id)
        if card:
            return card
        card = self.montecarlo_section.get_card(step_id)
        if card:
            return card
        return self.tick_montecarlo_section.get_card(step_id)

    def _find_section(self, step_id: str) -> Optional[WorkflowSection]:
        """Find which section contains the given step."""
        if self.data_section.get_card(step_id):
            return self.data_section
        if self.backtest_section.get_card(step_id):
            return self.backtest_section
        if self.montecarlo_section.get_card(step_id):
            return self.montecarlo_section
        if self.tick_montecarlo_section.get_card(step_id):
            return self.tick_montecarlo_section
        return None

    def _set_all_buttons_enabled(self, enabled: bool):
        self.data_section.set_all_buttons_enabled(enabled)
        self.backtest_section.set_all_buttons_enabled(enabled)
        self.montecarlo_section.set_all_buttons_enabled(enabled)
        self.tick_montecarlo_section.set_all_buttons_enabled(enabled)

    def _get_scripts_folder(self) -> str:
        """Get the folder where workflow scripts are located (same folder as this GUI/exe)."""
        # When running as a PyInstaller exe, __file__ points to a temp folder
        # Use sys.executable's directory instead
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller bundle
            return os.path.dirname(sys.executable)
        else:
            # Running as normal Python script
            return os.path.dirname(os.path.abspath(__file__))

    def _get_python_executable(self) -> str:
        """Get the Python executable path for running child scripts.
        
        When running as a PyInstaller exe, sys.executable points to the exe,
        so we need to find the actual Python installation.
        """
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller bundle - need to find Python
            import shutil
            
            # Try common Python locations
            candidates = [
                # Python in PATH
                shutil.which('python'),
                shutil.which('python3'),
                # Common Windows locations
                os.path.expandvars(r'%LOCALAPPDATA%\Programs\Python\Python314\python.exe'),
                os.path.expandvars(r'%LOCALAPPDATA%\Programs\Python\Python313\python.exe'),
                os.path.expandvars(r'%LOCALAPPDATA%\Programs\Python\Python312\python.exe'),
                os.path.expandvars(r'%LOCALAPPDATA%\Programs\Python\Python311\python.exe'),
                os.path.expandvars(r'%LOCALAPPDATA%\Programs\Python\Python310\python.exe'),
                r'C:\Python314\python.exe',
                r'C:\Python313\python.exe',
                r'C:\Python312\python.exe',
                r'C:\Python311\python.exe',
                r'C:\Python310\python.exe',
            ]
            
            for candidate in candidates:
                if candidate and os.path.isfile(candidate):
                    return candidate
            
            # Last resort - hope 'python' is in PATH
            return 'python'
        else:
            # Running as normal Python script
            return sys.executable

    def _handle_confirmation_step(self, step: WorkflowStep):
        """Handle a confirmation step - show dialog and mark complete if confirmed."""
        reply = QMessageBox.information(
            self,
            step.title,
            step.confirmation_message,
            QMessageBox.Ok | QMessageBox.Cancel,
            QMessageBox.Ok
        )
        
        if reply == QMessageBox.Ok:
            # Mark step as complete
            card = self._find_card(step.id)
            if card:
                card.set_status(StepStatus.COMPLETE)
            
            # Update dependencies in the section
            section = self._find_section(step.id)
            if section:
                section.on_step_completed(step.id)
            
            # Log confirmation
            self.log_panel.append_header(step.title)
            self.log_panel.append_success("User confirmed step completion")
            self.status_label.setText("Ready")

    def _on_run_step(self, step: WorkflowStep):
        if self.current_process is not None:
            QMessageBox.warning(
                self, "Step Running",
                "Another step is currently running. Please wait for it to complete."
            )
            return

        # Handle confirmation steps differently
        if step.is_confirmation:
            self._handle_confirmation_step(step)
            return

        # Check if this is the start of a sequential execution
        sequential_steps = self._get_sequential_steps()
        first_sequential_step_id = sequential_steps[0].id if sequential_steps else None
        
        if self._is_sequential_mode() and step.id == first_sequential_step_id and not self._sequential_running:
            # Starting a sequential batch
            self._sequential_running = True
            self.log_panel.clear()
            self.log_panel.append_success("=== STARTING SEQUENTIAL EXECUTION ===")
            self.log_panel.append_line("")

        # Get settings
        settings = self.settings_panel.get_settings()

        # Build command - scripts are in the same folder as this GUI
        scripts_folder = self._get_scripts_folder()
        script_path = os.path.join(scripts_folder, step.script_name)

        if not os.path.isfile(script_path):
            QMessageBox.critical(
                self, "Script Not Found",
                f"Could not find script:\n{script_path}\n\n"
                "Make sure the workflow scripts are in the same folder as this application."
            )
            return

        args = step.build_args(settings)

        # Update UI
        self.current_step_id = step.id
        card = self._find_card(step.id)
        if card:
            card.set_status(StepStatus.RUNNING)

        self._set_all_buttons_enabled(False)
        self.status_label.setText(f"Running: {step.title}")
        self.log_panel.set_title(f"Output — {step.title}")
        self.log_panel.append_header(step.title)
        self.log_panel.append_dim(f"Script: {script_path}")
        self.log_panel.append_dim(f"Arguments: {' '.join(args)}")
        self.log_panel.append_line("")

        # Start process using subprocess
        env = os.environ.copy()
        env['PYTHONUNBUFFERED'] = '1'  # Disable output buffering
        
        # Get Python executable (handles PyInstaller frozen exe case)
        python_exe = self._get_python_executable()
        
        # Hide console window on Windows (important when running as PyInstaller exe)
        creationflags = 0
        if sys.platform == 'win32':
            creationflags = subprocess.CREATE_NO_WINDOW
        
        try:
            self.current_process = subprocess.Popen(
                [python_exe, script_path] + args,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                bufsize=1,
                universal_newlines=True,
                env=env,
                cwd=scripts_folder,
                creationflags=creationflags
            )
            
            # Start thread to read output
            self._output_thread = threading.Thread(
                target=self._read_process_output,
                daemon=True
            )
            self._output_thread.start()
            
            # Start timer to check if process finished
            self._process_timer = self.startTimer(100)  # Check every 100ms
            
        except Exception as e:
            self.log_panel.append_error(str(e))
            self._cleanup_process(False)

    def _read_process_output(self):
        """Read process output in a separate thread."""
        try:
            if self.current_process and self.current_process.stdout:
                for line in self.current_process.stdout:
                    # Use signal to update UI from main thread
                    self._output_signal.emit(line)
        except Exception:
            pass

    def _append_output(self, text: str):
        """Append output to log panel (called from main thread via signal)."""
        self.log_panel.append(text)

    def timerEvent(self, event):
        """Check if process has finished."""
        if hasattr(self, '_process_timer') and event.timerId() == self._process_timer:
            if self.current_process is not None:
                ret = self.current_process.poll()
                if ret is not None:
                    self.killTimer(self._process_timer)
                    self._cleanup_process(ret == 0)

    def _cleanup_process(self, success: bool):
        completed_step_id = self.current_step_id
        card = self._find_card(completed_step_id) if completed_step_id else None
        section = self._find_section(completed_step_id) if completed_step_id else None

        if success:
            self.log_panel.append_line("")
            self.log_panel.append_success("Step completed successfully")
            self.status_label.setText("Ready")
            if card:
                card.set_status(StepStatus.COMPLETE)
            # Update dependencies
            if section:
                section.on_step_completed(completed_step_id)
        else:
            self.log_panel.append_line("")
            self.log_panel.append_error("Step failed")
            self.status_label.setText("Step failed")
            if card:
                card.set_status(StepStatus.FAILED)
            # Stop sequential execution on failure
            if self._sequential_running:
                self._sequential_running = False
                self.log_panel.append_error("=== SEQUENTIAL EXECUTION STOPPED (step failed) ===")
                self._update_sequential_mode()

        self.current_process = None
        self.current_step_id = None
        self._set_all_buttons_enabled(True)
        self.log_panel.set_title("Output Log")
        
        # Sequential mode: trigger next step if enabled and successful
        if success and self._sequential_running and self._is_sequential_mode():
            next_step = self._get_next_sequential_step(completed_step_id)
            if next_step:
                # Brief delay before starting next step
                from PySide6.QtCore import QTimer
                QTimer.singleShot(1000, lambda: self._run_sequential_step(next_step))
            else:
                # Sequence complete
                self._sequential_running = False
                self.log_panel.append_line("")
                self.log_panel.append_success("=== SEQUENTIAL EXECUTION COMPLETE ===")
                self._update_sequential_mode()

    def _run_sequential_step(self, step: WorkflowStep):
        """Run a step as part of sequential execution."""
        # Update the card to show it's now active
        card = self._find_card(step.id)
        if card:
            card.set_sequential_waiting(False)
        
        # Trigger the run
        self._on_run_step(step)


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)

    font = QFont("Segoe UI", 11)
    app.setFont(font)

    window = WorkflowWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
