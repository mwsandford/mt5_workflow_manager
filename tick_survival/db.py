"""SQLite schema and helpers for the Tick Survival store.

Uses stdlib sqlite3 (no SQLAlchemy dependency). Naming conventions mirror
Portfolio Tracker's SQLAlchemy models (PascalCase tables + columns,
{Entity}ID autoincrement PKs, ISO-UTC text timestamps).
"""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


SCHEMA_VERSION = 1


_SCHEMA = """
CREATE TABLE IF NOT EXISTS SchemaVersion (
    Version INTEGER PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS Strategies (
    StrategyID      INTEGER PRIMARY KEY AUTOINCREMENT,
    StrategyName    TEXT NOT NULL UNIQUE,
    Symbol          TEXT,
    Timeframe       TEXT,
    Direction       TEXT,
    EntryOrderType  TEXT,
    EntryPriceRefs  TEXT,          -- JSON array
    Style           TEXT,
    Indicators      TEXT,          -- JSON array
    SLType          TEXT,          -- ATR-based | Fixed | None | Unknown
    StopLossText    TEXT,
    ProfitTarget    INTEGER,
    TrailingStop    INTEGER,
    TSActivation    INTEGER,
    MoveSLToBE      INTEGER,
    ExitAfterBars   INTEGER,
    HasExitSignals  INTEGER,
    ExitSummary     TEXT,
    TimeFilter      TEXT,
    OrderValidBars  INTEGER,
    PseudoCodeHash  TEXT,
    PseudoCode      TEXT,
    FirstSeenAt     TEXT,
    UpdatedAt       TEXT
);

CREATE INDEX IF NOT EXISTS idx_Strategies_Symbol ON Strategies(Symbol);
CREATE INDEX IF NOT EXISTS idx_Strategies_Timeframe ON Strategies(Timeframe);

CREATE TABLE IF NOT EXISTS BacktestResults (
    ResultID            INTEGER PRIMARY KEY AUTOINCREMENT,
    StrategyID          INTEGER NOT NULL REFERENCES Strategies(StrategyID),
    BatchID             TEXT NOT NULL,
    RunTimestamp        TEXT NOT NULL,
    MC95RetDD_M1        REAL,
    Rank                INTEGER,
    CompositeScore      REAL,
    NetProfit           REAL,
    ProfitFactor        REAL,
    Sharpe              REAL,
    WinRatePct          REAL,
    Trades              INTEGER,
    RetDD               REAL,
    Recovery            REAL,
    DDDollar            REAL,
    DDPct               REAL,
    MC95RetDD_Tick      REAL,
    TickTested          INTEGER,
    TickPassed          INTEGER,
    TickThresholdUsed   REAL,
    SourceDashboardPath TEXT,
    CreatedAt           TEXT,
    UNIQUE(StrategyID, BatchID)
);

CREATE INDEX IF NOT EXISTS idx_BacktestResults_StrategyID ON BacktestResults(StrategyID);
CREATE INDEX IF NOT EXISTS idx_BacktestResults_BatchID ON BacktestResults(BatchID);
CREATE INDEX IF NOT EXISTS idx_BacktestResults_TickPassed ON BacktestResults(TickPassed);
"""


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')


def _json_or_none(value) -> str | None:
    if value is None:
        return None
    return json.dumps(value, ensure_ascii=False)


def _bool_int(value) -> int | None:
    if value is None:
        return None
    return 1 if value else 0


def connect(db_path: str | Path) -> sqlite3.Connection:
    """Open a connection with sane defaults."""
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(db_path: str | Path) -> sqlite3.Connection:
    """Open the DB and ensure schema exists."""
    conn = connect(db_path)
    conn.executescript(_SCHEMA)
    row = conn.execute("SELECT Version FROM SchemaVersion LIMIT 1").fetchone()
    if row is None:
        conn.execute("INSERT INTO SchemaVersion (Version) VALUES (?)", (SCHEMA_VERSION,))
    conn.commit()
    return conn


def upsert_strategy(conn: sqlite3.Connection, rec: dict[str, Any]) -> int:
    """Insert-or-update a Strategies row by natural key StrategyName.

    Returns the StrategyID.
    """
    now = _utc_now_iso()
    params = {
        'StrategyName':   rec['StrategyName'],
        'Symbol':         rec.get('Symbol'),
        'Timeframe':      rec.get('Timeframe'),
        'Direction':      rec.get('Direction'),
        'EntryOrderType': rec.get('EntryOrderType'),
        'EntryPriceRefs': _json_or_none(rec.get('EntryPriceRefs')),
        'Style':          rec.get('Style'),
        'Indicators':     _json_or_none(rec.get('Indicators')),
        'SLType':         rec.get('SLType'),
        'StopLossText':   rec.get('StopLossText'),
        'ProfitTarget':   _bool_int(rec.get('ProfitTarget')),
        'TrailingStop':   _bool_int(rec.get('TrailingStop')),
        'TSActivation':   _bool_int(rec.get('TSActivation')),
        'MoveSLToBE':     _bool_int(rec.get('MoveSLToBE')),
        'ExitAfterBars':  rec.get('ExitAfterBars'),
        'HasExitSignals': _bool_int(rec.get('HasExitSignals')),
        'ExitSummary':    rec.get('ExitSummary'),
        'TimeFilter':     rec.get('TimeFilter'),
        'OrderValidBars': rec.get('OrderValidBars'),
        'PseudoCodeHash': rec.get('PseudoCodeHash'),
        'PseudoCode':     rec.get('PseudoCode'),
        'UpdatedAt':      now,
    }

    sql = """
    INSERT INTO Strategies (
        StrategyName, Symbol, Timeframe, Direction, EntryOrderType, EntryPriceRefs,
        Style, Indicators, SLType, StopLossText, ProfitTarget, TrailingStop,
        TSActivation, MoveSLToBE, ExitAfterBars, HasExitSignals, ExitSummary,
        TimeFilter, OrderValidBars, PseudoCodeHash, PseudoCode,
        FirstSeenAt, UpdatedAt
    ) VALUES (
        :StrategyName, :Symbol, :Timeframe, :Direction, :EntryOrderType, :EntryPriceRefs,
        :Style, :Indicators, :SLType, :StopLossText, :ProfitTarget, :TrailingStop,
        :TSActivation, :MoveSLToBE, :ExitAfterBars, :HasExitSignals, :ExitSummary,
        :TimeFilter, :OrderValidBars, :PseudoCodeHash, :PseudoCode,
        :UpdatedAt, :UpdatedAt
    )
    ON CONFLICT(StrategyName) DO UPDATE SET
        Symbol         = excluded.Symbol,
        Timeframe      = excluded.Timeframe,
        Direction      = excluded.Direction,
        EntryOrderType = excluded.EntryOrderType,
        EntryPriceRefs = excluded.EntryPriceRefs,
        Style          = excluded.Style,
        Indicators     = excluded.Indicators,
        SLType         = excluded.SLType,
        StopLossText   = excluded.StopLossText,
        ProfitTarget   = excluded.ProfitTarget,
        TrailingStop   = excluded.TrailingStop,
        TSActivation   = excluded.TSActivation,
        MoveSLToBE     = excluded.MoveSLToBE,
        ExitAfterBars  = excluded.ExitAfterBars,
        HasExitSignals = excluded.HasExitSignals,
        ExitSummary    = excluded.ExitSummary,
        TimeFilter     = excluded.TimeFilter,
        OrderValidBars = excluded.OrderValidBars,
        PseudoCodeHash = excluded.PseudoCodeHash,
        PseudoCode     = excluded.PseudoCode,
        UpdatedAt      = excluded.UpdatedAt
    """
    conn.execute(sql, params)
    row = conn.execute(
        "SELECT StrategyID FROM Strategies WHERE StrategyName = ?",
        (rec['StrategyName'],),
    ).fetchone()
    return int(row['StrategyID'])


def upsert_backtest_result(conn: sqlite3.Connection, strategy_id: int, rec: dict[str, Any]) -> None:
    """Insert or update a BacktestResults row keyed by (StrategyID, BatchID)."""
    now = _utc_now_iso()
    params = {
        'StrategyID':          strategy_id,
        'BatchID':             rec['BatchID'],
        'RunTimestamp':        rec.get('RunTimestamp', now),
        'MC95RetDD_M1':        rec.get('MC95RetDD_M1'),
        'Rank':                rec.get('Rank'),
        'CompositeScore':      rec.get('CompositeScore'),
        'NetProfit':           rec.get('NetProfit'),
        'ProfitFactor':        rec.get('ProfitFactor'),
        'Sharpe':              rec.get('Sharpe'),
        'WinRatePct':          rec.get('WinRatePct'),
        'Trades':              rec.get('Trades'),
        'RetDD':               rec.get('RetDD'),
        'Recovery':            rec.get('Recovery'),
        'DDDollar':            rec.get('DDDollar'),
        'DDPct':               rec.get('DDPct'),
        'MC95RetDD_Tick':      rec.get('MC95RetDD_Tick'),
        'TickTested':          _bool_int(rec.get('TickTested')),
        'TickPassed':          _bool_int(rec.get('TickPassed')),
        'TickThresholdUsed':   rec.get('TickThresholdUsed'),
        'SourceDashboardPath': rec.get('SourceDashboardPath'),
        'CreatedAt':           now,
    }

    sql = """
    INSERT INTO BacktestResults (
        StrategyID, BatchID, RunTimestamp, MC95RetDD_M1, Rank, CompositeScore,
        NetProfit, ProfitFactor, Sharpe, WinRatePct, Trades, RetDD, Recovery,
        DDDollar, DDPct, MC95RetDD_Tick, TickTested, TickPassed, TickThresholdUsed,
        SourceDashboardPath, CreatedAt
    ) VALUES (
        :StrategyID, :BatchID, :RunTimestamp, :MC95RetDD_M1, :Rank, :CompositeScore,
        :NetProfit, :ProfitFactor, :Sharpe, :WinRatePct, :Trades, :RetDD, :Recovery,
        :DDDollar, :DDPct, :MC95RetDD_Tick, :TickTested, :TickPassed, :TickThresholdUsed,
        :SourceDashboardPath, :CreatedAt
    )
    ON CONFLICT(StrategyID, BatchID) DO UPDATE SET
        RunTimestamp        = excluded.RunTimestamp,
        MC95RetDD_M1        = excluded.MC95RetDD_M1,
        Rank                = excluded.Rank,
        CompositeScore      = excluded.CompositeScore,
        NetProfit           = excluded.NetProfit,
        ProfitFactor        = excluded.ProfitFactor,
        Sharpe              = excluded.Sharpe,
        WinRatePct          = excluded.WinRatePct,
        Trades              = excluded.Trades,
        RetDD               = excluded.RetDD,
        Recovery            = excluded.Recovery,
        DDDollar            = excluded.DDDollar,
        DDPct               = excluded.DDPct,
        MC95RetDD_Tick      = excluded.MC95RetDD_Tick,
        TickTested          = excluded.TickTested,
        TickPassed          = excluded.TickPassed,
        TickThresholdUsed   = excluded.TickThresholdUsed,
        SourceDashboardPath = excluded.SourceDashboardPath
    """
    conn.execute(sql, params)


def fetch_joined_records(conn: sqlite3.Connection) -> list[dict]:
    """Return Strategies joined to BacktestResults as a list of plain dicts.

    One row per BacktestResults entry — a single strategy with multiple
    runs appears multiple times. Downstream analytics handle aggregation.
    """
    sql = """
    SELECT
        s.StrategyID, s.StrategyName, s.Symbol, s.Timeframe, s.Direction,
        s.EntryOrderType, s.EntryPriceRefs, s.Style, s.Indicators, s.SLType,
        s.StopLossText, s.ProfitTarget, s.TrailingStop, s.TSActivation,
        s.MoveSLToBE, s.ExitAfterBars, s.HasExitSignals, s.ExitSummary,
        s.TimeFilter, s.OrderValidBars, s.FirstSeenAt,
        r.ResultID, r.BatchID, r.RunTimestamp, r.MC95RetDD_M1, r.Rank,
        r.CompositeScore, r.NetProfit, r.ProfitFactor, r.Sharpe, r.WinRatePct,
        r.Trades, r.RetDD, r.Recovery, r.DDDollar, r.DDPct,
        r.MC95RetDD_Tick, r.TickTested, r.TickPassed, r.TickThresholdUsed,
        r.SourceDashboardPath, r.CreatedAt
    FROM Strategies s
    JOIN BacktestResults r ON r.StrategyID = s.StrategyID
    ORDER BY r.RunTimestamp DESC, r.Rank ASC
    """
    rows = conn.execute(sql).fetchall()
    out = []
    for row in rows:
        rec = dict(row)
        # Decode JSON columns
        for k in ('EntryPriceRefs', 'Indicators'):
            if rec.get(k):
                try:
                    rec[k] = json.loads(rec[k])
                except (ValueError, TypeError):
                    rec[k] = []
            else:
                rec[k] = []
        out.append(rec)
    return out


def row_count(conn: sqlite3.Connection, table: str) -> int:
    row = conn.execute(f"SELECT COUNT(*) AS n FROM {table}").fetchone()
    return int(row['n'])
