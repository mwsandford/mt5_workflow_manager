"""Load strategy characteristics + backtest metrics from a run's Dashboard folder.

Primary source: Dashboard/strategies_data.json + Dashboard/index.html's embedded
`const DATA = {...}` block. These already contain the output of Step 7's
pseudo-code parser (strategy_codes) and full overviews.

Fallback: when strategy_codes is missing a given strategy, re-run Step 7's
parse_strategy_pseudo_code against any *.txt pseudo-code file we can find
under the run output folder.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


# ─────────────────────────────────────────────
# Step 7 parser fallback (lazy import)
# ─────────────────────────────────────────────
def _load_step7_parser():
    """Load parse_strategy_pseudo_code from Step7_Strategy_Ranking.py.

    Step 7 has heavy top-level imports (pandas, matplotlib). We only want the
    parser function, so we import it lazily and only when a fallback parse is
    actually needed.
    """
    scripts_dir = Path(__file__).resolve().parent.parent
    step7_path = scripts_dir / "Step7_Strategy_Ranking.py"
    if not step7_path.is_file():
        return None
    spec = importlib.util.spec_from_file_location("_step7_ranking", step7_path)
    if spec is None or spec.loader is None:
        return None
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
    except Exception as e:
        print(f"  WARNING: could not import Step7 for pseudo-code fallback: {e}")
        return None
    return getattr(mod, "parse_strategy_pseudo_code", None)


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
def _load_strategies_json(dashboard_folder: Path) -> dict:
    path = dashboard_folder / "strategies_data.json"
    if not path.is_file():
        raise FileNotFoundError(f"strategies_data.json not found at {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_embedded_data(dashboard_folder: Path) -> dict:
    """Extract the `const DATA = {...};` object from index.html."""
    path = dashboard_folder / "index.html"
    if not path.is_file():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()
    m = re.search(r"const\s+DATA\s*=\s*(.*?);\s*\n", html, re.DOTALL)
    if not m:
        return {}
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return {}


_SUFFIXES_TO_STRIP = (".QDM", "_pepperstone", ".a", ".m", ".M4", ".M5")


def _clean_symbol(sym: str | None) -> str | None:
    if not sym:
        return None
    s = sym.strip()
    for suf in _SUFFIXES_TO_STRIP:
        if s.endswith(suf):
            s = s[: -len(suf)]
    # Strip anything after first '.'
    if "." in s:
        s = s.split(".", 1)[0]
    if "_" in s:
        s = s.split("_", 1)[0]
    return s or None


_NAME_TF_RE = re.compile(
    r"\b(M1|M5|M15|M30|H1|H4|H6|H8|H12|D1|W1|MN1|Daily|Weekly|Monthly)\b",
    re.IGNORECASE,
)


def _timeframe_from_name(name: str, fallback: str | None = None) -> str | None:
    if fallback:
        return fallback.upper()
    m = _NAME_TF_RE.search(name or "")
    return m.group(1).upper() if m else None


def _symbol_from_name(name: str) -> str | None:
    # Expected pattern: "SQ <SYMBOL> <TF> <version>"
    tokens = (name or "").split()
    if len(tokens) >= 2 and tokens[0].upper() == "SQ":
        return tokens[1]
    return None


# Regex for the Stop Loss formula line in pseudo code.
_SL_FORMULA_RE = re.compile(r"Stop Loss\s*=\s*(.+?);", re.IGNORECASE)


def _classify_sl(raw_content: str) -> tuple[str, str | None]:
    """Return (SLType, StopLossText) for a pseudo code block.

    SLType ∈ {"ATR-based", "Fixed", "None", "Unknown"}.
    """
    if not raw_content:
        return "Unknown", None
    m = _SL_FORMULA_RE.search(raw_content)
    if not m:
        return "None", None
    text = m.group(1).strip()
    if re.search(r"\bATR\b", text, re.IGNORECASE):
        return "ATR-based", text
    if re.search(r"\b(pip|pips|point|points)\b", text, re.IGNORECASE) or re.search(r"^\s*[\d.]+\s*$", text):
        return "Fixed", text
    # Coefficient-only formulas with no ATR often resolve to fixed SL
    if re.search(r"Coef", text) and not re.search(r"ATR", text, re.IGNORECASE):
        return "Fixed", text
    return "Unknown", text


def _pseudo_code_hash(raw_content: str) -> str | None:
    if not raw_content:
        return None
    return hashlib.sha1(raw_content.encode("utf-8", errors="replace")).hexdigest()


def _find_pseudo_code_file(output_folder: Path, strategy_name: str) -> Path | None:
    """Best-effort search for a pseudo code .txt file matching `strategy_name`."""
    if not output_folder.is_dir():
        return None
    norm = strategy_name.replace("_", " ").replace(".", " ").strip()
    # Also build a version-only fingerprint like "1.1.173" to match filenames
    ver_match = re.search(r"(\d+[._]\d+[._]\d+(?:[._]\d+)?)", strategy_name)
    ver = ver_match.group(1) if ver_match else None
    for p in output_folder.rglob("*.txt"):
        fname = p.stem.replace("_", " ").replace(".", " ").strip()
        if norm in fname or fname in norm:
            return p
        if ver and ver.replace(".", "_") in p.stem.replace(".", "_"):
            return p
    return None


# ─────────────────────────────────────────────
# Batch ID
# ─────────────────────────────────────────────
def derive_batch_id(
    strategies_json: dict,
    strategies_json_path: Path,
    explicit: str | None = None,
) -> tuple[str, str]:
    """Return (batch_id, run_timestamp_iso).

    Priority:
      1. explicit CLI arg
      2. strategies_data.json's `generated_at`
      3. file mtime of strategies_data.json
      4. utc now
    """
    if explicit:
        ts = explicit
    else:
        ts = (strategies_json.get("generated_at") or "").strip() or None
        if not ts:
            try:
                mtime = datetime.fromtimestamp(
                    strategies_json_path.stat().st_mtime, tz=timezone.utc
                )
                ts = mtime.strftime("%Y-%m-%d %H:%M:%S")
            except OSError:
                ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    # Normalise into a compact batch id (keep full ts separately)
    compact = re.sub(r"[^0-9A-Za-z]+", "_", ts).strip("_")
    return compact, ts


# ─────────────────────────────────────────────
# Main load
# ─────────────────────────────────────────────
def load_run(
    dashboard_folder: str | Path,
    output_folder: str | Path | None = None,
    tick_threshold: float = 2.0,
    batch_id_override: str | None = None,
) -> dict:
    """Parse a run's Dashboard folder into strategy + result records.

    Returns a dict with:
        batch_id, run_timestamp, tick_threshold,
        records: list of {strategy: {...}, result: {...}}
    Only strategies whose ranking entry has a non-null `mc95_ret_dd_tick`
    are emitted (i.e. top-10 strategies that were tick-tested).
    """
    dashboard_folder = Path(dashboard_folder)
    if output_folder is None:
        output_folder = dashboard_folder.parent
    output_folder = Path(output_folder)

    sjson_path = dashboard_folder / "strategies_data.json"
    sjson = _load_strategies_json(dashboard_folder)
    embedded = _load_embedded_data(dashboard_folder)

    ranking = sjson.get("ranking", []) or []
    overviews = (embedded.get("overviews", {}) or {}) if embedded else {}
    strategy_codes = (embedded.get("strategy_codes", {}) or {}) if embedded else {}

    batch_id, run_timestamp = derive_batch_id(sjson, sjson_path, batch_id_override)

    step7_parser = None
    records = []

    for r in ranking:
        mc95_tick = r.get("mc95_ret_dd_tick")
        if mc95_tick is None:
            continue  # Not tick-tested — skip per spec

        name = (r.get("name") or "").strip()
        if not name:
            continue

        # Pull characteristics. Prefer strategy_codes from index.html.
        code = strategy_codes.get(name)
        raw_content = ""
        if code is None:
            # Fallback: re-parse a .txt file
            if step7_parser is None:
                step7_parser = _load_step7_parser()
            txt_file = _find_pseudo_code_file(output_folder, name)
            if txt_file and step7_parser:
                try:
                    code = step7_parser(str(txt_file)) or {}
                    with open(txt_file, "r", errors="replace") as f:
                        raw_content = f.read()
                    code["raw_content"] = raw_content
                except Exception as e:
                    print(f"  WARNING: fallback parse failed for {name}: {e}")
                    code = {}
            else:
                code = {}
        else:
            raw_content = code.get("raw_content", "") or ""

        sl_type, sl_text = _classify_sl(raw_content)

        # Symbol/Timeframe — try ranking first, then name, then embedded overview
        symbol = _clean_symbol(r.get("symbol"))
        if not symbol:
            symbol = _symbol_from_name(name)
        timeframe = _timeframe_from_name(name)
        if not timeframe:
            ov = overviews.get(name, {})
            # overview 'period' is like "H1 2012.01.19 - 2026.04.01"
            period = (ov.get("period") or "").strip()
            if period:
                tf_token = period.split()[0] if period.split() else ""
                if tf_token:
                    timeframe = tf_token.upper()

        strategy_rec = {
            "StrategyName":   name,
            "Symbol":         symbol,
            "Timeframe":      timeframe,
            "Direction":      code.get("direction") or None,
            "EntryOrderType": code.get("entry_type") or None,
            "EntryPriceRefs": code.get("entry_refs") or [],
            "Style":          code.get("style") or None,
            "Indicators":     code.get("indicators") or [],
            "SLType":         sl_type,
            "StopLossText":   sl_text,
            "ProfitTarget":   bool(code.get("profit_target")),
            "TrailingStop":   bool(code.get("trailing_stop")),
            "TSActivation":   bool(code.get("ts_activation")),
            "MoveSLToBE":     bool(code.get("move_sl_be")),
            "ExitAfterBars":  _as_int(code.get("exit_after_bars")),
            "HasExitSignals": bool(code.get("has_exit_signals")),
            "ExitSummary":    code.get("exit_summary") or None,
            "TimeFilter":     code.get("time_filter") or None,
            "OrderValidBars": _as_int(code.get("order_valid_bars")),
            "PseudoCodeHash": _pseudo_code_hash(raw_content),
            "PseudoCode":     raw_content or None,
        }

        tick_passed = mc95_tick is not None and float(mc95_tick) >= tick_threshold

        result_rec = {
            "BatchID":             batch_id,
            "RunTimestamp":        run_timestamp,
            "MC95RetDD_M1":        _as_float(r.get("mc95_ret_dd")),
            "Rank":                _as_int(r.get("rank")),
            "CompositeScore":      _as_float(r.get("score")),
            "NetProfit":           _as_float(r.get("net_profit")),
            "ProfitFactor":        _as_float(r.get("pf")),
            "Sharpe":              _as_float(r.get("sharpe")),
            "WinRatePct":          _as_float(r.get("win_rate")),
            "Trades":              _as_int(r.get("trades")),
            "RetDD":               _as_float(r.get("ret_dd")),
            "Recovery":            _as_float(r.get("recovery")),
            "DDDollar":            _as_float(r.get("dd_dollar")),
            "DDPct":               _as_float(r.get("dd_pct")),
            "MC95RetDD_Tick":      _as_float(mc95_tick),
            "TickTested":          True,
            "TickPassed":          tick_passed,
            "TickThresholdUsed":   float(tick_threshold),
            "SourceDashboardPath": str(dashboard_folder),
        }

        records.append({"strategy": strategy_rec, "result": result_rec})

    return {
        "batch_id": batch_id,
        "run_timestamp": run_timestamp,
        "tick_threshold": float(tick_threshold),
        "records": records,
    }


# ─────────────────────────────────────────────
# Value coercion
# ─────────────────────────────────────────────
def _as_float(v) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _as_int(v) -> int | None:
    if v is None or v == "":
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None
