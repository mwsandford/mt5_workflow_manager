"""Build the Tick Survival Analysis HTML dashboard.

Reads joined records from the persistent SQLite DB and produces a single
self-contained HTML file. Styling mirrors the Step 7 run-level Dashboard
(GitHub-dark palette, card layout, sortable tables).

All aggregation and filtering happens client-side over an embedded JSON blob,
so the threshold slider / symbol filter / date range respond instantly.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


# ─────────────────────────────────────────────
# CSS — matches Step 7 dark theme
# ─────────────────────────────────────────────
_CSS = """
:root {
    --bg-darkest: #0d1117;
    --bg-dark: #161b22;
    --bg-mid: #1c2333;
    --bg-light: #242d3d;
    --border: #2a3444;
    --border-light: #3a4858;
    --text-primary: #e6edf3;
    --text-secondary: #8b949e;
    --text-muted: #6e7681;
    --accent: #58a6ff;
    --accent-dim: #1a3a5c;
    --status-pass: #3fb950;
    --status-fail: #f85149;
    --status-warn: #f59e0b;
    --section-tick: #ec4899;
}

* { box-sizing: border-box; }
html, body {
    margin: 0;
    padding: 0;
    background: var(--bg-darkest);
    color: var(--text-primary);
    font-family: "Segoe UI", "SF Pro Text", "Helvetica Neue", sans-serif;
    font-size: 13px;
}

.wrapper { max-width: 1500px; margin: 0 auto; padding: 24px; }

.page-header {
    display: flex; align-items: baseline; justify-content: space-between;
    padding-bottom: 16px; border-bottom: 1px solid var(--border); margin-bottom: 24px;
}
.page-title { color: var(--section-tick); font-size: 22px; font-weight: 700; letter-spacing: 0.5px; }
.page-subtitle { color: var(--text-muted); font-size: 12px; }

.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 24px; }
.card {
    background: var(--bg-dark); border: 1px solid var(--border);
    border-radius: 8px; padding: 14px 16px;
}
.card-label { font-size: 10px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
.card-value { font-size: 22px; font-weight: 700; margin-top: 6px; color: var(--text-primary); }
.card-value.accent { color: var(--section-tick); }
.card-value.pass { color: var(--status-pass); }
.card-value.fail { color: var(--status-fail); }
.card-sub { font-size: 10px; color: var(--text-muted); margin-top: 4px; }

.filters {
    background: var(--bg-dark); border: 1px solid var(--border); border-radius: 8px;
    padding: 12px 16px; margin-bottom: 24px;
    display: flex; flex-wrap: wrap; gap: 16px; align-items: center;
}
.filter-group { display: flex; flex-direction: column; gap: 4px; }
.filter-group label { font-size: 10px; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; }
.filter-group select, .filter-group input {
    background: #1a2030; border: 1px solid var(--border); border-radius: 6px;
    color: var(--text-primary); padding: 6px 10px; font-size: 12px; min-width: 120px;
}
.filter-group input[type="number"] { width: 90px; min-width: 0; }
.filter-reset {
    background: var(--bg-mid); color: var(--accent); border: 1px solid var(--border);
    border-radius: 6px; padding: 6px 14px; font-weight: 700; font-size: 11px;
    cursor: pointer; margin-left: auto; height: 32px; align-self: flex-end;
}
.filter-reset:hover { background: var(--bg-light); }

.section {
    background: var(--bg-dark); border: 1px solid var(--border); border-radius: 8px;
    padding: 16px 20px; margin-bottom: 20px;
}
.section h2 {
    margin: 0 0 4px 0;
    font-size: 11px; letter-spacing: 1px; text-transform: uppercase;
    color: var(--section-tick); border-left: 3px solid var(--section-tick);
    padding-left: 10px;
}
.section .section-desc { color: var(--text-secondary); font-size: 12px; margin: 2px 0 14px 13px; }

table { width: 100%; border-collapse: collapse; font-size: 12px; }
th, td { padding: 7px 10px; text-align: left; border-bottom: 1px solid var(--border); }
th {
    font-size: 10px; letter-spacing: 0.5px; text-transform: uppercase;
    color: var(--text-secondary); font-weight: 600; background: var(--bg-mid);
    cursor: pointer; user-select: none; position: sticky; top: 0;
}
th:hover { color: var(--text-primary); }
th.sorted-asc::after { content: " ▲"; color: var(--accent); }
th.sorted-desc::after { content: " ▼"; color: var(--accent); }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
tbody tr:hover { background: var(--bg-mid); }

.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 700; letter-spacing: 0.3px; }
.badge.pass { color: var(--status-pass); background: rgba(63,185,80,0.12); }
.badge.fail { color: var(--status-fail); background: rgba(248,81,73,0.12); }
.badge.low { color: var(--status-warn); background: rgba(245,158,11,0.12); }
.badge.neutral { color: var(--text-secondary); background: rgba(110,118,129,0.1); }

.bar-cell { display: flex; align-items: center; gap: 8px; }
.bar-track { flex: 1; height: 6px; background: var(--bg-light); border-radius: 3px; overflow: hidden; min-width: 60px; }
.bar-fill { height: 100%; background: var(--status-pass); }
.bar-fill.mid { background: var(--status-warn); }
.bar-fill.low { background: var(--status-fail); }

details.strategy { background: var(--bg-mid); border: 1px solid var(--border); border-radius: 6px; margin-top: 6px; }
details.strategy > summary {
    cursor: pointer; padding: 8px 12px; font-weight: 600;
    display: flex; gap: 14px; align-items: center; list-style: none;
}
details.strategy > summary::-webkit-details-marker { display: none; }
details.strategy > summary::before { content: "▸"; color: var(--text-muted); margin-right: 4px; transition: transform 0.1s; }
details.strategy[open] > summary::before { content: "▾"; }
details.strategy pre {
    background: var(--bg-darkest); color: var(--text-secondary); padding: 12px;
    border-top: 1px solid var(--border); font-size: 11px; overflow-x: auto; margin: 0;
}
.strat-history { padding: 8px 12px; border-top: 1px solid var(--border); }
.strat-history table { font-size: 11px; }

.tag { display: inline-block; padding: 1px 7px; margin: 1px 2px; border-radius: 4px;
    background: var(--bg-light); color: var(--text-secondary); font-size: 10px; }

.empty { color: var(--text-muted); font-style: italic; padding: 14px; text-align: center; }

.footer { color: var(--text-muted); font-size: 11px; text-align: center; padding: 24px 0 8px 0; }
"""


# ─────────────────────────────────────────────
# Client-side JS — embedded analytics
# ─────────────────────────────────────────────
_JS = r"""
(function() {
    const fmt = (v, d = 2) => (v === null || v === undefined || isNaN(v)) ? '—' : Number(v).toFixed(d);
    const fmtPct = (v, d = 1) => (v === null || v === undefined || isNaN(v)) ? '—' : Number(v).toFixed(d) + '%';
    const fmtInt = v => (v === null || v === undefined || isNaN(v)) ? '—' : Math.round(v).toString();

    const state = {
        threshold: DATA.defaultThreshold,
        symbol: '',
        timeframe: '',
        direction: '',
        dateFrom: '',
        dateTo: '',
    };

    function filtered() {
        return DATA.records.filter(r => {
            if (state.symbol && r.Symbol !== state.symbol) return false;
            if (state.timeframe && r.Timeframe !== state.timeframe) return false;
            if (state.direction && r.Direction !== state.direction) return false;
            if (state.dateFrom && r.RunTimestamp && r.RunTimestamp < state.dateFrom) return false;
            if (state.dateTo && r.RunTimestamp && r.RunTimestamp > state.dateTo + ' 23:59:59') return false;
            return true;
        }).map(r => ({ ...r, _passed: r.MC95RetDD_Tick !== null && r.MC95RetDD_Tick >= state.threshold }));
    }

    // ─────────────────────────────────────────────
    // Aggregate by a key function
    // ─────────────────────────────────────────────
    function group(records, keyFn) {
        const m = new Map();
        for (const r of records) {
            const keys = keyFn(r);
            const arr = Array.isArray(keys) ? keys : [keys];
            for (const k of arr) {
                if (k === null || k === undefined || k === '') continue;
                if (!m.has(k)) m.set(k, []);
                m.get(k).push(r);
            }
        }
        return m;
    }

    function summarise(groupRecs) {
        const n = groupRecs.length;
        const passed = groupRecs.filter(r => r._passed).length;
        const passRate = n > 0 ? (passed / n) : 0;
        const m1s = groupRecs.map(r => r.MC95RetDD_M1).filter(v => v !== null && !isNaN(v));
        const ticks = groupRecs.map(r => r.MC95RetDD_Tick).filter(v => v !== null && !isNaN(v));
        const mean = arr => arr.length ? arr.reduce((a,b) => a+b, 0) / arr.length : null;
        // Pearson correlation between MC95_M1 and MC95_Tick within this group
        let corr = null;
        const pairs = groupRecs
            .map(r => [r.MC95RetDD_M1, r.MC95RetDD_Tick])
            .filter(p => p[0] !== null && p[1] !== null && !isNaN(p[0]) && !isNaN(p[1]));
        if (pairs.length >= 2) {
            const mx = pairs.reduce((a, p) => a + p[0], 0) / pairs.length;
            const my = pairs.reduce((a, p) => a + p[1], 0) / pairs.length;
            let num = 0, dx = 0, dy = 0;
            for (const [x, y] of pairs) {
                num += (x - mx) * (y - my);
                dx += (x - mx) * (x - mx);
                dy += (y - my) * (y - my);
            }
            corr = (dx > 0 && dy > 0) ? num / Math.sqrt(dx * dy) : null;
        }
        return { n, passed, passRate, meanM1: mean(m1s), meanTick: mean(ticks), corr };
    }

    function bar(rate) {
        const pct = (rate * 100).toFixed(1);
        const cls = rate >= 0.5 ? '' : (rate >= 0.25 ? 'mid' : 'low');
        return `<div class="bar-cell"><div class="bar-track"><div class="bar-fill ${cls}" style="width:${pct}%"></div></div><span style="min-width:50px;text-align:right">${pct}%</span></div>`;
    }

    function confidenceBadge(n) {
        if (n < 5) return ' <span class="badge low" title="Small sample — treat with caution">LOW n</span>';
        if (n < 10) return ' <span class="badge neutral" title="Moderate sample">MOD n</span>';
        return '';
    }

    // ─────────────────────────────────────────────
    // Render tables
    // ─────────────────────────────────────────────
    function renderGroupTable(tableId, records, keyFn) {
        const groups = group(records, keyFn);
        const rows = [];
        for (const [k, recs] of groups.entries()) {
            const s = summarise(recs);
            rows.push({ key: k, ...s });
        }
        rows.sort((a, b) => b.passRate - a.passRate || b.n - a.n);
        const tb = document.querySelector(`#${tableId} tbody`);
        if (!rows.length) { tb.innerHTML = '<tr><td colspan="7" class="empty">No data</td></tr>'; return; }
        tb.innerHTML = rows.map(r => `
            <tr>
                <td>${escapeHtml(String(r.key))}${confidenceBadge(r.n)}</td>
                <td class="num">${r.n}</td>
                <td class="num">${r.passed}</td>
                <td>${bar(r.passRate)}</td>
                <td class="num">${fmt(r.meanM1)}</td>
                <td class="num">${fmt(r.meanTick)}</td>
                <td class="num">${r.corr === null ? '—' : fmt(r.corr, 2)}</td>
            </tr>
        `).join('');
    }

    // ─────────────────────────────────────────────
    // Predictive ranking — |pass_rate − baseline|, Bayesian-shrunk
    // ─────────────────────────────────────────────
    function renderPredictive(records) {
        const baselineN = records.length;
        const baselinePassed = records.filter(r => r._passed).length;
        const baselineRate = baselineN > 0 ? baselinePassed / baselineN : 0;
        // Bayesian shrinkage: prior strength = 5 pseudo-observations at baseline rate
        const K = 5;
        const groupings = [
            { dim: 'Direction',      fn: r => r.Direction },
            { dim: 'EntryOrderType', fn: r => r.EntryOrderType },
            { dim: 'Style',          fn: r => (r.Style || '').split(' / ') },
            { dim: 'SLType',         fn: r => r.SLType },
            { dim: 'TrailingStop',   fn: r => r.TrailingStop ? 'Yes' : 'No' },
            { dim: 'TSActivation',   fn: r => r.TSActivation ? 'Yes' : 'No' },
            { dim: 'MoveSLToBE',     fn: r => r.MoveSLToBE ? 'Yes' : 'No' },
            { dim: 'ProfitTarget',   fn: r => r.ProfitTarget ? 'Yes' : 'No' },
            { dim: 'HasExitSignals', fn: r => r.HasExitSignals ? 'Yes' : 'No' },
            { dim: 'Symbol',         fn: r => r.Symbol },
            { dim: 'Timeframe',      fn: r => r.Timeframe },
            { dim: 'Indicator',      fn: r => r.Indicators || [] },
        ];
        const rows = [];
        for (const g of groupings) {
            const m = group(records, g.fn);
            for (const [k, recs] of m.entries()) {
                if (recs.length < 3) continue;  // Too few to be meaningful
                const s = summarise(recs);
                const shrunkRate = (s.passed + K * baselineRate) / (s.n + K);
                const lift = shrunkRate - baselineRate;
                rows.push({
                    dimension: g.dim, value: k, n: s.n, passed: s.passed,
                    passRate: s.passRate, shrunkRate, lift,
                });
            }
        }
        rows.sort((a, b) => Math.abs(b.lift) - Math.abs(a.lift));
        const tb = document.querySelector('#pred-table tbody');
        if (!rows.length) { tb.innerHTML = '<tr><td colspan="7" class="empty">Need at least 3 strategies in a group to rank</td></tr>'; return; }
        tb.innerHTML = rows.slice(0, 30).map(r => {
            const liftPct = (r.lift * 100);
            const liftCls = liftPct > 5 ? 'pass' : (liftPct < -5 ? 'fail' : 'neutral');
            const liftStr = (liftPct >= 0 ? '+' : '') + liftPct.toFixed(1) + '%';
            return `
                <tr>
                    <td>${escapeHtml(r.dimension)}</td>
                    <td>${escapeHtml(String(r.value))}${confidenceBadge(r.n)}</td>
                    <td class="num">${r.n}</td>
                    <td class="num">${r.passed}</td>
                    <td>${bar(r.passRate)}</td>
                    <td class="num">${fmt(r.shrunkRate * 100, 1)}%</td>
                    <td class="num"><span class="badge ${liftCls}">${liftStr}</span></td>
                </tr>
            `;
        }).join('');
    }

    // ─────────────────────────────────────────────
    // Per-strategy list (latest row per strategy; expand for history + pseudo code)
    // ─────────────────────────────────────────────
    function renderStrategies(records) {
        const byName = new Map();
        for (const r of records) {
            if (!byName.has(r.StrategyName)) byName.set(r.StrategyName, []);
            byName.get(r.StrategyName).push(r);
        }
        const entries = Array.from(byName.entries()).sort((a, b) =>
            (b[1][0].RunTimestamp || '').localeCompare(a[1][0].RunTimestamp || '')
        );
        const container = document.getElementById('strategy-list');
        if (!entries.length) { container.innerHTML = '<div class="empty">No strategies match current filters</div>'; return; }
        const html = entries.map(([name, recs]) => {
            const latest = recs[0];
            const passes = recs.filter(r => r._passed).length;
            const passStr = recs.length === 1
                ? (latest._passed ? '<span class="badge pass">PASS</span>' : '<span class="badge fail">FAIL</span>')
                : `<span class="badge ${passes > 0 ? 'pass' : 'fail'}">${passes}/${recs.length} PASS</span>`;
            const indTags = (latest.Indicators || []).map(i => `<span class="tag">${escapeHtml(i)}</span>`).join('');
            const historyRows = recs.map(r => `
                <tr>
                    <td>${escapeHtml(r.RunTimestamp || '')}</td>
                    <td>${escapeHtml(r.BatchID || '')}</td>
                    <td class="num">${fmtInt(r.Rank)}</td>
                    <td class="num">${fmt(r.MC95RetDD_M1)}</td>
                    <td class="num">${fmt(r.MC95RetDD_Tick)}</td>
                    <td>${r._passed ? '<span class="badge pass">PASS</span>' : '<span class="badge fail">FAIL</span>'}</td>
                    <td class="num">${fmt(r.NetProfit, 0)}</td>
                    <td class="num">${fmt(r.ProfitFactor, 2)}</td>
                    <td class="num">${fmt(r.Sharpe, 2)}</td>
                    <td class="num">${fmtPct(r.WinRatePct)}</td>
                </tr>
            `).join('');
            return `
                <details class="strategy">
                    <summary>
                        <span style="min-width:240px">${escapeHtml(name)}</span>
                        <span class="tag">${escapeHtml(latest.Symbol || '')}</span>
                        <span class="tag">${escapeHtml(latest.Timeframe || '')}</span>
                        <span class="tag">${escapeHtml(latest.Direction || '')}</span>
                        <span class="tag">${escapeHtml(latest.EntryOrderType || '')}</span>
                        <span class="tag">${escapeHtml(latest.Style || '')}</span>
                        <span class="tag">SL: ${escapeHtml(latest.SLType || '')}</span>
                        <span style="margin-left:auto">${passStr}</span>
                        <span style="min-width:90px;text-align:right">M1 ${fmt(latest.MC95RetDD_M1)} · Tick ${fmt(latest.MC95RetDD_Tick)}</span>
                    </summary>
                    <div class="strat-history">
                        <div style="margin-bottom:8px">Indicators: ${indTags || '<span class="tag">none</span>'}</div>
                        <table>
                            <thead><tr>
                                <th>Run</th><th>Batch</th><th class="num">Rank</th>
                                <th class="num">MC95 M1</th><th class="num">MC95 Tick</th>
                                <th>Result</th><th class="num">Net $</th>
                                <th class="num">PF</th><th class="num">Sharpe</th><th class="num">Win %</th>
                            </tr></thead>
                            <tbody>${historyRows}</tbody>
                        </table>
                    </div>
                </details>
            `;
        }).join('');
        container.innerHTML = html;
    }

    // ─────────────────────────────────────────────
    // Populate filter dropdowns + summary cards
    // ─────────────────────────────────────────────
    function unique(arr) { return Array.from(new Set(arr.filter(x => x !== null && x !== undefined && x !== ''))).sort(); }
    function fillSelect(id, values) {
        const sel = document.getElementById(id);
        const current = sel.value;
        sel.innerHTML = '<option value="">All</option>' + values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
        sel.value = values.includes(current) ? current : '';
    }
    function populateFilters() {
        fillSelect('f-symbol',    unique(DATA.records.map(r => r.Symbol)));
        fillSelect('f-timeframe', unique(DATA.records.map(r => r.Timeframe)));
        fillSelect('f-direction', unique(DATA.records.map(r => r.Direction)));
    }

    function renderCards(recs) {
        const strategies = new Set(recs.map(r => r.StrategyName));
        const passes = recs.filter(r => r._passed);
        const passRate = recs.length > 0 ? (passes.length / recs.length * 100) : 0;
        const medTick = (() => {
            const v = passes.map(r => r.MC95RetDD_Tick).filter(x => x !== null).sort((a, b) => a - b);
            if (!v.length) return null;
            const mid = Math.floor(v.length / 2);
            return v.length % 2 ? v[mid] : (v[mid-1] + v[mid]) / 2;
        })();
        const batches = new Set(recs.map(r => r.BatchID));
        document.getElementById('card-strategies').textContent = strategies.size;
        document.getElementById('card-runs').textContent = recs.length;
        document.getElementById('card-batches').textContent = batches.size;
        document.getElementById('card-passes').textContent = passes.length;
        document.getElementById('card-passrate').textContent = passRate.toFixed(1) + '%';
        document.getElementById('card-passrate').className = 'card-value ' + (passRate >= 30 ? 'pass' : 'fail');
        document.getElementById('card-medtick').textContent = medTick === null ? '—' : medTick.toFixed(2);
    }

    // ─────────────────────────────────────────────
    // Sortable table helper
    // ─────────────────────────────────────────────
    function makeSortable(table) {
        const ths = table.querySelectorAll('thead th');
        ths.forEach((th, idx) => {
            th.addEventListener('click', () => {
                const tb = table.querySelector('tbody');
                const rows = Array.from(tb.querySelectorAll('tr'));
                const asc = !th.classList.contains('sorted-asc');
                ths.forEach(t => t.classList.remove('sorted-asc', 'sorted-desc'));
                th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');
                const isNum = th.classList.contains('num');
                rows.sort((a, b) => {
                    const av = a.children[idx]?.innerText.trim() || '';
                    const bv = b.children[idx]?.innerText.trim() || '';
                    if (isNum) {
                        const an = parseFloat(av.replace(/[^0-9.\-]/g, '')) || 0;
                        const bn = parseFloat(bv.replace(/[^0-9.\-]/g, '')) || 0;
                        return asc ? an - bn : bn - an;
                    }
                    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
                });
                rows.forEach(r => tb.appendChild(r));
            });
        });
    }

    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    }

    // ─────────────────────────────────────────────
    // Main render loop
    // ─────────────────────────────────────────────
    function renderAll() {
        const recs = filtered();
        renderCards(recs);
        renderGroupTable('tbl-direction',    recs, r => r.Direction);
        renderGroupTable('tbl-entrytype',    recs, r => r.EntryOrderType);
        renderGroupTable('tbl-style',        recs, r => (r.Style || '').split(' / ').filter(Boolean));
        renderGroupTable('tbl-sltype',       recs, r => r.SLType);
        renderGroupTable('tbl-trailing',     recs, r => r.TrailingStop ? 'Trailing Stop' : 'No Trailing Stop');
        renderGroupTable('tbl-tsact',        recs, r => r.TSActivation ? 'TS Activation' : 'No TS Activation');
        renderGroupTable('tbl-besl',         recs, r => r.MoveSLToBE ? 'SL→BE' : 'No SL→BE');
        renderGroupTable('tbl-pt',           recs, r => r.ProfitTarget ? 'Has Profit Target' : 'No Profit Target');
        renderGroupTable('tbl-exitsig',      recs, r => r.HasExitSignals ? 'Has Exit Signals' : 'No Exit Signals');
        renderGroupTable('tbl-symbol',       recs, r => r.Symbol);
        renderGroupTable('tbl-timeframe',    recs, r => r.Timeframe);
        renderGroupTable('tbl-indicator',    recs, r => r.Indicators || []);
        renderPredictive(recs);
        renderStrategies(recs);
    }

    // Hook up filters
    function bind(id, key, evt) {
        document.getElementById(id).addEventListener(evt || 'change', (e) => {
            state[key] = e.target.value;
            renderAll();
        });
    }

    window.addEventListener('DOMContentLoaded', () => {
        populateFilters();
        document.getElementById('f-threshold').value = state.threshold;
        bind('f-symbol',    'symbol');
        bind('f-timeframe', 'timeframe');
        bind('f-direction', 'direction');
        bind('f-datefrom',  'dateFrom');
        bind('f-dateto',    'dateTo');
        document.getElementById('f-threshold').addEventListener('input', (e) => {
            const v = parseFloat(e.target.value);
            if (!isNaN(v)) { state.threshold = v; renderAll(); }
        });
        document.getElementById('f-reset').addEventListener('click', () => {
            state.threshold = DATA.defaultThreshold;
            state.symbol = ''; state.timeframe = ''; state.direction = '';
            state.dateFrom = ''; state.dateTo = '';
            document.getElementById('f-threshold').value = state.threshold;
            document.getElementById('f-symbol').value = '';
            document.getElementById('f-timeframe').value = '';
            document.getElementById('f-direction').value = '';
            document.getElementById('f-datefrom').value = '';
            document.getElementById('f-dateto').value = '';
            renderAll();
        });
        document.querySelectorAll('table').forEach(makeSortable);
        renderAll();
    });
})();
"""


# ─────────────────────────────────────────────
# HTML skeleton
# ─────────────────────────────────────────────
def _group_table(section_id: str, title: str, description: str, table_id: str, key_header: str) -> str:
    return f"""
    <div class="section">
        <h2>{title}</h2>
        <div class="section-desc">{description}</div>
        <table id="{table_id}">
            <thead><tr>
                <th>{key_header}</th>
                <th class="num">n</th>
                <th class="num">Pass</th>
                <th>Pass rate</th>
                <th class="num">Mean MC95 M1</th>
                <th class="num">Mean MC95 Tick</th>
                <th class="num">corr(M1, Tick)</th>
            </tr></thead>
            <tbody></tbody>
        </table>
    </div>
    """


def build_html(records: list[dict], default_threshold: float = 2.0) -> str:
    """Produce a full, self-contained HTML string."""
    data_blob = {
        "records": records,
        "defaultThreshold": default_threshold,
        "generatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
    }
    data_json = json.dumps(data_blob, default=str)
    total_records = len(records)
    total_strategies = len({r.get("StrategyName") for r in records})
    total_batches = len({r.get("BatchID") for r in records})

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Tick Survival Analysis</title>
<style>{_CSS}</style>
</head>
<body>
<div class="wrapper">
    <div class="page-header">
        <div>
            <div class="page-title">Tick Survival Analysis</div>
            <div class="page-subtitle">Persistent cross-run view · {total_strategies} strategies · {total_records} backtest results · {total_batches} batches · generated {data_blob["generatedAt"]}</div>
        </div>
    </div>

    <div class="cards">
        <div class="card"><div class="card-label">Strategies</div><div id="card-strategies" class="card-value"></div><div class="card-sub">unique in filter</div></div>
        <div class="card"><div class="card-label">Backtest runs</div><div id="card-runs" class="card-value"></div><div class="card-sub">one row per strategy × batch</div></div>
        <div class="card"><div class="card-label">Batches</div><div id="card-batches" class="card-value"></div><div class="card-sub">distinct runs captured</div></div>
        <div class="card"><div class="card-label">Passed</div><div id="card-passes" class="card-value pass"></div><div class="card-sub">MC95 Tick ≥ threshold</div></div>
        <div class="card"><div class="card-label">Overall pass rate</div><div id="card-passrate" class="card-value accent"></div><div class="card-sub">re-computed from threshold</div></div>
        <div class="card"><div class="card-label">Median MC95 Tick (passers)</div><div id="card-medtick" class="card-value accent"></div><div class="card-sub">only strategies ≥ threshold</div></div>
    </div>

    <div class="filters">
        <div class="filter-group"><label>Symbol</label><select id="f-symbol"></select></div>
        <div class="filter-group"><label>Timeframe</label><select id="f-timeframe"></select></div>
        <div class="filter-group"><label>Direction</label><select id="f-direction"></select></div>
        <div class="filter-group"><label>Run from</label><input id="f-datefrom" type="date"></div>
        <div class="filter-group"><label>Run to</label><input id="f-dateto" type="date"></div>
        <div class="filter-group"><label>Pass threshold (MC95 Tick)</label><input id="f-threshold" type="number" step="0.1" min="0"></div>
        <button id="f-reset" class="filter-reset">Reset filters</button>
    </div>

    {_group_table('s-direction', 'Survival by Direction', 'Long-only vs Short-only vs Bidirectional.', 'tbl-direction', 'Direction')}
    {_group_table('s-entrytype', 'Survival by Entry Order Type', 'Stop breakouts, Limit mean-reversion, Market entries.', 'tbl-entrytype', 'Entry Order Type')}
    {_group_table('s-style', 'Survival by Strategy Style', 'Style classification from pseudo code. Multi-style strategies appear in each.', 'tbl-style', 'Style')}
    {_group_table('s-sltype', 'Survival by Stop Loss Type', 'ATR-based vs fixed-pip stops.', 'tbl-sltype', 'SL Type')}
    {_group_table('s-trailing', 'Survival by Trailing Stop', 'Does the strategy use a trailing stop?', 'tbl-trailing', 'Trailing Stop')}
    {_group_table('s-tsact', 'Survival by TS Activation', 'Does the trailing stop have an activation level?', 'tbl-tsact', 'TS Activation')}
    {_group_table('s-besl', 'Survival by Move SL to BE', 'Does the strategy move SL to breakeven?', 'tbl-besl', 'SL→BE')}
    {_group_table('s-pt', 'Survival by Profit Target', 'Explicit profit target present?', 'tbl-pt', 'Profit Target')}
    {_group_table('s-exitsig', 'Survival by Exit Signals', 'Does the strategy have coded exit signals?', 'tbl-exitsig', 'Exit Signals')}
    {_group_table('s-symbol', 'Survival by Symbol', 'Breakdown by traded instrument.', 'tbl-symbol', 'Symbol')}
    {_group_table('s-timeframe', 'Survival by Timeframe', 'Breakdown by chart timeframe.', 'tbl-timeframe', 'Timeframe')}
    {_group_table('s-indicator', 'Survival by Indicator', 'Each indicator is counted for every strategy that uses it.', 'tbl-indicator', 'Indicator')}

    <div class="section">
        <h2>Predictive Ranking</h2>
        <div class="section-desc">Characteristics sorted by absolute deviation from the overall pass rate. Rates are Bayesian-shrunk toward the baseline (prior strength K=5) so tiny samples don't dominate.</div>
        <table id="pred-table">
            <thead><tr>
                <th>Dimension</th>
                <th>Value</th>
                <th class="num">n</th>
                <th class="num">Pass</th>
                <th>Raw pass rate</th>
                <th class="num">Shrunk rate</th>
                <th class="num">Lift vs baseline</th>
            </tr></thead>
            <tbody></tbody>
        </table>
    </div>

    <div class="section">
        <h2>Strategies</h2>
        <div class="section-desc">Click a row to see all backtest runs for that strategy and its pseudo-code characteristics.</div>
        <div id="strategy-list"></div>
    </div>

    <div class="footer">Tick Survival Analysis · MT5 Workflow Manager</div>
</div>

<script>const DATA = {data_json};</script>
<script>{_JS}</script>
</body>
</html>
"""
    return html


def write_html(records: list[dict], out_path: str | Path, default_threshold: float = 2.0) -> Path:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    html = build_html(records, default_threshold=default_threshold)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path
