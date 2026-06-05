"""
Trade Accumulator — Flask web frontend
"""

import os
import tempfile
from collections import defaultdict
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request, send_file

from trade_accumulator_v4 import (
    SH_LOG,
    append_trades,
    build_fresh_workbook,
    group_rows_into_trades,
)
from ocr_parser import parse_image_local
from openpyxl import load_workbook

app = Flask(__name__)

EXCEL_PATH = Path(__file__).parent / "trade_tally_v4.xlsx"


# ── HTML ───────────────────────────────────────────────────────────────────────

HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Trade Accumulator</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: Arial, sans-serif;
    font-size: 13px;
    background: #0d1117;
    color: #e6edf3;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }

  header {
    background: #1f3864;
    padding: 14px 24px;
    display: flex;
    align-items: center;
    gap: 12px;
    border-bottom: 2px solid #2e75b6;
  }
  header h1 { font-size: 17px; font-weight: bold; color: #fff; }
  header span { font-size: 12px; color: #a8c4e0; }

  .layout { display: flex; flex: 1; }

  /* ── Left panel ── */
  .left {
    width: 300px;
    min-width: 300px;
    background: #161b22;
    border-right: 1px solid #30363d;
    padding: 20px 16px;
    display: flex;
    flex-direction: column;
    gap: 14px;
  }

  label.field-label {
    display: block;
    font-size: 11px;
    font-weight: bold;
    color: #8b949e;
    text-transform: uppercase;
    letter-spacing: .05em;
    margin-bottom: 5px;
  }

  .drop-zone {
    border: 2px dashed #30363d;
    border-radius: 8px;
    padding: 24px 12px;
    text-align: center;
    cursor: pointer;
    transition: border-color .2s, background .2s;
    background: #0d1117;
  }
  .drop-zone:hover, .drop-zone.drag-over {
    border-color: #2e75b6;
    background: #111827;
  }
  .drop-zone .icon { font-size: 28px; }
  .drop-zone p { color: #8b949e; margin-top: 6px; font-size: 12px; }
  .drop-zone strong { color: #58a6ff; }

  #preview-img {
    display: none;
    width: 100%;
    border-radius: 6px;
    border: 1px solid #30363d;
    max-height: 180px;
    object-fit: contain;
    background: #000;
  }

  .btn {
    display: block;
    width: 100%;
    padding: 9px;
    border: none;
    border-radius: 6px;
    font-size: 13px;
    font-weight: bold;
    cursor: pointer;
    transition: opacity .15s;
  }
  .btn:disabled { opacity: .4; cursor: not-allowed; }
  .btn-primary { background: #1f6feb; color: #fff; }
  .btn-primary:hover:not(:disabled) { background: #388bfd; }
  .btn-success { background: #238636; color: #fff; }
  .btn-success:hover:not(:disabled) { background: #2ea043; }
  .btn-outline { background: transparent; color: #58a6ff; border: 1px solid #30363d; }
  .btn-outline:hover:not(:disabled) { border-color: #58a6ff; }

  .status-bar {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 6px;
    padding: 8px 10px;
    font-size: 12px;
    color: #8b949e;
    min-height: 36px;
    white-space: pre-wrap;
  }
  .status-bar.ok   { color: #3fb950; border-color: #238636; }
  .status-bar.err  { color: #f85149; border-color: #da3633; }
  .status-bar.info { color: #79c0ff; border-color: #1f6feb; }

  /* ── Right panel ── */
  .right { flex: 1; display: flex; flex-direction: column; overflow: hidden; }

  .tabs {
    display: flex;
    border-bottom: 1px solid #30363d;
    background: #161b22;
  }
  .tab {
    padding: 10px 20px;
    cursor: pointer;
    font-size: 13px;
    color: #8b949e;
    border-bottom: 2px solid transparent;
    transition: color .15s;
    user-select: none;
  }
  .tab:hover { color: #e6edf3; }
  .tab.active { color: #58a6ff; border-bottom-color: #1f6feb; }

  .tab-content {
    display: none;
    flex: 1;
    overflow: auto;
    padding: 16px;
  }
  .tab-content.active { display: flex; flex-direction: column; }

  /* ── Trade preview cards ── */
  #trades-container { display: flex; flex-direction: column; gap: 10px; }

  .trade-card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 8px;
    overflow: hidden;
  }
  .trade-card-header {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 12px;
    border-bottom: 1px solid #30363d;
  }
  .badge {
    font-size: 11px;
    font-weight: bold;
    border-radius: 4px;
    padding: 2px 7px;
  }
  .badge-outright { background: #21262d; color: #e6edf3; }
  .badge-spread   { background: #3d2b00; color: #f0b429; }
  .badge-flag     { background: #490202; color: #f85149; }

  .trade-card-body {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
    gap: 1px;
    background: #30363d;
    font-size: 12px;
  }
  .trade-field { background: #0d1117; padding: 6px 10px; }
  .trade-field dt { color: #8b949e; font-size: 11px; }
  .trade-field dd { color: #e6edf3; font-weight: bold; margin-top: 2px; }

  /* ── Trade log table ── */
  #log-table-wrap { overflow: auto; }

  /* ── Volume Tally ── */
  #tally-container { display: flex; flex-direction: column; gap: 0; }

  .tally-cc-banner {
    background: #1f3864;
    color: #fff;
    font-weight: bold;
    font-size: 13px;
    padding: 7px 12px;
    margin-top: 16px;
    border-radius: 4px 4px 0 0;
    letter-spacing: .04em;
  }
  .tally-cc-banner:first-child { margin-top: 0; }

  .tally-block { margin-bottom: 2px; }

  .tally-block-header {
    background: #d6e4f0;
    color: #1f3864;
    font-weight: bold;
    font-size: 12px;
    padding: 5px 12px;
  }

  .tally-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    font-family: monospace;
  }
  .tally-table thead th {
    background: #21262d;
    color: #8b949e;
    padding: 4px 10px;
    text-align: right;
    font-size: 11px;
    font-weight: normal;
    white-space: nowrap;
  }
  .tally-table thead th:first-child { text-align: left; }

  .tally-table tbody tr.trade-even { background: #ebf3fb22; }
  .tally-table tbody tr.trade-odd  { background: #0d1117; }
  .tally-table tbody tr.spread-row { background: #fff2cc18; }
  .tally-table tbody tr.summary-row {
    background: #e2efda22;
    font-weight: bold;
    border-top: 1px solid #30363d;
  }
  .tally-table td {
    padding: 4px 10px;
    white-space: nowrap;
    text-align: right;
    color: #e6edf3;
    border-bottom: 1px solid #1c2128;
  }
  .tally-table td:first-child { text-align: left; color: #8b949e; }
  .tally-table td.highlight { color: #58a6ff; }
  .tally-table td.summary-label { color: #3fb950; font-family: Arial, sans-serif; }

  /* shared */
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  thead th {
    background: #1f3864;
    color: #fff;
    padding: 7px 10px;
    text-align: left;
    white-space: nowrap;
    position: sticky;
    top: 0;
  }
  tbody tr:nth-child(even) { background: #161b22; }
  tbody tr:nth-child(odd)  { background: #0d1117; }
  tbody tr.spread { background: #2b2000; }
  tbody tr.flag   { background: #2b0000; }
  tbody td { padding: 5px 10px; white-space: nowrap; }
  .num { text-align: right; font-family: monospace; }

  .empty-state {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    color: #484f58;
    gap: 10px;
  }
  .empty-state .icon { font-size: 48px; }

  .spinner {
    display: inline-block;
    width: 14px; height: 14px;
    border: 2px solid #30363d;
    border-top-color: #58a6ff;
    border-radius: 50%;
    animation: spin .7s linear infinite;
    vertical-align: middle;
    margin-right: 6px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>

<header>
  <div>
    <h1>📊 Trade Accumulator</h1>
    <span>Blotter screengrab → Excel workbook</span>
  </div>
</header>

<div class="layout">

  <!-- ── LEFT PANEL ── -->
  <aside class="left">

    <div>
      <label class="field-label">Screenshot</label>
      <div class="drop-zone" id="drop-zone"
           onclick="document.getElementById('file-input').click()">
        <div class="icon">🖼️</div>
        <p>Drag &amp; drop a screenshot here<br>or <strong>click to browse</strong></p>
      </div>
      <input type="file" id="file-input" accept="image/*" style="display:none">
      <img id="preview-img" alt="preview">
    </div>

    <button class="btn btn-primary" id="parse-btn" disabled onclick="parseImage()">
      Parse Screenshot
    </button>

    <button class="btn btn-success" id="save-btn" disabled onclick="saveTrades()">
      Save to Excel
    </button>

    <button class="btn btn-outline" onclick="loadLog()">
      Refresh Trade Log
    </button>

    <button class="btn btn-outline" onclick="loadTally()">
      Refresh Volume Tally
    </button>

    <button class="btn btn-outline" onclick="downloadExcel()">
      Download Excel ↓
    </button>

    <div class="status-bar" id="status">Ready — load a screenshot to begin.</div>

  </aside>

  <!-- ── RIGHT PANEL ── -->
  <main class="right">

    <div class="tabs">
      <div class="tab active"  onclick="switchTab('preview')">Trade Preview</div>
      <div class="tab"         onclick="switchTab('log')">Trade Log</div>
      <div class="tab"         onclick="switchTab('tally')">Volume Tally</div>
    </div>

    <!-- Preview tab -->
    <div class="tab-content active" id="tab-preview">
      <div class="empty-state" id="preview-empty">
        <div class="icon">🔍</div>
        <div>Parse a screenshot to see extracted trades here</div>
      </div>
      <div id="trades-container" style="display:none"></div>
    </div>

    <!-- Log tab -->
    <div class="tab-content" id="tab-log">
      <div class="empty-state" id="log-empty">
        <div class="icon">📋</div>
        <div>Click "Refresh Trade Log" to load saved trades</div>
      </div>
      <div id="log-table-wrap" style="display:none"></div>
    </div>

    <!-- Tally tab -->
    <div class="tab-content" id="tab-tally">
      <div class="empty-state" id="tally-empty">
        <div class="icon">📈</div>
        <div>Click "Refresh Volume Tally" to load cumulative data</div>
      </div>
      <div id="tally-container" style="display:none"></div>
    </div>

  </main>
</div>

<script>
let parsedTrades = null;
let imageFile    = null;
const TAB_IDS    = ['preview', 'log', 'tally'];

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t, i) => {
    t.classList.toggle('active', TAB_IDS[i] === name);
  });
  document.querySelectorAll('.tab-content').forEach(c => {
    c.classList.toggle('active', c.id === 'tab-' + name);
  });
}

// ── File drag & drop ──────────────────────────────────────────────────────────
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) setFile(fileInput.files[0]); });

function setFile(f) {
  imageFile = f;
  const reader = new FileReader();
  reader.onload = e => {
    const img = document.getElementById('preview-img');
    img.src = e.target.result;
    img.style.display = 'block';
    dropZone.style.display = 'none';
  };
  reader.readAsDataURL(f);
  document.getElementById('parse-btn').disabled = false;
  setStatus('Image loaded: ' + f.name, 'info');
  parsedTrades = null;
  document.getElementById('save-btn').disabled = true;
}

// ── Status helper ─────────────────────────────────────────────────────────────
function setStatus(msg, type='') {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status-bar' + (type ? ' ' + type : '');
}

// ── Parse image ───────────────────────────────────────────────────────────────
async function parseImage() {
  if (!imageFile) { setStatus('Select a screenshot first.', 'err'); return; }
  const btn = document.getElementById('parse-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Parsing…';
  setStatus('Reading screenshot with local OCR…', 'info');

  const fd = new FormData();
  fd.append('image', imageFile);

  try {
    const resp = await fetch('/parse', { method: 'POST', body: fd });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    parsedTrades = data.trades;
    renderTrades(data.trades);
    document.getElementById('save-btn').disabled = false;
    setStatus(`✓ ${data.raw_count} rows → ${data.trades.length} trade(s) extracted.`, 'ok');
    switchTab('preview');
  } catch(e) {
    setStatus('Error: ' + e.message, 'err');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Parse Screenshot';
  }
}

// ── Render trade preview cards ────────────────────────────────────────────────
function renderTrades(trades) {
  const container = document.getElementById('trades-container');
  const empty     = document.getElementById('preview-empty');
  if (!trades.length) { container.style.display='none'; empty.style.display='flex'; return; }

  container.innerHTML = '';
  trades.forEach(t => {
    const isFlag   = t.notes && t.notes.includes('⚠');
    const isSpread = t.trade_type === 'SPREAD';
    const badgeCls = isFlag ? 'badge-flag' : isSpread ? 'badge-spread' : 'badge-outright';
    const badgeLbl = isFlag ? '⚠ ' + t.trade_type : t.trade_type;
    const fields   = [
      ['Timestamp', t.timestamp], ['CC', t.cc], ['Qty', t.qty], ['Hub', t.hub || '—'],
      ...(t.spread_price != null ? [['Spread Px', t.spread_price]] : []),
      ...t.legs.map((l, i) => [`Leg ${i+1}`, `${l.strip}  @  ${l.price}`]),
    ];
    container.insertAdjacentHTML('beforeend', `
      <div class="trade-card">
        <div class="trade-card-header">
          <span class="badge ${badgeCls}">${badgeLbl}</span>
          ${t.notes ? `<span style="font-size:11px;color:#f85149">${t.notes}</span>` : ''}
        </div>
        <dl class="trade-card-body">
          ${fields.map(([k,v]) => `<div class="trade-field"><dt>${k}</dt><dd>${v??'—'}</dd></div>`).join('')}
        </dl>
      </div>`);
  });
  container.style.display = 'flex';
  empty.style.display = 'none';
}

// ── Save trades ───────────────────────────────────────────────────────────────
async function saveTrades() {
  if (!parsedTrades) { setStatus('Parse a screenshot first.', 'err'); return; }
  const btn = document.getElementById('save-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Saving…';
  setStatus('Writing to Excel…', 'info');

  try {
    const resp = await fetch('/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ trades: parsedTrades, raw_count: parsedTrades.length, filename: imageFile.name }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    setStatus(`✓ Saved! ${data.added} new trade(s) added, ${data.skipped} duplicate(s) skipped.`, 'ok');
    parsedTrades = null;
    btn.disabled = true;
    loadTally();
    switchTab('tally');
  } catch(e) {
    setStatus('Error: ' + e.message, 'err');
    btn.disabled = false;
  }
  btn.textContent = 'Save to Excel';
}

// ── Trade log ─────────────────────────────────────────────────────────────────
async function loadLog() {
  setStatus('Loading trade log…', 'info');
  try {
    const resp = await fetch('/log');
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    renderLog(data.rows);
    setStatus(`Trade Log: ${data.rows.length} trade(s).`, 'ok');
    switchTab('log');
  } catch(e) { setStatus('Error: ' + e.message, 'err'); }
}

function renderLog(rows) {
  const wrap  = document.getElementById('log-table-wrap');
  const empty = document.getElementById('log-empty');
  if (!rows.length) { wrap.style.display='none'; empty.style.display='flex'; return; }

  const cols    = ['Timestamp','Type','Notes','CC','Qty','Hub','Spread Px',
                   'Leg1 Strip','Leg1 Px','Leg2 Strip','Leg2 Px','Leg3 Strip','Leg3 Px','Source'];
  const numCols = new Set([4,6,8,10,12]);

  wrap.innerHTML = `
    <table>
      <thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r => {
        const cls = (r[2]&&r[2].includes('⚠')) ? 'flag' : r[1]==='SPREAD' ? 'spread' : '';
        return `<tr class="${cls}">${r.map((v,i)=>`<td class="${numCols.has(i)?'num':''}">${v??''}</td>`).join('')}</tr>`;
      }).join('')}</tbody>
    </table>`;
  wrap.style.display  = 'block';
  empty.style.display = 'none';
}

// ── Volume Tally ──────────────────────────────────────────────────────────────
async function loadTally() {
  setStatus('Loading volume tally…', 'info');
  try {
    const resp = await fetch('/tally');
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    renderTally(data.groups);
    const total = data.groups.reduce((s,g) => s + g.blocks.reduce((b,bl) => b + bl.trades.length, 0), 0);
    setStatus(`Volume Tally: ${data.groups.length} CC(s), ${total} trade row(s).`, 'ok');
    switchTab('tally');
  } catch(e) { setStatus('Error: ' + e.message, 'err'); }
}

function fmt(v) {
  if (v == null) return '';
  const n = parseFloat(v);
  if (isNaN(n)) return v;
  // show up to 6 sig figs, strip trailing zeros
  return n % 1 === 0 ? n.toLocaleString() : parseFloat(n.toPrecision(6)).toString();
}

function renderTally(groups) {
  const container = document.getElementById('tally-container');
  const empty     = document.getElementById('tally-empty');
  if (!groups || !groups.length) { container.style.display='none'; empty.style.display='flex'; return; }

  container.innerHTML = '';

  groups.forEach(g => {
    container.insertAdjacentHTML('beforeend',
      `<div class="tally-cc-banner">⬛ ${g.cc}</div>`);

    g.blocks.forEach(bl => {
      const isSpread = bl.kind !== 'OUTRIGHT';
      const rows = bl.trades.map((t, i) => {
        const rowCls = isSpread
          ? 'spread-row'
          : i % 2 === 0 ? 'trade-even' : 'trade-odd';
        return `<tr class="${rowCls}">
          <td></td>
          <td>${t.ts}</td>
          <td>${fmt(t.qty)}</td>
          <td class="highlight">${fmt(t.price)}</td>
          <td>${fmt(t.cumvol)}</td>
          <td class="highlight">${fmt(t.vwap)}</td>
        </tr>`;
      }).join('');

      container.insertAdjacentHTML('beforeend', `
        <div class="tally-block">
          <div class="tally-block-header">${bl.label}</div>
          <table class="tally-table">
            <thead>
              <tr>
                <th style="text-align:left;width:30%">Strip / Spread</th>
                <th>Timestamp</th><th>Qty</th><th>Price</th>
                <th>Cumul. Vol</th><th>VWAP</th>
              </tr>
            </thead>
            <tbody>
              ${rows}
              <tr class="summary-row">
                <td class="summary-label" colspan="2">► Cumul. Vol / VWAP</td>
                <td>${fmt(bl.total_qty)}</td>
                <td></td>
                <td>${fmt(bl.total_qty)}</td>
                <td class="highlight">${fmt(bl.final_vwap)}</td>
              </tr>
            </tbody>
          </table>
        </div>`);
    });
  });

  container.style.display = 'flex';
  container.style.flexDirection = 'column';
  empty.style.display = 'none';
}

// ── Download ──────────────────────────────────────────────────────────────────
function downloadExcel() { window.location.href = '/download'; }
</script>
</body>
</html>
"""


# ── Tally computation (mirrors _rebuild_tally but returns JSON) ────────────────

def _compute_tally():
    if not EXCEL_PATH.exists():
        return []

    wb = load_workbook(str(EXCEL_PATH), read_only=True, data_only=True)
    ws = wb[SH_LOG]

    all_trades = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        ts, tt, cc, qty, sp = str(row[0]), str(row[1]), str(row[3] or ""), row[4] or 0, row[6]
        legs = []
        for i in range(3):
            base = 7 + i * 2
            s, p = row[base], row[base + 1]
            if s:
                legs.append({"strip": str(s), "price": float(p) if p else 0.0})
        all_trades.append({"ts": ts, "tt": tt, "cc": cc, "qty": float(qty), "sp": sp, "legs": legs})
    wb.close()

    if not all_trades:
        return []

    buckets = defaultdict(list)
    for t in all_trades:
        cc, qty, legs, tt = t["cc"], t["qty"], t["legs"], t["tt"]
        if tt == "OUTRIGHT" and legs:
            l = legs[0]
            buckets[(cc, l["strip"], "OUTRIGHT")].append(
                {"ts": t["ts"], "qty": qty, "price": l["price"]})
        elif tt == "SPREAD":
            for l in legs:
                buckets[(cc, l["strip"], "SPREAD_LEG")].append(
                    {"ts": t["ts"], "qty": qty, "price": l["price"]})
            if t["sp"] is not None:
                diff_label = " / ".join(l["strip"] for l in legs)
                buckets[(cc, diff_label, "SPREAD_DIFF")].append(
                    {"ts": t["ts"], "qty": qty, "price": float(t["sp"])})

    for k in buckets:
        buckets[k].sort(key=lambda x: x["ts"])

    KIND_ORDER = {"OUTRIGHT": 0, "SPREAD_LEG": 1, "SPREAD_DIFF": 2}
    KIND_LABEL = {
        "OUTRIGHT":    "Outright",
        "SPREAD_LEG":  "Spread — leg price",
        "SPREAD_DIFF": "Spread — differential",
    }
    sorted_keys = sorted(buckets, key=lambda k: (k[0], KIND_ORDER.get(k[2], 9), k[1]))

    # Build output grouped by CC
    cc_groups = {}
    for key in sorted_keys:
        cc_key, strip_label, kind = key
        recs = buckets[key]

        cumvol = 0.0
        vwap_num = 0.0
        trade_rows = []
        for rec in recs:
            q, p = rec["qty"], rec["price"]
            cumvol   += q
            vwap_num += q * p
            vwap = round(vwap_num / cumvol, 6) if cumvol else None
            trade_rows.append({"ts": rec["ts"], "qty": q, "price": p,
                                "cumvol": cumvol, "vwap": vwap})

        final_vwap = round(vwap_num / cumvol, 6) if cumvol else None

        block = {
            "label":      f"{strip_label}  [{KIND_LABEL[kind]}]",
            "kind":       kind,
            "trades":     trade_rows,
            "total_qty":  cumvol,
            "final_vwap": final_vwap,
        }

        cc_groups.setdefault(cc_key, []).append(block)

    return [{"cc": cc, "blocks": blocks} for cc, blocks in cc_groups.items()]


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/parse", methods=["POST"])
def parse():
    image = request.files.get("image")
    if not image:
        return jsonify(error="No image uploaded."), 400

    suffix = Path(image.filename).suffix or ".png"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        image.save(tmp.name)
        tmp_path = tmp.name

    try:
        raw_rows = parse_image_local(tmp_path)
        trades   = group_rows_into_trades(raw_rows)
        return jsonify(raw_count=len(raw_rows), trades=trades)
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        os.unlink(tmp_path)


@app.route("/save", methods=["POST"])
def save():
    data      = request.get_json()
    trades    = data.get("trades", [])
    raw_count = data.get("raw_count", len(trades))
    filename  = data.get("filename", "screengrab")

    try:
        if not EXCEL_PATH.exists():
            build_fresh_workbook(str(EXCEL_PATH))
        added, skipped = append_trades(str(EXCEL_PATH), trades, raw_count, filename)
        return jsonify(added=added, skipped=skipped)
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/log")
def log():
    if not EXCEL_PATH.exists():
        return jsonify(rows=[])
    try:
        wb   = load_workbook(str(EXCEL_PATH), read_only=True, data_only=True)
        ws   = wb[SH_LOG]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                rows.append([str(v) if v is not None else None for v in row])
        wb.close()
        return jsonify(rows=rows)
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/tally")
def tally():
    try:
        groups = _compute_tally()
        return jsonify(groups=groups)
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/download")
def download():
    if not EXCEL_PATH.exists():
        return "No Excel file yet.", 404
    return send_file(str(EXCEL_PATH), as_attachment=True, download_name="trade_tally.xlsx")


if __name__ == "__main__":
    print("Starting Trade Accumulator UI…")
    print("Open http://localhost:5001 in your browser")
    app.run(debug=False, port=5001)
