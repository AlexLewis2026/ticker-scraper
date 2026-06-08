"""
Trade Accumulator — Flask web frontend
"""

import base64
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
from ocr_parser import parse_image_local, ocr_raw_text
from openpyxl import load_workbook
import db as database

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB upload limit
database.init_db()

EXCEL_PATH = Path(__file__).parent / "trade_tally_v4.xlsx"

# Auto-reset on startup if data is from a previous business day
if database.is_new_business_day():
    print("New business day detected — archiving previous data and resetting.")
    database.reset_day(EXCEL_PATH)
ALLOWED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"}


def _allowed_image(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.errorhandler(413)
def upload_too_large(e):
    return jsonify(error="Image file is too large (max 20 MB)."), 413


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

  .layout { display: flex; flex: 1; overflow: hidden; }

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
    overflow-y: auto;
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
    cursor: pointer;
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
    flex-shrink: 0;
  }
  .tab {
    padding: 10px 20px;
    cursor: pointer;
    font-size: 13px;
    color: #8b949e;
    border-bottom: 2px solid transparent;
    transition: color .15s;
    user-select: none;
    white-space: nowrap;
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
  .badge { font-size: 11px; font-weight: bold; border-radius: 4px; padding: 2px 7px; }
  .badge-outright { background: #21262d; color: #e6edf3; }
  .badge-spread   { background: #3d2b00; color: #f0b429; }
  .badge-taps      { background: #0d2b3e; color: #58a6ff; }
  .badge-cancelled { background: #1a1a1a; color: #666; text-decoration: line-through; }
  .badge-flag      { background: #490202; color: #f85149; }

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

  /* ── History tab ── */
  .history-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    gap: 12px;
  }

  .history-card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 8px;
    overflow: hidden;
    cursor: pointer;
    transition: border-color .15s, transform .1s;
  }
  .history-card:hover { border-color: #58a6ff; transform: translateY(-1px); }
  .history-card.active-import { border-color: #2ea043; }

  .history-card img {
    width: 100%;
    height: 140px;
    object-fit: cover;
    object-position: top;
    background: #000;
    display: block;
  }
  .history-card-body { padding: 10px 12px; }
  .history-card-body .filename {
    font-weight: bold;
    color: #e6edf3;
    font-size: 12px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .history-card-body .ts { color: #8b949e; font-size: 11px; margin-top: 3px; }
  .history-card-body .pills { display: flex; gap: 6px; margin-top: 8px; flex-wrap: wrap; }
  .pill {
    font-size: 11px;
    border-radius: 10px;
    padding: 2px 8px;
    background: #21262d;
    color: #8b949e;
  }
  .pill.saved { background: #1a3a1a; color: #3fb950; }
  .pill.unsaved { background: #3a1a1a; color: #f85149; }

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

  /* ── shared table styles ── */
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  thead th {
    background: #1f3864;
    color: #fff;
    padding: 7px 10px;
    text-align: left;
    white-space: nowrap;
    position: sticky;
    top: 0;
    z-index: 1;
  }
  tbody tr:nth-child(even) { background: #161b22; }
  tbody tr:nth-child(odd)  { background: #0d1117; }
  tbody tr.spread    { background: #2b2000; }
  tbody tr.taps      { background: #0d1f2e; }
  tbody tr.flag      { background: #2b0000; }
  tbody tr.cancelled { background: #1a1a1a; opacity: 0.55; text-decoration: line-through; }
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

  /* ── Lightbox ── */
  #lightbox {
    display: none;
    position: fixed;
    inset: 0;
    background: #000000cc;
    z-index: 100;
    align-items: center;
    justify-content: center;
    cursor: zoom-out;
  }
  #lightbox.open { display: flex; }
  #lightbox img { max-width: 90vw; max-height: 90vh; border-radius: 6px; }

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
  <div style="flex:1">
    <h1>📊 Trade Accumulator</h1>
    <span>Blotter screengrab → Excel workbook</span>
  </div>
  <div id="day-info" style="text-align:right;font-size:12px;color:#a8c4e0"></div>
</header>

<!-- Stale-data banner (hidden until JS detects previous-day data) -->
<div id="stale-banner" style="display:none;background:#3a1a00;border-bottom:1px solid #f0b429;
     padding:10px 24px;display:none;align-items:center;gap:16px">
  <span style="color:#f0b429;font-weight:bold">⚠ Data from a previous trading day is loaded.</span>
  <button onclick="newTradingDay()"
          style="background:#c05000;color:#fff;border:none;border-radius:5px;
                 padding:6px 14px;font-weight:bold;cursor:pointer;font-size:12px">
    Start New Trading Day
  </button>
  <span style="color:#8b7050;font-size:11px">Previous data will be archived before wiping.</span>
</div>

<!-- Lightbox -->
<div id="lightbox" onclick="closeLightbox()">
  <img id="lightbox-img" src="" alt="screenshot">
</div>

<div class="layout">

  <!-- ── LEFT PANEL ── -->
  <aside class="left">

    <div>
      <label class="field-label">New Screenshot</label>
      <div class="drop-zone" id="drop-zone"
           onclick="document.getElementById('file-input').click()">
        <div class="icon">🖼️</div>
        <p>Drag &amp; drop a screenshot here<br>or <strong>click to browse</strong></p>
      </div>
      <input type="file" id="file-input" accept="image/*" style="display:none">
      <img id="preview-img" alt="preview" title="Click to enlarge"
           onclick="openLightbox(this.src)">
    </div>

    <button class="btn btn-primary" id="parse-btn" disabled onclick="parseImage()">
      Parse Screenshot
    </button>

    <button class="btn btn-outline" id="next-btn" style="display:none" onclick="resetForNext()">
      + Parse Next Screenshot
    </button>

    <button class="btn btn-success" id="save-btn" disabled onclick="saveTrades()">
      Save to Excel
    </button>

    <div style="border-top:1px solid #30363d;padding-top:4px">
      <label class="field-label" style="margin-bottom:8px">Saved data</label>
      <button class="btn" style="margin-bottom:8px;background:#c05000;color:#fff"
              onclick="newTradingDay()">
        🗓 New Trading Day
      </button>
      <button class="btn btn-outline" style="margin-bottom:8px" onclick="loadHistory()">
        Refresh History
      </button>
      <button class="btn btn-outline" style="margin-bottom:8px" onclick="loadLog()">
        Refresh Trade Log
      </button>
      <button class="btn btn-outline" style="margin-bottom:8px" onclick="loadTally()">
        Refresh Volume Tally
      </button>
      <button class="btn btn-outline" onclick="downloadExcel()">
        Download Excel ↓
      </button>
    </div>

    <div class="status-bar" id="status">Ready — load a screenshot to begin.</div>

  </aside>

  <!-- ── RIGHT PANEL ── -->
  <main class="right">

    <div class="tabs">
      <div class="tab active" onclick="switchTab('preview')">Trade Preview</div>
      <div class="tab"        onclick="switchTab('history')">History</div>
      <div class="tab"        onclick="switchTab('log')">Trade Log</div>
      <div class="tab"        onclick="switchTab('tally')">Volume Tally</div>
    </div>

    <!-- Preview -->
    <div class="tab-content active" id="tab-preview">
      <div class="empty-state" id="preview-empty">
        <div class="icon">🔍</div>
        <div>Parse a screenshot to see extracted trades here</div>
      </div>
      <div id="trades-container" style="display:none"></div>
    </div>

    <!-- History -->
    <div class="tab-content" id="tab-history">
      <div class="empty-state" id="history-empty">
        <div class="icon">🗂️</div>
        <div>No screenshots imported yet</div>
      </div>
      <div class="history-grid" id="history-grid" style="display:none"></div>
    </div>

    <!-- Trade Log -->
    <div class="tab-content" id="tab-log">
      <div class="empty-state" id="log-empty">
        <div class="icon">📋</div>
        <div>Click "Refresh Trade Log" to load saved trades</div>
      </div>
      <div id="log-table-wrap" style="display:none"></div>
    </div>

    <!-- Volume Tally -->
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
let parsedTrades  = null;
let currentImport = null;   // import_id of the currently previewed import
let imageFile     = null;
const TAB_IDS     = ['preview', 'history', 'log', 'tally'];

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t, i) =>
    t.classList.toggle('active', TAB_IDS[i] === name));
  document.querySelectorAll('.tab-content').forEach(c =>
    c.classList.toggle('active', c.id === 'tab-' + name));
}

// ── Lightbox ──────────────────────────────────────────────────────────────────
function openLightbox(src) {
  document.getElementById('lightbox-img').src = src;
  document.getElementById('lightbox').classList.add('open');
}
function closeLightbox() {
  document.getElementById('lightbox').classList.remove('open');
}
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeLightbox(); });

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
  parsedTrades  = null;
  currentImport = null;
  document.getElementById('save-btn').disabled = true;
}

// ── Status helper ─────────────────────────────────────────────────────────────
function setStatus(msg, type='') {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status-bar' + (type ? ' ' + type : '');
}

// ── Reset for next screenshot ─────────────────────────────────────────────────
function resetForNext() {
  // Reset file selection only — leave the trade preview intact until new trades arrive
  imageFile     = null;
  currentImport = null;

  const img = document.getElementById('preview-img');
  img.src   = '';
  img.style.display = 'none';
  document.getElementById('drop-zone').style.display = 'block';
  document.getElementById('parse-btn').disabled = true;
  document.getElementById('save-btn').disabled  = true;
  document.getElementById('next-btn').style.display = 'none';
  document.getElementById('file-input').value = '';
  setStatus('Ready — load the next screenshot.', '');
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

    parsedTrades  = data.trades;
    currentImport = data.import_id;

    renderTrades(data.trades);
    document.getElementById('save-btn').disabled = false;
    const dupMsg = data.dup_count > 0 ? `, ${data.dup_count} duplicate row(s) ignored` : '';
    setStatus(`✓ ${data.raw_count} rows scanned → ${data.new_count} new → ${data.trades.length} trade(s) extracted${dupMsg}.`, 'ok');
    document.getElementById('next-btn').style.display = 'block';
    switchTab('preview');
    // Refresh history badge without switching tab
    _refreshHistoryBackground();
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
  if (!trades || !trades.length) {
    container.style.display = 'none';
    empty.style.display = 'flex';
    return;
  }
  container.innerHTML = '';
  trades.forEach(t => {
    const isFlag      = t.notes && t.notes.includes('⚠');
    const isSpread    = t.trade_type === 'SPREAD';
    const isTaps      = t.trade_type === 'TAPS';
    const isCancelled = t.trade_type === 'CANCELLED';
    const badgeCls    = isFlag ? 'badge-flag' : isSpread ? 'badge-spread' : isTaps ? 'badge-taps' : isCancelled ? 'badge-cancelled' : 'badge-outright';
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
  if (!parsedTrades || !parsedTrades.length) { setStatus('Parse a screenshot first.', 'err'); return; }
  const btn = document.getElementById('save-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Saving…';
  setStatus('Writing to Excel…', 'info');

  try {
    const resp = await fetch('/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        trades:    parsedTrades,
        raw_count: parsedTrades.length,
        filename:  imageFile.name,
        import_id: currentImport,
      }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    setStatus(`✓ Saved! ${data.added} new trade(s) added, ${data.skipped} duplicate(s) skipped.`, 'ok');
    btn.disabled = true;
    _refreshHistoryBackground();
    loadTally();
  } catch(e) {
    setStatus('Error: ' + e.message, 'err');
    btn.disabled = false;
  }
  btn.textContent = 'Save to Excel';
}

// ── History ───────────────────────────────────────────────────────────────────
let _historyCache = [];

async function loadHistory() {
  setStatus('Loading history…', 'info');
  try {
    const resp = await fetch('/history');
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    _historyCache = data.imports;
    renderHistory(data.imports);
    setStatus(`History: ${data.imports.length} import(s).`, 'ok');
    switchTab('history');
  } catch(e) { setStatus('Error: ' + e.message, 'err'); }
}

async function _refreshHistoryBackground() {
  try {
    const resp = await fetch('/history');
    const data = await resp.json();
    if (resp.ok) { _historyCache = data.imports; renderHistory(data.imports); }
  } catch(_) {}
}

function renderHistory(imports) {
  const grid  = document.getElementById('history-grid');
  const empty = document.getElementById('history-empty');
  if (!imports.length) { grid.style.display='none'; empty.style.display='flex'; return; }

  grid.innerHTML = '';
  imports.forEach(imp => {
    const saved    = imp.added != null;
    const pill     = saved
      ? `<span class="pill saved">✓ ${imp.added} saved, ${imp.skipped} skipped</span>`
      : `<span class="pill unsaved">Not yet saved to Excel</span>`;
    const isCur    = imp.id === currentImport;

    grid.insertAdjacentHTML('beforeend', `
      <div class="history-card ${isCur ? 'active-import' : ''}"
           onclick="loadImport(${imp.id})">
        <img src="/screenshot/${imp.id}" alt="${imp.original_name}"
             onerror="this.style.display='none'"
             onclick="event.stopPropagation(); openLightbox(this.src)">
        <div class="history-card-body">
          <div class="filename" title="${imp.original_name}">${imp.original_name}</div>
          <div class="ts">${imp.imported_at}</div>
          <div class="pills">
            <span class="pill">${imp.raw_row_count} rows</span>
            <span class="pill">${imp.trade_count} trades</span>
            ${pill}
          </div>
        </div>
      </div>`);
  });
  grid.style.display  = 'grid';
  empty.style.display = 'none';
}

async function loadImport(importId) {
  setStatus('Loading past import…', 'info');
  try {
    const resp = await fetch(`/import/${importId}/trades`);
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);

    currentImport = importId;
    parsedTrades  = data.trades;

    // Show the stored screenshot in the left panel
    const img = document.getElementById('preview-img');
    img.src   = `/screenshot/${importId}`;
    img.style.display = 'block';
    document.getElementById('drop-zone').style.display = 'none';

    // Enable Save button (lets user re-save if needed)
    document.getElementById('save-btn').disabled = false;

    renderTrades(data.trades);
    setStatus(`Loaded import #${importId} — ${data.trades.length} trade(s).`, 'ok');
    switchTab('preview');

    // Highlight active card
    document.querySelectorAll('.history-card').forEach(c =>
      c.classList.toggle('active-import', parseInt(c.getAttribute('onclick').match(/\d+/)[0]) === importId));
  } catch(e) { setStatus('Error: ' + e.message, 'err'); }
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
        const cls = (r[2]&&r[2].includes('⚠')) ? 'flag' : r[1]==='SPREAD' ? 'spread' : r[1]==='TAPS' ? 'taps' : r[1]==='CANCELLED' ? 'cancelled' : '';
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
    const total = data.groups.reduce((s,g) => s + g.blocks.reduce((b,bl)=>b+bl.trades.length,0), 0);
    setStatus(`Volume Tally: ${data.groups.length} CC(s), ${total} trade row(s).`, 'ok');
    switchTab('tally');
  } catch(e) { setStatus('Error: ' + e.message, 'err'); }
}

function fmt(v) {
  if (v == null) return '';
  const n = parseFloat(v);
  if (isNaN(n)) return v;
  return n % 1 === 0 ? n.toLocaleString() : parseFloat(n.toPrecision(6)).toString();
}

function renderTally(groups) {
  const container = document.getElementById('tally-container');
  const empty     = document.getElementById('tally-empty');
  if (!groups || !groups.length) { container.style.display='none'; empty.style.display='flex'; return; }

  container.innerHTML = '';
  groups.forEach(g => {
    container.insertAdjacentHTML('beforeend', `<div class="tally-cc-banner">⬛ ${g.cc}</div>`);
    g.blocks.forEach(bl => {
      const isSpread = bl.kind !== 'OUTRIGHT';
      const rows = bl.trades.map((t, i) => {
        const rc = isSpread ? 'spread-row' : i%2===0 ? 'trade-even' : 'trade-odd';
        return `<tr class="${rc}">
          <td></td><td>${t.ts}</td>
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
            <thead><tr>
              <th style="text-align:left;width:30%">Strip / Spread</th>
              <th>Timestamp</th><th>Qty</th><th>Price</th>
              <th>Cumul. Vol</th><th>VWAP</th>
            </tr></thead>
            <tbody>
              ${rows}
              <tr class="summary-row">
                <td class="summary-label" colspan="2">► Cumul. Vol / VWAP</td>
                <td>${fmt(bl.total_qty)}</td><td></td>
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

// ── Day status ────────────────────────────────────────────────────────────────
async function checkDayStatus() {
  try {
    const resp = await fetch('/day-status');
    const d    = await resp.json();
    const info = document.getElementById('day-info');
    info.textContent = `Today: ${d.today}` + (d.last_import_date ? `  |  Last data: ${d.last_import_date}` : '');
    const banner = document.getElementById('stale-banner');
    banner.style.display = d.new_business_day ? 'flex' : 'none';
  } catch(_) {}
}

async function newTradingDay() {
  if (!confirm('Archive and wipe all current data to start a fresh trading day?')) return;
  setStatus('Archiving and resetting…', 'info');
  try {
    const resp = await fetch('/reset', { method: 'POST' });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error);

    // Clear UI state
    parsedTrades = null; currentImport = null; imageFile = null;
    document.getElementById('preview-img').style.display = 'none';
    document.getElementById('drop-zone').style.display   = 'block';
    document.getElementById('parse-btn').disabled = true;
    document.getElementById('save-btn').disabled  = true;
    document.getElementById('next-btn').style.display = 'none';
    document.getElementById('trades-container').innerHTML = '';
    document.getElementById('preview-empty').style.display = 'flex';
    document.getElementById('log-table-wrap').style.display  = 'none';
    document.getElementById('log-empty').style.display       = 'flex';
    document.getElementById('tally-container').innerHTML = '';
    document.getElementById('tally-empty').style.display = 'flex';

    _historyCache = [];
    renderHistory([]);
    document.getElementById('stale-banner').style.display = 'none';
    checkDayStatus();
    setStatus('✓ New trading day started. Previous data archived.', 'ok');
    switchTab('preview');
  } catch(e) {
    setStatus('Reset failed: ' + e.message, 'err');
  }
}

// ── Auto-load history on page load ───────────────────────────────────────────
window.addEventListener('load', () => {
  checkDayStatus();
  _refreshHistoryBackground().then(() => {
    if (_historyCache.length) { renderHistory(_historyCache); switchTab('history'); }
  });
});
</script>
</body>
</html>
"""


# ── Tally computation ──────────────────────────────────────────────────────────

def _compute_tally():
    all_trades = database.get_all_trades()
    if not all_trades:
        return []

    buckets = defaultdict(list)
    for t in all_trades:
        cc   = t.get("cc", "")
        qty  = float(t.get("qty") or 0)
        legs = t.get("legs", [])
        tt   = t.get("trade_type", "")
        ts   = t.get("timestamp", "")
        sp   = t.get("spread_price")

        if tt == "CANCELLED":
            continue
        if tt in ("OUTRIGHT", "TAPS") and legs:
            l = legs[0]
            buckets[(cc, l["strip"], tt)].append(
                {"ts": ts, "qty": qty, "price": float(l["price"])})
        elif tt == "SPREAD" and sp is not None:
            # One entry per spread trade using the differential price.
            # Individual legs are not shown separately to avoid duplication.
            diff_label = " / ".join(l["strip"] for l in legs)
            buckets[(cc, diff_label, "SPREAD")].append(
                {"ts": ts, "qty": qty, "price": float(sp)})

    for k in buckets:
        buckets[k].sort(key=lambda x: x["ts"])

    KIND_ORDER = {"OUTRIGHT": 0, "TAPS": 1, "SPREAD": 2}
    KIND_LABEL = {"OUTRIGHT": "Outright", "TAPS": "TAPS / MOC", "SPREAD": "Spread"}
    sorted_keys = sorted(buckets, key=lambda k: (k[0], KIND_ORDER.get(k[2], 9), k[1]))

    cc_groups = {}
    for key in sorted_keys:
        cc_key, strip_label, kind = key
        recs = buckets[key]
        cumvol = vwap_num = 0.0
        trade_rows = []
        for rec in recs:
            q, p = rec["qty"], rec["price"]
            cumvol   += q
            vwap_num += q * p
            vwap = round(vwap_num / cumvol, 6) if cumvol else None
            trade_rows.append({"ts": rec["ts"], "qty": q, "price": p,
                                "cumvol": cumvol, "vwap": vwap})
        final_vwap = round(vwap_num / cumvol, 6) if cumvol else None
        cc_groups.setdefault(cc_key, []).append({
            "label":      f"{strip_label}  [{KIND_LABEL[kind]}]",
            "kind":       kind,
            "trades":     trade_rows,
            "total_qty":  cumvol,
            "final_vwap": final_vwap,
        })

    return [{"cc": cc, "blocks": blocks} for cc, blocks in cc_groups.items()]


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/parse", methods=["POST"])
def parse():
    image = request.files.get("image")
    if not image or not image.filename:
        return jsonify(error="No image uploaded."), 400
    if not _allowed_image(image.filename):
        return jsonify(error=f"Unsupported file type. Allowed: {', '.join(ALLOWED_EXTENSIONS)}"), 400

    suffix = Path(image.filename).suffix.lower() or ".png"
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            image.save(tmp.name)
            tmp_path = tmp.name

        all_rows            = parse_image_local(tmp_path)
        raw_rows, dup_count = database.filter_new_rows(all_rows)
        trades              = group_rows_into_trades(raw_rows)
        import_id           = database.save_import(image.filename, tmp_path, raw_rows, trades)
        return jsonify(raw_count=len(all_rows), new_count=len(raw_rows),
                       dup_count=dup_count, trades=trades, import_id=import_id)
    except (FileNotFoundError, ValueError) as e:
        return jsonify(error=str(e)), 400
    except Exception as e:
        return jsonify(error=f"Parse failed: {e}"), 500
    finally:
        if tmp_path:
            Path(tmp_path).unlink(missing_ok=True)


@app.route("/save", methods=["POST"])
def save():
    data = request.get_json(silent=True)
    if not data:
        return jsonify(error="Invalid or missing JSON body."), 400

    trades    = data.get("trades", [])
    raw_count = data.get("raw_count", len(trades))
    filename  = data.get("filename", "screengrab")
    import_id = data.get("import_id")

    if not isinstance(trades, list):
        return jsonify(error="'trades' must be a list."), 400

    try:
        if not EXCEL_PATH.exists():
            build_fresh_workbook(str(EXCEL_PATH))
        added, skipped = append_trades(str(EXCEL_PATH), trades, raw_count, filename)
        if import_id:
            database.update_save_counts(import_id, added, skipped)
        return jsonify(added=added, skipped=skipped)
    except PermissionError:
        return jsonify(error="Excel file is open in another program. Close it and try again."), 409
    except Exception as e:
        return jsonify(error=f"Save failed: {e}"), 500


@app.route("/parse-debug", methods=["POST"])
def parse_debug():
    """Return raw Tesseract OCR text for a given image — useful for diagnosing parse failures."""
    image = request.files.get("image")
    if not image or not image.filename:
        return jsonify(error="No image uploaded."), 400
    if not _allowed_image(image.filename):
        return jsonify(error="Unsupported file type."), 400

    suffix = Path(image.filename).suffix.lower() or ".png"
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            image.save(tmp.name)
            tmp_path = tmp.name
        raw = ocr_raw_text(tmp_path)
        return jsonify(raw_text=raw)
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        if tmp_path:
            Path(tmp_path).unlink(missing_ok=True)


@app.route("/history")
def history():
    try:
        return jsonify(imports=database.list_imports())
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/import/<int:import_id>/trades")
def import_trades(import_id):
    try:
        trades = database.get_import_trades(import_id)
        return jsonify(trades=trades)
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/screenshot/<int:import_id>")
def screenshot(import_id):
    path = database.get_screenshot_path(import_id)
    if not path or not path.exists():
        return "Not found", 404
    return send_file(str(path))


@app.route("/log")
def log():
    try:
        trades = database.get_all_trades()
        rows = []
        for t in trades:
            legs = t.get("legs", [])
            def leg(i, f): return legs[i][f] if i < len(legs) else None
            rows.append([
                t.get("timestamp"), t.get("trade_type"), t.get("notes"),
                t.get("cc"), t.get("qty"), t.get("hub"), t.get("spread_price"),
                leg(0, "strip"), leg(0, "price"),
                leg(1, "strip"), leg(1, "price"),
                leg(2, "strip"), leg(2, "price"),
                t.get("source_file"),
            ])
        return jsonify(rows=rows)
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/tally")
def tally():
    try:
        return jsonify(groups=_compute_tally())
    except Exception as e:
        return jsonify(error=str(e)), 500


@app.route("/day-status")
def day_status():
    """Return today's date, last import date, and whether a reset is suggested."""
    from datetime import date
    return jsonify(
        today=str(date.today()),
        last_import_date=database.get_last_import_date(),
        new_business_day=database.is_new_business_day(),
    )


@app.route("/reset", methods=["POST"])
def reset():
    """Archive the current Excel file and wipe the DB + screenshots."""
    try:
        database.reset_day(EXCEL_PATH)
        return jsonify(ok=True)
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
