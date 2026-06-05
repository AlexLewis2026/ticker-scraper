"""
Trade Accumulator — Flask web frontend
"""

import json
import os
import tempfile
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request, send_file

# Re-use all logic from the existing accumulator
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


# ── HTML (single-file for simplicity) ─────────────────────────────────────────

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

  .layout {
    display: flex;
    flex: 1;
    gap: 0;
  }

  /* ── Left panel ── */
  .left {
    width: 340px;
    min-width: 340px;
    background: #161b22;
    border-right: 1px solid #30363d;
    padding: 20px 16px;
    display: flex;
    flex-direction: column;
    gap: 18px;
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

  input[type=text], input[type=password] {
    width: 100%;
    background: #0d1117;
    border: 1px solid #30363d;
    border-radius: 6px;
    color: #e6edf3;
    padding: 7px 10px;
    font-size: 13px;
  }
  input[type=text]:focus, input[type=password]:focus {
    outline: none;
    border-color: #2e75b6;
  }

  .drop-zone {
    border: 2px dashed #30363d;
    border-radius: 8px;
    padding: 28px 16px;
    text-align: center;
    cursor: pointer;
    transition: border-color .2s, background .2s;
    background: #0d1117;
  }
  .drop-zone:hover, .drop-zone.drag-over {
    border-color: #2e75b6;
    background: #111827;
  }
  .drop-zone .icon { font-size: 32px; }
  .drop-zone p { color: #8b949e; margin-top: 8px; font-size: 12px; }
  .drop-zone strong { color: #58a6ff; }

  #preview-img {
    display: none;
    width: 100%;
    border-radius: 6px;
    border: 1px solid #30363d;
    max-height: 200px;
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
  .btn:disabled { opacity: .45; cursor: not-allowed; }
  .btn-primary   { background: #1f6feb; color: #fff; }
  .btn-primary:hover:not(:disabled)   { background: #388bfd; }
  .btn-success   { background: #238636; color: #fff; }
  .btn-success:hover:not(:disabled)   { background: #2ea043; }
  .btn-outline   { background: transparent; color: #58a6ff;
                   border: 1px solid #30363d; }
  .btn-outline:hover:not(:disabled)   { border-color: #58a6ff; }

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
  .right {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

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

  /* Trade preview cards */
  #trades-container {
    display: flex;
    flex-direction: column;
    gap: 10px;
  }

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
  .trade-field {
    background: #0d1117;
    padding: 6px 10px;
  }
  .trade-field dt { color: #8b949e; font-size: 11px; }
  .trade-field dd { color: #e6edf3; font-weight: bold; margin-top: 2px; }

  /* Log table */
  #log-table-wrap { overflow: auto; }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
  }
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
  tbody tr.spread  { background: #2b2000; }
  tbody tr.flag    { background: #2b0000; }
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

    <button class="btn btn-primary" id="parse-btn" disabled
            onclick="parseImage()">
      Parse Screenshot
    </button>

    <button class="btn btn-success" id="save-btn" disabled
            onclick="saveTrades()">
      Save to Excel
    </button>

    <button class="btn btn-outline" onclick="loadLog()">
      Refresh Trade Log
    </button>

    <button class="btn btn-outline" onclick="downloadExcel()">
      Download Excel ↓
    </button>

    <div class="status-bar" id="status">Ready — load a screenshot to begin.</div>

  </aside>

  <!-- ── RIGHT PANEL ── -->
  <main class="right">

    <div class="tabs">
      <div class="tab active" onclick="switchTab('preview')">Trade Preview</div>
      <div class="tab" onclick="switchTab('log')">Trade Log</div>
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

  </main>

</div>

<script>
let parsedTrades = null;
let imageFile    = null;

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t, i) => {
    const ids = ['preview', 'log'];
    t.classList.toggle('active', ids[i] === name);
  });
  document.querySelectorAll('.tab-content').forEach(c => {
    c.classList.toggle('active', c.id === 'tab-' + name);
  });
}

// ── File drag & drop ──────────────────────────────────────────────────────────
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) setFile(fileInput.files[0]);
});

function setFile(f) {
  imageFile = f;
  const preview = document.getElementById('preview-img');
  const reader  = new FileReader();
  reader.onload  = e => {
    preview.src   = e.target.result;
    preview.style.display = 'block';
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
    renderTrades(data.trades, data.raw_count);
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

// ── Render trade cards ────────────────────────────────────────────────────────
function renderTrades(trades, rawCount) {
  const container = document.getElementById('trades-container');
  const empty     = document.getElementById('preview-empty');

  if (!trades.length) {
    container.style.display = 'none';
    empty.style.display     = 'flex';
    return;
  }

  container.innerHTML = '';
  trades.forEach(t => {
    const isFlag   = t.notes && t.notes.includes('⚠');
    const isSpread = t.trade_type === 'SPREAD';
    const badgeCls = isFlag ? 'badge-flag' : isSpread ? 'badge-spread' : 'badge-outright';
    const badgeLbl = isFlag ? '⚠ ' + t.trade_type : t.trade_type;

    const fields = [
      ['Timestamp', t.timestamp],
      ['CC',        t.cc],
      ['Qty',       t.qty],
      ['Hub',       t.hub || '—'],
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
          ${fields.map(([k,v]) => `
            <div class="trade-field">
              <dt>${k}</dt><dd>${v ?? '—'}</dd>
            </div>`).join('')}
        </dl>
      </div>`);
  });

  container.style.display = 'flex';
  empty.style.display     = 'none';
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
      body: JSON.stringify({
        trades:    parsedTrades,
        raw_count: parsedTrades.length,
        filename:  imageFile.name,
      }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);

    setStatus(
      `✓ Saved! ${data.added} new trade(s) added, ${data.skipped} duplicate(s) skipped.`,
      'ok'
    );
    parsedTrades = null;
    btn.disabled = true;
    loadLog();
    switchTab('log');
  } catch(e) {
    setStatus('Error: ' + e.message, 'err');
    btn.disabled = false;
  }
  btn.textContent = 'Save to Excel';
}

// ── Load trade log ────────────────────────────────────────────────────────────
async function loadLog() {
  setStatus('Loading trade log…', 'info');
  try {
    const resp = await fetch('/log');
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || resp.statusText);
    renderLog(data.rows);
    setStatus(`Trade Log: ${data.rows.length} trade(s).`, 'ok');
    switchTab('log');
  } catch(e) {
    setStatus('Error loading log: ' + e.message, 'err');
  }
}

function renderLog(rows) {
  const wrap  = document.getElementById('log-table-wrap');
  const empty = document.getElementById('log-empty');

  if (!rows.length) {
    wrap.style.display  = 'none';
    empty.style.display = 'flex';
    return;
  }

  const cols = ['Timestamp','Type','Notes','CC','Qty','Hub',
                 'Spread Px','Leg1 Strip','Leg1 Px',
                 'Leg2 Strip','Leg2 Px','Leg3 Strip','Leg3 Px','Source'];
  const numCols = new Set([4,6,8,10,12]);

  wrap.innerHTML = `
    <table>
      <thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead>
      <tbody>
        ${rows.map(r => {
          const isFlag   = r[2] && r[2].includes('⚠');
          const isSpread = r[1] === 'SPREAD';
          const cls      = isFlag ? 'flag' : isSpread ? 'spread' : '';
          return `<tr class="${cls}">${r.map((v,i) =>
            `<td class="${numCols.has(i)?'num':''}">${v??''}</td>`
          ).join('')}</tr>`;
        }).join('')}
      </tbody>
    </table>`;

  wrap.style.display  = 'block';
  empty.style.display = 'none';
}

// ── Download Excel ────────────────────────────────────────────────────────────
function downloadExcel() {
  window.location.href = '/download';
}
</script>
</body>
</html>
"""


# ── Routes ──────────────────────────────────────────────────────────────────────

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
        added, skipped = append_trades(
            str(EXCEL_PATH), trades, raw_count, filename)
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


@app.route("/download")
def download():
    if not EXCEL_PATH.exists():
        return "No Excel file yet.", 404
    return send_file(str(EXCEL_PATH),
                     as_attachment=True,
                     download_name="trade_tally.xlsx")


if __name__ == "__main__":
    print("Starting Trade Accumulator UI…")
    print("Open http://localhost:5001 in your browser")
    app.run(debug=False, port=5001)
