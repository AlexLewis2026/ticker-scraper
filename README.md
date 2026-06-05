# Trade Blotter Accumulator

Reads commodity trade blotter screenshots using local OCR (no API key or cost required) and accumulates volume and pricing data into a structured Excel workbook throughout the trading day.

---

## What it does

- Drag and drop a blotter screenshot into the browser UI
- OCR reads every visible trade row automatically
- Trades are classified as **Outright**, **Spread**, or **Flagged** (needs review)
- Results are saved into `trade_tally_v4.xlsx` with three sheets:
  - **Trade Log** — every trade, colour-coded
  - **Volume Tally** — cumulative volume and running VWAP per strip
  - **Import Log** — audit trail of every screenshot processed
- Duplicate detection: re-running the same screenshot skips already-imported trades

---

## One-time setup

### 1. Install Python

Download and install Python 3.11 or later from https://www.python.org/downloads/

During installation, tick **"Add Python to PATH"**.

### 2. Install Tesseract OCR

Tesseract reads the text from screenshots. Download the Windows installer from:

https://github.com/UB-Mannheim/tesseract/wiki

Run the installer and note where it installs (default: `C:\Users\<you>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe`).

If your path is different, open `ocr_parser.py` and update this line near the top:

```python
TESSERACT_CMD = r"C:\Users\AlexLewis\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
```

### 3. Clone the repo

```
git clone https://github.com/AlexLewis2026/ticker-scraper.git
cd ticker-scraper
```

### 4. Install Python dependencies

```
pip install -r requirements.txt
```

---

## Running the app

```
python app.py
```

Then open **http://localhost:5001** in your browser.

> On Windows, if `python` is not recognised, use the full path printed during install, e.g.:
> `C:\Users\AlexLewis\AppData\Local\Python\bin\python3.exe app.py`

---

## Daily workflow

1. Start the app: `python app.py`
2. Open http://localhost:5001
3. Drag a blotter screenshot onto the upload area (or click to browse)
4. Click **Parse Screenshot** — trades appear as cards in the Preview tab
5. Review the cards. Flagged trades (highlighted red) need manual verification.
6. Click **Save to Excel** — trades are written to `trade_tally_v4.xlsx`
7. Repeat throughout the day as new screenshots come in. Duplicates are skipped automatically.
8. Click **Download Excel** at any time to save a copy of the workbook.

---

## Keeping the repo up to date

Pull the latest changes before each session:

```
git pull
```

If dependencies have changed (you'll see changes to `requirements.txt` after a pull), reinstall them:

```
pip install -r requirements.txt
```

---

## Files

| File | Purpose |
|---|---|
| `app.py` | Flask web server and browser UI |
| `ocr_parser.py` | Tesseract OCR — reads blotter screenshots |
| `trade_accumulator_v4.py` | Core logic: trade grouping, Excel writing, deduplication |
| `requirements.txt` | Python package dependencies |
| `trade_tally_v4.xlsx` | Output workbook (created on first save, not tracked in git) |

---

## Troubleshooting

**"tesseract is not installed or not in PATH"**
Update `TESSERACT_CMD` in `ocr_parser.py` with the full path to your `tesseract.exe`.

**Trades missing from a screenshot**
The OCR parser skips rows it cannot confidently parse (e.g. rows partially hidden by a filter dropdown). Check the Preview tab — the number of parsed rows is shown in the status bar.

**"Port 5001 already in use"**
Another process is using that port. Edit the last line of `app.py` and change `port=5001` to another number (e.g. `5002`), then open that port in your browser instead.

**Duplicate trades not being skipped**
The deduplication key is `(timestamp, trade type, CC, leg 1 strip)`. Two legitimate trades with identical values on all four fields in the same second would clash. If that happens, the second trade needs to be entered manually into the Trade Log sheet.
