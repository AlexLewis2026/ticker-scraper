"""
auto_capture.py  —  Automatic screengrab + parse for the Light Ends ticker.

Takes a screenshot of a named window at a regular interval and POSTs it to
the Flask parser (http://localhost:5001/parse).

Usage:
    python auto_capture.py                        # defaults below
    python auto_capture.py --window "Light Ends"  # match any window containing this text
    python auto_capture.py --interval 300         # every 5 minutes
    python auto_capture.py --interval 600 --url http://localhost:5001/parse
    python auto_capture.py --once                 # single capture then exit

Requirements (add to requirements.txt if not already present):
    pygetwindow
    requests
    Pillow  (already required)
"""

import argparse
import io
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

import requests

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")),
        logging.FileHandler("auto_capture.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ── defaults ──────────────────────────────────────────────────────────────────
DEFAULT_WINDOW   = "Light Ends"
DEFAULT_INTERVAL = 900          # seconds between captures (15 minutes)
DEFAULT_URL      = "http://localhost:5001/parse"
SNAPSHOT_DIR     = Path("snapshots")   # saved for audit trail; set to None to skip


def find_window(title_fragment: str):
    """Return the first window whose title contains title_fragment (case-insensitive)."""
    try:
        import pygetwindow as gw
    except ImportError:
        log.error("pygetwindow not installed — run: pip install pygetwindow")
        return None

    matches = [w for w in gw.getAllWindows()
               if title_fragment.lower() in w.title.lower() and w.visible]
    if not matches:
        log.warning("No visible window found matching '%s'", title_fragment)
        return None
    if len(matches) > 1:
        log.info("Multiple matches — using first: '%s'", matches[0].title)
    return matches[0]


def capture_window(win) -> bytes | None:
    """Capture the given window and return PNG bytes, or None on failure."""
    import mss
    from PIL import Image

    try:
        # Bring window to foreground so it isn't obscured
        try:
            if win.isMinimized:
                win.restore()
            win.activate()
        except Exception:
            # Fallback: use ctypes directly
            import ctypes
            hwnd = win._hWnd
            ctypes.windll.user32.ShowWindow(hwnd, 9)       # SW_RESTORE
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.4)   # allow window to fully paint

        monitor = {
            "left":   win.left,
            "top":    win.top,
            "width":  win.width,
            "height": win.height,
        }
        with mss.MSS() as sct:
            shot = sct.grab(monitor)
            img  = Image.frombytes("RGB", shot.size, shot.rgb)
            buf  = io.BytesIO()
            img.save(buf, format="PNG")
            return buf.getvalue()

    except Exception as exc:
        log.error("Capture failed: %s", exc)
        return None


def save_snapshot(png_bytes: bytes, label: str):
    """Optionally persist the PNG for audit / replay."""
    if SNAPSHOT_DIR is None:
        return
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = SNAPSHOT_DIR / f"{ts}_{label}.png"
    path.write_bytes(png_bytes)
    log.debug("Saved snapshot: %s", path)


def post_to_parser(png_bytes: bytes, url: str, source_label: str) -> bool:
    """POST image bytes to the Flask /parse endpoint. Returns True on success."""
    filename = f"{source_label}.png"
    try:
        resp = requests.post(
            url,
            files={"image": (filename, io.BytesIO(png_bytes), "image/png")},
            timeout=60,
        )
        data = resp.json()
        if resp.status_code == 200:
            new   = data.get("new_count",  "?")
            skip  = data.get("dup_count",  "?")
            total = data.get("raw_count",  "?")
            log.info("Parsed OK — new: %s  skipped: %s  raw rows: %s", new, skip, total)
            return True
        else:
            log.warning("Parser returned %s: %s", resp.status_code,
                        data.get("error", resp.text[:120]))
            return False
    except requests.exceptions.ConnectionError:
        log.error("Cannot reach Flask server at %s — is it running?", url)
        return False
    except Exception as exc:
        log.error("POST failed: %s", exc)
        return False


def run_once(window_title: str, url: str) -> bool:
    win = find_window(window_title)
    if win is None:
        return False

    log.info("Capturing '%s'", win.title)
    png = capture_window(win)
    if png is None:
        return False

    label = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_snapshot(png, label)
    return post_to_parser(png, url, label)


def replay_snapshots(url: str, snapshot_dir: Path):
    """Re-parse every PNG in snapshot_dir in chronological order."""
    if not snapshot_dir.exists():
        log.error("Snapshot directory '%s' does not exist.", snapshot_dir)
        return

    snapshots = sorted(snapshot_dir.glob("*.png"))
    if not snapshots:
        log.warning("No snapshots found in '%s'.", snapshot_dir)
        return

    log.info("Replaying %d snapshot(s) from '%s'", len(snapshots), snapshot_dir)
    ok = fail = 0
    for path in snapshots:
        log.info("  → %s", path.name)
        png = path.read_bytes()
        if post_to_parser(png, url, path.stem):
            ok += 1
        else:
            fail += 1

    log.info("Replay complete — OK: %d  Failed: %d", ok, fail)


def main():
    ap = argparse.ArgumentParser(description="Auto-capture Light Ends ticker and parse")
    ap.add_argument("--window",   default=DEFAULT_WINDOW,
                    help=f"Window title fragment to capture (default: '{DEFAULT_WINDOW}')")
    ap.add_argument("--interval", type=int, default=DEFAULT_INTERVAL,
                    help=f"Seconds between captures (default: {DEFAULT_INTERVAL})")
    ap.add_argument("--url",      default=DEFAULT_URL,
                    help=f"Flask parse endpoint (default: {DEFAULT_URL})")
    ap.add_argument("--once",     action="store_true",
                    help="Capture once and exit")
    ap.add_argument("--replay",   action="store_true",
                    help="Re-parse all saved snapshots in chronological order then exit")
    ap.add_argument("--replay-dir", default=str(SNAPSHOT_DIR),
                    help=f"Folder to replay from (default: {SNAPSHOT_DIR})")
    args = ap.parse_args()

    log.info("═" * 55)
    log.info("Auto-capture starting")
    log.info("  Window   : %s", args.window)
    log.info("  Interval : %ss", args.interval)
    log.info("  Endpoint : %s", args.url)
    log.info("  Snapshots: %s", SNAPSHOT_DIR or "disabled")
    log.info("═" * 55)

    if args.replay:
        replay_snapshots(args.url, Path(args.replay_dir))
        return

    if args.once:
        run_once(args.window, args.url)
        return

    while True:
        try:
            run_once(args.window, args.url)
        except KeyboardInterrupt:
            log.info("Stopped by user.")
            break
        except Exception as exc:
            log.error("Unexpected error: %s", exc)

        log.info("Next capture in %ds — press Ctrl+C to stop", args.interval)
        try:
            time.sleep(args.interval)
        except KeyboardInterrupt:
            log.info("Stopped by user.")
            break


if __name__ == "__main__":
    main()
