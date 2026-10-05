"""S02 independent Excel speed stand (outside the working launcher).

Compares preview routes on identical SYNTHETIC workbooks and prints, per
route, the end artifact + engine time + serve (file-read) time separately:

- H-full : existing server.excel_sheet_html (full styles, up to 2000 rows)
- H-fast : excel_speed.get_fast_html (first 50x20 viewport, no styles)
- I-png  : excel_speed.get_fast_png  (direct PIL PNG, no COM/PDF/browser)
- P-com  : Excel COM -> PDF (existing scripts/convert_excel_to_pdf.ps1)
           + pdftoppm PNG (only small book, guarded by timeout)

Screen-paint time is NOT equated with engine time: engineMs (build) and
serveMs (bytes ready for the viewer) are reported separately, and one real
viewing check (gallery over local HTTP) is done explicitly.

Usage:
    python scripts/excel_speed_stand.py [--out DIR] [--skip-com]
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

WORKTREE_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WORKTREE_ROOT))

from app.backend import excel_speed  # noqa: E402


def make_workbook(path: Path, rows: int, cols: int, sheets: int = 1) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    for s in range(sheets):
        ws = wb.active if s == 0 else wb.create_sheet(f"Лист {s + 1}")
        ws.title = f"Смета {s + 1}" if sheets > 1 else "Смета"
        ws.append([f"Колонка {c + 1}" for c in range(cols)])
        for r in range(rows):
            ws.append([f"поз. {r + 1}-{c + 1} итого {r * c + 7.5}" for c in range(cols)])
    wb.save(path)
    wb.close()


def bench_h_full(path: Path) -> dict:
    """Existing full-HTML route. End artifact: per-sheet HTML file."""
    from app.backend.server import excel_sheet_html

    t0 = time.perf_counter()
    page, info = excel_sheet_html(path, 0)
    engine_ms = (time.perf_counter() - t0) * 1000
    blob = page.encode("utf-8")
    return {
        "route": "H-full",
        "artifact": "HTML (лист 1, полные стили)",
        "engineMs": round(engine_ms, 1),
        "serveMs": 0.0,
        "bytes": len(blob),
        "info": info,
    }


def bench_h_fast(path: Path, base: Path) -> dict:
    t0 = time.perf_counter()
    res = excel_speed.get_fast_html(path, base)
    wall = (time.perf_counter() - t0) * 1000
    return {
        "route": "H-fast",
        "artifact": "HTML (первые 50 строк x 20 кол.)",
        "engineMs": round(res["engineMs"], 1),
        "serveMs": round(res["serveMs"], 2),
        "wallMs": round(wall, 1),
        "bytes": res["bytes"],
        "cacheHit": res["cacheHit"],
        "path": res["path"],
    }


def bench_i_png(path: Path, base: Path) -> dict:
    t0 = time.perf_counter()
    res = excel_speed.get_fast_png(path, base)
    wall = (time.perf_counter() - t0) * 1000
    return {
        "route": "I-png",
        "artifact": "PNG (прямой PIL, первый вид)",
        "engineMs": round(res["engineMs"], 1),
        "serveMs": round(res["serveMs"], 2),
        "wallMs": round(wall, 1),
        "bytes": res["bytes"],
        "cacheHit": res["cacheHit"],
        "path": res["path"],
    }


def bench_p_com(path: Path, out_dir: Path, timeout: int = 150) -> dict:
    """Existing COM route. End artifacts: PDF + first-page PNG."""
    script = WORKTREE_ROOT / "scripts" / "convert_excel_to_pdf.ps1"
    if not script.exists():
        return {"route": "P-com", "artifact": "PDF->PNG", "error": "script missing"}
    pdf = out_dir / (path.stem + ".pdf")
    t0 = time.perf_counter()
    try:
        proc = subprocess.run(
            [
                "powershell",
                "-STA",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-InputPath",
                str(path),
                "-OutputPath",
                str(pdf),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return {
            "route": "P-com",
            "artifact": "PDF->PNG",
            "error": f"timeout {timeout}s (Excel busy?)",
        }
    engine_ms = (time.perf_counter() - t0) * 1000
    if proc.returncode != 0:
        return {
            "route": "P-com",
            "artifact": "PDF->PNG",
            "engineMs": round(engine_ms, 1),
            "error": (proc.stderr or proc.stdout).strip()[:300],
        }
    pdf_bytes = pdf.stat().st_size if pdf.exists() else 0
    # First-page PNG via pdftoppm (existing pdf pipeline tool).
    png = out_dir / (path.stem + "-p1.png")
    pdftoppm = shutil.which("pdftoppm")
    png_ms, png_bytes = 0.0, 0
    if pdftoppm and pdf.exists():
        prefix = out_dir / (path.stem + "-p1")
        t1 = time.perf_counter()
        try:
            subprocess.run(
                [pdftoppm, "-png", "-r", "100", "-f", "1", "-l", "1",
                 "-singlefile", str(pdf), str(prefix)],
                capture_output=True,
                timeout=120,
            )
            png_ms = (time.perf_counter() - t1) * 1000
            if png.exists():
                png_bytes = png.stat().st_size
        except Exception as err:  # noqa: BLE001
            return {
                "route": "P-com",
                "artifact": "PDF + PNG p.1",
                "engineMs": round(engine_ms, 1),
                "pdfBytes": pdf_bytes,
                "error": f"pdftoppm: {err}",
            }
    return {
        "route": "P-com",
        "artifact": "PDF + PNG стр.1 (100dpi)",
        "engineMs": round(engine_ms, 1),
        "pngMs": round(png_ms, 1),
        "pdfBytes": pdf_bytes,
        "pngBytes": png_bytes,
    }


def write_gallery(out_dir: Path, entries: list[dict]) -> Path:
    import shutil as _shutil

    cards = []
    for e in entries:
        p = Path(e["path"])
        local = out_dir / f"{e.get('tag', p.parent.name)}-{p.name}"
        _shutil.copyfile(p, local)
        if local.suffix == ".png":
            cards.append(
                f"<div><h3>{local.name} ({e.get('bytes', 0)} B)</h3>"
                f"<img src='{local.name}' style='max-width:640px;border:1px solid #ccc'></div>"
            )
        else:
            cards.append(
                f"<div><h3>{local.name} ({e.get('bytes', 0)} B)</h3>"
                f"<iframe src='{local.name}' style='width:640px;height:420px;border:1px solid #ccc'></iframe></div>"
            )
    gallery = out_dir / "gallery.html"
    gallery.write_text(
        "<!doctype html><html lang=ru><head><meta charset=utf-8>"
        "<title>S02 gallery</title></head><body><h1>S02 Excel stand</h1>"
        + "".join(cards) + "</body></html>",
        encoding="utf-8",
    )
    return gallery


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="")
    ap.add_argument("--skip-com", action="store_true")
    args = ap.parse_args()

    tmp = Path(args.out) if args.out else Path(tempfile.mkdtemp(prefix="excel_s02_"))
    tmp.mkdir(parents=True, exist_ok=True)
    cache_base = tmp / "cache"
    small = tmp / "book_small.xlsx"
    medium = tmp / "book_medium.xlsx"
    make_workbook(small, rows=30, cols=8)
    make_workbook(medium, rows=500, cols=15, sheets=3)

    results: dict = {"books": {}, "cache": {}, "note": ""}
    for name, book in (("small(30x8)", small), ("medium(500x15x3)", medium)):
        row: dict = {}
        for rep in range(3):
            row.setdefault("H-fast", []).append(bench_h_fast(book, cache_base))
            # drop cache between reps 1..2 only for cold measure: keep rep0 cold
            if rep == 0:
                shutil.rmtree(cache_base, ignore_errors=True)
        # steady cold/warm pair on clean keys
        shutil.rmtree(cache_base, ignore_errors=True)
        cold_h = bench_h_fast(book, cache_base)
        warm_h = bench_h_fast(book, cache_base)
        cold_i = bench_i_png(book, cache_base)
        # warm png: second call hits cache
        warm_i = bench_i_png(book, cache_base)
        entry = {
            "H-fast cold": cold_h,
            "H-fast warm": warm_h,
            "I-png cold": cold_i,
            "I-png warm": warm_i,
            "H-full": bench_h_full(book),
        }
        if name.startswith("small") and not args.skip_com:
            entry["P-com"] = bench_p_com(book, tmp)
        row = entry
        results["books"][name] = row

    # Restart reuse: fresh python process, same cache dir -> must hit.
    probe = tmp / "restart_probe.py"
    probe.write_text(
        "import sys; sys.path.insert(0, r'%s');"
        "from pathlib import Path;"
        "from app.backend import excel_speed;"
        "r1 = excel_speed.get_fast_html(Path(r'%s'), Path(r'%s'));"
        "r2 = excel_speed.get_fast_png(Path(r'%s'), Path(r'%s'));"
        "print('HTML_HIT' if r1['cacheHit'] else 'HTML_MISS');"
        "print('PNG_HIT' if r2['cacheHit'] else 'PNG_MISS')"
        % (WORKTREE_ROOT, medium, cache_base, medium, cache_base),
        encoding="utf-8",
    )
    proc = subprocess.run(
        [sys.executable, str(probe)], capture_output=True, text=True, timeout=120
    )
    results["cache"]["restart"] = proc.stdout.strip().splitlines()

    # Source change -> new key, rebuild, no hit.
    small.stat()
    with open(small, "ab") as f:
        f.write(b" ")
    # bumping mtime reliably
    time.sleep(0.05)
    small.touch()
    after_h = excel_speed.get_fast_html(small, cache_base)
    after_i = excel_speed.get_fast_png(small, cache_base)
    results["cache"]["afterChange"] = {
        "htmlHit": after_h["cacheHit"],
        "pngHit": after_i["cacheHit"],
    }

    gallery = write_gallery(
        tmp,
        [
            {"tag": "small-h", "path": after_h["path"], "bytes": after_h["bytes"]},
            {"tag": "small-p", "path": after_i["path"], "bytes": after_i["bytes"]},
        ],
    )
    results["gallery"] = str(gallery)
    results["cacheDir"] = str(cache_base)
    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
