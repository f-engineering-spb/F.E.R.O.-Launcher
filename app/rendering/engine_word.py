"""Движок Word: быстрый DOCX→HTML первично, Word COM — только фолбэк .doc.

Замер на боевом договоре (35 стр.): COM→PDF 20.5 с против 1.27 с прямого
HTML с полным сохранением пунктов, таблиц реквизитов и подписей.
"""

from __future__ import annotations

import base64
import html as _html
import json
import subprocess
import time
from pathlib import Path

try:
    from docx import Document
except ImportError:
    Document = None

WORD_COM_TIMEOUT_SECONDS = 120


def _runs_html(paragraph) -> str:
    out = []
    for run in paragraph.runs:
        text = _html.escape(run.text)
        if not text:
            continue
        if run.bold:
            text = f"<b>{text}</b>"
        if run.italic:
            text = f"<i>{text}</i>"
        if run.underline:
            text = f"<u>{text}</u>"
        out.append(text)
    return "".join(out) or _html.escape(paragraph.text)


def _cell_shading(cell) -> str | None:
    try:
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = tc_pr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd")
        if shd is not None:
            fill = shd.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
            if fill not in (None, "auto", "ffffff"):
                return "#" + fill[-6:]
    except Exception:
        pass
    return None


def _cell_span(cell) -> str:
    try:
        grid_span = cell._tc.get_or_add_tcPr().find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridSpan")
        if grid_span is not None:
            val = int(grid_span.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or 1)
            if val > 1:
                return f' colspan="{val}"'
    except Exception:
        pass
    return ""


def docx_to_html_string(src: Path) -> tuple[str, dict]:
    """Распарсить DOCX в standalone HTML. Возвращает (html, stats)."""
    if Document is None:
        raise RuntimeError("Для просмотра DOCX нужен пакет python-docx")
    t0 = time.perf_counter()
    doc = Document(str(src))
    t_load = (time.perf_counter() - t0) * 1000
    t1 = time.perf_counter()
    parts = ['<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">'
             f'<title>{_html.escape(src.stem)}</title><style>'
             'body{font-family:"Times New Roman",serif;max-width:900px;margin:24px auto;'
             'padding:0 20px;color:#111;line-height:1.5;font-size:15px}'
             'h1{font-size:21px;text-align:center}h2{font-size:17px}'
             'table{border-collapse:collapse;margin:14px 0;width:100%}'
             'td,th{border:1px solid #555;padding:5px 10px;vertical-align:top}'
             'img{max-width:100%}.c{text-align:center}.r{text-align:right}.j{text-align:justify}'
             '</style></head><body>']
    rels = doc.part.rels
    n_paragraphs = 0
    for p in doc.paragraphs:
        if not p.text.strip() and not p._p.xpath(".//w:drawing|.//w:pict"):
            continue
        style = (p.style.name or "").lower()
        if style.startswith("heading 1"):
            tag = "h1"
        elif style.startswith("heading 2"):
            tag = "h2"
        elif style.startswith("heading 3"):
            tag = "h3"
        else:
            tag = "p"
        align = {1: "c", 2: "r", 3: "j"}.get(p.alignment, "")
        cls = f' class="{align}"' if align else ""
        imgs = ""
        for blip in p._p.xpath(".//a:blip"):
            rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006"
                           "/relationships}embed")
            if rid in rels:
                blob = rels[rid].target_part.blob
                ctype = rels[rid].target_part.content_type
                imgs += (f'<img src="data:{ctype};base64,'
                         f'{base64.b64encode(blob).decode()}">')
        parts.append(f"<{tag}{cls}>{_runs_html(p)}{imgs}</{tag}>")
        n_paragraphs += 1
    n_tables = 0
    for tbl in doc.tables:
        n_tables += 1
        parts.append("<table>")
        for row in tbl.rows:
            parts.append("<tr>")
            for cell in row.cells:
                bg = _cell_shading(cell)
                style_attr = f' style="background:{bg}"' if bg else ""
                text = "<br>".join(_runs_html(p) for p in cell.paragraphs)
                parts.append(f"<td{style_attr}{_cell_span(cell)}>{text}</td>")
            parts.append("</tr>")
        parts.append("</table>")
    parts.append("</body></html>")
    t_gen = (time.perf_counter() - t1) * 1000
    stats = {
        "paragraphs": n_paragraphs,
        "tables": n_tables,
        "parseMs": round(t_load, 1),
        "generateMs": round(t_gen, 1),
        "totalMs": round(t_load + t_gen, 1),
    }
    return "".join(parts), stats


def ensure_docx_preview(src: Path, cache_dir: Path) -> dict:
    """Закэшированный DOCX→HTML. Повтор — дисковый хит без парсинга."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    out = cache_dir / "preview.html"
    meta_path = cache_dir / "preview.json"
    stat = src.stat()
    if out.exists() and meta_path.exists():
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            if (meta.get("sourceMtimeNs") == stat.st_mtime_ns
                    and meta.get("sourceSize") == stat.st_size):
                meta["cacheHit"] = True
                return meta
        except (OSError, json.JSONDecodeError):
            pass
    page, stats = docx_to_html_string(src)
    out.write_text(page, encoding="utf-8")
    meta = {
        "url": None,  # URL подставляет сервер (знает маршрут /cache/)
        "file": out.name,
        "bytes": out.stat().st_size,
        "sourceMtimeNs": stat.st_mtime_ns,
        "sourceSize": stat.st_size,
        "cacheHit": False,
        **stats,
    }
    meta_path.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    return meta


def ensure_word_pdf(src: Path, target_dir: Path, cache_key: str,
                    convert_script: Path,
                    timeout: int = WORD_COM_TIMEOUT_SECONDS) -> tuple[Path, bool]:
    """COM→PDF с дисковым кэшем (порт word_to_pdf из server.py 1:1 по семантике).

    Манифест и именование PDF идентичны старым, старые кэши переиспользуются.
    """
    from datetime import datetime as _dt
    if not src.exists():
        raise FileNotFoundError(f"Word-файл не найден: {src}")
    if not src.is_file() or src.suffix.casefold() not in {".doc", ".docx"}:
        raise ValueError(f"Это не Word-файл: {src}")
    if not convert_script.exists():
        raise RuntimeError(f"Скрипт конвертации Word не найден: {convert_script}")
    target_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = target_dir / f"{src.stem}.pdf"
    manifest_path = target_dir / "manifest.json"
    if pdf_path.exists() and pdf_path.stat().st_size > 0 and manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            if (manifest.get("sourcePath") == str(src)
                    and manifest.get("cacheKey") == cache_key
                    and manifest.get("sourceMtimeNs") == src.stat().st_mtime_ns
                    and manifest.get("sourceSize") == src.stat().st_size):
                return pdf_path, True
        except (OSError, json.JSONDecodeError):
            pass
    info = word_to_pdf_via_com(src, pdf_path, convert_script, timeout)
    manifest_path.write_text(
        json.dumps(
            {
                "sourcePath": str(src),
                "sourceName": src.name,
                "sourceMtimeNs": src.stat().st_mtime_ns,
                "sourceSize": src.stat().st_size,
                "cacheKey": cache_key,
                "pdfPath": str(pdf_path),
                "convertedAt": _dt.now().isoformat(timespec="seconds"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return Path(info["pdf"]), False


def word_to_pdf_via_com(src: Path, dst_pdf: Path, convert_script: Path,
                        timeout: int = WORD_COM_TIMEOUT_SECONDS) -> dict:
    """Фолбэк для старых .doc: Word COM через PowerShell-скрипт (как раньше).

    Пути манифеста/кэша — на вызывающей стороне; здесь только конвертация.
    """
    if dst_pdf.exists():
        dst_pdf.unlink()
    t0 = time.perf_counter()
    process = subprocess.run(
        ["powershell", "-STA", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(convert_script),
         "-InputPath", str(src), "-OutputPath", str(dst_pdf)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout,
    )
    ms = (time.perf_counter() - t0) * 1000
    if process.returncode != 0:
        message = (process.stderr.strip() or process.stdout.strip()
                   or "Word не смог конвертировать документ в PDF")
        raise RuntimeError(message)
    if not dst_pdf.exists() or dst_pdf.stat().st_size <= 0:
        raise RuntimeError("Word не создал PDF для preview")
    return {"pdf": str(dst_pdf), "bytes": dst_pdf.stat().st_size,
            "convertMs": round(ms, 1)}
