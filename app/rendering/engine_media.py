"""Движок медиа: фото (Pillow + безопасный EXIF) и текст (автокодировка + ссылки).

Безопасный EXIF-поворот: тег Orientation применяется, только если он не
противоречит уже портретному кадру (h > w). Протухший тег после кадрирования
(кейс «Горяинов ВБ стр. 3»: хранится 600x800 + Orientation=8) — игнорируется,
кадр остается портретным, EXIF сбрасывается при сохранении.
"""

from __future__ import annotations

import html as _html
import re
import time
from pathlib import Path

try:
    from PIL import Image, ImageOps
except ImportError:
    Image = None
    ImageOps = None

PHOTO_MAX_SIDE = 1920
PHOTO_HUGE_SIDE = 2500
PHOTO_QUALITY = 85

_URL_RE = re.compile(r"(https?://[^\s<\"']+)")


def _needs_transpose(image) -> tuple[bool, str]:
    """Решить, применять ли EXIF-поворот. Возвращает (применять, причина)."""
    try:
        exif = image.getexif()
    except Exception:
        return False, "no-exif"
    orientation = exif.get(274)
    if not orientation or orientation == 1:
        return False, "orientation-1-or-missing"
    w, h = image.size
    if h > w and orientation in (5, 6, 7, 8):
        # Кадр уже портретный, а тег требует ландшафт — тег протухший
        # (кадрирование после съемки). Не переворачивать.
        return False, f"stale-tag-{orientation}-kept-portrait"
    return True, f"orientation-{orientation}"


def photo_preview(src: Path, dst: Path, max_side: int = PHOTO_MAX_SIDE,
                  quality: int = PHOTO_QUALITY) -> dict:
    """Превью фото: safe-EXIF + даунскейл огромных кадров до Full HD."""
    if Image is None:
        raise RuntimeError("Для просмотра фото нужен пакет Pillow")
    t0 = time.perf_counter()
    timings = {}
    with Image.open(src) as im:
        im.load()
        timings["openMs"] = round((time.perf_counter() - t0) * 1000, 1)
        src_size = (im.width, im.height)
        animated = getattr(im, "is_animated", False)

        t1 = time.perf_counter()
        if animated:
            # Анимированный GIF: уменьшаем кадры, сохраняем анимацию.
            frames = []
            try:
                for frame in range(getattr(im, "n_frames", 1)):
                    im.seek(frame)
                    fr = im.convert("RGB" if frame == 0 else "P")
                    if max(fr.size) > max_side:
                        fr.thumbnail((max_side, max_side), Image.LANCZOS)
                    frames.append(fr.copy())
            except Exception:
                pass
            dst.parent.mkdir(parents=True, exist_ok=True)
            if frames:
                frames[0].save(dst, "GIF", save_all=True,
                               append_images=frames[1:], loop=0)
            transpose_note = "animated-gif-frames-resized"
            out_format = "GIF"
            out_size = frames[0].size if frames else src_size
        else:
            apply, reason = _needs_transpose(im)
            if apply:
                im = ImageOps.exif_transpose(im)
            if max(im.size) > max_side:
                # Тяжелые кадры (>2500px и любые больше экрана) — до Full HD.
                im.thumbnail((max_side, max_side), Image.LANCZOS)
            dst.parent.mkdir(parents=True, exist_ok=True)
            has_alpha = (im.mode in ("RGBA", "LA")
                         or (im.mode == "P" and "transparency" in im.info))
            if src.suffix.casefold() == ".png" or has_alpha:
                if im.mode == "P":
                    im = im.convert("RGBA" if has_alpha else "RGB")
                im.save(dst, "PNG", optimize=True)
                out_format = "PNG"
            else:
                if im.mode != "RGB":
                    im = im.convert("RGB")
                im.save(dst, "JPEG", quality=quality, optimize=True)
                out_format = "JPEG"
            transpose_note = reason
            out_size = (im.width, im.height)
        timings["processMs"] = round((time.perf_counter() - t1) * 1000, 1)

    total = round((timings["openMs"] + timings["processMs"]), 1)
    return {
        "path": str(dst),
        "format": out_format if not animated else "GIF",
        "srcSize": f"{src_size[0]}x{src_size[1]}",
        "outSize": f"{out_size[0]}x{out_size[1]}",
        "srcBytes": src.stat().st_size,
        "outBytes": dst.stat().st_size,
        "exif": transpose_note,
        "totalMs": total,
        **timings,
    }


def detect_text_encoding(raw: bytes) -> tuple[str, str]:
    """Вернуть (encoding, text): UTF-8 (с/без BOM) → иначе CP1251."""
    if raw[:3] == b"\xef\xbb\xbf":
        return "UTF-8-SIG", raw[3:].decode("utf-8", errors="replace")
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "UTF-16", raw.decode("utf-16", errors="replace")
    try:
        return "UTF-8", raw.decode("utf-8")
    except UnicodeDecodeError:
        return "CP1251", raw.decode("cp1251", errors="replace")


def text_preview(src: Path, dst_html: Path) -> dict:
    """Текст в безопасный HTML <pre> с кликабельными ссылками."""
    t0 = time.perf_counter()
    raw = src.read_bytes()
    encoding, text = detect_text_encoding(raw)
    body = _html.escape(text)
    body, n_links = _URL_RE.subn(r'<a href="\1" target="_blank">\1</a>', body)
    page = ('<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">'
            f'<title>{_html.escape(src.name)}</title><style>'
            'body{font-family:Consolas,"Courier New",monospace;margin:24px;'
            'background:#fff;color:#1c2530}pre{background:#f4f6f8;padding:16px;'
            'border-radius:8px;white-space:pre-wrap;word-break:break-word}'
            'a{color:#0b5ed7}.meta{font-family:"Segoe UI",Arial,sans-serif;'
            'color:#607080;font-size:12px;margin-bottom:8px}</style></head><body>'
            f'<div class="meta">{_html.escape(src.name)} · {encoding} · '
            f'{len(text.splitlines())} строк</div><pre>{body}</pre></body></html>')
    dst_html.parent.mkdir(parents=True, exist_ok=True)
    dst_html.write_text(page, encoding="utf-8")
    ms = round((time.perf_counter() - t0) * 1000, 1)
    return {
        "path": str(dst_html),
        "encoding": encoding,
        "lines": len(text.splitlines()),
        "links": n_links,
        "bytes": dst_html.stat().st_size,
        "totalMs": ms,
    }
