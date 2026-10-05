"""Движок Word: быстрый DOCX→HTML первично, Word COM — только фолбэк .doc.

Рендер v2 (DOCX_HTML_RENDER_VERSION):
- блоки тела (абзацы и таблицы) идут в исходном порядке документа;
- нумерация вычисляется по определениям самого DOCX (numId/abstractNum/ilvl,
  numFmt/lvlText, start, overrides, перезапуски, нумерация через стили);
- гиперссылки сохраняют текст и адрес (исполняемые схемы блокируются без потери текста);
- изображения — из абзацев и из ячеек таблиц (DrawingML и VML);
- непустые колонтитулы показываются в помеченных блоках (без фиктивных номеров страниц).
"""

from __future__ import annotations

import base64
import html as _html
import json
import os
import re
import subprocess
import tempfile
import time
from pathlib import Path

try:
    from docx import Document
except ImportError:
    Document = None

WORD_COM_TIMEOUT_SECONDS = 120

# Версия HTML-рендера DOCX. Поднимается при любом изменении выдачи docx_to_html_string.
# ensure_docx_preview сверяет её в метаданных и не отдаёт старый HTML за новый.
DOCX_HTML_RENDER_VERSION = 4

# Сколько секунд чужой lock-файл считается живым (защита от зависших генераций).
CACHE_LOCK_STALE_SECONDS = 180

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
W = "{%s}" % W_NS
A = "{%s}" % A_NS
R_ID = "{%s}id" % R_NS
R_EMBED = "{%s}embed" % R_NS

# Схемы, которые разрешено активировать как ссылки. Остальное (javascript:, data:,
# vbscript:, file:, ftp: и т.п.) показывается текстом без href — текст не удаляется.
ALLOWED_LINK_SCHEMES = {"http", "https", "mailto"}

# Известные маркеры Symbol/Wingdings (private use) -> Unicode.
BULLET_CHAR_MAP = {
    "\uf0b7": "\u2022",  # Symbol metastatic dot -> bullet
    "\uf0a7": "\u25aa",  # square-ish -> small black square
    "\uf06c": "\u25aa",
    "\uf0fc": "\u2713",  # check mark
}


def _fmt_roman(value: int, upper: bool) -> str:
    table = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
             (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
             (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
    out = []
    rest = value
    for arabic, glyph in table:
        while rest >= arabic:
            out.append(glyph)
            rest -= arabic
    text = "".join(out)
    return text if upper else text.lower()


def _fmt_letters(value: int, upper: bool) -> str:
    out = []
    rest = value
    while rest > 0:
        rest, rem = divmod(rest - 1, 26)
        out.append(chr(ord("A") + rem))
    text = "".join(reversed(out))
    return text if upper else text.lower()


class _Numbering:
    """Счётчики списков Word по определениям numbering.xml.

    Перезапуски уровней — по документации w:lvlRestart:
    - 0 означает НИКОГДА не перезапускать этот уровень;
    - положительное значение — 1-based номер уровня-триггера
      (1 -> ilvl 0, 2 -> ilvl 1): уровень сбрасывается при его использовании;
    - отсутствие атрибута — перезапуск при использовании высшего уровня;
    - некорректные значения — без сброса (тоже никогда).
    Текущий уровень при собственном наступлении никогда не сбрасывается.
    Состояние независимо для каждого numId. Неподдержанные форматы фиксируются
    в self.unsupported, а не теряются молча.
    """

    def __init__(self, doc, diagnostics: list):
        self._diag = diagnostics
        self._abstracts: dict = {}
        self._nums: dict = {}
        self._counters: dict = {}
        self._last: tuple | None = None
        try:
            numbering = doc.part.numbering_part.numbering_definitions._numbering
        except Exception:
            return
        for abstract in numbering.findall(W + "abstractNum"):
            aid = abstract.get(W + "abstractNumId")
            levels = {}
            for lvl in abstract.findall(W + "lvl"):
                ilvl_raw = lvl.get(W + "ilvl")
                try:
                    ilvl = int(ilvl_raw)
                except (TypeError, ValueError):
                    continue
                levels[ilvl] = self._parse_level(lvl)
            self._abstracts[aid] = levels
        for num in numbering.findall(W + "num"):
            nid = num.get(W + "numId")
            aid_el = num.find(W + "abstractNumId")
            starts = {}
            level_ov = {}
            for override in num.findall(W + "lvlOverride"):
                try:
                    oilvl = int(override.get(W + "ilvl"))
                except (TypeError, ValueError):
                    continue
                start_el = override.find(W + "startOverride")
                if start_el is not None:
                    try:
                        starts[oilvl] = int(start_el.get(W + "val"))
                    except (TypeError, ValueError):
                        pass
                nested = override.find(W + "lvl")
                if nested is not None:
                    level_ov[oilvl] = self._parse_level(nested)
            self._nums[nid] = {
                "abstract": aid_el.get(W + "val") if aid_el is not None else None,
                "starts": starts,
                "levels": level_ov,
            }

    @staticmethod
    def _parse_level(lvl_el) -> dict:
        get = lambda tag: (lvl_el.find(W + tag).get(W + "val")
                           if lvl_el.find(W + tag) is not None else None)
        try:
            start = int(get("start") or 1)
        except ValueError:
            start = 1
        return {
            "start": start,
            "numFmt": get("numFmt") or "decimal",
            "lvlText": get("lvlText") or "",
            "lvlRestart": get("lvlRestart"),
        }

    def _restart_trigger(self, num_id, abstract_id, other: int) -> int | None:
        """Уровень-триггер для сброса уровня other либо None (= никогда).

        По Microsoft Open XML триггером может быть только существующий высший
        уровень: 0 <= trigger < other. Значение 0, отрицательные, нечисловые,
        указывающие на несуществующий уровень (val > 9) либо на текущий
        и более глубокие уровни — недопустимы и дают «никогда» (с записью
        в диагностику). Учитывает вложенное переопределение уровня из lvlOverride.
        """
        definition = self._level_def(num_id, abstract_id, other)
        raw = (definition or {}).get("lvlRestart")
        if raw is None:
            return other - 1 if other > 0 else None
        try:
            value = int(raw)
        except (TypeError, ValueError):
            self._diag.append(f"lvlRestart:{raw} не число (numId={num_id} ilvl={other}) — без сброса")
            return None
        if value <= 0:
            return None
        trigger = value - 1
        if trigger < 0 or trigger >= other or trigger > 8:
            self._diag.append(
                f"lvlRestart:{raw} недопустим для ilvl={other} (numId={num_id}) — без сброса")
            return None
        return trigger

    def _level_def(self, num_id, abstract_id, ilvl):
        override = ((self._nums.get(str(num_id)) or {}).get("levels", {}) or {}).get(ilvl)
        if override is not None:
            return override
        levels = self._abstracts.get(abstract_id)
        if levels is None:
            return None
        return levels.get(ilvl)

    def _format_value(self, num_fmt: str, value: int, num_id, ilvl) -> str:
        if num_fmt == "decimal":
            return str(value)
        if num_fmt == "upperRoman":
            return _fmt_roman(value, True) if 0 < value < 4000 else str(value)
        if num_fmt == "lowerRoman":
            return _fmt_roman(value, False) if 0 < value < 4000 else str(value)
        if num_fmt == "upperLetter":
            return _fmt_letters(value, True)
        if num_fmt == "lowerLetter":
            return _fmt_letters(value, False)
        if num_fmt in ("bullet", "none"):
            return ""
        self._diag.append(f"numFmt:{num_fmt} (numId={num_id} ilvl={ilvl}) — показан десятичным")
        return str(value)

    def number_for(self, num_id: str | None, ilvl: int | None) -> str | None:
        """Следующий видимый номер списка. Возвращает None, если numbering неизвестна."""
        if num_id is None or num_id == "0":
            return None
        num = self._nums.get(str(num_id))
        if num is None or num["abstract"] is None:
            self._diag.append(f"numId:{num_id} без abstractNum — номер пропущен")
            return None
        abstract_id = num["abstract"]
        level = int(ilvl) if ilvl is not None else 0
        state = self._counters.setdefault(str(num_id), {"levels": [None] * 9})
        counters = state["levels"]
        # Перезапуск уровня other — только когда наступает его триггер
        # (_restart_trigger). Текущий уровень здесь никогда не сбрасывается.
        for other in range(9):
            if other == level:
                continue
            if self._restart_trigger(num_id, abstract_id, other) == level:
                counters[other] = None

        def level_start(idx: int) -> int:
            override = num["starts"].get(idx)
            if override is not None:
                return override
            definition = self._level_def(num_id, abstract_id, idx)
            return (definition or {}).get("start", 1) or 1

        if counters[level] is None:
            counters[level] = level_start(level)
        else:
            counters[level] += 1
        self._last = (str(num_id), level)

        definition = self._level_def(num_id, abstract_id, level)
        if definition is None:
            self._diag.append(f"numId:{num_id} ilvl:{level} без lvl — номер пропущен")
            return None
        num_fmt = definition.get("numFmt") or "decimal"
        if num_fmt == "none":
            return ""
        if num_fmt == "bullet":
            return BULLET_CHAR_MAP.get(definition.get("lvlText") or "",
                                       definition.get("lvlText") or "\u2022")

        def formatted(idx: int) -> str:
            idx_def = self._level_def(num_id, abstract_id, idx) or {}
            value = counters[idx]
            if value is None:
                value = idx_def.get("start", 1) or 1
            return self._format_value(idx_def.get("numFmt") or "decimal",
                                      value, num_id, idx)

        template = definition.get("lvlText") or ""
        return re.sub(r"%([1-9])", lambda m: formatted(int(m.group(1)) - 1), template)


def _split_num_pr(num_pr_el) -> tuple:
    """(numId|None, ilvl|None) — None сохраняется, чтобы вызыватель мог дополнить."""
    if num_pr_el is None:
        return None, None
    num_id_el = num_pr_el.find(W + "numId")
    ilvl_el = num_pr_el.find(W + "ilvl")
    num_id = num_id_el.get(W + "val") if num_id_el is not None else None
    ilvl_raw = ilvl_el.get(W + "val") if ilvl_el is not None else None
    try:
        ilvl = int(ilvl_raw) if ilvl_raw is not None else None
    except ValueError:
        ilvl = None
    return num_id, ilvl


def _paragraph_num_pr(paragraph, style_cache: dict | None = None) -> tuple | None:
    """(numId, ilvl) абзаца.

    Прямой numPr дополняется нумерацией из цепочки стилей поатрибутно:
    недостающий numId/ilvl берётся из стиля. Обход цепочки кэшируется по styleId.
    """
    direct_id, direct_ilvl = None, None
    style_id = None
    pPr = paragraph._p.find(W + "pPr")
    if pPr is not None:
        direct_id, direct_ilvl = _split_num_pr(pPr.find(W + "numPr"))
        p_style = pPr.find(W + "pStyle")
        if p_style is not None:
            style_id = p_style.get(W + "val")
    if direct_id is not None and direct_ilvl is not None:
        return direct_id, direct_ilvl
    if style_cache is not None and style_id in style_cache:
        style_ref = style_cache[style_id]
    else:
        style_ref = _style_chain_num_pr(paragraph)
        if style_cache is not None:
            style_cache[style_id] = style_ref
    style_id_v, style_ilvl = style_ref or (None, None)
    num_id = direct_id if direct_id is not None else style_id_v
    if num_id is None:
        return None
    ilvl = direct_ilvl if direct_ilvl is not None else style_ilvl
    return num_id, ilvl if ilvl is not None else 0


def _style_chain_num_pr(paragraph) -> tuple | None:
    style = getattr(paragraph, "style", None)
    seen = set()
    while style is not None and style.name not in seen:
        seen.add(style.name)
        try:
            style_pPr = style.element.find(W + "pPr")
        except Exception:
            break
        if style_pPr is not None:
            num_id, ilvl = _split_num_pr(style_pPr.find(W + "numPr"))
            if num_id is not None or ilvl is not None:
                return num_id, ilvl
        try:
            style = style.base_style
        except Exception:
            break
    return None


def _heading_tag(paragraph, style_cache: dict | None = None) -> str:
    """Тег по стилю заголовка. Имя стиля кэшируется по styleId (дорогой резолв)."""
    pPr = paragraph._p.find(W + "pPr")
    style_id = None
    if pPr is not None:
        p_style = pPr.find(W + "pStyle")
        if p_style is not None:
            style_id = p_style.get(W + "val")
    if style_cache is not None and style_id in style_cache:
        return style_cache[style_id]
    try:
        name = (paragraph.style.name or "").lower()
    except Exception:
        name = ""
    if name.startswith("heading 1"):
        tag = "h1"
    elif name.startswith("heading 2"):
        tag = "h2"
    elif name.startswith("heading 3"):
        tag = "h3"
    else:
        tag = "p"
    if style_cache is not None:
        style_cache[style_id] = tag
    return tag


def _has_content(p_element) -> bool:
    """Есть ли видимое содержимое (текст/рисунок) — один проход, без .text и xpath."""
    for node in p_element.iter():
        tag = node.tag
        if tag == W + "t":
            if (node.text or "").strip():
                return True
        elif tag == W + "drawing" or tag == W + "pict":
            return True
    return False


def _run_html(run_el, stats: dict) -> str:
    """HTML одного w:r. Текст экранируется сразу, разметка <br> — нет."""
    parts = []
    for node in run_el:
        if node.tag == W + "t":
            if node.text:
                parts.append(_html.escape(node.text))
        elif node.tag == W + "tab":
            parts.append("  ")
        elif node.tag == W + "br":
            parts.append("<br>")
    text = "".join(parts)
    if not text:
        return ""
    rPr = run_el.find(W + "rPr")
    bold = italic = underline = False
    if rPr is not None:
        bold = _is_on(rPr.find(W + "b"))
        italic = _is_on(rPr.find(W + "i"))
        underline = _underline_on(rPr.find(W + "u"))
    if bold:
        text = f"<b>{text}</b>"
    if italic:
        text = f"<i>{text}</i>"
    if underline:
        text = f"<u>{text}</u>"
    return text


def _is_on(el) -> bool:
    """ST_OnOff: отсутствие = вкл; явные false/0/off/no/none = выкл."""
    if el is None:
        return False
    val = (el.get(W + "val") or "true").strip().lower()
    return val not in ("false", "0", "off", "no", "none", "f", "n")


def _underline_on(el) -> bool:
    if el is None:
        return False
    val = (el.get(W + "val") or "single").strip().lower()
    return val != "none"


def _styled_text(text: str, rPr) -> str:
    """Единая реализация b/i/u для обычного текста и гиперссылок."""
    if not text:
        return ""
    if rPr is not None:
        if _is_on(rPr.find(W + "b")):
            text = f"<b>{text}</b>"
        if _is_on(rPr.find(W + "i")):
            text = f"<i>{text}</i>"
        if _underline_on(rPr.find(W + "u")):
            text = f"<u>{text}</u>"
    return text


def _sanitize_href(target: str) -> str | None:
    """Разрешённый адрес или None (текст ссылки при этом сохраняется)."""
    if target.startswith("#"):
        return target
    scheme = target.split(":", 1)[0].lower() if ":" in target else ""
    if scheme in ALLOWED_LINK_SCHEMES:
        return target
    return None


class _BlockRenderer:
    """Построчный рендер содержимого: runs, гиперссылки, закладки, рисунки, поля."""

    def __init__(self, doc, stats: dict):
        self._doc_rels = doc.part.rels
        self._stats = stats

    def _rels_for(self, rels):
        return rels if rels is not None else self._doc_rels

    def images_in(self, element, rels=None) -> str:
        rels = self._rels_for(rels)
        out = []
        for node in element.iter():
            if node.tag == A + "blip":
                rid = node.get(R_EMBED)
            elif node.tag == "{%s}imagedata" % V_NS:
                rid = node.get(R_ID)
            else:
                continue
            if rid in rels:
                try:
                    blob = rels[rid].target_part.blob
                    ctype = rels[rid].target_part.content_type
                except Exception:
                    continue
                if ctype in ("image/emf", "image/wmf"):
                    self._stats.setdefault("imagesUnsupported", []).append(ctype)
                    out.append(f'<span class="img-unsupported" title="Формат {ctype} '
                               f"не показывается браузером\">[рисунок {ctype}]</span>")
                    continue
                out.append(f'<img src="data:{ctype};base64,'
                           f'{base64.b64encode(blob).decode()}">')
        if out:
            self._stats["images"] = self._stats.get("images", 0) + len(out)
        return "".join(out)

    def content_html(self, p_element, *, allow_fields_marker: bool, rels=None) -> str:
        """Внутренний HTML абзаца: гиперссылки, закладки, рисунки, поля страниц.

        Runs разбираются за один проход по дочерним узлам (без повторных поисков).
        rels — relationships части-владельца (у колонтитулов свои rels).
        """
        out = []
        in_field = False
        field_name: str | None = None
        field_has_marker = False
        for child in p_element:
            if child.tag == W + "bookmarkStart":
                name = child.get(W + "name")
                if name and not name.startswith("_"):
                    out.append(f'<span class="anchor" id="bm-{_html.escape(name)}"></span>')
                    self._stats["anchors"] = self._stats.get("anchors", 0) + 1
                continue
            if child.tag == W + "hyperlink":
                out.append(self._hyperlink_html(child, rels))
                continue
            if child.tag != W + "r":
                continue
            parts: list = []
            run_rPr = None
            run_instr: str | None = None
            run_fld: str | None = None
            for node in child:
                ntag = node.tag
                if ntag == W + "t":
                    if node.text:
                        parts.append(_html.escape(node.text))
                elif ntag == W + "tab":
                    parts.append("  ")
                elif ntag == W + "br":
                    parts.append("<br>")
                elif ntag == W + "rPr":
                    run_rPr = node
                elif ntag == W + "instrText":
                    code = (node.text or "").strip().split()
                    if code:
                        run_instr = code[0].upper()
                elif ntag == W + "fldChar":
                    run_fld = node.get(W + "fldCharType")
            if run_instr is not None:
                in_field = True
                field_name = run_instr
                field_has_marker = False
                continue
            if run_fld is not None:
                if run_fld == "begin":
                    in_field = True
                    field_has_marker = False
                elif run_fld == "end":
                    in_field = False
                    field_name = None
                continue
            if in_field and field_name in ("PAGE", "NUMPAGES", "PAGENUM", "SECTIONPAGES"):
                if allow_fields_marker and not field_has_marker:
                    out.append(f'<span class="page-field" title="Поле Word {field_name}: '
                               f"номер не вычисляется\">[{field_name}]</span>")
                    field_has_marker = True
                    marked = self._stats.setdefault("fieldsMarked", [])
                    if field_name not in marked:
                        marked.append(field_name)
                continue
            text = _styled_text("".join(parts), run_rPr)
            if not text:
                continue
            out.append(text)
        return "".join(out)

    def _hyperlink_html(self, link_el, rels=None) -> str:
        rels = self._rels_for(rels)
        inner_runs = "".join(_run_html(r, self._stats) for r in link_el.findall(W + "r"))
        inner = inner_runs or _html.escape("".join(
            t.text or "" for t in link_el.findall(".//" + W + "t")))
        if not inner:
            return ""
        self._stats["links"] = self._stats.get("links", 0) + 1
        rid = link_el.get(R_ID)
        anchor = link_el.get(W + "anchor")
        if anchor:
            return f'<a href="#bm-{_html.escape(anchor)}">{inner}</a>'
        if rid and rid in rels:
            try:
                target = rels[rid].target_ref or ""
            except Exception:
                target = ""
            href = _sanitize_href(target)
            if href is not None:
                return f'<a href="{_html.escape(href, quote=True)}">{inner}</a>'
            self._stats.setdefault("blockedUrls", []).append(target.split(":", 1)[0].lower())
            return (f'<span class="blocked-link" title="Адрес не активируется, '
                    f'текст сохранён">{inner}</span>')
        self._stats.setdefault("linksUnsupported", []).append("hyperlink-without-target")
        return inner


def _runs_html(paragraph) -> str:
    """Совместимость: runs абзаца (без ссылок/картинок верхнего уровня)."""
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
        shd = tc_pr.find(W + "shd")
        if shd is not None:
            fill = shd.get(W + "fill")
            if fill not in (None, "auto", "ffffff"):
                return "#" + fill[-6:]
    except Exception:
        pass
    return None


def _cell_span(cell) -> str:
    try:
        grid_span = cell._tc.get_or_add_tcPr().find(W + "gridSpan")
        if grid_span is not None:
            val = int(grid_span.get(W + "val") or 1)
            if val > 1:
                return f' colspan="{val}"'
    except Exception:
        pass
    return ""


def _is_vmerge_continue(cell) -> bool:
    try:
        tc_pr = cell._tc.find(W + "tcPr")
        if tc_pr is None:
            return False
        vmerge = tc_pr.find(W + "vMerge")
        if vmerge is None:
            return False
        return (vmerge.get(W + "val") or "continue") != "restart"
    except Exception:
        return False


def _is_hmerge_continue(cell) -> bool:
    try:
        tc_pr = cell._tc.find(W + "tcPr")
        if tc_pr is None:
            return False
        hmerge = tc_pr.find(W + "hMerge")
        if hmerge is None:
            return False
        return (hmerge.get(W + "val") or "continue") != "restart"
    except Exception:
        return False


def _row_grid_map(table, row_idx: int) -> list:
    """Карта EQUAL grid-колонок строки: список (cell, start_col, span).

    Используются истинные <w:tc> строки (row._tr), а не row.cells: прокси
    python-docx подменяют vMerge-продолжения корневой ячейкой и размножают
    ячейку по grid_span — оба поведения дают дублирование содержимого.
    """
    from docx.table import _Cell
    mapping = []
    col = 0
    for tc in table.rows[row_idx]._tr.findall(W + "tc"):
        cell = _Cell(tc, table)
        try:
            grid_span = tc.get_or_add_tcPr().find(W + "gridSpan")
            span = int(grid_span.get(W + "val") or 1) if grid_span is not None else 1
        except Exception:
            span = 1
        mapping.append((cell, col, span))
        col += span
    return mapping


def _vmerge_rowspan(table, row_idx: int, start_col: int, span: int) -> int:
    """Сколько строк вниз занимает вертикально объединённая ячейка."""
    rowspan = 1
    for other_idx in range(row_idx + 1, len(table.rows)):
        covered = 0
        for cell, col, cell_span in _row_grid_map(table, other_idx):
            overlap = min(col + cell_span, start_col + span) - max(col, start_col)
            if overlap > 0:
                if _is_vmerge_continue(cell):
                    covered += overlap
                else:
                    return rowspan
        if covered < span:
            return rowspan
        rowspan += 1
    return rowspan


def _render_table(table, renderer: _BlockRenderer, numbering: _Numbering,
                  num_cache: dict, tag_cache: dict, rels=None) -> str:
    parts = ["<table>"]
    for row_idx, row in enumerate(table.rows):
        parts.append("<tr>")
        grid = _row_grid_map(table, row_idx)
        for cell, start_col, span in grid:
            # Продолжения объединений не дублируют содержимое исходной ячейки.
            if _is_vmerge_continue(cell) or _is_hmerge_continue(cell):
                continue
            bg = _cell_shading(cell)
            attrs = f' style="background:{bg}"' if bg else ""
            if span > 1:
                attrs += f' colspan="{span}"'
            vspan = _vmerge_rowspan(table, row_idx, start_col, span)
            if vspan > 1:
                attrs += f' rowspan="{vspan}"'
            chunks = []
            for paragraph in cell.paragraphs:
                chunks.append(_render_paragraph_inner(
                    paragraph, renderer, numbering, num_cache, tag_cache,
                    in_table=True, rels=rels))
            parts.append(f"<td{attrs}>{'<br>'.join(c for c in chunks if c)}</td>")
        parts.append("</tr>")
    parts.append("</table>")
    return "".join(parts)


def _alignment_class(paragraph) -> str:
    try:
        align = {1: "c", 2: "r", 3: "j"}.get(paragraph.alignment, "")
    except Exception:
        align = ""
    return f' class="{align}"' if align else ""


def _render_paragraph_inner(paragraph, renderer: _BlockRenderer,
                            numbering: _Numbering, num_cache: dict, tag_cache: dict,
                            *, in_table: bool, rels=None) -> str:
    num_ref = _paragraph_num_pr(paragraph, num_cache)
    marker = ""
    if num_ref is not None:
        number = numbering.number_for(*num_ref)
        if number:
            marker = f'<span class="wnum">{_html.escape(number)}</span> '
    return marker + renderer.content_html(paragraph._p, allow_fields_marker=True,
                                          rels=rels) \
        + renderer.images_in(paragraph._p, rels)


def _render_paragraph_block(paragraph, renderer: _BlockRenderer, numbering: _Numbering,
                            num_cache: dict, tag_cache: dict) -> str | None:
    if not _has_content(paragraph._p):
        return None
    tag = _heading_tag(paragraph, tag_cache)
    inner = _render_paragraph_inner(paragraph, renderer, numbering, num_cache, tag_cache,
                                    in_table=False, rels=None)
    return f"<{tag}{_alignment_class(paragraph)}>{inner}</{tag}>"


def _walk_body(doc):
    """Блоки тела в исходном порядке: ('p', paragraph) / ('tbl', table)."""
    tables = {id(t): t for t in doc.tables}

    def table_for(el):
        from docx.table import Table
        key = None
        for ident, table in tables.items():
            if table._tbl is el:
                key = ident
                break
        if key is not None:
            return tables.pop(key)
        return Table(el, doc)

    def visit(container, out):
        for child in container:
            if child.tag == W + "p":
                from docx.text.paragraph import Paragraph
                out.append(("p", Paragraph(child, doc)))
            elif child.tag == W + "tbl":
                out.append(("tbl", table_for(child)))
            elif child.tag == W + "sdt":
                content = child.find(W + "sdtContent")
                if content is not None:
                    visit(content, out)

    blocks: list = []
    visit(doc.element.body, blocks)
    return blocks


def _render_header_footer(container, renderer: _BlockRenderer, numbering: _Numbering,
                          num_cache: dict, tag_cache: dict,
                          kind: str, section_idx: int) -> str | None:
    try:
        part_rels = container.part.rels
    except Exception:
        part_rels = None
    chunks = []
    for paragraph in container.paragraphs:
        if not _has_content(paragraph._p):
            continue
        inner = _render_paragraph_inner(paragraph, renderer, numbering, num_cache, tag_cache,
                                        in_table=False, rels=part_rels)
        chunks.append(f"<p>{inner}</p>")
    for table in container.tables:
        chunks.append(_render_table(table, renderer, numbering, num_cache, tag_cache,
                                    rels=part_rels))
    if not chunks:
        return None
    label = "Верхний колонтитул" if kind == "header" else "Нижний колонтитул"
    return (f'<div class="word-{kind}" data-section="{section_idx}">'
            f'<div class="wkh-label">{label}</div>{"".join(chunks)}</div>')


def docx_to_html_string(src: Path) -> tuple[str, dict]:
    """Распарсить DOCX в standalone HTML. Возвращает (html, stats)."""
    if Document is None:
        raise RuntimeError("Для просмотра DOCX нужен пакет python-docx")
    diagnostics: list = []
    stats: dict = {"renderVersion": DOCX_HTML_RENDER_VERSION}
    t0 = time.perf_counter()
    doc = Document(str(src))
    t_load = (time.perf_counter() - t0) * 1000
    t1 = time.perf_counter()
    numbering = _Numbering(doc, diagnostics)
    renderer = _BlockRenderer(doc, stats)
    num_cache: dict = {}
    tag_cache: dict = {}
    parts = ['<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">'
             f'<title>{_html.escape(src.stem)}</title><style>'
             'body{font-family:"Times New Roman",serif;max-width:900px;margin:24px auto;'
             'padding:0 20px;color:#111;line-height:1.5;font-size:15px}'
             'h1{font-size:21px;text-align:center}h2{font-size:17px}'
             'table{border-collapse:collapse;margin:14px 0;width:100%}'
             'td,th{border:1px solid #555;padding:5px 10px;vertical-align:top}'
             'img{max-width:100%}.c{text-align:center}.r{text-align:right}.j{text-align:justify}'
             '.wnum{font-weight:bold;margin-right:.4em}'
             '.word-header,.word-footer{border:1px dashed #888;padding:8px 12px;margin:14px 0;'
             'background:#f7f7f7}.wkh-label{font-size:12px;color:#555;margin-bottom:6px}'
             '.page-field{color:#555;border-bottom:1px dotted #555}'
             '.blocked-link{border-bottom:1px dotted #555}'
             '.img-unsupported{color:#555}'
             '</style></head><body>']
    n_paragraphs = 0
    n_tables = 0
    for kind, block in _walk_body(doc):
        if kind == "p":
            html_block = _render_paragraph_block(block, renderer, numbering, num_cache, tag_cache)
            if html_block is None:
                continue
            parts.append(html_block)
            n_paragraphs += 1
        else:
            parts.append(_render_table(block, renderer, numbering, num_cache, tag_cache))
            n_tables += 1
    n_headers = n_footers = 0
    try:
        sections = doc.sections
    except Exception:
        sections = []
    for idx, section in enumerate(sections):
        try:
            sect_pr = section._sectPr
        except Exception:
            sect_pr = None
        if sect_pr is not None:
            for ref in list(sect_pr.findall(W + "headerReference")) + \
                    list(sect_pr.findall(W + "footerReference")):
                ref_type = ref.get(W + "type")
                if ref_type in ("first", "even"):
                    limited = stats.setdefault("headersUnsupported", [])
                    if ref_type not in limited:
                        limited.append(ref_type)
        try:
            header = section.header
            linked_h = bool(header.is_linked_to_previous) and idx > 0
        except Exception:
            header, linked_h = None, False
        if header is not None and not linked_h:
            rendered = _render_header_footer(header, renderer, numbering,
                                             num_cache, tag_cache, "header", idx)
            if rendered:
                parts.append(rendered)
                n_headers += 1
        try:
            footer = section.footer
            linked_f = bool(footer.is_linked_to_previous) and idx > 0
        except Exception:
            footer, linked_f = None, False
        if footer is not None and not linked_f:
            rendered = _render_header_footer(footer, renderer, numbering,
                                             num_cache, tag_cache, "footer", idx)
            if rendered:
                parts.append(rendered)
                n_footers += 1
    parts.append("</body></html>")
    t_gen = (time.perf_counter() - t1) * 1000
    stats.update({
        "paragraphs": n_paragraphs,
        "tables": n_tables,
        "headers": n_headers,
        "footers": n_footers,
        "parseMs": round(t_load, 1),
        "generateMs": round(t_gen, 1),
        "totalMs": round(t_load + t_gen, 1),
    })
    if diagnostics:
        seen = []
        for item in diagnostics:
            if item not in seen:
                seen.append(item)
        stats["numberingUnsupported"] = seen
    for key in ("blockedUrls", "linksUnsupported", "imagesUnsupported",
                "fieldsMarked", "headersUnsupported",
                "links", "images", "anchors"):
        if key in stats and not stats[key]:
            del stats[key]
    return "".join(parts), stats


def _try_acquire_lock(cache_dir: Path) -> bool:
    """Одноразовая попытка взять per-документ lock. True — генерация за нами."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    try:
        fd = os.open(str(cache_dir / "preview.lock"), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return False
    try:
        os.write(fd, str(os.getpid()).encode())
    finally:
        os.close(fd)
    return True


def _release_doc_lock(cache_dir: Path) -> None:
    try:
        (cache_dir / "preview.lock").unlink()
    except OSError:
        pass


def _lock_age(cache_dir: Path) -> float | None:
    try:
        return time.time() - (cache_dir / "preview.lock").stat().st_mtime
    except OSError:
        return None


def _read_valid_meta(meta_path: Path, html_path: Path, stat) -> dict | None:
    """Валидная мета: JSON цел, HTML существует и непуст, источник и версия совпали."""
    try:
        if not html_path.exists() or html_path.stat().st_size == 0:
            return None
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if (meta.get("sourceMtimeNs") == stat.st_mtime_ns
            and meta.get("sourceSize") == stat.st_size
            and meta.get("renderVersion") == DOCX_HTML_RENDER_VERSION):
        return meta
    return None


def _atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
        os.replace(tmp_name, path)
    except BaseException:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise


def ensure_docx_preview(src: Path, cache_dir: Path) -> dict:
    """Закэшированный DOCX→HTML. Повтор — дисковый хит без парсинга.

    Старый HTML (без renderVersion) за новый не выдаётся. Запись атомарна через
    временные файлы; одновременная генерация одного документа сериализуется
    per-документ lock-файлом. Ожидающий видит только валидный кэш нужной версии
    (устаревшая пара HTML/JSON готовностью не считается); при превышении ожидания
    — обоснованный TimeoutError, а не RuntimeError из-за чужой генерации.
    """
    cache_dir = Path(cache_dir)
    out = cache_dir / "preview.html"
    meta_path = cache_dir / "preview.json"
    stat = src.stat()
    meta = _read_valid_meta(meta_path, out, stat)
    if meta is not None:
        meta["cacheHit"] = True
        return meta
    if _try_acquire_lock(cache_dir):
        owned = True
    else:
        meta = _wait_valid_or_reacquire(cache_dir, out, meta_path, stat)
        if meta is not None:
            meta["cacheHit"] = True
            return meta
        owned = True  # блокировка уже за нами (см. _wait_valid_or_reacquire)
    try:
        # Повторная проверка после взятия блокировки (гонка двух генераторов).
        meta = _read_valid_meta(meta_path, out, stat)
        if meta is not None:
            meta["cacheHit"] = True
            return meta
        page, stats = docx_to_html_string(src)
        _atomic_write_bytes(out, page.encode("utf-8"))
        meta = {
            "url": None,  # URL подставляет сервер (знает маршрут /cache/)
            "file": out.name,
            "bytes": out.stat().st_size,
            "sourceMtimeNs": stat.st_mtime_ns,
            "sourceSize": stat.st_size,
            "cacheHit": False,
            **stats,
        }
        _atomic_write_bytes(meta_path, json.dumps(meta, ensure_ascii=False).encode("utf-8"))
        return meta
    finally:
        if owned:
            _release_doc_lock(cache_dir)


def _wait_valid_or_reacquire(cache_dir: Path, out: Path, meta_path: Path, stat,
                             timeout: int = CACHE_LOCK_STALE_SECONDS) -> dict | None:
    """Ждать чужую генерацию.

    Возвращает валидный кэш (брать как хит) либо None — тогда блокировка уже
    захвачена нами и генерация за вызывающим. Устаревшая пара HTML/JSON
    готовностью не считается. При превышении ожидания — TimeoutError.
    """
    deadline = time.monotonic() + timeout
    while True:
        meta = _read_valid_meta(meta_path, out, stat)
        if meta is not None:
            return meta
        age = _lock_age(cache_dir)
        if age is None or age > CACHE_LOCK_STALE_SECONDS:
            if age is not None:
                try:
                    (cache_dir / "preview.lock").unlink()
                except OSError:
                    pass
            if _try_acquire_lock(cache_dir):
                return None
        if time.monotonic() >= deadline:
            raise TimeoutError(
                "Preview не готов за отведённое ожидание чужой генерации")
        time.sleep(0.2)


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
