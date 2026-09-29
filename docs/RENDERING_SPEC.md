# Стандарт рендеринга F-Engineering Launcher v1.0

Ядро: `app/rendering/` (единый шлюз `dispatcher.get_file_preview()`).
Сервер `app/backend/server.py` держит только тонкие шимы и HTTP-роуты,
вся тяжелая логика живет в движках.

## 1. Золотая 5-ка (превью внутри окна, мгновенно)

| Формат | Движок | Стандарт | Лимит |
|---|---|---|---|
| PDF | `engine_dwg_pdf` (PyMuPDF) | **150 DPI PNG**, стр. 1 сразу, остальные лениво | растр стр.1 < 1 с |
| DWG | DWG-демон → PDF, растр `engine_dwg_pdf` | 150 DPI PNG | как PDF |
| XLSX/XLSM | `engine_excel` (openpyxl) | **1 разбор книги → все вкладки**, HTML-сетка, переключение — кэш-хит | книга 800 КБ < 10 с (все листы), вкладка из кэша < 50 мс |
| XLS (BIFF8) | `engine_excel` (xlrd) | HTML-таблица напрямую, **без COM** | ~10 мс |
| DOCX | `engine_word` (python-docx) | первичный рендер — **чистый HTML** (заголовки, таблицы, картинки base64) | договор 35 стр. ~1.2 с |
| JPG/PNG/GIF | `engine_media` (Pillow) | **safe EXIF-поворот** + даунскейл >1920px до Full HD, JPEG q85 | < 1 с |
| TXT/LOG | `engine_media` | **UTF-8→CP1251 автодетект**, `<pre>` + кликабельные ссылки | < 10 мс |

Кэш: `runtime/cache/rendering/<тип>/…` + существующие кэши сервера
(`runtime/cache/{pdf,word,excel,…}`); ключ — `(путь, mtime, размер, параметры)`.

## 2. COM — только фолбэк

- Word COM (`convert_word_to_pdf.ps1`, таймаут 120 с) — **только старые `.doc`**
  и pixel-perfect случаи (`engine_word.ensure_word_pdf`).
- Excel COM — только legacy `.xls→xlsx` конвертация и PDF-фолбэк.
- Правило: COM никогда не блокирует первичное превью Золотой 5-ки.

## 3. Нативное открытие (специфические форматы)

`.mpp .mov .mp4 .avi .mkv .mpeg .zip .rar .7z .rvt .ifc .ppt …` —
диспетчер возвращает карточку `native_app` (`openAction: /api/open-file`),
фронт показывает заглушку с кнопкой «Открыть в программе».
Исполнение — существующий маршрут `POST /api/open-file`
(ShellExecuteW → `os.startfile` fallback, контракт NATIVE_FILE_OPENING_CONTRACT).

## 4. Безопасность EXIF (обязательно)

Тег `Orientation` применяется, только если не противоречит уже портретному
кадру (`h > w` + тег 6/8 → тег протухший после кадрирования, игнорировать).
Кейс: «Горяинов ВБ стр. 3» (600×800 + Orientation=8) остается портретным.

## 5. HTTP-маршруты превью

- `POST /api/preview {file}` → единый шлюз диспетчера (все типы).
- `POST /api/word/preview {file}` → быстрый DOCX→HTML (`/cache/word/html/…`).
- `POST /api/excel/workbook|/sheet`, `/api/word|excel/render|page` — без изменений
  (работают через шимы ядра, HTML-оболочка iframe и протокол
  `launcher-sheet-*` сохранены побайтово).
- `POST /api/open-file` — нативное открытие (без изменений).

## 6. Замеры-эталоны (боевые файлы, 09.2026)

- Смета XLSX 833 КБ/5 листов: старый лист-0 — 4869 мс; ядро — все 5 за ~8.4 с,
  повтор/вкладки — ~4 мс; побайтовая идентичность HTML подтверждена.
- Договор DOCX 35 стр.: COM 20.5 с → HTML 1.2 с (≈16×).
- Счет XLS 35 КБ: xlrd→HTML 11 мс. Фото 6.5 МБ → 546 КБ за 642 мс.
