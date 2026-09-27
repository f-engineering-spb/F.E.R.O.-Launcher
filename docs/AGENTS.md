# Agent instructions for F-Engineering Launcher

## Source of truth

This repository is the source of truth for Launcher v3 application code.

Do not treat old local Launcher folders, Google Drive virtual disks, exported demos, or browser-served experiments as authoritative unless explicitly named as donor material.

## Runtime safety

- Keep live runtime execution on a local Windows `C:` workspace.
- Do not use Google Drive `G:` or `H:` as the live writable runtime root.
- Do not commit generated caches, manifests, rendered previews, logs, object files, customer documents, secrets, or temporary outputs.

## Development discipline

- Make small, verifiable changes.
- Keep frontend, backend, render pipeline, and runtime data separated.
- After each meaningful change, run a local smoke test and record the result in the handoff/summary.
- Preserve accepted viewer behavior unless the user explicitly changes it.
- Treat Cyrillic text as a first-class project requirement. All source, docs, JSON manifests, HTML, CSS, JS, and Python files must be UTF-8.
- Do not judge Cyrillic correctness from raw PowerShell output alone: Windows console encoding can display valid UTF-8 as mojibake. Verify text with UTF-8-aware checks, browser rendering, or Python `unicode_escape` inspection when needed.
- Static text assets served by the backend must include `charset=utf-8` for text content types.
- If mojibake appears in committed source text, stop and fix encoding before adding product logic.
- See `docs/ENCODING.md` before editing Russian UI labels, backend messages, manifests, or generated text.
- Always start Launcher v3 through `scripts/start_windows.cmd`; do not start the backend with ad-hoc PowerShell snippets when Russian paths, labels, or Google Drive paths are involved.
- Always run `scripts/check_encoding.cmd` before and after changes that touch Russian text, Windows paths, backend messages, frontend labels, or docs.

## Product direction

Launcher v3 should be built from verified bricks:

1. object list and object import/update/exclude;
2. tree and format filters;
3. accepted PDF render/cache pipeline: mass overview at `150 DPI`, active page at `300 DPI`;
4. accepted viewer controls: thumbnails, zoom, fit, pan, hand/arrow, medium/full modes;
5. later: Word, Excel, images, DWG strategy, and modules.

The intermediate product goal is documented in
`docs/PRODUCT_DIRECTION_VISUAL_TRIAGE_AND_MODULES.md`: Launcher is a local
visual-selection shell between project folders, native applications, and later
bounded processing modules.  It must not be positioned or implemented as a
replacement for CAD, Office, a full DMS, or a generic AI chat.

- Preserve original files and source folder structure; cache is derived and
  separately disposable.
- Prioritize time to first useful preview, transparent progress, warm-cache
  speed, and safe cache invalidation over an unrealistic promise that every
  cold document set is fully rendered in seconds.
- A future module must have declared inputs, output, confidence, and a
  reproducible log; never add an opaque “AI button”.

## Session handoff

Before the next product-development session, read `docs/HANDOFF_2026-08-12.md` and `docs/ROADMAP.md`.

The accepted state at the end of 2026-08-12:

- object shell works;
- wide tree browsing for long file names works;
- PDF overview/quality split works;
- viewer quality indicator works;
- DWG/PDF pairing works and is accepted for the current stage;
- DWG opens through the Windows native/default application;
- Vortex Word preview works and is accepted for the current stage;
- next direction is Excel preview support, plus optional DWG pair diagnostics and later image previews.

## Accepted DWG rules

- Do not build a custom CAD viewer inside Launcher v3 at this stage.
- Use PDF pairs as visual previews for DWG.
- Keep DWG files without PDF pairs visible; show a clear empty preview state instead of hiding them.
- Open DWG files through Windows default application from Launcher.
- Pairing must normalize project codes, separators, Cyrillic text, and service prefixes.
- Pairing must prefer missing a doubtful pair over connecting files from different sections.
- Explicit `Часть 1` / `Часть 2`, `узел`, and `сечение` conflicts must not be paired.

## Accepted Vortex Word rules

- Use the name `Вортекс` / `Vortex` for the universal preview module.
- Do not build a custom Word viewer.
- Preview `.doc` and `.docx` by converting through Microsoft Word to PDF cache, then use the accepted PDF render pipeline.
- Keep native opening available as `Открыть Word`.
- Treat `.gdoc` as a Google Docs shortcut, not as a local Word document.
- `.gdoc` must be visible in filters/tree and openable, but it should show a no-local-preview placeholder instead of being sent through Microsoft Word.
- Next Vortex target is Excel: `.xls` / `.xlsx`.

## Accepted Vortex Excel rules — 2026-08-13 checkpoint

Excel is a read-only navigation surface, not a second Excel editor.

- Support `.xlsx` / `.xlsm` through an HTML cache per workbook sheet.  Keep
  `.xls` visible and natively openable; add conversion support only after a
  separate compatibility test.
- The large view is the source of truth: it preserves sheet tabs, authored
  column widths, row heights, merges, basic fonts, fills, scrolling, zoom and
  hand panning.
- A workbook card/thumbnail must be a reduced viewport of that exact cached
  HTML sheet.  Never create a second simplified raster/table renderer for
  Excel thumbnails.  Separate renderers caused wrong text baselines, clipped
  rows and a mismatch between card and large view.
- Build the first selected sheet synchronously; prepare remaining sheets in
  the background, one at a time.  A cached sheet request must not reopen the
  workbook merely to rediscover its title.
- Do not show technical notices inside a worksheet canvas.  Show the complete
  workbook file name in the Launcher chrome above its sheet tabs.  Card labels
  must not use ellipses: wrap or allow horizontal reading of the full name.
- Before declaring Excel work accepted, run: Python compile, JS syntax check,
  `scripts/check_encoding.cmd`, then real browser QA with several ordinary
  workbooks, including one multi-sheet file.  Do not generalize from damaged
  or unusually large workbooks that Office itself struggles to open.

## Accepted Excel tab rule

- Excel sheet tabs are ordinary HTML buttons. Keep the proven rounded-button
  implementation with immediate active-state feedback.
- Do not reintroduce experimental tab shapes, overlapping tabs, SVG tab
  backgrounds, or pseudo-element browser-cap effects: they caused visual
  artefacts and delayed active-state feedback.

## Accepted viewer control and layer rules — 2026-08-13 checkpoint

- The bottom viewer panel is one shared control set for PDF, DWG previews,
  Word previews, and Excel sheets.  It must retain zoom, `Вписать`, rotate,
  hand/arrow, and the three viewing modes.  Do not make a reduced Excel or
  Word variant.
- The same panel must show a native-open action whenever an original file is
  available: `Открыть PDF`, `Открыть DWG`, `Открыть Word`, or `Открыть Excel`.
  Do not omit it for PDF/Word merely because their preview uses the PDF stage.
- `Вписать` must always have enough fixed width for its full Russian label;
  it must never be clipped or wrapped because another format has a different
  panel layout.
- Excel HTML is an overlay layer.  Its semantic `hidden` state must map to
  `display: none`; otherwise an inactive Excel iframe can cover a correctly
  rendered PDF, DWG preview, or Word page and make the large image look blank.
- Before changing a shared viewer control or layer, inspect both an Excel
  workbook and a PDF/DWG preview in the real browser.  Do not infer cross-
  format behaviour from one mode alone.

## Office handoff: multi-machine rules (2026-09-15, office PC)

These rules enable task transfer between the office PC, the laptop and
different agents without context loss. They supplement (never override)
the product rules above.

### 1. Purpose of the Launcher

Fast visual triage of project documents: open an object folder, skim the
tree, preview PDF / Word / Excel / images / DWG-via-PDF-pairs, open the
original in its native application. It is a viewing shell, not a CAD,
Office, DMS or AI-chat replacement.

### 2. Three interface zones (see `app/frontend/index.html`)

1. Left sidebar (`#sidebar`): object list (`#objectListState`) and file
   tree with search + format filters (`#treeState`, `#objectTree`).
2. Center viewer (`#pdfViewer` / `#pdfStage`): large page/sheet view,
   including the Excel overlay (`#excelViewer`).
3. Bottom strip: thumbnail ribbon (`#pdfThumbs`), shared viewer controls
   (`#viewerControls`) and the operation progress panel (`#progressPanel`).

### 3. Address and ports

- The only allowed application address is `http://127.0.0.1:8780/`
  (backend default `--port 8780`, see `app/backend/server.py`).
- Never use port 8000 for checking this project.

### 4. Sources of truth

- GitHub (`f-engineering-spb/f-engineering-launcher`) is the only source
  of code and commit history. Working branch: `dvg-main`.
- The local copy on SSD (e.g.
  `C:\Users\u301\Documents\Codex\f-engineering-launcher`) is the only
  working copy of the code.
- Google Drive `E:\Общие диски\021_F-Engineering_Knowledge_Library\09_Launcher`
  is a library/archive: documents, tasks, reports, artefacts. It is NOT
  a Git repository — never `git init` inside `E:` and never hand-copy
  code between `E:` and the local repo when GitHub can provide it.

### 5. Live data protection

- Do NOT touch live data on `H:` / `E:` on your own initiative (no scan,
  import, move, rename, delete of real object folders).
- Test fixtures live only locally (e.g. `C:\tmp\launcher_check_src`);
  real DWG/object folders are opened only on explicit instruction.

### 6. Heavy CAD ban (default)

- Do NOT run heavy DWG conversion or AutoCAD (daemon, oneshot, model
  space renders) without a separate explicit permission — every
  conversion opens AutoCAD on the user's machine.
- Lightweight checks (PDF, Word, Excel, images, tree, progress panel)
  need no extra permission.

### 7. Mandatory verification and reporting

- After every meaningful change run a local smoke test on
  `http://127.0.0.1:8780/` and record the result.
- Assignments follow `docs/TASK_TEMPLATE.md`; every session ends with a
  report in the shape of `docs/REPORT_TEMPLATE.md` (final message only,
  no intermediate questions).

### 8. Do not distract the user

- Never ask the user to run git commands, copy files, hunt processes or
  open a console. Do everything checkable yourself on the machine.
- Ask the user only for decisions you cannot resolve by inspection
  (credentials, live-data access, heavy CAD runs, product direction).
- Frontend edits require F5 in the browser (stale JS trap); backend
  Russian text requires `scripts/check_encoding.cmd` before/after.

### 9. Forbidden by default

- No health-check endpoints polling schemes, no cron/schedule tasks, no
  auto-sync and no background monitors unless a separate decision
  explicitly allows them.

### 10. Clean baseline freeze (2026-09-17, branch `review/v3.3-updates`)

- Before touching UI or key routes, read the docs: base flow is
  `docs/LAUNCHER_BASELINE_WORKFLOW.md`, the Load button is a frozen
  contract in `docs/LOAD_BUTTON_CONTRACT.md`.
- Without an explicit user order it is forbidden to: change main-button
  behaviour, add dialogs / built-in explorers / file browsers, replace
  system Windows dialogs with custom UI, "improve" the working flow, or
  resurrect previously removed mechanisms (`choose_folder.*`,
  `/api/browse`, `browse_filesystem()`, legacy COM openers,
  `/api/open-explorer`).
- A feature not described in the base flow is not assumed needed: ask
  or propose it separately first.
- DWG work must not break: object loading, file tree, ribbon + viewer,
  native operations, the Load button. Keep DWG changes in small
  separate commits.

### 11. Sole current baseline (dvg-main)

ЕДИНСТВЕННАЯ АКТУАЛЬНАЯ БАЗА

Единственная актуальная ветка проекта — dvg-main.

Работать нужно только из текущего репозитория и только после:

git switch dvg-main

История до текущего состояния намеренно не используется.
Старые ветки, теги, коммиты, архивы и внешние копии не являются источниками кода.
Не искать и не восстанавливать прежние версии.
Не сравнивать текущий код со старыми версиями.
Не использовать _BACKUP как рабочую базу.

Внешняя папка _BACKUP не является актуальной рабочей базой.
Единственная актуальная база — ветка dvg-main текущего репозитория.
Старые локальные копии и архивы запрещено использовать без прямого указания пользователя.

Актуальная рабочая папка — `C:\Users\a9379\Documents\Codex\FEngineering_Launcher_v3_CURRENT_COPY`.
Папку `FEngineering_Launcher_v3_OLD_DO_NOT_USE` не использовать.
Старые Git-истории не искать, старые версии не восстанавливать.
Единственная база — текущий HEAD ветки dvg-main.

### 12. Стандарты запуска CLI-агента (agy)

- **Исполняемый файл:** `agy`
- **Команда пакетного запуска:**
  ```powershell
  agy -p "$(Get-Content prompt.txt -Raw)" --dangerously-skip-permissions --mode accept-edits
  ```
- **Рабочая ветка:** `dvg-main`
- **Актуальная рабочая папка:** `C:\Users\a9379\Documents\Codex\FEngineering_Launcher_v3_CURRENT_COPY`
- **Запрещённая папка:** `C:\Users\a9379\Documents\Codex\FEngineering_Launcher_v3_OLD_DO_NOT_USE`
- **Хелпер запуска:** В корне репозитория создан скрипт `RUN_AGENT.cmd` для удобного запуска CLI-агента со считыванием задачи из `prompt.txt`.

### 13. Контракт нативного открытия файлов (NATIVE FILE OPENING CONTRACT)

Строго соблюдать правила из `docs/NATIVE_FILE_OPENING_CONTRACT.md`:
1. **Приоритет A:** Если в Параметрах сохранён валидный путь к EXE, запускать **ТОЛЬКО** этот EXE отдельными аргументами `[exe, file]`. Никакого fallback на Windows default, никакой подмены на AutoCAD/DWG TrueView.
2. **Приоритет B:** Если настройка пуста, использовать системную ассоциацию Windows (`ShellExecuteW` / `os.startfile`).
3. **Приоритет C:** Если настройка заполнена, но EXE не существует/недоступен, возвращать ошибку `configured_app_missing` **БЕЗ ТИХОГО FALLBACK**. Фронтенд выводит уведомление и открывает окно Параметров.
4. **Запрет AcLauncher:** Autodetect обязан фильтровать `AcLauncher.exe` и `ZwLauncher.exe` для `.dwg`/`.dxf`.
5. **Обязательный тест:** Перед коммитом всегда запускать:
   ```powershell
   python -m unittest tests/test_native_open_contract.py -v
   ```


