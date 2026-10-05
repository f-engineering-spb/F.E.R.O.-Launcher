// Perf-маяк в DevTools Console (замеры фронта, видны на вкладке Console, F12).
function beaconPerf(label, ms, extra = "") {
  try {
    const text = `[perf] ${label} ${Math.round(ms)}ms ${extra}`.trim();
    if (typeof console !== "undefined" && typeof console.log === "function") console.log(text);
  } catch (_) {}
}

const state = {
  objects: [],
  selectedObjectId: null,
  currentManifest: null,
  activeFilters: new Set(),
  selectedPaths: new Set(),
  lastSelectedIndex: null,
  visibleRows: [],
  collapsedFolders: new Set(),
  progressTimer: null,
  progressCancelled: false,
  renderEpoch: 0,
  activeTxtPath: "",
  operationController: null,
  operationControllers: [],
  renderedPages: [],
  activePageUrl: "",
  activePageKey: "",
  activeNativePath: "",
  appSettings: null,
  revealedPath: "",
  excelWorkbook: null,
  excelWorkbooks: [],
  excelWorkbookIndex: 0,
  excelSheetIndex: 0,
  excelScale: 1,
  wordDoc: null,
  highQualityPages: new Map(),
  pdfPairIndex: new Map(),
  pairless: false,
  diffStatus: new Map(),
  diffRemovedNodes: [],
  diffSummary: null,
  diffFilter: false,
  diffPollTimer: null,
  viewMode: "standard",
  activeNavZone: "tree",
  thumbScale: 1,
  treeScale: 1,
  view: {
    scale: 1,
    fitScale: 1,
    rotation: 0,
    userZoomed: false,
    panX: 0,
    panY: 0,
    panMode: true,
    dragging: false,
    dragStartX: 0,
    dragStartY: 0,
    dragPanX: 0,
    dragPanY: 0,
  },
};

const els = {
  shell: document.querySelector(".shell"),
  sidebar: document.getElementById("sidebar"),
  load: document.getElementById("loadObject"),
  refresh: document.getElementById("refreshObject"),
  display: document.getElementById("displayObject"),
  exclude: document.getElementById("excludeObject"),
  cancel: document.getElementById("cancelOperation"),
  progressPanel: document.getElementById("progressPanel"),
  progressLabel: document.getElementById("progressLabel"),
  progressValue: document.getElementById("progressValue"),
  progressFill: document.getElementById("progressFill"),
  progressDetail: document.getElementById("progressDetail"),
  objectListState: document.getElementById("objectListState"),
  treeState: document.getElementById("treeState"),
  objectList: document.getElementById("objectList"),
  objectStats: document.getElementById("objectStats"),
  backToObjects: document.getElementById("backToObjects"),
  backToTree: document.getElementById("backToTree"),
  treeSearch: document.getElementById("treeSearch"),
  formatStrip: document.getElementById("formatStrip"),
  objectTree: document.getElementById("objectTree"),
  pdfThumbs: document.getElementById("pdfThumbs"),
  pdfViewer: document.getElementById("pdfViewer"),
  pdfStage: document.getElementById("pdfStage"),
  pdfPageImage: document.getElementById("pdfPageImage"),
  txtViewer: document.getElementById("txtViewer"),
  txtContent: document.getElementById("txtContent"),
  excelViewer: document.getElementById("excelViewer"),
  excelBookTitle: document.getElementById("excelBookTitle"),
  excelTabs: document.getElementById("excelTabs"),
  excelTabsLeft: document.getElementById("excelTabsLeft"),
  excelTabsRight: document.getElementById("excelTabsRight"),
  excelSheetFrame: document.getElementById("excelSheetFrame"),
  wordViewer: document.getElementById("wordViewer"),
  wordDocTitle: document.getElementById("wordDocTitle"),
  wordMeta: document.getElementById("wordMeta"),
  wordDocFrame: document.getElementById("wordDocFrame"),
  wordOpenNative: document.getElementById("wordOpenNative"),
  viewerEmpty: document.getElementById("viewerEmpty"),
  viewerControls: document.getElementById("viewerControls"),
  qualityBadge: document.getElementById("qualityBadge"),
  pagePosition: document.getElementById("pagePosition"),
  viewZoomOut: document.getElementById("viewZoomOut"),
  viewZoomIn: document.getElementById("viewZoomIn"),
  viewFit: document.getElementById("viewFit"),
  viewRotate: document.getElementById("viewRotate"),
  viewOpenNative: document.getElementById("viewOpenNative"),
  viewPanMode: document.getElementById("viewPanMode"),
  contextMenu: document.getElementById("contextMenu"),
  viewStandardMode: document.getElementById("viewStandardMode"),
  viewFullMode: document.getElementById("viewFullMode"),
  scaleWidget: document.getElementById("scaleWidget"),
  scaleResetBtn: document.getElementById("scaleResetBtn"),
  scaleMinusBtn: document.getElementById("scaleMinusBtn"),
  scalePlusBtn: document.getElementById("scalePlusBtn"),
  btnSettings: document.getElementById("btnSettings"),
  settingsModal: document.getElementById("settingsModal"),
  settingsCloseBtn: document.getElementById("settingsCloseBtn"),
  settingsCancelBtn: document.getElementById("settingsCancelBtn"),
  settingsFillEmptyBtn: document.getElementById("settingsFillEmptyBtn"),
  settingsReplaceBtn: document.getElementById("settingsReplaceBtn"),
  settingsOpenWinSettingsBtn: document.getElementById("settingsOpenWinSettingsBtn"),
  settingsDwgHint: document.getElementById("settingsDwgHint"),
  settingsSaveBtn: document.getElementById("settingsSaveBtn"),

};

function text(value, fallback = "") {
  return value == null ? fallback : String(value);
}

function formatExtensions(extensions = {}) {
  const entries = Object.entries(extensions).sort((left, right) => left[0].localeCompare(right[0], "ru"));
  return entries.length ? entries.map(([ext, count]) => `${ext}: ${count}`).join(" · ") : "форматов нет";
}

function selectedObject() {
  return state.objects.find((item) => item.id === state.selectedObjectId) || null;
}

function inTreeMode() {
  return els.treeState.classList.contains("active");
}

function setMode(mode) {
  els.objectListState.classList.toggle("active", mode === "objects");
  els.treeState.classList.toggle("active", mode === "tree");
  els.shell.classList.toggle("tree-browse", false);
  if (els.backToTree) els.backToTree.disabled = true;
  if (mode === "objects") {
    els.pdfViewer.classList.add("empty");
    updateViewerBanner('Политика "Локальный компьютер"');
  }
  updateObjectButtons();
}

function setTreeBrowseMode(enabled) {
  els.shell.classList.toggle("tree-browse", Boolean(enabled));
  if (els.backToTree) els.backToTree.disabled = Boolean(enabled) || !state.currentManifest?.tree;
}

function updateObjectButtons() {
  const hasSelection = Boolean(state.selectedObjectId);
  const tree = inTreeMode();
  if (els.backToObjects) els.backToObjects.disabled = !tree;
  if (els.backToTree) els.backToTree.disabled = !tree;
  els.load.disabled = false;
  els.refresh.disabled = !hasSelection;
  els.display.disabled = !hasSelection && !tree;
  els.exclude.disabled = !hasSelection;
}

function updateObjectStats() {
  const object = selectedObject();
  if (!object) {
    els.objectStats.textContent = "Объект не выбран";
    return;
  }
  els.objectStats.innerHTML = [
    `<strong>${text(object.name)}</strong>`,
    `Папок: ${object.statistics?.folders || 0}`,
    `Файлов: ${object.statistics?.files || 0}`,
    formatExtensions(object.statistics?.extensions || {}),
    text(object.rootPath),
  ].join("<br>");
}

function renderObjectList() {
  els.objectList.replaceChildren();
  if (!state.objects.length) {
    const empty = document.createElement("div");
    empty.className = "empty-note";
    empty.textContent = "Объекты ещё не загружены.";
    els.objectList.append(empty);
    updateObjectStats();
    updateObjectButtons();
    return;
  }

  state.objects.forEach((object) => {
    const row = document.createElement("div");
    row.className = "object-row";
    row.dataset.objectId = object.id;
    row.classList.toggle("selected", object.id === state.selectedObjectId);
    row.title = object.rootPath;

    const main = document.createElement("div");
    main.className = "object-main";

    const icon = document.createElement("span");
    icon.className = "object-icon";
    icon.innerHTML = `<svg viewBox="0 0 16 16" width="15" height="15" fill="#f59e0b"><path d="M1.5 3A1.5 1.5 0 0 0 0 4.5v7A1.5 1.5 0 0 0 1.5 13h13a1.5 1.5 0 0 0 1.5-1.5v-5A1.5 1.5 0 0 0 14.5 5H7.707L6.354 3.646A1.5 1.5 0 0 0 5.293 3.207L1.5 3z"/></svg>`;

    const title = document.createElement("span");
    title.className = "object-title";
    title.textContent = object.name;

    const badge = document.createElement("span");
    badge.className = "object-badge";
    badge.textContent = `(${object.statistics?.files || 0})`;

    main.append(icon, title, badge);

    row.append(main);
    row.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      event.stopPropagation();
      showFileContextMenu(event.clientX, event.clientY, {
        path: object.rootPath,
        isDir: true,
        ext: "",
      });
    });

    row.addEventListener("click", (event) => {
      event.stopPropagation();
      state.selectedObjectId = state.selectedObjectId === object.id ? null : object.id;
      updateObjectListSelection();
      updateObjectButtons();
      updateObjectStats();
      if (state.selectedObjectId) {
        updateViewerBanner(object.name);
      } else {
        updateViewerBanner('Политика "Локальный компьютер"');
      }
    });
    row.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      state.selectedObjectId = object.id;
      updateObjectListSelection();
      updateObjectButtons();
      updateObjectStats();
      openSelectedObject().catch(showOperationError);
    });
    els.objectList.append(row);
  });
  updateObjectStats();
  updateObjectButtons();
}

function updateObjectListSelection() {
  [...els.objectList.querySelectorAll(".object-row")].forEach((row) => {
    row.classList.toggle("selected", row.dataset.objectId === state.selectedObjectId);
  });
}

async function loadObjectSummaries() {
  const response = await fetch("/api/objects");
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || "Список объектов не получен");
  state.objects = payload.items || [];
  if (state.selectedObjectId && !state.objects.some((item) => item.id === state.selectedObjectId)) {
    state.selectedObjectId = null;
  }
  renderObjectList();
}

function createOperationController() {
  if (state.operationController) state.operationController.abort();
  state.operationController = new AbortController();
  return state.operationController;
}

function startProgress(label, detail) {
  stopProgress(false, false);
  state.progressCancelled = false;
  els.progressFill.classList.remove("is-error");
  els.progressPanel.hidden = false;
  els.progressLabel.textContent = label;
  els.progressDetail.textContent = detail;
  els.progressValue.textContent = "0%";
  els.progressFill.style.width = "0%";
  // Полоса движется всегда, пока идёт операция: пакетные циклы выставляют
  // state.progressEase { base, spent0, estMs } — оценку оставшегося времени,
  // тик плавно тянет шкалу к 99 по elapsed/ETA (мёртвой зоны на 92 нет).
  // Реальные завершения выставляют точные цифры через refreshRenderProgress.
  // Строка деталей всегда честная (счётчики файлов/пачек), плюс CSS-пульс.
  state.progressShown = 0;
  state.progressEase = null;
  els.progressFill.classList.add("is-busy");
  state.progressTimer = setInterval(() => {
    if (state.progressCancelled) return;
    const ease = state.progressEase;
    let target;
    if (ease && ease.estMs > 0) {
      const spent = Date.now() - ease.spent0;
      target = ease.base + (99 - ease.base) * (spent / (spent + ease.estMs));
    } else {
      // Без оценки времени полоска медленно ползёт к 99 и не замирает на 92.
      target = Math.min(99, (state.progressShown || 0) + 0.4);
    }
    state.progressShown = Math.max(state.progressShown, Math.min(99, target));
    els.progressValue.textContent = `${Math.floor(state.progressShown)}%`;
    els.progressFill.style.width = `${state.progressShown}%`;
  }, 500);
}

// Оценка оставшегося времени пачки по измеренной скорости (мс).
// До первых завершений — дефолт 10 с/файл (облако), минимум 5 с.
function estimateBatchRemainMs(elapsedMs, doneFiles, totalFiles) {
  const remaining = Math.max(0, totalFiles - doneFiles);
  if (remaining <= 0) return 0;
  const perFile = doneFiles > 0 ? elapsedMs / doneFiles : 10000;
  return Math.min(600000, Math.max(5000, perFile * remaining));
}

function finishProgress(detail = "Готово") {
  if (state.progressTimer) clearInterval(state.progressTimer);
  state.progressTimer = null;
  state.progressEase = null;
  state.operationController = null;
  els.progressFill.classList.remove("is-busy");
  els.progressDetail.textContent = detail;
  els.progressValue.textContent = "100%";
  els.progressFill.style.width = "100%";
  setTimeout(() => {
    if (!state.progressTimer) els.progressPanel.hidden = true;
  }, 1200);
}

function stopProgress(cancelled = true, abortRequest = true) {
  if (state.progressTimer) clearInterval(state.progressTimer);
  state.progressTimer = null;
  state.progressEase = null;
  els.progressFill.classList.remove("is-busy");
  state.progressCancelled = cancelled;
  if (abortRequest) {
    if (state.operationController) {
      try { state.operationController.abort(); } catch (_) {}
      state.operationController = null;
    }
    if (state.operationControllers?.length) {
      state.operationControllers.forEach((controller) => {
        try { controller.abort(); } catch (_) {}
      });
      state.operationControllers = [];
    }
  }
  els.progressPanel.hidden = true;
  if (cancelled) {
    els.progressFill.style.width = "0%";
    els.progressValue.textContent = "";
    if (state._ribbonScrollObserver) {
      state._ribbonScrollObserver.disconnect();
      state._ribbonScrollObserver = null;
    }
    // Отмена возвращает в нейтральное состояние: снимаем выделение и фильтры,
    // чтобы не оставалось залипшей подсветки и следующий показ шёл с чистого листа.
    state.selectedPaths.clear();
    state.activeFilters.clear();
    state.revealedPath = "";
    renderFormats();
    renderTree();
  }
}

function showOperationError(error) {
  if (error?.name === "AbortError") return;
  // Never disguise a failed render as a completed 100% operation.  Keep the
  // factual error on screen long enough for both the user and QA to see it.
  if (state.progressTimer) clearInterval(state.progressTimer);
  state.progressTimer = null;
  state.operationController = null;
  els.progressLabel.textContent = "Операция не выполнена";
  els.progressDetail.textContent = `Ошибка: ${error.message || error}`;
  els.progressValue.textContent = "Ошибка";
  els.progressFill.style.width = "100%";
  els.progressFill.classList.add("is-error");
  setTimeout(() => {
    if (!state.progressTimer) {
      els.progressPanel.hidden = true;
      els.progressFill.classList.remove("is-error");
    }
  }, 8000);
}

// Non-blocking warning (not an error): shows why a command may have silently
// misbehaved, e.g. a Windows long-path limit, without failing the operation.
function showNotice(message) {
  if (state.progressTimer) clearInterval(state.progressTimer);
  state.progressTimer = null;
  els.progressFill.classList.remove("is-busy");
  els.progressLabel.textContent = "Предупреждение";
  els.progressDetail.textContent = message;
  els.progressValue.textContent = "";
  els.progressFill.style.width = "100%";
  els.progressPanel.hidden = false;
  setTimeout(() => {
    if (!state.progressTimer) {
      els.progressPanel.hidden = true;
    }
  }, 9000);
}



function getNativeAppLabel(ext = "") {
  const custom = settingsCustomExe(ext);
  if (custom) {
    const filename = custom.split(/[\\/]/).pop();
    const base = filename.replace(/\.exe$/i, "");
    const baseLower = base.toLowerCase();
    if (baseLower === "acad") return "AutoCAD";
    if (baseLower.includes("trueview")) return "DWG TrueView";
    if (baseLower === "chrome") return "Google Chrome";
    if (baseLower === "msedge") return "Microsoft Edge";
    if (baseLower === "winword") return "Word";
    if (baseLower === "excel") return "Excel";
    if (baseLower === "powerpnt") return "PowerPoint";
    if (baseLower === "acrobat" || baseLower === "acrord32") return "Adobe Acrobat";
    if (baseLower === "vlc") return "VLC";
    if (baseLower === "mspaint") return "Paint";
    if (baseLower === "notepad") return "Блокнот";
    return base;
  }
  const e = String(ext).toUpperCase().replace(/^\./, "");
  switch (e) {
    case "DWG":
    case "DXF": return "AutoCAD / DWG";
    case "PDF": return "PDF-просмотрщике";
    case "DOC":
    case "DOCX":
    case "RTF":
    case "ODT": return "Word";
    case "XLS":
    case "XLSX":
    case "XLSM":
    case "XLSB":
    case "CSV":
    case "ODS": return "Excel";
    case "PPT":
    case "PPTX":
    case "ODP": return "PowerPoint";
    case "JPG":
    case "JPEG":
    case "PNG":
    case "BMP":
    case "WEBP":
    case "SVG":
    case "TIF":
    case "TIFF":
    case "GIF":
    case "ICO": return "Просмотре фото";
    case "MP4":
    case "AVI":
    case "MOV":
    case "MKV":
    case "WMV":
    case "MP3":
    case "WAV": return "Медиаплеере";
    case "ZIP":
    case "RAR":
    case "7Z":
    case "TAR":
    case "GZ": return "Архиваторе";
    case "TXT":
    case "LOG":
    case "INI":
    case "CFG":
    case "JSON":
    case "XML": return "Блокноте";
    default: return e ? `${e}` : "программе";
  }
}


let isOpeningFile = false;
let lastOpenedPath = "";
let lastOpenedTime = 0;

async function openFileByPath(path, action) {
  if (!path) return;
  const now = Date.now();
  if (isOpeningFile && (now - lastOpenedTime < 1800)) {
    console.warn("[Launcher] openFileByPath ignored: already in progress", path);
    return;
  }
  if (lastOpenedPath === path && (now - lastOpenedTime < 1800)) {
    console.warn("[Launcher] openFileByPath ignored: rapid repeat call for same path", path);
    return;
  }
  isOpeningFile = true;
  lastOpenedPath = path;
  lastOpenedTime = now;

  const act = String(action || "explorer");
  let actLabel = "Проводнике Windows";
  if (act === "system" || act === "native") {
    actLabel = "программе по умолчанию";
  }
  startProgress("Запуск: " + actLabel, path);
  try {
    const response = await fetch("/api/open-file", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path, action: act }),
    });
    const payload = await response.json();
    if (!response.ok) {
      if (payload.error_code === "configured_app_missing") {
        const ext = (payload.details && payload.details.extension)
          ? payload.details.extension.toUpperCase().replace(".", "")
          : "DWG";
        const msg = `Для файлов ${ext} указано недоступное приложение. Проверьте Параметры → Приложения.`;
        showNotice(msg);
        showToast(msg);
        openSettingsModal(payload.details && payload.details.extension);
        return;
      }
      throw new Error(payload.error || "Не удалось открыть файл");
    }
    finishProgress("Открыто: " + actLabel);
    showToast("Файл запущен: " + actLabel);
    if (payload.longPathWarning) showNotice(payload.longPathWarning);
  } catch (error) {
    console.error("[Launcher] Failed to open file:", error);
    showOperationError(error);
  } finally {
    setTimeout(() => {
      isOpeningFile = false;
    }, 1200);
  }
}

function escapeHtml(value = "") {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

let isPickerOpen = false;

async function triggerSystemPicker() {
  // Единственный механизм загрузки: системный диалог выбора папки Windows.
  // Никаких самописных окон, прогресса, HTML-ответов и «запросов к агенту»:
  // один клик -> одно окно FolderBrowserDialog -> importObjectByPath.
  if (isPickerOpen) return "";
  isPickerOpen = true;
  if (els.load) els.load.disabled = true;
  try {
    const res = await fetch("/api/choose-folder", {
      method: "POST",
      headers: { Accept: "application/json" },
    });
    const contentType = res.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      throw new Error("Сервер вернул не JSON при выборе папки");
    }
    const data = await res.json();
    if (data && data.path) {
      await importObjectByPath(data.path);
      return data.path;
    }
    if (data && data.error) {
      showOperationError(new Error(data.error));
    }
    return "";
  } catch (err) {
    showOperationError(err);
    return "";
  } finally {
    isPickerOpen = false;
    if (els.load) els.load.disabled = false;
  }
}

async function importObjectByPath(path, forceRefresh = false) {
  if (!path || !path.trim()) return;
  const controller = createOperationController();
  startProgress(forceRefresh ? "Обновление объекта" : "Загрузка объекта", path);
  state.progressEase = {
    base: 8,
    spent0: Date.now(),
    estMs: 120000,
  };
  try {
    const response = await fetch("/api/objects/import", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: path.trim() }),
      signal: controller.signal,
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Объект не загружен");
    if (state.progressCancelled) return;
    state.selectedObjectId = payload.id;
    await loadObjectSummaries();
    finishProgress(forceRefresh ? "Объект обновлён" : "Объект загружен");
  } catch (error) {
    showOperationError(error);
  }
}

async function importObject(forceRefresh = false) {
  if (forceRefresh) {
    const object = selectedObject();
    if (object?.rootPath) {
      await importObjectByPath(object.rootPath, true);
    }
    return;
  }
  // Единственный сценарий: системный диалог выбора папки Windows.
  await triggerSystemPicker().catch((err) => {
    showOperationError(err);
  });
}


async function excludeSelectedObject() {
  const object = selectedObject();
  if (!object) return;
  const controller = createOperationController();
  startProgress("Исключение объекта", object.name);
  const response = await fetch("/api/objects/exclude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: object.id }),
    signal: controller.signal,
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || "Объект не исключён");
  state.selectedObjectId = null;
  state.currentManifest = null;
  await loadObjectSummaries();
  finishProgress("Объект исключён");
}

function getCachedFlatNodes() {
  if (!state.currentManifest?.tree) return [];
  if (!state.currentManifest._flatNodes) {
    state.currentManifest._flatNodes = flattenTree(state.currentManifest.tree);
  }
  return state.currentManifest._flatNodes;
}

function flattenTree(rootNode, result = []) {
  if (!rootNode) return result;
  const stack = [rootNode];
  const seen = new Set();
  while (stack.length) {
    const node = stack.pop();
    if (!node) continue;
    const key = node.path || node;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(node);
    const children = node.children;
    if (children && children.length) {
      for (let i = children.length - 1; i >= 0; i--) {
        stack.push(children[i]);
      }
    }
  }
  return result;
}

function fileStem(name = "") {
  const value = String(name);
  const dotIndex = value.lastIndexOf(".");
  return dotIndex > 0 ? value.slice(0, dotIndex) : value;
}

function normalizePairName(name = "") {
  return fileStem(name)
    .toLocaleLowerCase("ru")
    .replace(/ё/g, "е")
    .replace(/\b(?:ap|ar)\s*(?=\d)/g, "ар")
    .replace(/ар\s*(?=\d)/g, "ар ")
    .replace(/проеомв/g, "проемов")
    .replace(/срп\s*[_-]?\s*7/g, " ")
    .replace(/лист/g, " ")
    .replace(/[+]/g, " ")
    .replace(/[_\-.()[\]{},]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function pairTokens(name = "") {
  return normalizePairName(name)
    .split(" ")
    .map((token) => token.trim())
    .filter((token) => token.length > 1)
    .filter((token) => !["срп", "для", "всех", "отм", "осях", "оси", "на"].includes(token));
}

function escapeRegExp(value = "") {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractCodeParts(normalized = "") {
  const codeMatch = normalized.match(/(?:^|\s)ар\s*(\d)\s+(\d)\s+(\d)(?=\s|$)/u)
    || normalized.match(/(?:^|\s)ар\s*(\d)\s+(\d)(?=\s|$)/u);
  return codeMatch ? codeMatch.slice(1).filter(Boolean) : [];
}

function stripPairServiceNoise(normalized = "") {
  let value = normalized
    .replace(/\bсрп\s*7\b/g, " ")
    .replace(/\bлист\b/g, " ");
  const codeParts = extractCodeParts(value);
  if (codeParts.length) {
    const codePattern = new RegExp(`(?:^|\\s)ар\\s*${codeParts.map(escapeRegExp).join("\\s+")}(?=\\s|$)`, "g");
    value = value.replace(codePattern, " ");
  }
  return value
    .replace(/\b50\s+\d{1,3}\b/g, " ")
    .replace(/^\s*\d{1,3}\s+/, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stemPairToken(token = "") {
  if (token.length <= 4) return token;
  return token.replace(/(иями|ями|ами|ого|его|ому|ему|ыми|ими|ых|их|ая|яя|ое|ее|ые|ие|ый|ий|ой|ую|юю|ом|ем|ах|ях|ов|ев|ей|ам|ям|а|я|ы|и|е|у|ю|о)$/u, "");
}

function semanticPairTokens(normalized = "") {
  return stripPairServiceNoise(normalized)
    .split(" ")
    .map((token) => stemPairToken(token.trim()))
    .filter((token) => token.length > 1)
    .filter((token) => !["срп", "для", "всех", "отм", "осях", "оси", "на"].includes(token));
}

function tokenBigrams(tokens = []) {
  const result = new Set();
  for (let index = 0; index < tokens.length - 1; index += 1) {
    result.add(`${tokens[index]} ${tokens[index + 1]}`);
  }
  return result;
}

function longestCommonRun(left, right) {
  if (!left?.length || !right?.length) return 0;
  let best = 0;
  const previous = new Array(right.length + 1).fill(0);
  const current = new Array(right.length + 1).fill(0);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      current[rightIndex] = left[leftIndex - 1] === right[rightIndex - 1] ? previous[rightIndex - 1] + 1 : 0;
      if (current[rightIndex] > best) best = current[rightIndex];
    }
    previous.splice(0, previous.length, ...current);
    current.fill(0);
  }
  return best;
}

function pairFingerprint(name = "") {
  const normalized = normalizePairName(name);
  const codeParts = extractCodeParts(normalized);
  const code = codeParts.length ? `ар${codeParts.join(".")}` : "";
  const sheetMatch = normalized.match(/\b50\s+(\d{1,3})\b/);
  const sheet = sheetMatch ? String(Number(sheetMatch[1])) : "";
  const codeNumberPattern = codeParts.length
    ? new RegExp(`(?:^|\\s)ар\\s*${codeParts.join("\\s+")}\\s+(\\d{1,3})(?=\\s|$)`, "u")
    : /(?:^|\s)ар\s*(?:\d\s*)+\s+(\d{1,3})(?=\s|$)/u;
  const codeNumberMatch = normalized.match(codeNumberPattern);
  const codeNumber = codeNumberMatch ? String(Number(codeNumberMatch[1])) : "";
  const detailMatches = [...normalized.matchAll(/(?:^|\s)(узел|сечение|часть)\s+([a-zа-я0-9]+(?:\s*[-–]\s*[a-zа-я0-9]+)?)(?=\s|$)/gu)]
    .map((match) => `${match[1]} ${match[2].replace(/\s+/g, "")}`);
  const details = new Set(detailMatches);
  const floorMatch = normalized.match(/\b(\d+\s*[-–]\s*\d+|\d+)\s+(?:го\s+)?этаж/u);
  const floor = floorMatch ? floorMatch[1].replace(/\s+/g, "") : "";
  const technical = /техническ|тех\s*пространств/u.test(normalized);
  const types = new Set();
  if (/маркировочный\s+план/u.test(normalized)) types.add("маркировочный план");
  if (/общие\s+данные/u.test(normalized)) types.add("общие данные");
  if (/ведомост/u.test(normalized) && /отделк/u.test(normalized)) types.add("ведомость отделки");
  if (/экспликац/u.test(normalized) && /полов/u.test(normalized)) types.add("экспликация полов");
  if (/схем/u.test(normalized) && /двер/u.test(normalized)) types.add("схемы дверей");
  if (/колористическ/u.test(normalized)) types.add("колористическое решение");
  if (/фасад/u.test(normalized)) types.add("фасад");
  const axisTokens = new Set([...normalized.matchAll(/\b7\s+2\s+[0-9а-я]+\b/gu)].map((match) => match[0].replace(/\s+/g, ".")));
  const tokens = new Set(pairTokens(name));
  const semanticText = stripPairServiceNoise(normalized);
  const semanticTokensList = semanticPairTokens(normalized);
  const semanticTokensSet = new Set(semanticTokensList);
  return {
    normalizedName: normalized,
    semanticText,
    semanticTokensList,
    semanticTokensSet,
    semanticBigrams: tokenBigrams(semanticTokensList),
    code,
    sheet: sheet || codeNumber,
    details,
    floor,
    technical,
    types,
    axisTokens,
    tokens,
  };
}

function buildPdfPairIndex(nodes) {
  const list = [];
  const byName = new Map();
  for (const node of nodes) {
    if (node.type === "file" && node.extension === "PDF") {
      const normalizedName = normalizePairName(node.name || "");
      const entry = {
        node,
        normalizedName,
        _fingerprint: null,
        get fingerprint() {
          if (!this._fingerprint) {
            this._fingerprint = pairFingerprint(this.node.name);
          }
          return this._fingerprint;
        },
      };
      list.push(entry);
      if (normalizedName) {
        let arr = byName.get(normalizedName);
        if (!arr) {
          arr = [];
          byName.set(normalizedName, arr);
        }
        arr.push(entry);
      }
    }
  }
  list.byName = byName;
  return list;
}

// Сколько хвостовых сегментов папок совпадает у DWG и кандидата PDF.
// DVG/5.1/КМД/x.dwg + PDF/5.1/КМД/x.pdf → 2, чужой каталог → 0.
function mirrorScore(dwgPath = "", pdfPath = "") {
  const dw = String(dwgPath).replace(/\//g, "\\").toLowerCase().split("\\").slice(0, -1);
  const pf = String(pdfPath).replace(/\//g, "\\").toLowerCase().split("\\").slice(0, -1);
  let score = 0;
  while (score < dw.length && score < pf.length
    && dw[dw.length - 1 - score] === pf[pf.length - 1 - score]) {
    score += 1;
  }
  return score;
}

// Общий бюджет нечёткого сопоставления DWG→PDF (мс) на один проход
// подбора превью. Полный перебор на объекте вида «2687 DWG × 6195 PDF» —
// это ~16 млн сравнений ≈ 9 минут мёртвого окна (замер 20:32, объект 03_ЖК МОД).
const PAIR_FUZZY_BUDGET_MS = 2500;
let pairFuzzyBudgetUntil = Infinity;

function intersectionSize(left, right) {
  let count = 0;
  for (const item of left) {
    if (right.has(item)) count += 1;
  }
  return count;
}

function comparePairFingerprints(dwg, pdf) {
  if (dwg.normalizedName === pdf.normalizedName) return { score: 100, confidence: "exact" };
  if (dwg.code && pdf.code && dwg.code !== pdf.code) return null;
  const detailOverlap = intersectionSize(dwg.details, pdf.details);
  if (dwg.details.size && pdf.details.size && detailOverlap === 0) return null;

  let score = 0;
  if (dwg.code && pdf.code && dwg.code === pdf.code) score += 32;
  if (dwg.sheet && pdf.sheet && dwg.sheet === pdf.sheet) score += 8;
  if (dwg.sheet && pdf.sheet && dwg.sheet !== pdf.sheet) score -= 4;
  if (detailOverlap) score += 34;
  if (dwg.floor && pdf.floor && dwg.floor === pdf.floor) score += 28;
  if (dwg.technical && pdf.technical) score += 28;

  const typeOverlap = intersectionSize(dwg.types, pdf.types);
  score += typeOverlap * 16;

  const axisOverlap = intersectionSize(dwg.axisTokens, pdf.axisTokens);
  score += Math.min(18, axisOverlap * 6);

  const sharedTokens = intersectionSize(dwg.tokens, pdf.tokens);
  const tokenRatio = sharedTokens / Math.max(dwg.tokens.size, pdf.tokens.size, 1);
  score += Math.round(tokenRatio * 28);

  const semanticOverlap = intersectionSize(dwg.semanticTokensSet, pdf.semanticTokensSet);
  const semanticRatio = semanticOverlap / Math.max(dwg.semanticTokensSet.size, pdf.semanticTokensSet.size, 1);
  score += Math.round(semanticRatio * 46);

  const bigramOverlap = intersectionSize(dwg.semanticBigrams, pdf.semanticBigrams);
  score += Math.min(36, bigramOverlap * 12);

  const commonRun = longestCommonRun(dwg.semanticTokensList, pdf.semanticTokensList);
  if (commonRun >= 5) score += 42;
  else if (commonRun >= 4) score += 32;
  else if (commonRun >= 3) score += 22;

  const commonSubstring = longestCommonRun(dwg.semanticText, pdf.semanticText);
  if (commonSubstring >= 32) score += 32;
  else if (commonSubstring >= 22) score += 22;
  else if (commonSubstring >= 16) score += 12;

  if (dwg.semanticText && pdf.semanticText && dwg.semanticText === pdf.semanticText) score += 52;

  const hasSemanticAnchor = Boolean(
    detailOverlap
    || (dwg.floor && pdf.floor && dwg.floor === pdf.floor)
    || (dwg.technical && pdf.technical)
    || typeOverlap >= 1
    || axisOverlap >= 2
    || semanticRatio >= 0.5
    || bigramOverlap >= 1
    || commonRun >= 3
    || commonSubstring >= 22
  );
  if (!hasSemanticAnchor && score < 82) return null;
  if (score < 68) return null;
  return { score, confidence: score >= 78 ? "strong" : "probable" };
}

function findPdfPairForDwg(dwgNode, pdfIndex, options = {}) {
  if (!dwgNode || dwgNode.extension !== "DWG") return null;
  const list = Array.isArray(pdfIndex) ? pdfIndex : [];
  if (!list.length) return null;
  // Точное имя ищется по всему объекту, папка не важна:
  // пара обычно в соседней папке PDF. Среди тёзок — зеркальный путь.
  const key = normalizePairName(dwgNode.name || "");
  const candidates = (key && pdfIndex?.byName?.get(key)) || [];
  if (candidates.length) {
    let pick = candidates[0];
    if (candidates.length > 1) {
      let bestScore = -1;
      for (const cand of candidates) {
        const sc = mirrorScore(dwgNode.path, cand.node?.path);
        if (sc > bestScore) { bestScore = sc; pick = cand; }
      }
    }
    return { node: pick.node, score: 100, confidence: "exact" };
  }
  // Массовый показ: только точные пары (быстро и детерминированно).
  // Нечёткий поиск оставлен одиночным файлам — там он занимает ~0,2 с.
  if (options.fuzzy === false) return null;
  const dwgFingerprint = pairFingerprint(dwgNode.name);
  // Нечёткий поиск — в пределах общего бюджета. Когда бюджет исчерпан
  // (гигантский объект), возвращаем лучшее из найденного или null:
  // тогда файл пойдёт по честному пути DWG_MODEL, а не повесит окно.
  let best = null;
  const budgetUntil = Number.isFinite(pairFuzzyBudgetUntil) ? pairFuzzyBudgetUntil : Date.now() + 100;
  for (let i = 0; i < list.length; i++) {
    if ((i & 63) === 0 && Date.now() > budgetUntil) break;
    const comparison = comparePairFingerprints(dwgFingerprint, list[i].fingerprint);
    if (comparison && (!best || comparison.score > best.score)) best = { node: list[i].node, ...comparison };
  }
  return best;
}

let nodeMatchCache = null;

function nodeMatches(node, visiting = new Set()) {
  if (!node) return false;
  const query = els.treeSearch.value.trim().toLocaleLowerCase("ru");
  const hasFormatFilter = state.activeFilters.size > 0;
  // Если поиск не ведётся и фильтры отключены — любой узел подходит сразу без рекурсии
  if (!query && !hasFormatFilter && !state.diffFilter) {
    return true;
  }
  if (nodeMatchCache && nodeMatchCache.has(node)) {
    return nodeMatchCache.get(node);
  }
  if (visiting.has(node)) {
    return false;
  }
  visiting.add(node);

  const value = `${node.name} ${node.extension || ""} ${node.path || ""}`.toLocaleLowerCase("ru");
  const searchOk = !query || value.includes(query);
  const formatOk = !hasFormatFilter || node.type === "folder" || state.activeFilters.has(node.extension);
  if (node.type === "file") {
    if (state.diffFilter && !state.diffStatus.has(node.path)) {
      if (nodeMatchCache) nodeMatchCache.set(node, false);
      visiting.delete(node);
      return false;
    }
    const res = searchOk && formatOk;
    if (nodeMatchCache) nodeMatchCache.set(node, res);
    visiting.delete(node);
    return res;
  }

  // Для папок: если имя совпало с поиском (и нет фильтра расширений), сразу подходит
  if (!state.diffFilter && searchOk && !hasFormatFilter) {
    if (nodeMatchCache) nodeMatchCache.set(node, true);
    visiting.delete(node);
    return true;
  }

  const childMatch = (node.children || []).some((child) => nodeMatches(child, visiting));
  const res = state.diffFilter ? childMatch : (childMatch || searchOk);
  if (nodeMatchCache) nodeMatchCache.set(node, res);
  visiting.delete(node);
  return res;
}

function countDiffDescendants(node, result = { added: 0, changed: 0, removed: 0 }) {
  if (!node) return result;
  const stack = [node];
  const seen = new Set();
  while (stack.length) {
    const current = stack.pop();
    if (!current || seen.has(current)) continue;
    seen.add(current);
    for (const child of current.children || []) {
      if (child.type === "file") {
        const status = state.diffStatus.get(child.path);
        if (status && status in result) result[status] += 1;
      } else {
        stack.push(child);
      }
    }
  }
  return result;
}

const diffCountsCache = new Map();
function countDiffDescendantsCached(node) {
  if (diffCountsCache.has(node.path)) return diffCountsCache.get(node.path);
  const counts = countDiffDescendants(node);
  diffCountsCache.set(node.path, counts);
  return counts;
}

function diffStatusText(status) {
  return status === "added" ? "новый" : status === "changed" ? "изменён" : "удалён";
}

function buildDiffState(manifest, previousManifest) {
  diffCountsCache.clear();
  pdfPairCache.clear();
  if (previousManifest) state.pdfPairIndex = null;
  state.diffStatus.clear();
  state.diffRemovedNodes = [];
  state.diffSummary = null;
  const diff = manifest?.lastDiff;
  if (!diff) return;
  for (const path of diff.added || []) state.diffStatus.set(path, "added");
  for (const path of diff.changed || []) state.diffStatus.set(path, "changed");
  const removedSet = new Set(diff.removed || []);
  if (previousManifest?.tree && removedSet.size) {
    flattenTree(previousManifest.tree, []).forEach((node) => {
      if (node.type === "file" && removedSet.has(node.path)) {
        state.diffStatus.set(node.path, "removed");
        state.diffRemovedNodes.push(node);
      }
    });
  }
  state.diffSummary = {
    added: (diff.added || []).length,
    changed: (diff.changed || []).length,
    removed: state.diffRemovedNodes.length,
    unchanged: (diff.unchanged || []).length,
  };
}

function startDiffPolling() {
  stopDiffPolling();
  if (!state.selectedObjectId) return;
  state.diffPollTimer = setInterval(async () => {
    if (!inTreeMode() || !state.currentManifest) return;
    try {
      const controller = new AbortController();
      const response = await fetch("/api/objects/diff", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: state.selectedObjectId }),
        signal: controller.signal,
      });
      const payload = await response.json();
      if (!response.ok || state.progressCancelled) return;
      const diff = payload.lastDiff;
      const hasChanges = diff && ((diff.added || []).length || (diff.changed || []).length || (diff.removed || []).length);
      if (hasChanges) await refreshObjectInPlace({ reRender: false });
    } catch (error) {
      if (error.name === "AbortError") return;
    }
  }, DIFF_POLL_INTERVAL_MS);
}

function stopDiffPolling() {
  if (state.diffPollTimer) {
    clearInterval(state.diffPollTimer);
    state.diffPollTimer = null;
  }
}

async function refreshObjectInPlace(options = {}) {
  const object = selectedObject();
  if (!object) return;
  const previousManifest = state.currentManifest;
  const controller = createOperationController();
  startProgress("Проверка изменений", object.name);
  const response = await fetch("/api/objects/import", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: object.rootPath }),
    signal: controller.signal,
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || "Объект не обновлён");
  if (state.progressCancelled) return;
  state.selectedObjectId = payload.id;
  state.currentManifest = payload;
  const removedPaths = new Set((previousManifest?.lastDiff?.removed || []).map((p) => p.casefold ? p.toLocaleLowerCase("ru") : p));
  state.selectedPaths = new Set([...state.selectedPaths].filter((path) => !removedPaths.has(path.toLocaleLowerCase("ru"))));
  buildDiffState(payload, previousManifest);
  await loadObjectSummaries();
  renderFormats();
  renderTree();
  finishProgress(hasDiffChanges(payload) ? "Обнаружены изменения" : "Изменений нет");
  if (options.reRender !== false && state.renderedPages.length && state.selectedPaths.size) {
    renderSelectedFiles().catch(showOperationError);
  }
}

function hasDiffChanges(manifest) {
  const diff = manifest?.lastDiff;
  return Boolean(diff && ((diff.added || []).length || (diff.changed || []).length || (diff.removed || []).length));
}

function selectNode(node, event = {}, options = {}) {
  const index = state.visibleRows.findIndex((item) => item.path === node.path);
  if (event.shiftKey && state.lastSelectedIndex !== null) {
    const start = Math.min(state.lastSelectedIndex, index);
    const end = Math.max(state.lastSelectedIndex, index);
    state.selectedPaths.clear();
    state.visibleRows.slice(start, end + 1).forEach((item) => state.selectedPaths.add(item.path));
  } else if (event.ctrlKey || event.metaKey) {
    if (state.selectedPaths.has(node.path)) state.selectedPaths.delete(node.path);
    else state.selectedPaths.add(node.path);
    state.lastSelectedIndex = index;

  } else {
    state.selectedPaths.clear();
    state.selectedPaths.add(node.path);
    state.lastSelectedIndex = index;
  }
  if (state.selectedPaths.size === 1) {
    const selectedPath = Array.from(state.selectedPaths)[0];
    const selectedItem = state.visibleRows.find((item) => item.path === selectedPath) || node;
    if (selectedItem && selectedItem.type === "file") {
      setActiveNativePath(selectedItem.path);
      state.revealedPath = selectedItem.path;
      if (!options.deferPreview) syncSelectionToPreview(selectedItem);
    } else {
      setActiveNativePath("");
      state.revealedPath = "";
    }
  } else {
    setActiveNativePath("");
    state.revealedPath = "";
  }
  updateTreeSelectionHighlight();
}

// Выбор из ленты миниатюр: та же логика, что в дереве (одиночный /
// Shift-диапазон / Ctrl), но без автооткрытия превью — его запускает
// сам обработчик карточки. Якорь общий (state.lastSelectedIndex).
function selectRailPath(path, event = {}) {
  if (!path) return false;
  const norm = String(path).replace(/\//g, "\\").toLowerCase();
  const node = state.visibleRows.find(
    (item) => String(item.path || "").replace(/\//g, "\\").toLowerCase() === norm
  ) || null;
  if (!node) {
    // Файла нет в видимых строках: честная одиночная метка без якоря.
    if (!event.shiftKey && !event.ctrlKey && !event.metaKey) {
      state.selectedPaths.clear();
      state.selectedPaths.add(path);
      state.lastSelectedIndex = null;
      setActiveNativePath(path);
      state.revealedPath = path;
      updateTreeSelectionHighlight();
    }
    return true;
  }
  selectNode(node, event, { deferPreview: true });
  return true;
}

function syncSelectionToPreview(fileNode) {
  if (!fileNode || fileNode.type !== "file") return;
  const path = fileNode.path;
  const norm = path.replace(/\//g, "\\").toLowerCase();

  // 1. Проверяем уже отрендеренные страницы PDF / Word / DWG / изображений
  if (state.renderedPages?.length) {
    const matchedPage = state.renderedPages.find((p) => {
      const pPath = (p.previewFor?.path || p.sourcePath || p.documentPath || p.path || "").replace(/\//g, "\\").toLowerCase();
      return pPath === norm;
    });

    if (matchedPage) {
      const isStage = Boolean(els.pdfViewer?.classList.contains("stage-active"));
      showPdfPage(matchedPage, { skipTreeScroll: true, activateStage: isStage });
      const key = pageKey(matchedPage);
      let thumbEl = Array.from(els.pdfThumbs?.querySelectorAll(".pdf-thumb") || []).find(
        (el) => el.dataset.pageKey === key
      );
      if (!thumbEl) {
        thumbEl = Array.from(els.pdfThumbs?.querySelectorAll(".pdf-thumb") || []).find((el) => {
          const p = state.renderedPages.find((x) => {
            try { return pageKey(x) === el.dataset.pageKey; } catch (_) { return false; }
          });
          if (!p) return false;
          const pPath = (p.previewFor?.path || p.sourcePath || p.documentPath || p.path || "").replace(/\//g, "\\").toLowerCase();
          return pPath === norm;
        });
      }
      if (thumbEl) {
        thumbEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
      }
      return;
    }
  }

  // 2. Проверяем открытые книги Excel
  if (state.excelWorkbooks?.length) {
    const wbIndex = state.excelWorkbooks.findIndex((wb) => {
      const wbPath = (wb.path || wb.sourcePath || "").replace(/\//g, "\\").toLowerCase();
      return wbPath === norm;
    });
    if (wbIndex >= 0) {
      activateExcelWorkbook(wbIndex);
      const wbThumb = els.pdfThumbs?.querySelector(`.excel-book-thumb[data-workbook-index="${wbIndex}"]`);
      if (wbThumb) {
        wbThumb.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
      }
      return;
    }
  }

  // 3. Если файл ещё не отрендерен — не блокируем дерево фоновым рендером.
  // Для полного открытия файла служит двойной щелчок, а для пакета — кнопка «Отобразить».
}

async function previewFileDirectly(node, options = {}) {
  if (!node || node.type !== "file") return;
  setActiveNativePath(node.path);
  state.revealedPath = node.path;
  const ext = (node.extension || "").toUpperCase();

  els.pdfViewer.classList.remove("empty");
  // Позиции: 1 (дерево) -> 2 (лента), 2 -> 3 (большое окно).
  // Одиночка всегда открывается в ленте, кроме явного fullView.
  // TXT — как все: сначала миниатюра в ленте, крупно — только по двойному
  // клику на миниатюре.
  if (options.fullView) setViewerMode("full");
  else setViewerMode("standard");

  // 1. Excel: мгновенное открытие таблицы в просмотрщике
  if (["XLS", "XLSX", "XLSM", "CSV"].includes(ext)) {
    const excelItem = {
      ...node,
      previewType: "EXCEL",
      previewFor: {
        type: ext,
        name: node.name,
        path: node.path,
      },
    };
    await renderExcelWorkbooks([excelItem], { singleFile: true, fullView: options.fullView });
    if (options.fullView) setViewerMode("full");
    return;
  }

  // 1a. DOCX: быстрый HTML через ядро rendering (без Word COM).
  if (ext === "DOCX") {
    await showWordPreviewFast(node, options);
    return;
  }

  // 2. Если файл уже отрендерен в памяти ЦЕЛИКОМ — мгновенно переключаемся.
  // Если в памяти только титульник (рендер из папки), проваливаемся ниже
  // к полному рендеру, иначе пользователь навсегда останется на 1-й странице.
  const norm = node.path.replace(/\//g, "\\").toLowerCase();
  if (state.renderedPages?.length) {
    const matchedPage = state.renderedPages.find((p) => {
      const pPath = (p.previewFor?.path || p.sourcePath || p.documentPath || p.path || "").replace(/\//g, "\\").toLowerCase();
      return pPath === norm;
    });
    if (matchedPage) {
      const total = matchedPage.totalPages || 1;
      const mDoc = matchedPage.documentPath || "";
      const mSrc = matchedPage.sourcePath || "";
      const loaded = state.renderedPages.filter((p) => (mDoc && p.documentPath === mDoc) || (mSrc && p.sourcePath === mSrc)).length;
      if (loaded >= total) {
        showPdfPage(matchedPage, { skipTreeScroll: true });
        if (options.fullView) setViewerMode("full");
        const key = pageKey(matchedPage);
        const thumbEl = els.pdfThumbs?.querySelector(`.pdf-thumb[data-page-key="${CSS.escape(key)}"]`);
        if (thumbEl) thumbEl.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "center" });
        return;
      }
    }
  }

  // 3. Подготовка элемента для рендеринга одиночного документа
  let item = node;

  if (["DOC", "RTF"].includes(ext)) {
    item = {
      ...node,
      previewType: "WORD",
      previewFor: { type: ext, name: node.name, path: node.path },
    };
  } else if (ext === "DWG") {
    if (!state.pairless && !state.pdfPairIndex && state.currentManifest?.tree) {
      state.pdfPairIndex = buildPdfPairIndex(getCachedFlatNodes());
    }
    const pair = findPdfPairForDwg(node, state.pdfPairIndex || new Map(), { fuzzy: false });
    if (pair) {
      item = {
        ...pair.node,
        previewFor: { type: "DWG", name: node.name, path: node.path, confidence: pair.confidence },
      };
    } else {
      item = {
        ...node,
        previewType: "DWG_MODEL",
        previewFor: { type: "DWG", name: node.name, path: node.path },
      };
    }
  } else if (["JPG", "JPEG", "PNG", "BMP", "WEBP", "SVG", "GIF", "JFIF", "TIF", "TIFF", "ICO"].includes(ext)) {
    item = {
      ...node,
      previewType: "IMAGE",
      url: `/api/file/raw?path=${encodeURIComponent(node.path)}`,
      previewFor: { type: ext, name: node.name, path: node.path },
    };
  } else if (ext === "PDF") {
    item = node;
  } else if (ext === "TXT") {
    item = {
      ...node,
      previewType: "TXT",
      sourcePath: node.path,
      previewFor: { type: "TXT", name: node.name, path: node.path },
    };
  } else {
    // Нативный формат: карточка с метаданными ядра + кнопка открытия.
    await showNativeAppCard(node, options);
    return;
  }

  await renderSelectedPdfFiles([item], { singleFile: true, fullView: options.fullView });
  if (options.fullView) setViewerMode("full");
}

function updateTreeSelectionHighlight() {
  if (!els.objectTree) return;
  const revealedNorm = (state.revealedPath || "").replace(/\//g, "\\").toLowerCase();
  const selectedNorms = new Set(Array.from(state.selectedPaths).map((p) => (p || "").replace(/\//g, "\\").toLowerCase()));
  const rows = els.objectTree.querySelectorAll(".tree-row");
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const pNorm = (r.dataset.path || "").replace(/\//g, "\\").toLowerCase();
    const isSelected = selectedNorms.has(pNorm);
    const isRevealed = Boolean(revealedNorm && pNorm === revealedNorm);
    r.classList.toggle("selected", isSelected);
    r.classList.toggle("thumb-reveal", isRevealed);
  }
  updateRailSelectionHighlight();
}

// Подсветка выделения на карточках ленты: карточка подсвечена, если её
// файл входит в state.selectedPaths. Активная (текущий просмотр) рамка
// (.active) живёт отдельно и не трогается.
function updateRailSelectionHighlight() {
  if (!els.pdfThumbs) return;
  const norms = new Set(
    Array.from(state.selectedPaths).map((p) => String(p || "").replace(/\//g, "\\").toLowerCase())
  );
  els.pdfThumbs.querySelectorAll(".pdf-thumb").forEach((thumb) => {
    let p = "";
    try {
      if (thumb.classList.contains("excel-book-thumb") && thumb.dataset.workbookIndex !== undefined) {
        p = state.excelWorkbooks?.[Number(thumb.dataset.workbookIndex)]?.path || "";
      } else if (thumb.dataset.pageKey) {
        const page = state.renderedPages.find((x) => {
          try { return pageKey(x) === thumb.dataset.pageKey; } catch (_) { return false; }
        });
        p = page ? thumbPathForPage(page) : "";
      }
    } catch (_) {
      p = "";
    }
    const hit = Boolean(p) && norms.has(String(p).replace(/\//g, "\\").toLowerCase());
    thumb.classList.toggle("selected", hit);
  });
}

function rebuildVisibleRows() {
  state.visibleRows = [];
  if (!state.currentManifest?.tree) return;
  const stack = [state.currentManifest.tree];
  const seen = new Set();
  while (stack.length) {
    const node = stack.pop();
    if (!node || seen.has(node)) continue;
    seen.add(node);
    if (!nodeMatches(node)) continue;
    state.visibleRows.push(node);
    if (node.type === "folder" && !state.collapsedFolders.has(node.path) && node.children) {
      for (let i = node.children.length - 1; i >= 0; i--) {
        stack.push(node.children[i]);
      }
    }
  }
  if (state.diffRemovedNodes?.length && !state.diffFilter) {
    state.diffRemovedNodes.forEach((node) => state.visibleRows.push(node));
  }
}

let toastTimer = null;
function showToast(message) {
  let toast = document.getElementById("toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "toast";
    toast.className = "toast";
    document.body.append(toast);
  }
  toast.innerHTML = `<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="#34d399" stroke-width="2"><path d="M3 8.5l3.5 3.5 6.5-7.5"/></svg><span>${escapeHtml(message)}</span>`;
  toast.hidden = false;
  requestAnimationFrame(() => {
    toast.classList.add("visible");
  });
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("visible");
    setTimeout(() => {
      if (!toast.classList.contains("visible")) toast.hidden = true;
    }, 200);
  }, 2200);
}

async function copyPathToClipboard(rawPath) {
  if (!rawPath) return;
  // Убеждаемся, что системный путь Windows использует обратные слэши
  const path = String(rawPath).replace(/\//g, "\\");
  let success = false;

  // 1. Попытка через нативный Win32 API бэкенда (работает везде без ограничений webview)
  try {
    const resp = await fetch("/api/clipboard", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: path }),
    });
    if (resp.ok) {
      const data = await resp.json();
      if (data.ok) success = true;
    }
  } catch (err) {
  }

  // 2. Если бэкенд недоступен — пробуем современный navigator.clipboard
  if (!success && navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(path);
      success = true;
    } catch (e) {
    }
  }

  // 3. Резервный браузерный фоллбэк через textarea без readonly и с явным фокусом
  if (!success) {
    try {
      const textarea = document.createElement("textarea");
      textarea.value = path;
      textarea.style.position = "fixed";
      textarea.style.top = "0";
      textarea.style.left = "0";
      textarea.style.width = "2em";
      textarea.style.height = "2em";
      textarea.style.padding = "0";
      textarea.style.border = "none";
      textarea.style.outline = "none";
      textarea.style.boxShadow = "none";
      textarea.style.background = "transparent";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      textarea.setSelectionRange(0, path.length);
      success = document.execCommand("copy");
      document.body.removeChild(textarea);
    } catch (e) {
    }
  }

  if (success) {
    const preview = path.length > 55 ? "..." + path.slice(-52) : path;
    showToast(`Путь скопирован:\n${preview}`);
  } else {
    showToast("Не удалось скопировать путь");
  }
}

function revealPathInTree(path, options = {}) {
  // Файл активной миниатюры в дереве подсвечен постоянно, пока стоим на нём:
  // одна запись в дереве — один подсвеченный файл, сколько бы страниц
  // у него ни было в ленте. Маркер синхронизируем с просмотром,
  // при этом выделение (state.selectedPaths) связываем с текущим документом!
  if (!path) return;
  const shouldUpdateSelection = options.updateSelection !== false;
  state.revealedPath = path;
  if (shouldUpdateSelection) {
    state.selectedPaths.clear();
    state.selectedPaths.add(path);
  }

  const fileName = path.split(/[\\/]/).pop();
  if (fileName) updateViewerBanner(fileName);

  if (!state.currentManifest?.tree || !els.objectTree) return;

  const norm = String(path).replace(/\//g, "\\").toLowerCase();
  const nodeIndex = state.visibleRows.findIndex((item) => String(item.path || "").replace(/\//g, "\\").toLowerCase() === norm);
  if (nodeIndex >= 0 && shouldUpdateSelection) state.lastSelectedIndex = nodeIndex;

  let row = null;
  const pathFileName = path.split(/[\/\\]/).pop().toLowerCase().trim();
  const rows = els.objectTree.querySelectorAll(".tree-row");
  for (const r of rows) {
    const rPath = String(r.dataset.path || "").replace(/\//g, "\\").toLowerCase();
    if (rPath === norm) {
      row = r;
      break;
    }
  }
  if (!row && pathFileName) {
    for (const r of rows) {
      const rPath = String(r.dataset.path || "").replace(/\//g, "\\").toLowerCase();
      const rName = rPath.split(/[\/\\]/).pop().toLowerCase().trim();
      if (rName === pathFileName) {
        row = r;
        break;
      }
    }
  }

  // Если строка ещё не в DOM (родительская папка свёрнута) — раскрываем только нужных родителей
  // Раскрытие свёрнутых папок — только для явной навигации (клик, стрелки,
  // показ страницы). Наведение мыши на миниатюры дерево не трогает вообще,
  if (!row) {
    let uncollapsedAny = false;
    for (const folderPath of [...state.collapsedFolders]) {
      const folderNorm = String(folderPath).replace(/\//g, "\\").toLowerCase();
      if (norm.startsWith(folderNorm + "\\")) {
        state.collapsedFolders.delete(folderPath);
        uncollapsedAny = true;
      }
    }
    if (uncollapsedAny) {
      renderTree();
      try {
        row = els.objectTree.querySelector(`.tree-row[data-path="${CSS.escape(path)}"]`);
        if (!row) {
          const rows = els.objectTree.querySelectorAll(".tree-row");
          for (const r of rows) {
            if (String(r.dataset.path || "").replace(/\//g, "\\").toLowerCase() === norm) {
              row = r;
              break;
            }
          }
        }
      } catch {
        row = null;
      }
    }
  }

  updateTreeSelectionHighlight();

  if (row && !options.skipScroll) {
    row.scrollIntoView({ block: "center", behavior: "smooth" });
  }
}

const FOLDER_ICON_SVG = `<svg viewBox="0 0 16 16" width="15" height="15"><path fill="#d89e13" d="M1.5 2A1.5 1.5 0 0 0 0 3.5v9A1.5 1.5 0 0 0 1.5 14h13a1.5 1.5 0 0 0 1.5-1.5v-7A1.5 1.5 0 0 0 14.5 4H7.5L6.146 2.646A1.5 1.5 0 0 0 5.086 2.207L1.5 2z"/><path fill="#ffcb30" d="M1 5h14v7.5a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 1 12.5V5z"/><path fill="#ffe277" d="M1 5h14v1H1z"/></svg>`;

function getFileIconSvg(ext = "") {
  const e = String(ext).toUpperCase();
  if (e === "PDF") {
    return `<svg viewBox="0 0 16 16" width="15" height="15" fill="#dc2626"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h5l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-7.5A1.5 1.5 0 0 1 3 14.5v-13zm6.5.5v3h3l-3-3z"/><text x="3.8" y="12" font-size="5" font-weight="bold" fill="#ffffff" font-family="Segoe UI, sans-serif">PDF</text></svg>`;
  }
  if (e === "DWG" || e === "DXF") {
    return `<svg viewBox="0 0 16 16" width="15" height="15" fill="#2563eb"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h5l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-7.5A1.5 1.5 0 0 1 3 14.5v-13zm6.5.5v3h3l-3-3z"/><text x="3" y="12" font-size="4.5" font-weight="bold" fill="#ffffff" font-family="Segoe UI, sans-serif">DWG</text></svg>`;
  }
  if (e === "XLSX" || e === "XLS" || e === "CSV") {
    return `<svg viewBox="0 0 16 16" width="15" height="15" fill="#16a34a"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h5l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-7.5A1.5 1.5 0 0 1 3 14.5v-13zm6.5.5v3h3l-3-3z"/><text x="3.8" y="12" font-size="5" font-weight="bold" fill="#ffffff" font-family="Segoe UI, sans-serif">XLS</text></svg>`;
  }
  if (e === "DOCX" || e === "DOC") {
    return `<svg viewBox="0 0 16 16" width="15" height="15" fill="#1d4ed8"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h5l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-7.5A1.5 1.5 0 0 1 3 14.5v-13zm6.5.5v3h3l-3-3z"/><text x="3.2" y="12" font-size="4.5" font-weight="bold" fill="#ffffff" font-family="Segoe UI, sans-serif">DOC</text></svg>`;
  }
  return `<svg viewBox="0 0 16 16" width="15" height="15" fill="#94a3b8"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h5l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-7.5A1.5 1.5 0 0 1 3 14.5v-13zm6.5.5v3h3l-3-3z"/></svg>`;
}

const pdfPairCache = new Map();
function getPdfPairForDwgCached(node) {
  if (!node || node.extension !== "DWG") return null;
  if (pdfPairCache.has(node.path)) return pdfPairCache.get(node.path);
  if (state.pairless) return null;
  if (!state.pdfPairIndex && state.currentManifest?.tree) {
    state.pdfPairIndex = buildPdfPairIndex(getCachedFlatNodes());
  }
  const pair = findPdfPairForDwg(node, state.pdfPairIndex, { fuzzy: false });
  pdfPairCache.set(node.path, pair);
  return pair;
}

function toggleFolderNode(node, nodeEl) {
  if (!node || node.type !== "folder" || !nodeEl) return;
  const isCurrentlyCollapsed = state.collapsedFolders.has(node.path);
  const willBeCollapsed = !isCurrentlyCollapsed;

  if (willBeCollapsed) {
    state.collapsedFolders.add(node.path);
  } else {
    state.collapsedFolders.delete(node.path);
  }

  const chevron = nodeEl.querySelector(":scope > .tree-row .tree-chevron");
  if (chevron) {
    chevron.classList.toggle("expanded", !willBeCollapsed);
    chevron.title = willBeCollapsed ? "Развернуть папку" : "Свернуть папку";
  }
  nodeEl.classList.toggle("expanded", !willBeCollapsed);

  let childrenEl = nodeEl.querySelector(":scope > .tree-children");
  if (!willBeCollapsed) {
    if (!childrenEl && node.children?.length) {
      childrenEl = document.createElement("div");
      childrenEl.className = "tree-children";
      node.children.forEach((child) => renderTreeNode(child, childrenEl));
      nodeEl.append(childrenEl);
    } else if (childrenEl) {
      childrenEl.classList.remove("collapsed");
    }
  } else {
    if (childrenEl) {
      childrenEl.classList.add("collapsed");
    }
  }

  state.selectedPaths.clear();
  state.selectedPaths.add(node.path);
  updateTreeSelectionHighlight();
  rebuildVisibleRows();
  if (state.currentManifest?.name) {
    updateViewerBanner(state.currentManifest.name);
  }
}

function treeFilteringActive() {
  // Поиск, фильтр расширений или фильтр изменений: дерево показываем
  // раскрытым, иначе совпадения внутри свёрнутых папок не найти.
  return Boolean(els.treeSearch.value.trim() || state.activeFilters.size || state.diffFilter);
}

function collapseAllFolders() {
  state.collapsedFolders.clear();
  if (!state.currentManifest?.tree) return;
  const stack = [state.currentManifest.tree];
  const seen = new Set();
  while (stack.length) {
    const node = stack.pop();
    if (!node || seen.has(node)) continue;
    seen.add(node);
    for (const child of node.children || []) {
      if (child.type === "folder") {
        state.collapsedFolders.add(child.path);
        stack.push(child);
      }
    }
  }
}

function renderTreeNode(rootNode, rootParent) {
  // Итеративный обход: стек хранит {node, parent} — что вставить и куда.
  const stack = [{ node: rootNode, parent: rootParent }];
  while (stack.length) {
    const { node, parent } = stack.pop();
    if (!nodeMatches(node)) continue;

    const nodeEl = document.createElement("div");
    nodeEl.className = "tree-node";
    const collapsed = state.collapsedFolders.has(node.path) && !treeFilteringActive();
    if (node.type === "folder" && !collapsed) {
      nodeEl.classList.add("expanded");
    }

    const row = document.createElement("div");
    row.className = `tree-row ${node.type}`;
    row.dataset.path = node.path;
    if (node.path === state.revealedPath) row.classList.add("thumb-reveal");
    row.classList.toggle("selected", state.selectedPaths.has(node.path));
    row.title = node.path;

    const content = document.createElement("div");
    content.className = "tree-row-content";

    const chevron = document.createElement("span");
    chevron.className = "tree-chevron";
    if (node.type === "folder") {
      chevron.textContent = "›";
      if (!collapsed) chevron.classList.add("expanded");
      chevron.title = collapsed ? "Развернуть папку" : "Свернуть папку";
      chevron.addEventListener("click", ((n, el) => (e) => {
        e.stopPropagation();
        toggleFolderNode(n, el);
      })(node, nodeEl));
    } else {
      chevron.classList.add("leaf");
      chevron.innerHTML = "&nbsp;";
    }
    content.append(chevron);

    const iconEl = document.createElement("span");
    iconEl.className = "tree-icon";
    if (node === state.currentManifest?.tree) {
      iconEl.innerHTML = `<svg viewBox="0 0 16 16" width="15" height="15"><path fill="#ffffff" stroke="#718096" stroke-width="0.9" d="M3 1.5A1.5 1.5 0 0 1 4.5 0h6l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-9A1.5 1.5 0 0 1 3 14.5v-13z"/><path fill="#3b82f6" d="M5 6h6v1.2H5zm0 3h6v1.2H5zm0 3h4v1.2H5z"/></svg>`;
    } else if (node.type === "folder") {
      iconEl.innerHTML = FOLDER_ICON_SVG;
    } else {
      iconEl.innerHTML = getFileIconSvg(node.extension);
    }
    content.append(iconEl);

    const nameEl = document.createElement("span");
    nameEl.className = "tree-name";
    nameEl.textContent = node.name;

    const diffStatus = state.diffStatus.get(node.path);
    if (node.type === "file" && diffStatus) {
      const diffBadge = document.createElement("span");
      diffBadge.className = `diff-badge diff-${diffStatus}`;
      diffBadge.textContent = diffStatusText(diffStatus);
      nameEl.append(" ", diffBadge);
    }
    if (node.type === "folder" && state.diffSummary) {
      const counts = countDiffDescendantsCached(node);
      const parts = [];
      if (counts.added) parts.push(`+${counts.added}`);
      if (counts.changed) parts.push(`~${counts.changed}`);
      if (counts.removed) parts.push(`-${counts.removed}`);
      if (parts.length) {
        const diffFolderBadge = document.createElement("span");
        diffFolderBadge.className = "diff-badge diff-folder";
        diffFolderBadge.textContent = parts.join(" ");
        nameEl.append(" ", diffFolderBadge);
      }
    }
    if (node.type === "file" && node.extension === "DWG") {
      const pdfPair = getPdfPairForDwgCached(node);
      if (pdfPair) {
        const pairBadge = document.createElement("span");
        pairBadge.className = "pair-badge";
        pairBadge.textContent = "↔";
        pairBadge.title = `Связан с PDF-превью: ${pdfPair.node.name}${pdfPair.confidence === "probable" ? " (вероятная пара)" : ""}`;
        nameEl.append(" ", pairBadge);
      }
    }
    content.append(nameEl);

    if (node.type === "folder" && node.children) {
      const metaEl = document.createElement("span");
      metaEl.className = "tree-meta";
      metaEl.textContent = `(${node.children.length})`;
      content.append(metaEl);
    }

    row.append(content);
    row.addEventListener("contextmenu", ((n) => (event) => {
      event.preventDefault();
      event.stopPropagation();
      showFileContextMenu(event.clientX, event.clientY, {
        path: n.path,
        isDir: n.type === "folder",
        ext: n.extension || "",
      });
    })(node));

    row.addEventListener("click", ((n, el) => (event) => {
      event.stopPropagation();
      if (n.type === "folder") {
        const wasCollapsed = state.collapsedFolders.has(n.path);
        if (wasCollapsed && !event.shiftKey && !event.ctrlKey && !event.metaKey) {
          toggleFolderNode(n, el);
        }
        selectNode(n, event);
        return;
      }
      selectNode(n, event);
    })(node, nodeEl));

    row.addEventListener("dblclick", ((n) => (event) => {
      event.stopPropagation();
      if (n.type === "folder") {
        toggleFolderNode(n, nodeEl);
        return;
      }
      const index = state.visibleRows.findIndex((item) => item.path === n.path);
      state.selectedPaths.clear();
      state.selectedPaths.add(n.path);
      state.lastSelectedIndex = index;
      updateTreeSelectionHighlight();
      previewFileDirectly(n, { fullView: false }).catch(showOperationError);
    })(node));

    nodeEl.append(row);
    state.visibleRows.push(node);

    if (node.type === "folder" && node.children?.length && !collapsed) {
      const children = document.createElement("div");
      children.className = "tree-children";
      nodeEl.append(children);
      // Дочерние элементы добавляем в стек в обратном порядке,
      // чтобы при извлечении из стека порядок остался правильным
      for (let i = node.children.length - 1; i >= 0; i--) {
        stack.push({ node: node.children[i], parent: children });
      }
    }

    parent.append(nodeEl);
  }
}

function renderFormats() {
  els.formatStrip.replaceChildren();
  const extensions = state.currentManifest?.statistics?.extensions || {};
  const preferred = ["PDF", "DWG", "XLSX", "XLS", "XLSM", "GSHEET", "DOCX", "DOC", "GDOC", "TXT", "PNG", "JPG", "JPEG", "PPTX", "SVG"];
  const all = Object.keys(extensions).sort((a, b) => a.localeCompare(b, "ru"));
  const formats = [...new Set([...preferred.filter((ext) => ext in extensions), ...all])];
  if (!formats.length) {
    const empty = document.createElement("span");
    empty.className = "format-empty";
    empty.textContent = "Форматов нет";
    els.formatStrip.append(empty);
    return;
  }
  formats.forEach((ext) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "format-chip";
    button.classList.toggle("active", state.activeFilters.has(ext));
    button.textContent = `${ext} ${extensions[ext] || 0}`;
    let chipClickTimer = null;
    button.addEventListener("click", () => {
      if (chipClickTimer) {
        clearTimeout(chipClickTimer);
        chipClickTimer = null;
        return;
      }
      chipClickTimer = setTimeout(() => {
        chipClickTimer = null;
      }, 260);

      // Одиночный клик по чипу формата: фильтрация дерева по расширению.
      // Не выделяем все файлы скопом (все файлы выделяет и рендерит только двойной клик по чипу).
      const wasSoloActive = state.activeFilters.size === 1 && state.activeFilters.has(ext);
      state.activeFilters.clear();
      state.selectedPaths.clear();
      if (!wasSoloActive) {
        state.activeFilters.add(ext);
      }
      renderFormats();
      renderTree();
      // Якорь для Shift+клика: первый выделенный файл в видимых строках.
      // Без этого Shift сразу после фильтра брал бы старый якорь.
      const firstChipSel = Array.from(state.selectedPaths)[0];
      const firstChipIdx = firstChipSel
        ? state.visibleRows.findIndex((item) => item.path === firstChipSel)
        : -1;
      state.lastSelectedIndex = firstChipIdx >= 0 ? firstChipIdx : null;
      updateRailSelectionHighlight();
    });
    button.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      event.preventDefault();
      if (chipClickTimer) {
        clearTimeout(chipClickTimer);
        chipClickTimer = null;
      }
      state.activeFilters.clear();
      state.activeFilters.add(ext);
      state.selectedPaths.clear();
      if (state.currentManifest?.tree) {
        getCachedFlatNodes()
          .filter((node) => node.type === "file" && node.extension === ext)
          .forEach((node) => state.selectedPaths.add(node.path));
      }
      renderFormats();
      renderTree();
      // Якорь для Shift+клика: первый выделенный файл в видимых строках.
      const firstDblSel = Array.from(state.selectedPaths)[0];
      const firstDblIdx = firstDblSel
        ? state.visibleRows.findIndex((item) => item.path === firstDblSel)
        : -1;
      state.lastSelectedIndex = firstDblIdx >= 0 ? firstDblIdx : null;
      updateRailSelectionHighlight();
      renderSelectedFiles().catch(showOperationError);
    });
    els.formatStrip.append(button);
  });
}

function updateViewerBanner(title) {
  // Заголовок вьювера упразднен в пользу минималистичного интерфейса
}

function renderTree() {
  state.visibleRows = [];
  if (!state.currentManifest?.tree) {
    els.objectTree.replaceChildren();
    const empty = document.createElement("div");
    empty.className = "empty-note";
    empty.textContent = "Дерево ещё не открыто.";
    els.objectTree.append(empty);
    return;
  }
  nodeMatchCache = new Map();
  if (!state.pairless && !state.pdfPairIndex) {
    state.pdfPairIndex = buildPdfPairIndex(getCachedFlatNodes());
  }
  const fragment = document.createDocumentFragment();
  renderTreeNode(state.currentManifest.tree, fragment);
  if (state.diffRemovedNodes.length && !state.diffFilter) {
    const removedTitle = document.createElement("div");
    removedTitle.className = "diff-removed-title";
    removedTitle.textContent = "Удалённые файлы";
    fragment.append(removedTitle);
    state.diffRemovedNodes.forEach((node) => {
      const row = document.createElement("div");
      row.className = "tree-row file diff-removed-row";
      row.dataset.path = node.path;
      row.title = `${node.path}\nФайл удалён из исходной папки, но превью сохранено.`;

      const content = document.createElement("div");
      content.className = "tree-row-content";

      const chevron = document.createElement("span");
      chevron.className = "tree-chevron leaf";
      chevron.innerHTML = "&nbsp;";
      content.append(chevron);

      const iconEl = document.createElement("span");
      iconEl.className = "tree-icon";
      iconEl.innerHTML = getFileIconSvg(node.extension);
      content.append(iconEl);

      const nameEl = document.createElement("span");
      nameEl.className = "tree-name";
      nameEl.textContent = node.name;

      const badge = document.createElement("span");
      badge.className = "diff-badge diff-removed";
      badge.textContent = "удалён";
      nameEl.append(" ", badge);
      content.append(nameEl);

      row.append(content);
      row.addEventListener("click", (event) => {
        event.stopPropagation();
        selectNode(node, event);
      });
      row.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        event.stopPropagation();
        showFileContextMenu(event.clientX, event.clientY, {
          path: node.path,
          isDir: false,
          ext: node.extension || "",
        });
      });
      fragment.append(row);
      state.visibleRows.push(node);
    });
  }
  if (!state.visibleRows.length) {
    const empty = document.createElement("div");
    empty.className = "empty-note";
    empty.textContent = "Поиск или фильтр ничего не нашли.";
    fragment.append(empty);
  }
  els.objectTree.replaceChildren(fragment);
  nodeMatchCache = null;
  updateTreeSelectionHighlight();
}

async function openSelectedObject() {
  const object = selectedObject();
  if (!object) return;
  const controller = createOperationController();
  startProgress("Открытие структуры", object.name);
  const _ooT0 = performance.now();
  const response = await fetch(`/api/objects/${encodeURIComponent(object.id)}`, { signal: controller.signal });
  const manifest = await response.json();
  beaconPerf("object-open", performance.now() - _ooT0, `id=${object.id}`);
  if (!response.ok) throw new Error(manifest.error || "Объект не открыт");
  state.currentManifest = manifest;
  state.selectedPaths.clear();
  state.revealedPath = "";
  state.activeFilters.clear();
  state.diffFilter = false;
  collapseAllFolders();
  state.renderedPages = [];
  pdfPairCache.clear();
  diffCountsCache.clear();
  // Pairless (ситуация 2): в объекте нет PDF вовсе — индекс пар не нужен,
  // все DWG сразу пойдут по пути DWG_MODEL. Проверка по статистике O(1),
  // без обхода дерева. Иначе индекс строится один раз при открытии объекта
  // (~0,3 с на 14 тыс. файлов), дальше все показы и дерево его переиспользуют.
  buildDiffState(manifest, null);
  const _pairExts = (manifest.statistics || {}).extensions || {};
  const _pdfCount = Number(_pairExts.PDF ?? _pairExts.pdf ?? 0);
  state.pairless = _pdfCount === 0;
  const _piT0 = performance.now();
  state.pdfPairIndex = state.pairless ? null : buildPdfPairIndex(getCachedFlatNodes());
  beaconPerf("pair-index", performance.now() - _piT0, `files=${(manifest.statistics || {}).files || 0} pdf=${_pdfCount} pairless=${state.pairless}`);
  setMode("tree");
  updateViewerBanner(manifest.name || object.name);
  setTreeBrowseMode(true);
  renderFormats();
  const _trT0 = performance.now();
  renderTree();
  beaconPerf("tree-render", performance.now() - _trT0, `rows=${state.visibleRows.length}`);
  startDiffPolling();
  finishProgress("Структура открыта");
}

function collectPreviewFilesForDisplay() {
  if (!state.currentManifest?.tree) return [];
  const allNodes = getCachedFlatNodes();
  const selectedNodes = allNodes.filter((node) => state.selectedPaths.has(node.path));
  // Индекс один на объект (построен при открытии), не пересобираем при каждом показе.
  // Pairless-объекты (без PDF) индекс не строят вовсе.
  if (!state.pairless && !state.pdfPairIndex) state.pdfPairIndex = buildPdfPairIndex(allNodes);
  const pdfIndex = state.pdfPairIndex;
  const result = [];
  const directPdfPaths = new Set();
  const directSourcePaths = new Set();

  function addPreviewFile(node) {
    if (!node || node.type !== "file") return;
    if (node.name.startsWith("~$") || node.name.startsWith(".~")) return;
    const ext = (node.extension || "").toUpperCase();

    if (ext === "PDF" && !directPdfPaths.has(node.path)) {
      directPdfPaths.add(node.path);
      result.push(node);
    } else if (["DOC", "DOCX", "RTF"].includes(ext) && !directSourcePaths.has(node.path)) {
      directSourcePaths.add(node.path);
      result.push({
        ...node,
        previewType: "WORD",
        previewFor: {
          type: ext,
          name: node.name,
          path: node.path,
        },
      });
    } else if (["XLS", "XLSX", "XLSM", "CSV"].includes(ext) && !directSourcePaths.has(node.path)) {
      directSourcePaths.add(node.path);
      result.push({
        ...node,
        previewType: "EXCEL",
        previewFor: {
          type: ext,
          name: node.name,
          path: node.path,
        },
      });
    } else if (ext === "DWG") {
      // Массовый показ: только точные пары (папка не важна — пара обычно
      // в соседней папке PDF). При отсутствии пары используем мгновенный thumbnail.
      const pair = findPdfPairForDwg(node, pdfIndex, { fuzzy: false });
      if (pair) {
        result.push({
          ...pair.node,
          previewFor: {
            type: "DWG",
            name: node.name,
            path: node.path,
            confidence: pair.confidence,
          },
        });
      } else {
        result.push({
          ...node,
          previewType: "DWG_MODEL",
          previewFor: {
            type: "DWG",
            name: node.name,
            path: node.path,
          },
        });
      }
    } else if (ext === "TXT" && !directSourcePaths.has(node.path)) {
      directSourcePaths.add(node.path);
      result.push({
        ...node,
        previewType: "TXT",
        sourcePath: node.path,
        previewFor: { type: "TXT", name: node.name, path: node.path },
      });
    } else if (["JPG", "JPEG", "PNG", "BMP", "WEBP", "SVG", "GIF", "JFIF", "TIF", "TIFF", "ICO"].includes(ext) && !directSourcePaths.has(node.path)) {
      directSourcePaths.add(node.path);
      result.push({
        ...node,
        previewType: "IMAGE",
        url: `/api/file/raw?path=${encodeURIComponent(node.path)}`,
        previewFor: {
          type: ext,
          name: node.name,
          path: node.path,
        },
      });
    } else if (["GDOC", "GSHEET", "GSLIDES"].includes(ext) && !directSourcePaths.has(node.path)) {
      directSourcePaths.add(node.path);
      const isSheet = ext === "GSHEET";
      const isSlides = ext === "GSLIDES";
      const label = isSheet ? "Google Sheets" : (isSlides ? "Google Slides" : "Google Docs");
      result.push({
        type: "native-file",
        name: node.name,
        documentPath: node.path,
        sourcePath: node.path,
        sourceType: ext,
        message: `${label}: документ Google Drive. Нажмите кнопку, чтобы открыть в браузере.`,
      });
    } else if (["ZIP", "RAR", "7Z", "TAR", "GZ"].includes(ext) && !directSourcePaths.has(node.path)) {
      directSourcePaths.add(node.path);
      result.push({
        type: "native-file",
        name: node.name,
        documentPath: node.path,
        sourcePath: node.path,
        sourceType: ext,
        message: `Архив ${ext}: нажмите кнопку, чтобы открыть в Проводнике Windows или архиваторе.`,
      });
    // Дубликат PDF (папка + файл внутри неё выбраны вместе): путь уже занят
    // directPdfPaths выше, но провалился бы в native-карточку — отсекаем.
    } else if (!directSourcePaths.has(node.path) && !directPdfPaths.has(node.path)) {
      directSourcePaths.add(node.path);
      result.push({
        type: "native-file",
        name: node.name,
        documentPath: node.path,
        sourcePath: node.path,
        sourceType: ext,
        message: `Файл .${ext.toLowerCase()}: нажмите кнопку для открытия в программе Windows по умолчанию.`,
      });
    }
  }

  // Bulk pairing guard: cap fuzzy DWG->PDF search (see PAIR_FUZZY_BUDGET_MS).
  pairFuzzyBudgetUntil = Date.now() + PAIR_FUZZY_BUDGET_MS;
  try {
    if (selectedNodes.length) {
      selectedNodes.forEach((node) => {
        if (node.type === "file") addPreviewFile(node);
        if (node.type === "folder") {
          // Строго внутри папки: через разделитель, иначе папка «Договор»
          // захватывала бы соседнюю папку «Договор №2» как свою часть.
          const prefixBack = `${node.path}\\`;
          const prefixSlash = `${node.path}/`;
          allNodes
            .filter((candidate) => candidate.path !== node.path
              && (candidate.path.startsWith(prefixBack) || candidate.path.startsWith(prefixSlash)))
            .forEach(addPreviewFile);
        }
      });
    }
  } finally {
    pairFuzzyBudgetUntil = Infinity;
  }

  return result;
}

function clearExcelViewer() {
  state.excelWorkbook = null;
  state.excelWorkbooks = [];
  state.excelWorkbookIndex = 0;
  state.excelSheetIndex = 0;
  state.excelScale = 1;
  state.view.rotation = 0;
  els.excelViewer.hidden = true;
  els.excelBookTitle.textContent = "";
  els.excelBookTitle.title = "";
  els.excelTabs.replaceChildren();
  els.excelSheetFrame.removeAttribute("src");
  clearWordViewer();
}

function clearWordViewer() {
  state.wordDoc = null;
  if (els.wordViewer) els.wordViewer.hidden = true;
  if (els.wordDocTitle) { els.wordDocTitle.textContent = ""; els.wordDocTitle.title = ""; }
  if (els.wordMeta) els.wordMeta.replaceChildren();
  if (els.wordDocFrame) els.wordDocFrame.removeAttribute("src");
}

function formatFileSize(bytes) {
  const n = Number(bytes);
  if (!Number.isFinite(n) || n < 0) return "";
  if (n >= 1048576) return `${(n / 1048576).toFixed(2)} МБ`;
  if (n >= 1024) return `${(n / 1024).toFixed(1)} КБ`;
  return `${n} Б`;
}

function nativeTypeLabel(ext) {
  const e = String(ext || "").toUpperCase();
  if (["MPP", "MPT"].includes(e)) return "Проект Microsoft Project";
  if (["MOV", "MP4", "M4V", "AVI", "MKV", "WEBM", "MPEG", "MPG"].includes(e)) return "Видеофайл";
  if (["ZIP", "RAR", "7Z", "TAR", "GZ"].includes(e)) return "Архив";
  if (["RVT", "RFA", "NWC", "NWD", "IFC"].includes(e)) return "BIM-модель";
  if (["PPT", "PPTX"].includes(e)) return "Презентация";
  if (["VSD", "VSDX"].includes(e)) return "Схема Visio";
  return `Файл .${e.toLowerCase()}`;
}

async function showWordPreviewFast(node, options = {}) {
  // Быстрый HTML-рендер DOCX через ядро rendering (~1 с вместо ~20 с COM).
  // При любой ошибке — фолбэк на старый COM-путь через PDF.
  clearExcelViewer();
  resetPdfPreview();
  setActiveNativePath(node.path);
  state.revealedPath = node.path;
  if (options.fullView) setViewerMode("full");
  else setViewerMode("standard");
  els.pdfViewer.classList.remove("empty");
  els.pdfPageImage.hidden = true;
  els.pdfPageImage.removeAttribute("src");
  els.viewerEmpty.hidden = true;
  els.viewerControls.hidden = false;
  els.viewRotate.hidden = true;
  els.viewPanMode.hidden = true;
  els.wordDocTitle.textContent = node.name;
  els.wordDocTitle.title = node.name;
  els.wordMeta.replaceChildren();
  els.wordViewer.hidden = false;
  els.wordOpenNative.onclick = () => openFileByPath(node.path, "native");
  revealPathInTree(node.path);
  try {
    startProgress("Документ Word (быстрое превью)", node.path);
    const response = await fetch("/api/word/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file: node.path }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Не удалось построить HTML-превью");
    state.wordDoc = { path: node.path, name: node.name, ...payload };
    const chip = document.createElement("span");
    chip.className = "excel-tab";
    chip.textContent = `${payload.paragraphs ?? "?"} абз. · ${payload.tables ?? "?"} табл. · ${formatFileSize(payload.bytes)}`;
    els.wordMeta.append(chip);
    els.wordDocFrame.src = payload.url;
    finishProgress("Документ готов");
  } catch (error) {
    console.warn("[Launcher] fast Word preview failed, fallback to COM:", error);
    const item = {
      ...node,
      previewType: "WORD",
      previewFor: { type: "DOCX", name: node.name, path: node.path },
    };
    await renderSelectedPdfFiles([item], { singleFile: true, fullView: options.fullView });
  }
  if (options.fullView) setViewerMode("full");
}

async function showNativeAppCard(node, options = {}) {
  // Карточка нативного формата: метаданные из ядра (/api/preview),
  // открытие — через /api/open-file. При недоступности ядра — generic-карточка.
  resetPdfPreview();
  let card = {
    type: "native-file",
    name: node.name,
    sourcePath: node.path,
    documentPath: node.path,
    sourceType: (node.extension || "").toUpperCase(),
    message: "Файл открывается через ассоциации Windows или настроенную нативную программу.",
  };
  try {
    const response = await fetch("/api/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file: node.path }),
    });
    const info = await response.json();
    const native = response.ok && info
      ? (info.type === "native_app" ? info : info.fallback || null)
      : null;
    if (native) {
      const label = nativeTypeLabel(native.extension);
      card = {
        type: "native-file",
        name: native.name || node.name,
        sourcePath: node.path,
        documentPath: node.path,
        sourceType: String(native.extension || node.extension || "").toUpperCase(),
        message: `${label} · ${formatFileSize(native.bytes)}. ${native.hint || "Нажмите кнопку, чтобы открыть."}`,
      };
    }
  } catch (error) {
    console.warn("[Launcher] /api/preview unavailable, generic native card:", error);
  }
  showPdfPage(card);
  if (options.fullView) setViewerMode("full");
}

function resetPdfPreview() {
  state.view.userZoomed = false;
  state.activePageKey = "";
  state.activePageUrl = "";
  state.renderedPages = [];
  state.activeTxtPath = "";
  if (els.txtViewer) els.txtViewer.hidden = true;
  // Новый показ стирает ленту целиком: старые миниатюры не смешиваются с новыми.
  els.pdfThumbs.replaceChildren();
  els.pdfPageImage.hidden = true;
  els.pdfPageImage.removeAttribute("src");
  els.viewerEmpty.hidden = true;
  els.viewerEmpty.textContent = "";
  updatePagePosition(null);
  updateViewTransform();
  updateViewerBanner("");
}

async function activateExcelSheet(index) {
  const workbook = state.excelWorkbook;
  const sheet = workbook?.sheets?.[index];
  if (!sheet) return;
  if (sheet.warming) await sheet.warming;
  if (!sheet.url) {
    const response = await fetch("/api/excel/sheet", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file: workbook.path, sheetIndex: index }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Лист Excel не подготовлен");
    Object.assign(sheet, payload);
  }
  state.excelSheetIndex = index;
  [...els.excelTabs.querySelectorAll(".excel-tab")].forEach((tab) => {
    tab.classList.toggle("active", Number(tab.dataset.sheetIndex) === index);
  });
  els.excelSheetFrame.src = sheet.url;
  forwardExcelSheetContextMenu();
}

// ПКМ внутри таблицы: события из same-origin iframe не всплывают в основной
// документ, поэтому меню исходного файла пробрасываем вручную. На чужом
// документе тихо выходим — там остаётся обычное меню браузера.
function forwardExcelSheetContextMenu() {
  let doc = null;
  try {
    doc = els.excelSheetFrame?.contentDocument;
  } catch (_) {
    return;
  }
  if (!doc || doc.__launcherCtxForwarded) return;
  try {
    doc.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      try { event.stopPropagation(); } catch (_) {}
      hideFileContextMenu();
      const info = getActiveSourceInfo();
      const wb = state.excelWorkbook;
      const path = (info && info.path) || (wb && wb.path) || "";
      if (!path) return;
      showFileContextMenu(event.clientX, event.clientY, { path, isDir: false, ext: extOfPath(path) });
    });
    doc.__launcherCtxForwarded = true;
  } catch (_) {}
}

// Ждём первой отрисовки листа во фрейме (с таймаутом): иначе большое окно
// открывается пустым раньше таблицы, а прогресс уже погашен. Метка
// data-painted-url отличает свежую навигацию от повторного показа того же URL.
function waitForExcelSheetPaint(timeoutMs = 20000) {
  return new Promise((resolve) => {
    const frame = els.excelSheetFrame;
    const want = frame?.getAttribute("src") || "";
    if (!frame || !want) {
      resolve(false);
      return;
    }
    let settled = false;
    const done = (ok) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      try { frame.removeEventListener("load", onLoad); } catch (_) {}
      if (ok) {
        try { frame.dataset.paintedUrl = want; } catch (_) {}
      }
      resolve(ok);
    };
    const onLoad = () => done(true);
    const timer = setTimeout(() => done(false), timeoutMs);
    let painted = "";
    try { painted = frame.dataset.paintedUrl || ""; } catch (_) {}
    try {
      const doc = frame.contentDocument;
      if (doc && doc.readyState === "complete" && painted === want) {
        done(true);
        return;
      }
    } catch (_) {}
    try {
      frame.addEventListener("load", onLoad, { once: true });
    } catch (_) {
      done(false);
    }
  });
}

let _activeExcelWarmingWorkbook = null;

async function warmExcelWorkbook(workbook, activeIndex) {
  _activeExcelWarmingWorkbook = workbook;
  const totalSheets = (workbook.sheets || []).length;
  // Прогреваем в фоне только ближайшие соседние вкладки последовательно,
  // чтобы не подвешивать сервер и не тратить гигабайты памяти на тяжелых файлах.
  const candidates = [activeIndex + 1, activeIndex - 1].filter((idx) => idx >= 0 && idx < totalSheets);
  for (const index of candidates) {
    if (_activeExcelWarmingWorkbook !== workbook || state.excelWorkbook !== workbook || els.excelViewer?.hidden) {
      return;
    }
    const sheet = workbook.sheets[index];
    if (!sheet || sheet.url || sheet.warming) continue;
    sheet.warming = fetch("/api/excel/sheet", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file: workbook.path, sheetIndex: index }),
    })
      .then(async (response) => {
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Excel sheet preparation failed");
        Object.assign(sheet, payload);
      })
      .catch(() => {})
      .finally(() => { delete sheet.warming; });
    try {
      await sheet.warming;
    } catch (_) {}
  }
}

function renderExcelWorkbookRail() {
  const existingThumbs = els.pdfThumbs.querySelectorAll(".excel-book-thumb");
  if (existingThumbs.length === state.excelWorkbooks.length && existingThumbs.length > 0) {
    // Карточки уже построены: обновляем только активную рамку и не трогаем работающие iframes
    existingThumbs.forEach((thumb, index) => {
      thumb.classList.toggle("active", index === state.excelWorkbookIndex);
    });
    updateRailSelectionHighlight();
    return;
  }
  els.pdfThumbs.replaceChildren();
  state.excelWorkbooks.forEach((workbook, index) => {
    const wrap = document.createElement("div");
    wrap.className = "thumb-wrap";
    const thumb = document.createElement("button");
    thumb.type = "button";
    thumb.className = "pdf-thumb excel-book-thumb";
    thumb.dataset.workbookIndex = String(index);
    thumb.classList.toggle("active", index === state.excelWorkbookIndex);
    thumb.title = workbook.name;
    const preview = document.createElement("div");
    preview.className = "excel-book-preview";
    const frame = document.createElement("iframe");
    frame.src = workbook.thumbnailUrl || "about:blank";
    frame.title = `Миниатюра ${workbook.name}`;
    frame.tabIndex = -1;
    preview.append(frame);
    const label = document.createElement("span");
    label.textContent = workbook.name;
    thumb.append(preview, label);
    let excelClickTimer = null;
    thumb.addEventListener("mouseenter", () => {
      state.activeNavZone = "thumbs";
      if (workbook.path) {
        revealPathInTree(workbook.path, { skipScroll: false });
      }
    });
    thumb.addEventListener("click", (event) => {
      event.stopPropagation();
      state.activeNavZone = "thumbs";
      thumb.focus();
      // Выбор как в дереве; превью — только по обычному клику (как раньше).
      if (workbook.path) selectRailPath(workbook.path, event);
      if (event.shiftKey || event.ctrlKey || event.metaKey) return;
      if (excelClickTimer) return;
      excelClickTimer = setTimeout(() => {
        excelClickTimer = null;
        activateExcelWorkbook(index).catch(showOperationError);
      }, 220);
    });
    thumb.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      if (excelClickTimer) {
        clearTimeout(excelClickTimer);
        excelClickTimer = null;
      }
      activateExcelWorkbook(index).catch(showOperationError);
      setViewerMode("full");
    });
    wrap.append(thumb);
    wrap.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      event.stopPropagation();
      showFileContextMenu(event.clientX, event.clientY, {
        path: workbook.path || "",
        isDir: false,
        ext: workbook.extension || "",
      });
    });
    els.pdfThumbs.append(wrap);
  });
  updateRailSelectionHighlight();
}

async function activateExcelWorkbook(index) {
  const workbook = state.excelWorkbooks[index];
  if (!workbook) return;
  setStageActive(true);
  state.excelWorkbook = workbook;
  state.excelWorkbookIndex = index;
  state.excelSheetIndex = 0;
  state.excelScale = 1;
  els.excelBookTitle.textContent = workbook.name;
  els.excelBookTitle.title = workbook.name;
  renderExcelWorkbookRail();
  els.pdfViewer.classList.remove("empty");
  els.pdfPageImage.hidden = true;
  els.pdfPageImage.removeAttribute("src");
  els.viewerEmpty.hidden = true;
  els.qualityBadge.hidden = true;
  els.excelTabs.replaceChildren();
  (workbook.sheets || []).forEach((sheet) => {
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = "excel-tab";
    tab.dataset.sheetIndex = String(sheet.index);
    tab.title = `${sheet.name} · ${sheet.rows} строк · ${sheet.columns} столбцов`;
    tab.textContent = sheet.name;
    tab.addEventListener("click", () => activateExcelSheet(sheet.index).catch(showOperationError));
    els.excelTabs.append(tab);
  });
  els.excelViewer.hidden = false;
  if (els.wordViewer) els.wordViewer.hidden = true;
  els.viewerControls.hidden = false;
  els.viewRotate.hidden = true;
  els.viewPanMode.hidden = false;
  setActiveNativePath(workbook.path);
  revealPathInTree(workbook.path);
  updateViewTransform();
  await activateExcelSheet(0);
  window.setTimeout(() => warmExcelWorkbook(workbook, 0), 250);
}

async function showExcelWorkbooks(workbooks, options = {}) {
  state.excelWorkbooks = workbooks;
  renderExcelWorkbookRail();
  els.pdfViewer?.classList.remove("empty");
  if (els.viewerEmpty) els.viewerEmpty.hidden = true;
  if (options.isSingle || workbooks.length === 1) {
    setStageActive(true);
    await activateExcelWorkbook(0);
    if (options.fullView) setViewerMode("full");
  } else {
    setStageActive(false);
    if (els.excelViewer) els.excelViewer.hidden = true;
    if (els.wordViewer) els.wordViewer.hidden = true;
  }
}

function createPageThumbElement(page) {
  const key = pageKey(page);
  const wrap = document.createElement("div");
  wrap.className = "thumb-wrap";
  const thumb = document.createElement("button");
  thumb.type = "button";
  thumb.className = "pdf-thumb";
  thumb.classList.toggle("missing-preview", page.type === "missing-preview");
  thumb.dataset.pageKey = key;
  thumb.classList.toggle("active", key === state.activePageKey);
  thumb.title = `${page.name} · клик — выбор листа, двойной клик — на весь экран`;

  if (page.previewType === "TXT") {
    const prev = document.createElement("div");
    prev.className = "txt-thumb-preview";
    prev.textContent = "TXT…";
    thumb.append(prev);
    fillTxtThumbPreview(page, prev);
  } else if (page.type === "missing-preview" || page.type === "native-file") {
    const missing = document.createElement("div");
    missing.className = "missing-preview-card";
    const ext = (page.sourceType || "").toUpperCase();
    let icon = "📄";
    if (ext === "GDOC") icon = "🌐";
    else if (ext === "GSHEET") icon = "📊";
    else if (ext === "GSLIDES") icon = "📽️";
    else if (["ZIP", "RAR", "7Z"].includes(ext)) icon = "📦";
    else if (["MPP", "MPT"].includes(ext)) icon = "📊";
    else if (["MOV", "MP4", "M4V", "AVI", "MKV", "WEBM", "MPEG", "MPG"].includes(ext)) icon = "🎬";
    else if (["RVT", "RFA", "NWC", "NWD", "IFC"].includes(ext)) icon = "🏗️";
    else if (["PPT", "PPTX"].includes(ext)) icon = "📽️";
    else if (["VSD", "VSDX"].includes(ext)) icon = "📝";
    else if (ext === "DWG") icon = "📐";
    else if (["DOC", "DOCX"].includes(ext)) icon = "📘";
    else if (["XLS", "XLSX"].includes(ext)) icon = "📗";
    missing.textContent = `${icon} ${ext || "Файл"}`;
    thumb.append(missing);
  } else {
    const img = document.createElement("img");
    img.src = page.url;
    img.alt = page.name;
    img.draggable = false;
    thumb.append(img);
  }
  const label = document.createElement("span");
  label.textContent = page.name;

  thumb.addEventListener("mouseenter", () => {
    state.activeNavZone = "thumbs";
    const docPath = thumbPathForPage(page);
    if (docPath && state.revealedPath !== docPath) {
      revealPathInTree(docPath, { updateSelection: false, skipScroll: false });
    }
  });

  thumb.addEventListener("click", (event) => {
    event.stopPropagation();
    state.activeNavZone = "thumbs";
    // Выбор как в дереве; превью — только по обычному клику (как раньше).
    const docPath = thumbPathForPage(page);
    if (docPath) selectRailPath(docPath, event);
    if (event.shiftKey || event.ctrlKey || event.metaKey) return;
    showPdfPage(page);
    thumb.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
  });

  thumb.addEventListener("dblclick", (event) => {
    event.stopPropagation();
    state.activeNavZone = "thumbs";
    showPdfPage(page);
    setViewerMode("full");
  });
  wrap.append(thumb);
  wrap.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const menuPath = thumbPathForPage(page);
    showFileContextMenu(event.clientX, event.clientY, {
      path: menuPath,
      isDir: false,
      ext: extOfPath(menuPath),
    });
  });
  return wrap;
}

async function appendPagesToViewer(pages, insertAfterKey = null, epoch = null) {
  if (!pages || !pages.length) return;
  if (epoch !== null && state.renderEpoch !== epoch) return;
  clearExcelViewer();
  els.pdfViewer.classList.toggle("empty", false);
  els.viewerEmpty.hidden = true;

  // Снимки существующих ключей один раз: проверка каждого нового
  // элемента идёт за O(1), а не сканированием всего DOM и массива.
  const knownStateKeys = new Set();
  for (const p of state.renderedPages) {
    try { knownStateKeys.add(pageKey(p)); } catch (_) {}
  }
  const existingDomKeys = new Set();
  els.pdfThumbs.querySelectorAll(".pdf-thumb").forEach((el) => {
    if (el.dataset.pageKey) existingDomKeys.add(el.dataset.pageKey);
  });

  // Якорь вставки вычисляем один раз на весь вызов.
  let anchorNode = null;
  let appendAtEnd = true;
  if (insertAfterKey) {
    const afterThumb = els.pdfThumbs.querySelector(`.pdf-thumb[data-page-key="${CSS.escape(insertAfterKey)}"]`);
    const afterWrap = afterThumb ? (afterThumb.closest(".thumb-wrap") || afterThumb) : null;
    if (afterWrap && afterWrap.nextSibling) {
      anchorNode = afterWrap.nextSibling;
      appendAtEnd = false;
    } else if (!afterWrap) {
      const loadMoreBtn = els.pdfThumbs.querySelector(".ribbon-load-more");
      if (loadMoreBtn) { anchorNode = loadMoreBtn; appendAtEnd = false; }
    }
  } else {
    const loadMoreBtn = els.pdfThumbs.querySelector(".ribbon-load-more");
    if (loadMoreBtn) { anchorNode = loadMoreBtn; appendAtEnd = false; }
  }

  // Строим DOM порциями и отдаём управление между ними, чтобы страница
  // не висела на сотнях миниатюр. Отмена прерывает добавление.
  const CHUNK_PAGES = 60;
  for (let start = 0; start < pages.length; start += CHUNK_PAGES) {
    if (state.progressCancelled) return;
    if (epoch !== null && state.renderEpoch !== epoch) return;
    const fragment = document.createDocumentFragment();
    let lastInserted = null;
    const slice = pages.slice(start, start + CHUNK_PAGES);
    for (const page of slice) {
      let key = "";
      try { key = pageKey(page); } catch (_) { continue; }
      if (!knownStateKeys.has(key)) {
        knownStateKeys.add(key);
        if (insertAfterKey) {
          const afterIdx = state.renderedPages.findIndex((p) => { try { return pageKey(p) === insertAfterKey; } catch (_) { return false; } });
          if (afterIdx >= 0) {
            state.renderedPages.splice(afterIdx + 1, 0, page);
          } else {
            state.renderedPages.push(page);
          }
        } else {
          state.renderedPages.push(page);
        }
      }
      if (existingDomKeys.has(key)) {
        continue;
      }
      existingDomKeys.add(key);
      let wrap = null;
      try {
        wrap = createPageThumbElement(page);
      } catch (_) {
        continue;
      }
      fragment.append(wrap);
      lastInserted = wrap;
    }
    if (lastInserted) {
      if (appendAtEnd || !anchorNode) {
        els.pdfThumbs.append(fragment);
      } else {
        els.pdfThumbs.insertBefore(fragment, anchorNode);
      }
      anchorNode = lastInserted.nextSibling;
      appendAtEnd = !anchorNode;
    }
    if (start + CHUNK_PAGES < pages.length) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  if ((!state.activePageKey || !knownStateKeys.has(state.activePageKey)) && state.renderedPages.length) {
    showPdfPage(state.renderedPages[0]);
  } else {
    updateActivePdfThumb();
  }
  updateRailSelectionHighlight();
}

function updateRibbonLoadMore() {
  const existing = els.pdfThumbs.querySelector(".ribbon-load-more");
  if (existing) existing.remove();

  // Убираем старый observer
  if (state._ribbonScrollObserver) {
    state._ribbonScrollObserver.disconnect();
    state._ribbonScrollObserver = null;
  }

  if (state.remainingPreviewItems && state.remainingPreviewItems.length > 0) {
    const sentinel = document.createElement("div");
    sentinel.className = "ribbon-load-more";
    const shown = state.totalPreviewCount - state.remainingPreviewItems.length;
    sentinel.textContent = `Загружено ${shown} из ${state.totalPreviewCount} · прокрутите для подгрузки...`;
    els.pdfThumbs.append(sentinel);

    const LOAD_MORE_SIZE = 200;
    let loading = false;
    // Страж принадлежит текущему показу: если пользователь тем временем
    // начал новый показ, молча снимаемся, чтобы не подмешивать старые страницы.
    const obsEpoch = state.renderEpoch || 0;

    const observer = new IntersectionObserver(async (entries) => {
      if (!entries[0]?.isIntersecting || loading) return;
      if (state.progressCancelled) return;
      if (state.renderEpoch !== obsEpoch) {
        observer.disconnect();
        sentinel.remove();
        return;
      }
      if (!state.remainingPreviewItems || !state.remainingPreviewItems.length) {
        observer.disconnect();
        sentinel.remove();
        return;
      }
      loading = true;
      sentinel.textContent = `Подгрузка следующих ${Math.min(LOAD_MORE_SIZE, state.remainingPreviewItems.length)} файлов...`;
      const nextBatch = state.remainingPreviewItems.splice(0, LOAD_MORE_SIZE);
      await loadMorePdfFiles(nextBatch);
      loading = false;
    }, { root: els.pdfThumbs, rootMargin: "0px 300px 0px 0px", threshold: 0 });

    observer.observe(sentinel);
    state._ribbonScrollObserver = observer;
  }
}

function pageKey(page) {
  if (page.type === "missing-preview" || page.type === "native-file") return `native|${page.sourcePath || page.documentPath || page.name}`;
  // У картинок нет documentPath/page — без своей ветки все они получали
  // один ключ "|" и в ленту вставала только первая (остальные скипались
  // как «дубликаты»).
  if (page.previewType === "IMAGE") return `image|${page.previewFor?.path || page.path || page.url || page.name}`;
  if (page.previewType === "TXT") return `txt|${page.previewFor?.path || page.path || page.name}`;
  return `${page.documentPath || ""}|${page.page || ""}`;
}

function updateActivePdfThumb() {
  [...els.pdfThumbs.querySelectorAll(".pdf-thumb")].forEach((thumb) => {
    thumb.classList.toggle("active", thumb.dataset.pageKey === state.activePageKey);
  });
}

function updatePagePosition(page) {
  // Счётчик «Стр. N из M» слева вверху сцены: всегда видно, на каком листе
  // какого документа стоим. Для одиночных страниц и карточек — скрыт.
  if (!els.pagePosition) return;
  const num = Number(page?.page) || 0;
  const total = Number(page?.totalPages) || 0;
  if (!num || total < 2) {
    els.pagePosition.hidden = true;
    return;
  }
  els.pagePosition.hidden = false;
  els.pagePosition.textContent = `Стр. ${num} из ${total}`;
}

function setActiveNativePath(path) {
  state.activeNativePath = path || "";
  if (els.viewerControls && state.activeNativePath) els.viewerControls.hidden = false;
}

function thumbPathForPage(page) {
  if (!page) return "";
  return page.previewFor?.path || page.sourcePath || page.documentPath || page.path || "";
}

// Источник текущего кадра для ПКМ в полном экране: всегда исходный файл,
// никогда PNG-кэш. DWG через пару -> путь DWG, Word -> .docx и т.д.
function getActiveSourceInfo() {
  const excelPath = (state.excelWorkbook && state.excelWorkbook.path) || "";
  if (excelPath) return { path: excelPath, ext: extOfPath(excelPath) };
  if (state.activeTxtPath) return { path: state.activeTxtPath, ext: "TXT" };
  const page = (state.renderedPages || []).find((p) => {
    try { return pageKey(p) === state.activePageKey; } catch (_) { return false; }
  });
  const filePath = thumbPathForPage(page);
  if (filePath) return { path: filePath, ext: extOfPath(filePath) };
  return { path: "", ext: "" };
}

function extOfPath(path) {
  const base = String(path || "").split(/[/\\]/).pop() || "";
  const dot = base.lastIndexOf(".");
  return dot > 0 ? base.slice(dot + 1).toUpperCase() : "";
}

async function requestHighQualityPage(page) {
  if (!page.documentPath || !page.page || page.dpi >= PDF_QUALITY_DPI) return;
  const key = pageKey(page);
  if (state.highQualityPages.has(key)) {
    const cached = state.highQualityPages.get(key);
    if (state.activePageKey === key) {
      state.activePageUrl = cached.url;
      els.pdfPageImage.src = cached.url;
        }
    return;
  }
  if (state.activePageKey === key) {
    }
  try {
    const endpoint = page.previewType === "WORD"
      ? "/api/word/page"
      : page.previewType === "DWG_MODEL"
        ? "/api/dwg/model-page"
        : page.previewType === "EXCEL"
          ? "/api/excel/page"
          : "/api/pdf/page";
    const sourceFile = ["WORD", "DWG_MODEL", "EXCEL"].includes(page.previewType)
      ? page.previewFor?.path
      : page.documentPath;
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file: sourceFile, page: page.page, dpi: PDF_QUALITY_DPI }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "PDF page render failed");
    const highPage = {
      ...page,
      ...payload.item,
      documentPath: payload.path,
      dpi: payload.dpi,
      cacheHit: payload.cacheHit,
    };
    state.highQualityPages.set(key, highPage);
    if (state.activePageKey === key) {
      state.activePageUrl = highPage.url;
      els.pdfPageImage.src = highPage.url;
        }
  } catch (error) {
    if (state.activePageKey === key) {
        }
  }
}

// Текстовый кэш TXT (маленькие файлы целиком, не более ~20 штук).
const txtCache = new Map();

function decodeTxtBytes(buf) {
  const bytes = new Uint8Array(buf);
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (_) {
    return new TextDecoder("windows-1251").decode(bytes);
  }
}

async function fetchTxtText(sourcePath) {
  let text = txtCache.get(sourcePath);
  if (text !== undefined) return text;
  const response = await fetch(`/api/file/raw?path=${encodeURIComponent(sourcePath)}`);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const buf = await response.arrayBuffer();
  if (buf.byteLength > 1000000) throw new Error("Файл больше 1 МБ — откройте в программе.");
  text = decodeTxtBytes(buf);
  if (txtCache.size > 20) txtCache.clear();
  txtCache.set(sourcePath, text);
  return text;
}

async function showTxtContent(sourcePath) {
  const box = els.txtContent;
  const view = els.txtViewer;
  if (!box || !view || !sourcePath) return;
  view.hidden = false;
  box.textContent = "Загрузка текста…";
  try {
    const text = await fetchTxtText(sourcePath);
    if (state.activeTxtPath !== sourcePath) return;
    box.textContent = text;
  } catch (error) {
    if (state.activeTxtPath !== sourcePath) return;
    box.textContent = `Не удалось показать текст: ${error?.message || error}`;
  }
}

async function fillTxtThumbPreview(page, el) {
  const src = page.previewFor?.path || page.path || page.sourcePath;
  if (!src) return;
  try {
    const text = await fetchTxtText(src);
    if (!el.isConnected) return;
    const snippet = text.split(/\r?\n/).slice(0, 3).join("\n").slice(0, 180);
    el.textContent = snippet || page.name;
  } catch (_) {
    // Остаётся заглушка «TXT…», вид не ломаем.
  }
}

function showPdfPage(page, options = {}) {
  const key = pageKey(page);
  state.activePageKey = key;
  const skipTreeScroll = Boolean(options?.skipTreeScroll);
  if (els.txtViewer) els.txtViewer.hidden = true;
  state.activeTxtPath = "";
  if (els.wordViewer) els.wordViewer.hidden = true;

  if (page.previewType === "IMAGE") {
    clearExcelViewer();
    const sourcePath = page.previewFor?.path || page.path || page.sourcePath;
    setActiveNativePath(sourcePath);
    revealPathInTree(sourcePath, { skipScroll: skipTreeScroll });
    state.activePageUrl = page.url;
    els.pdfPageImage.src = page.url;
    els.pdfPageImage.hidden = false;
    els.viewerEmpty.hidden = true;
    els.pdfViewer.classList.remove("empty");
    els.viewerControls.hidden = false;
    els.viewRotate.hidden = false;
    els.viewPanMode.hidden = false;
    updateActivePdfThumb();
    if (els.pdfPageImage.complete && els.pdfPageImage.naturalWidth) applyPageView();
    return;
  }

  if (page.previewType === "TXT") {
    clearExcelViewer();
    const sourcePath = page.previewFor?.path || page.path || page.sourcePath;
    setActiveNativePath(sourcePath);
    state.activeTxtPath = sourcePath;
    revealPathInTree(sourcePath, { skipScroll: skipTreeScroll });
    state.activePageUrl = "";
    els.pdfPageImage.hidden = true;
    els.pdfPageImage.removeAttribute("src");
    els.viewerEmpty.hidden = true;
    els.pdfViewer.classList.remove("empty");
    els.viewerControls.hidden = false;
    els.viewRotate.hidden = true;
    els.viewPanMode.hidden = true;
    updateActivePdfThumb();
    updatePagePosition(null);
    showTxtContent(sourcePath);
    return;
  }

  if (page.type === "missing-preview" || page.type === "native-file") {
    clearExcelViewer();
    const sourcePath = page.previewFor?.path || page.sourcePath || page.documentPath || "";
    setActiveNativePath(sourcePath);
    revealPathInTree(sourcePath, { skipScroll: skipTreeScroll });
    state.activePageUrl = "";
    els.pdfPageImage.hidden = true;
    els.pdfPageImage.removeAttribute("src");
    els.viewerEmpty.hidden = false;

    const ext = (page.sourceType || "").toUpperCase();
    let icon = "📄";
    if (ext === "GDOC") icon = "🌐";
    else if (ext === "GSHEET") icon = "📊";
    else if (ext === "GSLIDES") icon = "📽️";
    else if (["ZIP", "RAR", "7Z"].includes(ext)) icon = "📦";
    else if (["MPP", "MPT"].includes(ext)) icon = "📊";
    else if (["MOV", "MP4", "M4V", "AVI", "MKV", "WEBM", "MPEG", "MPG"].includes(ext)) icon = "🎬";
    else if (["RVT", "RFA", "NWC", "NWD", "IFC"].includes(ext)) icon = "🏗️";
    else if (["PPT", "PPTX"].includes(ext)) icon = "📽️";
    else if (["VSD", "VSDX"].includes(ext)) icon = "📝";
    else if (ext === "DWG") icon = "📐";
    else if (["DOC", "DOCX"].includes(ext)) icon = "📘";
    else if (["XLS", "XLSX"].includes(ext)) icon = "📗";

    els.viewerEmpty.innerHTML = `
      <div class="native-file-card">
        <div class="native-file-icon">${icon}</div>
        <div class="native-file-name">${escapeHtml(page.name)}</div>
        <div class="native-file-message">${escapeHtml(page.message || "")}</div>
        <div class="native-file-path" title="${escapeHtml(sourcePath)}">${escapeHtml(sourcePath)}</div>
        <button type="button" class="native-card-open-btn" id="nativeCardOpenBtn">Открыть в программе по умолчанию</button>
        <div class="native-file-hint">Правая кнопка мыши: открыть в программе по умолчанию · открыть в проводнике · скопировать путь</div>
      </div>
    `;

    const nativeCard = els.viewerEmpty.querySelector(".native-file-card");
    if (nativeCard) {
      const openBtn = nativeCard.querySelector("#nativeCardOpenBtn");
    if (openBtn) {
      openBtn.addEventListener("click", () => openFileByPath(sourcePath, "native"));
    }
    nativeCard.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        event.stopPropagation();
        showFileContextMenu(event.clientX, event.clientY, {
          path: sourcePath,
          isDir: false,
          ext: extOfPath(sourcePath),
        });
      });
    }

    els.pdfViewer.classList.remove("empty");
    els.viewerControls.hidden = false;
    els.viewRotate.hidden = true;
    els.viewPanMode.hidden = true;
    updateActivePdfThumb();
    updatePagePosition(null);
    return;
  }

  clearExcelViewer();
  const highPage = state.highQualityPages.get(key);
  const displayPage = highPage || page;
  const targetDocPath = page.previewFor?.path || page.sourcePath || page.documentPath || "";
  setActiveNativePath(targetDocPath);
  revealPathInTree(targetDocPath, { skipScroll: skipTreeScroll });
  state.activePageUrl = displayPage.url;
  els.pdfPageImage.src = displayPage.url;
  els.pdfPageImage.hidden = false;
  els.viewerEmpty.hidden = true;
  els.pdfViewer.classList.remove("empty");
  els.viewerControls.hidden = false;
  els.viewRotate.hidden = false;
  els.viewPanMode.hidden = false;
  updateActivePdfThumb();
  if (els.pdfPageImage.complete && els.pdfPageImage.naturalWidth) applyPageView();
  requestHighQualityPage(page);
  updatePagePosition(page);
}

function clampPan() {
  // Ровные границы перетаскивания: лист можно увести за любой край,
  // но пара сантиметров (70px) всегда остаётся в видимой зоне,
  // чтобы его можно было вернуть обратно.
  const view = state.view;
  if (els.pdfPageImage.hidden) return;
  const naturalWidth = els.pdfPageImage.naturalWidth;
  const naturalHeight = els.pdfPageImage.naturalHeight;
  if (!naturalWidth || !naturalHeight) return;
  const stage = els.pdfStage.getBoundingClientRect();
  if (!stage.width || !stage.height) return;
  const rotated = Math.abs(view.rotation % 180) === 90;
  const width = (rotated ? naturalHeight : naturalWidth) * view.scale;
  const height = (rotated ? naturalWidth : naturalHeight) * view.scale;
  const margin = 70;
  const maxX = Math.max(0, width / 2 + stage.width / 2 - margin);
  const maxY = Math.max(0, height / 2 + stage.height / 2 - margin);
  view.panX = Math.min(maxX, Math.max(-maxX, view.panX));
  view.panY = Math.min(maxY, Math.max(-maxY, view.panY));
}

function updateViewTransform() {
  const view = state.view;
  clampPan();
  els.pdfPageImage.style.transform = [
    "translate(-50%, -50%)",
    `translate(${view.panX}px, ${view.panY}px)`,
    `rotate(${view.rotation}deg)`,
    `scale(${view.scale})`,
  ].join(" ");
  els.pdfStage.classList.toggle("pan-mode", view.panMode);
  els.pdfStage.classList.toggle("dragging", view.dragging);
  els.viewPanMode.classList.toggle("active", view.panMode);
  if (state.excelWorkbook) {
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-hand", value: view.panMode }, "*");
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-rotate", value: view.rotation }, "*");
  }
}

function fitPdfPage() {
  if (state.excelWorkbook) {
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-fit" }, "*");
    return;
  }
  if (!els.pdfPageImage.naturalWidth || !els.pdfPageImage.naturalHeight) return;
  const stage = els.pdfStage.getBoundingClientRect();
  // Сцена ещё не разложена (нулевые размеры): не портим масштаб крошечным
  // вписыванием, дождёмся следующего кадра.
  if (stage.width < 10 || stage.height < 10) {
    requestAnimationFrame(() => requestAnimationFrame(fitPdfPage));
    return;
  }
  const rotated = Math.abs(state.view.rotation % 180) === 90;
  const imageWidth = rotated ? els.pdfPageImage.naturalHeight : els.pdfPageImage.naturalWidth;
  const imageHeight = rotated ? els.pdfPageImage.naturalWidth : els.pdfPageImage.naturalHeight;
  const availableWidth = Math.max(100, stage.width - 36);
  const availableHeight = Math.max(100, stage.height - 86);
  const fit = Math.min(availableWidth / imageWidth, availableHeight / imageHeight);
  state.view.fitScale = fit || 1;
  state.view.scale = state.view.fitScale;
  state.view.panX = 0;
  state.view.panY = 0;
  updateViewTransform();
}

function zoomPdf(factor) {
  const isStageViewing = state.viewMode === "full" || els.pdfViewer?.classList.contains("stage-active");
  if (!isStageViewing) {
    zoomThumbs(factor);
    return;
  }
  if (state.excelWorkbook) {
    state.excelScale = Math.min(3, Math.max(0.35, state.excelScale * factor));
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-zoom", value: state.excelScale }, "*");
    return;
  }
  if (els.pdfPageImage.hidden) return;
  state.view.scale = Math.min(8, Math.max(0.05, state.view.scale * factor));
  state.view.userZoomed = true;
  updateViewTransform();
}

function updateScaleIndicator() {
  if (els.scaleResetBtn) {
    const pct = Math.round((state.treeScale || 1) * 100);
    els.scaleResetBtn.textContent = `${pct}%`;
  }
}

function zoomThumbs(factor) {
  state.thumbScale = Math.min(5, Math.max(0.4, Number((state.thumbScale * factor).toFixed(2))));
  els.pdfThumbs.style.setProperty("--thumb-scale", String(state.thumbScale));
  try {
    localStorage.setItem("launcher_thumb_scale", String(state.thumbScale));
  } catch {}
  updateScaleIndicator();
}

function zoomTree(factor) {
  state.treeScale = Math.min(2.5, Math.max(0.75, Number(((state.treeScale || 1) * factor).toFixed(2))));
  document.documentElement.style.setProperty("--tree-scale", String(state.treeScale));
  try {
    localStorage.setItem("launcher_tree_scale", String(state.treeScale));
  } catch {}
  updateScaleIndicator();
}

function applyPageView() {
  if (state.view.userZoomed) {
    state.view.panX = 0;
    state.view.panY = 0;
    updateViewTransform();
    return;
  }
  fitPdfPage();
}

function setStageActive(active) {
  const isAct = Boolean(active);
  if (els.pdfViewer) els.pdfViewer.classList.toggle("stage-active", isAct);
  if (els.pdfStage) els.pdfStage.classList.toggle("active-stage", isAct);
  const showControls = state.viewMode === "full" || isAct;
  if (els.viewerControls) els.viewerControls.hidden = !showControls;
  if (els.viewRotate) els.viewRotate.hidden = !showControls || Boolean(state.excelWorkbook);
  if (els.viewPanMode) els.viewPanMode.hidden = !showControls;
  if (isAct) {
    requestAnimationFrame(() => requestAnimationFrame(fitPdfPage));
  }
}

function setViewerMode(mode) {
  const wasFull = state.viewMode === "full";
  state.viewMode = mode;
  state.view.userZoomed = false;
  els.shell.classList.toggle("full-view", mode === "full");
  els.viewStandardMode.classList.toggle("active", mode === "standard");
  els.viewFullMode.classList.toggle("active", mode === "full");
  const showControls = mode === "full" || els.pdfViewer?.classList.contains("stage-active");
  if (els.viewRotate) els.viewRotate.hidden = !showControls || Boolean(state.excelWorkbook);
  if (els.viewPanMode) els.viewPanMode.hidden = !showControls;
  if (els.viewerControls) els.viewerControls.hidden = !showControls;
  if (mode === "full") {
    setStageActive(true);
  }
  if (state.revealedPath) revealPathInTree(state.revealedPath);
  // Возврат из полноэкранного: лента прокручивается к активной миниатюре,
  // чтобы вернуться в ту же позицию, а не неизвестно куда.
  if (wasFull && mode !== "full" && state.activePageKey) {
    try {
      const thumbEl = els.pdfThumbs?.querySelector(`.pdf-thumb[data-page-key="${CSS.escape(state.activePageKey)}"]`);
      if (thumbEl) thumbEl.scrollIntoView({ behavior: "auto", block: "nearest", inline: "center" });
    } catch (e) {}
  }
  requestAnimationFrame(() => requestAnimationFrame(() => {
    fitPdfPage();
  }));
}

function chunkPreviewItems(items) {
  const chunks = [];
  let currentChunk = [];
  let currentType = null;
  const maxForType = (type) => (type === "DWG_MODEL" ? 1 : type === "WORD" ? 1 : type === "EXCEL" ? 1 : 10);

  for (const item of items) {
    const type = item.previewType || "PDF";
    if (currentChunk.length > 0 && (currentType !== type || currentChunk.length >= maxForType(type))) {
      chunks.push(currentChunk);
      currentChunk = [];
    }
    currentType = type;
    currentChunk.push(item);
  }
  if (currentChunk.length > 0) {
    chunks.push(currentChunk);
  }
  return chunks;
}

const PDF_RENDER_CONCURRENCY = 3;
// DWG рендерится через общий AutoCAD COM: параллельные запуски дерутся
// за CAD-процесс, поэтому пачки DWG_MODEL идут строго по одной, а PDF —
// по-прежнему до трёх параллельно. Семафор общий на все показы: эпохи
// и отмена проверяются внутри renderBatch, зависший слот отдаётся в finally.
const DWG_RENDER_CONCURRENCY = 1;
let dwgActiveCount = 0;
const dwgWaiters = [];
function acquireDwgSlot() {
  if (dwgActiveCount < DWG_RENDER_CONCURRENCY) {
    dwgActiveCount += 1;
    return Promise.resolve();
  }
  return new Promise((resolve) => dwgWaiters.push(resolve));
}
function releaseDwgSlot() {
  dwgActiveCount = Math.max(0, dwgActiveCount - 1);
  const next = dwgWaiters.shift();
  if (next) {
    dwgActiveCount += 1;
    next();
  }
}
const PDF_FETCH_TIMEOUT_MS = 300000;
// 15 секунд молотили /api/objects/diff каждые 15 с (живое облако H:, полный
// скан на каждое изменение) и могли перезагружать дерево под руками. 10 минут.
const DIFF_POLL_INTERVAL_MS = 600000;
const PDF_PREVIEW_DPI = 150;
const PDF_QUALITY_DPI = 300;

function formatBytes(bytes) {
  if (!bytes || bytes <= 0) return "";
  if (bytes < 1024) return `${bytes} Б`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} КБ`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} МБ`;
}

function formatEta(ms) {
  if (!ms || ms < 0) return "";
  const sec = Math.ceil(ms / 1000);
  if (sec < 60) return `~${sec} сек.`;
  const min = Math.ceil(sec / 60);
  return `~${min} мин.`;
}

async function renderExcelWorkbooks(excelFiles, options = {}) {
  setTreeBrowseMode(false);
  const isSingle = Boolean(options.singleFile) || excelFiles.length === 1;
  if (options.fullView) {
    setViewerMode("full");
  } else if (state.viewMode !== "full") {
    setViewerMode("standard");
  }
  setStageActive(isSingle);
  resetPdfPreview();
  startProgress("Подготовка Excel", `${excelFiles.length} книг · HTML-просмотр без редактирования`);
  // Долгое чтение: полоска не должна замирать на 92.
  // Цифры на шкале — только движение, не обещание «почти готово».
  state.progressEase = {
    base: 8,
    spent0: Date.now(),
    estMs: 90000,
  };
  const workbooks = [];
  const failed = [];
  for (let index = 0; index < excelFiles.length; index += 1) {
    if (state.progressCancelled) {
      beaconPerf("excel-open-cancelled", 0, `before-fetch ${excelFiles.length} files`);
      return;
    }
    const excelFile = excelFiles[index];
    els.progressDetail.textContent = `Книга ${index + 1} из ${excelFiles.length}: ${excelFile.name} — читается…`;
    const controller = createOperationController();
    const timeoutSec = 300;
    const timeoutId = setTimeout(() => controller.abort(), timeoutSec * 1000);
    try {
      const response = await fetch("/api/excel/workbook", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file: excelFile.path }),
        signal: controller.signal,
      });
      clearTimeout(timeoutId);
      if (state.progressCancelled) return;
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || `Не удалось прочитать файл: ${excelFile.name}`);
      workbooks.push(payload);
      const doneShare = workbooks.length / excelFiles.length;
      state.progressShown = Math.max(state.progressShown || 0, 8 + doneShare * 80);
      state.progressEase = {
        base: state.progressShown,
        spent0: Date.now(),
        estMs: 60000,
      };
      els.progressDetail.textContent = `Готово ${workbooks.length} из ${excelFiles.length}: ${excelFile.name}`;
    } catch (err) {
      clearTimeout(timeoutId);
      if (state.progressCancelled) {
        beaconPerf("excel-open-cancelled", 0, excelFile.name || "");
        return;
      }
      const isTimeout = err?.name === "AbortError";
      const errMsg = isTimeout
        ? `Не удалось прочитать файл (превышено время ожидания ${timeoutSec} с)`
        : (err?.message || `Не удалось прочитать файл: ${excelFile.name}`);
      failed.push(`${excelFile.name}: ${errMsg}`);
      els.progressDetail.textContent = `Ошибка в книге ${index + 1} из ${excelFiles.length}: ${excelFile.name} — пропускаем`;
    }
  }
  if (state.progressCancelled) {
    beaconPerf("excel-open-cancelled", 0, `before-show ${workbooks.length} books`);
    return;
  }
  if (!workbooks.length) {
    const firstError = failed[0] || `Не удалось прочитать файл: ${excelFiles[0]?.name || "Excel"}`;
    showOperationError(new Error(firstError));
    return;
  }
  await showExcelWorkbooks(workbooks, { isSingle, fullView: options.fullView });
  // Ждать рисунок листа только если сразу открывается большое окно.
  // Несколько книг: лента уже должна быть видна, полоску на 20 секунд не держать.
  if (!state.progressCancelled && isSingle) {
    await waitForExcelSheetPaint(20000).catch(() => {});
  }
  const sheets = workbooks.reduce((total, workbook) => total + (workbook.sheets?.length || 0), 0);
  const doneText = `Готово ${workbooks.length} из ${excelFiles.length} книг · ${sheets} листов · лимит ${workbooks[0]?.maxRows || 2000} строк на лист`;
  finishProgress(failed.length ? `${doneText} · ошибок: ${failed.length}` : doneText);
}

async function renderSelectedFiles() {
  const previewItems = collectPreviewFilesForDisplay();
  if (!previewItems.length) {
    startProgress("Превью не найдено", "Выберите папку или файлы (PDF, DWG, Word, Excel).");
    finishProgress("Превью не найдено");
    return;
  }
  const excelItems = previewItems.filter((item) => item.previewType === "EXCEL");
  if (excelItems.length === previewItems.length && excelItems.length <= 10) {
    await renderExcelWorkbooks(excelItems);
    return;
  }
  await renderSelectedPdfFiles(previewItems);
}

async function renderSelectedPdfFiles(previewItems = collectPreviewFilesForDisplay(), options = {}) {
  if (!previewItems.length) {
    startProgress("Превью не найдено", "Выберите папку или файлы (PDF, DWG, Word, Excel).");
    finishProgress("Превью не найдено");
    return;
  }

  const isSingle = options.singleFile || previewItems.length === 1;
  // Новый показ гасит предыдущий: поднимаем эпоху и добиваем старые запросы.
  // Иначе два параллельных рендера месят одну ленту и вешают страницу
  // (типично: PDF ещё грузится, а уже дважды щёлкнули DWG).
  stopProgress(false, true);
  const myEpoch = (state.renderEpoch = (state.renderEpoch || 0) + 1);
  // Показ выделение не меняет: оно уже выставлено вызывающим
  // (чип — все файлы, файл — один, «Отобразить» — текущее).
  // Лишь перекрашиваем дерево на месте, без перестройки.
  updateTreeSelectionHighlight();
  setTreeBrowseMode(false);
  // TXT-одиночка — тоже сначала в ленту (как PDF/Word/Excel):
  // крупно только по двойному клику на миниатюре.
  if (options.fullView) {
    setViewerMode("full");
  } else if (state.viewMode !== "full") {
    setViewerMode("standard");
  }

  // Большие папки: загружаем первые 200 файлов, остальные автоподгрузкой при скролле
  const FIRST_PAGE_SIZE = 200;
  let itemsToRender = previewItems;
  if (!isSingle && previewItems.length > FIRST_PAGE_SIZE) {
    state.remainingPreviewItems = previewItems.slice(FIRST_PAGE_SIZE);
    state.totalPreviewCount = previewItems.length;
    itemsToRender = previewItems.slice(0, FIRST_PAGE_SIZE);
  } else if (!isSingle) {
    state.remainingPreviewItems = [];
    state.totalPreviewCount = previewItems.length;
  } else {
    state.remainingPreviewItems = [];
    state.totalPreviewCount = 1;
  }

  resetPdfPreview();

  const batches = chunkPreviewItems(itemsToRender);
  const pageGroups = batches.map((batch) => {
    return batch.filter((item) => item?.type === "missing-preview" || item?.type === "native-file" || item?.previewType === "IMAGE" || item?.previewType === "TXT");
  });
  const renderableBatchCount = batches.filter((batch) => {
    return batch.some((item) => item?.type !== "missing-preview" && item?.type !== "native-file" && item?.previewType !== "IMAGE" && item?.previewType !== "TXT");
  }).length;

  const allErrors = [];
  let totalPages = 0;
  let renderedPages = 0;
  let nextBatchIndex = 0;
  let completedBatches = batches.length - renderableBatchCount;
  let completedFiles = 0;
  // Счётчики только для DWG: пачки DWG_MODEL всегда по одному файлу,
  // текст «Модель N из M» опирается на них, а не на общий completedFiles.
  const totalDwgFiles = itemsToRender.filter((item) => item?.previewType === "DWG_MODEL").length;
  let completedDwgFiles = 0;
  const renderStartTime = Date.now();

  const firstItem = itemsToRender[0];
  const isCloud = (firstItem?.path || "").startsWith("H:") || (firstItem?.path || "").includes("Общие диски");
  const singleDetail = `${firstItem?.name || "Файл"} — создаём страницы для просмотра${isCloud ? " · Файл расположен в облачном диске" : ""}`;

  startProgress(
    isSingle ? "Подготовка документа" : "Подготовка превью",
    isSingle ? singleDetail : `${itemsToRender.length} файлов · подготовка...`
  );
  // Ensure UI updates before heavy processing
  await new Promise(requestAnimationFrame);

  function refreshRenderProgress() {
    if (state.progressCancelled) return;
    const percent = Math.min(99, Math.round((completedFiles / itemsToRender.length) * 100));
    state.progressShown = percent;
    const elapsed = Date.now() - renderStartTime;
    state.progressEase = {
      base: percent,
      spent0: Date.now(),
      estMs: estimateBatchRemainMs(elapsed, completedFiles, itemsToRender.length),
    };
    els.progressValue.textContent = `${percent}%`;
    els.progressFill.style.width = `${percent}%`;
    if (isSingle) {
      els.progressDetail.textContent = singleDetail;
    } else {
      const eta = completedFiles > 0 ? formatEta((elapsed / completedFiles) * (itemsToRender.length - completedFiles)) : "";
      const totalLabel = state.totalPreviewCount > itemsToRender.length
        ? ` (из ${state.totalPreviewCount} всего)`
        : "";
      const batchCount = Math.max(1, batches.length);
      const curBatch = batches[Math.min(completedBatches, batchCount - 1)] || [];
      const firstIdx = completedFiles + 1;
      const lastIdx = Math.min(itemsToRender.length, completedFiles + curBatch.length);
      const curName = (curBatch[0] && curBatch[0].name) || "";
      els.progressDetail.textContent =
        `Пачка ${Math.min(completedBatches + 1, batchCount)} из ${batchCount}`
        + ` · Файлы ${firstIdx}–${lastIdx} из ${itemsToRender.length}${totalLabel}`
        + (curName ? ` · Сейчас: ${curName}` : "")
        + ` · Готово: ${completedFiles} файлов`
        + ` · стр. ${renderedPages}${eta ? ` · ${eta}` : ""}`;
    }
    // Лента пополняется инкрементально из renderBatch (appendPagesToViewer):
    // полная перестройка здесь больше не нужна и вешала страницу.
  }

  async function renderBatch(batchIndex) {
    if (state.progressCancelled || state.renderEpoch !== myEpoch) return;
    const batch = batches[batchIndex];
    const itemsToFetch = batch.filter((item) => item?.type !== "missing-preview" && item?.type !== "native-file" && item?.previewType !== "IMAGE" && item?.previewType !== "TXT");
    if (!itemsToFetch.length) {
      // Готовые элементы (картинки, карточки) — сразу в ленту, fetch не нужен.
      const ready = pageGroups[batchIndex] || [];
      if (ready.length) await appendPagesToViewer(ready, null, myEpoch);
      return;
    }
    if (state.renderEpoch !== myEpoch) return;

    const isDwg = itemsToFetch[0]?.previewType === "DWG_MODEL";
    const isWord = itemsToFetch[0]?.previewType === "WORD";
    const isExcel = itemsToFetch[0]?.previewType === "EXCEL";
    // DWG: сначала быстрая первая страница для мгновенного первого просмотра,
    // затем полный проход в фоне для остальных листов чертежа.
    const dwgFirstOnlyPasses = isDwg ? [true, false] : [false];
    // Облачный диск отвечает медленно: таймаут пачки масштабируем числом
    // файлов (до PDF_FETCH_TIMEOUT_MS), а оборванную по таймауту пачку
    // повторяем один раз — сервер тем временем продолжает рендер и греет
    // кэш, повтор обычно забирает уже готовые страницы.
    const batchTimeout = isDwg ? 600000 : isWord ? 180000 : isExcel ? 180000
      : Math.max(180000, Math.min(PDF_FETCH_TIMEOUT_MS, 60000 * Math.max(1, itemsToFetch.length)));
    let controller = null;
    let batchDone = false;
    for (const firstOnlyPass of dwgFirstOnlyPasses) {
      if (state.progressCancelled || state.renderEpoch !== myEpoch) return;
      let attempt = 0;
      while (!batchDone && attempt < 2) {
      attempt += 1;
      if (attempt > 1 && !isSingle) {
        els.progressDetail.textContent = (isDwg && totalDwgFiles > 0)
          ? `Модель ${completedDwgFiles + 1} из ${totalDwgFiles}: ${itemsToFetch[0]?.name || "чертёж"} — повторная попытка…`
          : `PDF ${completedFiles} из ${itemsToRender.length} · повторная попытка: ${itemsToFetch[0]?.name || "пачка"}`;
      }
      if (attempt === 1 && isDwg && totalDwgFiles > 0 && !state.progressCancelled) {
        els.progressDetail.textContent = `Модель ${completedDwgFiles + 1} из ${totalDwgFiles}: ${itemsToFetch[0]?.name || "чертёж"} — рендерится…`;
      }
    try {
      controller = new AbortController();
      state.operationControllers.push(controller);
      const timeoutId = setTimeout(() => controller.abort(), batchTimeout);

      const endpoint = isWord
        ? "/api/word/render"
        : isDwg
          ? "/api/dwg/model-render"
          : isExcel
            ? "/api/excel/render"
            : "/api/pdf/render";

      const requestBody = {
        files: itemsToFetch.map((file) => file.path),
        dpi: PDF_PREVIEW_DPI,
        firstPageOnly: firstOnlyPass,
      };

      const _batchT0 = performance.now();
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody),
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      if (state.progressCancelled) return;
      const payload = await response.json();
      const _hits = (payload.documents || []).filter((d) => d.cacheHit).length;
      beaconPerf("display-batch", performance.now() - _batchT0, `files=${itemsToFetch.length} hits=${_hits}/${itemsToFetch.length}`);
      if (!response.ok) {
        if (state.progressCancelled) return;
        if (firstOnlyPass) break;
        // Пачка целиком отклонена сервером (файл пропал, 400/500): считаем
        // ошибкой, иначе финиш соврёт «Готово: 0 страниц» при пустой ленте.
        allErrors.push({
          document: itemsToFetch.map((f) => f.name).join(", "),
          error: payload.error || `HTTP ${response.status}`,
        });
        const errCards = itemsToFetch.map((item) => ({
          type: "missing-preview",
          name: item?.name || "Документ",
          documentPath: item?.path || "",
          sourcePath: item?.path || "",
          sourceType: item?.extension || "PDF",
          message: payload.error || "Ошибка превью. Откройте файл через кнопку «Открыть».",
        }));
        pageGroups[batchIndex] = errCards;
        await appendPagesToViewer(errCards, null, myEpoch);
        batchDone = true;
      } else {
        if (state.progressCancelled || state.renderEpoch !== myEpoch) return;
        if (!firstOnlyPass) {
          totalPages += payload.totalPages || 0;
          renderedPages += payload.renderedPages || 0;
        }
        allErrors.push(...(payload.errors || []));

        const batchPages = payload.documents.flatMap((document) => {
          const matchedItem = itemsToFetch.find((f) => f.path === (document.sourcePath || document.path))
            || itemsToFetch.find((f) => f.name === (document.sourceName || document.name))
            || itemsToFetch[0];
          const realSourcePath = document.sourcePath || matchedItem?.previewFor?.path || matchedItem?.path || document.path;
          const realName = matchedItem?.previewFor ? matchedItem.previewFor.name : (document.sourceName || document.name);
          return document.items.map((page) => ({
            ...page,
            name: `${realName} · стр. ${page.page}`,
            documentPath: document.path,
            sourcePath: realSourcePath,
            sourceName: realName,
            sourceType: document.sourceType || matchedItem?.extension || "",
            dpi: payload.dpi,
            cacheHit: document.cacheHit,
            previewFor: matchedItem?.previewFor || (realSourcePath ? { path: realSourcePath, name: realName } : null),
            previewType: matchedItem?.previewType || document.sourceType || "",
            totalPages: document.pages || 1,
            rawDoc: document,
          }));
        });
        pageGroups[batchIndex] = batchPages;
        await appendPagesToViewer(batchPages, null, myEpoch);
        if (firstOnlyPass) {
          const docPages = payload.documents?.[0]?.pages || 1;
          if (docPages <= 1) { batchDone = true; }
          break;
        }
        batchDone = true;
      }
    } catch (error) {
      if (state.progressCancelled || state.renderEpoch !== myEpoch) return;
      const isTimeout = error?.name === "AbortError";
      if (isTimeout && attempt < 2) continue;
      if (firstOnlyPass) break;
      // Финальный провал пачки (включая повтор): тоже ошибка для честности
      // финишной строки, иначе будет «Готово: 0 страниц» при пустой ленте.
      allErrors.push({
        document: itemsToFetch.map((f) => f.name).join(", "),
        error: isTimeout ? "timeout" : String(error?.message || error),
      });
      const errCards = itemsToFetch.map((item) => {
        const sz = item?.size ? ` (${formatBytes(item.size)})` : "";
        return {
          type: "missing-preview",
          name: item?.name || "Документ",
          documentPath: item?.path || "",
          sourcePath: item?.path || "",
          sourceType: item?.extension || "PDF",
          message: isTimeout
            ? `Файл${sz} не дождался ответа облачного диска Google Drive (включая повторную попытку). Нажмите «Открыть», чтобы запустить его напрямую.`
            : "Ошибка рендеринга. Откройте файл через кнопку «Открыть».",
        };
      });
      pageGroups[batchIndex] = errCards;
      await appendPagesToViewer(errCards, null, myEpoch);
      batchDone = true;
    } finally {
      if (controller) {
        state.operationControllers = state.operationControllers.filter((item) => item !== controller);
        controller = null;
      }
    }
    } // end while
    } // end for firstOnlyPasses
  if (!state.progressCancelled && state.renderEpoch === myEpoch) {
    completedBatches += 1;
    completedFiles += batch.length;
    if ((batch[0]?.previewType) === "DWG_MODEL") completedDwgFiles += itemsToFetch.length;
    refreshRenderProgress();
  }
}

  async function worker() {
    while (!state.progressCancelled && state.renderEpoch === myEpoch && nextBatchIndex < batches.length) {
      if (state.progressCancelled || state.renderEpoch !== myEpoch) break;
      const batchIndex = nextBatchIndex++;
      // Пачки DWG_MODEL — строго по одной через общий семафор (CAD-COM),
      // остальные типы разбираются параллельно как раньше.
      if ((batches[batchIndex][0]?.previewType) === "DWG_MODEL") {
        await acquireDwgSlot();
        try {
          await renderBatch(batchIndex);
        } finally {
          releaseDwgSlot();
        }
      } else {
        await renderBatch(batchIndex);
      }
    }
  }

  refreshRenderProgress();
  const workerCount = Math.min(PDF_RENDER_CONCURRENCY, Math.max(1, renderableBatchCount));
  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  if (state.progressCancelled || state.renderEpoch !== myEpoch) return;

  refreshRenderProgress();
  updateRibbonLoadMore();

  const errorCount = allErrors.length;
  const isSingleDwg = isSingle && itemsToRender[0]?.previewType === "DWG_MODEL";
  const detail = isSingle
    ? (errorCount
      ? (isSingleDwg && allErrors[0]?.error ? `Не удалось открыть документ: ${allErrors[0].error}` : "Не удалось открыть документ")
      : `Готово: ${totalPages} страниц`)
    : (errorCount
      ? `Показано: ${renderedPages} из ${totalPages} стр. · ошибок: ${errorCount}`
      : `Готово: ${state.totalPreviewCount || itemsToRender.length} файлов · стр. ${renderedPages} из ${totalPages}`);
  finishProgress(detail);
}

async function loadMorePdfFiles(itemsToRender) {
  startProgress("Подгрузка страниц", `${itemsToRender.length} файлов — подгружаем порциями`);
  const batches = chunkPreviewItems(itemsToRender);
  let nextBatchIndex = 0;
  let completedBatches = 0;
  let completedFiles = 0;
  // Те же DWG-счётчики, что в основном показе; плюс счётчик ошибок пачек
  // для честной строки финиша.
  const totalDwgFiles = itemsToRender.filter((item) => item?.previewType === "DWG_MODEL").length;
  let completedDwgFiles = 0;
  let loadErrors = 0;
  const loadStartTime = Date.now();
  // Продолжение текущего показа: фиксируем его эпоху, чтобы подгрузка
  // сама гасла, если пользователь тем временем начал новый показ.
  const myEpoch = state.renderEpoch || 0;

  function refreshRenderProgress() {
    if (state.progressCancelled) return;
    const percent = Math.min(99, Math.round((completedFiles / itemsToRender.length) * 100));
    state.progressShown = percent;
    const elapsed = Date.now() - loadStartTime;
    state.progressEase = {
      base: percent,
      spent0: Date.now(),
      estMs: estimateBatchRemainMs(elapsed, completedFiles, itemsToRender.length),
    };
    els.progressValue.textContent = `${percent}%`;
    els.progressFill.style.width = `${percent}%`;
    const eta = completedFiles > 0 ? formatEta((elapsed / completedFiles) * (itemsToRender.length - completedFiles)) : "";
    const remaining = state.remainingPreviewItems?.length || 0;
    const batchCount = Math.max(1, batches.length);
    const curBatch = batches[Math.min(completedBatches, batchCount - 1)] || [];
    const firstIdx = completedFiles + 1;
    const lastIdx = Math.min(itemsToRender.length, completedFiles + curBatch.length);
    const curName = (curBatch[0] && curBatch[0].name) || "";
    els.progressDetail.textContent =
      `Пачка ${Math.min(completedBatches + 1, batchCount)} из ${batchCount}`
      + ` · Файлы ${firstIdx}–${lastIdx} из ${itemsToRender.length} (в очереди ещё ${remaining})`
      + (curName ? ` · Сейчас: ${curName}` : "")
      + ` · Готово: ${completedFiles} файлов${eta ? ` · ${eta}` : ""}`;
  }

  async function renderBatch(batchIndex) {
    if (state.progressCancelled || state.renderEpoch !== myEpoch) return;
    const batch = batches[batchIndex];
    const itemsToFetch = batch.filter((item) => item?.type !== "missing-preview" && item?.type !== "native-file" && item?.previewType !== "IMAGE" && item?.previewType !== "TXT");
    if (!itemsToFetch.length) {
      if (batch.length) await appendPagesToViewer(batch, null, myEpoch);
      completedBatches += 1;
      return;
    }
    if (state.renderEpoch !== myEpoch) {
      completedBatches += 1;
      return;
    }
    let controller = null;
    const isDwg = itemsToFetch[0]?.previewType === "DWG_MODEL";
    const isWord = itemsToFetch[0]?.previewType === "WORD";
    const isExcel = itemsToFetch[0]?.previewType === "EXCEL";
    // Тот же масштабируемый таймаут и один автоповтор, что в основном показе:
    // облачный диск отвечает медленно, сервер продолжает рендер и греет кэш.
    const batchTimeout = isDwg ? 600000 : isWord ? 180000 : isExcel ? 180000
      : Math.min(PDF_FETCH_TIMEOUT_MS, 30000 * Math.max(1, itemsToFetch.length));
    let attempt = 0;
    let batchDone = false;
    while (!batchDone && attempt < 2) {
      attempt += 1;
      if (attempt === 1 && isDwg && totalDwgFiles > 0 && !state.progressCancelled) {
        els.progressDetail.textContent = `Модель ${completedDwgFiles + 1} из ${totalDwgFiles}: ${itemsToFetch[0]?.name || "чертёж"} — рендерится…`;
      }
    try {
      controller = new AbortController();
      state.operationControllers.push(controller);
      const timeoutId = setTimeout(() => controller.abort(), batchTimeout);

      const endpoint = isWord
        ? "/api/word/render"
        : isDwg
          ? "/api/dwg/model-render"
          : "/api/pdf/render";

      const _batchT0 = performance.now();
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ files: itemsToFetch.map((file) => file.path), dpi: PDF_PREVIEW_DPI, firstPageOnly: false }),
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      if (state.progressCancelled) return;
      const payload = await response.json();
      const _hits = (payload.documents || []).filter((d) => d.cacheHit).length;
      beaconPerf("display-batch", performance.now() - _batchT0, `files=${itemsToFetch.length} hits=${_hits}/${itemsToFetch.length}`);
      if (response.ok) {
        const batchPages = payload.documents.flatMap((document) => {
          const matchedItem = itemsToFetch.find((f) => f.path === document.path) || itemsToFetch[0];
          return document.items.map((page) => ({
            ...page,
            name: `${matchedItem?.previewFor ? matchedItem.previewFor.name : document.name} · стр. ${page.page}`,
            documentPath: document.path,
            dpi: payload.dpi,
            cacheHit: document.cacheHit,
            previewFor: matchedItem?.previewFor || null,
            previewType: matchedItem?.previewType || "",
          }));
        });
        if (state.renderEpoch !== myEpoch) return;
        await appendPagesToViewer(batchPages, null, myEpoch);
        batchDone = true;
      } else {
        const errCards = itemsToFetch.map((item) => ({
          type: "missing-preview",
          name: item?.name || "Документ",
          documentPath: item?.path || "",
          sourcePath: item?.path || "",
          sourceType: item?.extension || "PDF",
          message: payload.error || "Ошибка превью. Откройте файл через кнопку «Открыть».",
        }));
      if (state.renderEpoch === myEpoch) await appendPagesToViewer(errCards, null, myEpoch);
      loadErrors += itemsToFetch.length;
      batchDone = true;
      }
    } catch (err) {
      if (state.progressCancelled || state.renderEpoch !== myEpoch) return;
      const isTimeout = err?.name === "AbortError";
      if (isTimeout && attempt < 2) continue;
      const errCards = itemsToFetch.map((item) => ({
        type: "missing-preview",
        name: item?.name || "Документ",
        documentPath: item?.path || "",
        sourcePath: item?.path || "",
        sourceType: item?.extension || "PDF",
          message: isTimeout
            ? "Файл не дождался ответа облачного диска Google Drive (включая повторную попытку). Нажмите «Открыть» для прямого запуска."
            : "Ошибка рендеринга. Откройте файл через кнопку «Открыть».",
        }));
        if (state.renderEpoch === myEpoch) await appendPagesToViewer(errCards, null, myEpoch);
        loadErrors += itemsToFetch.length;
        batchDone = true;
      } finally {
      if (controller) {
        state.operationControllers = state.operationControllers.filter((item) => item !== controller);
        controller = null;
      }
    }
  }
  if (!state.progressCancelled && state.renderEpoch === myEpoch) {
    completedBatches += 1;
    completedFiles += batch.length;
    if ((batch[0]?.previewType) === "DWG_MODEL") completedDwgFiles += itemsToFetch.length;
    refreshRenderProgress();
  }
}

  async function worker() {
    while (!state.progressCancelled && state.renderEpoch === myEpoch && nextBatchIndex < batches.length) {
      if (state.progressCancelled || state.renderEpoch !== myEpoch) break;
      const batchIndex = nextBatchIndex++;
      // Пачки DWG_MODEL — строго по одной через общий семафор (CAD-COM),
      // остальные типы разбираются параллельно как раньше.
      if ((batches[batchIndex][0]?.previewType) === "DWG_MODEL") {
        await acquireDwgSlot();
        try {
          await renderBatch(batchIndex);
        } finally {
          releaseDwgSlot();
        }
      } else {
        await renderBatch(batchIndex);
      }
    }
  }

  const workerCount = Math.min(PDF_RENDER_CONCURRENCY, Math.max(1, batches.length));
  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  if (state.progressCancelled) return;

  updateRibbonLoadMore();
  finishProgress(loadErrors > 0 ? `Дополнительные превью загружены · ошибок: ${loadErrors}` : "Дополнительные превью загружены");
}


els.load.addEventListener("click", () => importObject(false).catch(showOperationError));
els.refresh.addEventListener("click", () => {
  if (inTreeMode()) refreshObjectInPlace().catch(showOperationError);
  else importObject(true).catch(showOperationError);
});
els.display.addEventListener("click", () => {
  if (inTreeMode()) renderSelectedFiles().catch(showOperationError);
  else openSelectedObject().catch(showOperationError);
});
els.exclude.addEventListener("click", () => {
  stopDiffPolling();
  excludeSelectedObject().catch(showOperationError);
});
els.cancel.addEventListener("click", () => stopProgress(true));
els.backToObjects.addEventListener("click", () => {
  stopDiffPolling();
  setMode("objects");
});
els.backToTree.addEventListener("click", () => {
  if (!state.currentManifest?.tree) return;
  if (state.viewMode === "full") {
    setViewerMode("standard");
    if (state.excelWorkbooks?.length > 1 && els.excelViewer) {
      setStageActive(false);
      els.excelViewer.hidden = true;
      if (els.wordViewer) els.wordViewer.hidden = true;
      els.pdfViewer?.classList.remove("empty");
      renderExcelWorkbookRail();
    }
    return;
  }
  if (els.pdfViewer?.classList.contains("stage-active")) {
    setStageActive(false);
    if (state.excelWorkbooks?.length > 1 && els.excelViewer) {
      els.excelViewer.hidden = true;
      if (els.wordViewer) els.wordViewer.hidden = true;
      els.pdfViewer?.classList.remove("empty");
      renderExcelWorkbookRail();
    }
    return;
  }
  setViewerMode("standard");
  setTreeBrowseMode(true);
  renderTree();
});
els.treeSearch.addEventListener("input", () => {
  renderTree();
});
els.pdfPageImage.addEventListener("load", () => applyPageView());
els.excelSheetFrame.addEventListener("load", () => {
  forwardExcelSheetContextMenu();
  if (!state.excelWorkbook) return;
  els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-zoom", value: state.excelScale }, "*");
});
els.excelTabsLeft.addEventListener("click", () => {
  els.excelTabs.scrollBy({ left: -Math.max(180, els.excelTabs.clientWidth * .72), behavior: "smooth" });
});
els.excelTabsRight.addEventListener("click", () => {
  els.excelTabs.scrollBy({ left: Math.max(180, els.excelTabs.clientWidth * .72), behavior: "smooth" });
});
els.excelTabs.addEventListener("wheel", (event) => {
  if (!event.shiftKey || !event.deltaY) return;
  event.preventDefault();
  els.excelTabs.scrollLeft += event.deltaY;
}, { passive: false });
els.viewZoomOut.addEventListener("click", () => zoomPdf(0.82));
els.viewZoomIn.addEventListener("click", () => zoomPdf(1.22));
// «Вписать» возвращает всё в нормальное состояние: масштаб картинки,
// сдвиг и поворот. Панель при этом никуда не уезжает.
if (els.viewOpenNative) {
  els.viewOpenNative.addEventListener("click", () => {
    const info = getActiveSourceInfo();
    const wb = state.excelWorkbook;
    const targetPath = (info && info.path) || (wb && wb.path) || state.activeNativePath || "";
    if (targetPath) {
      openFileByPath(targetPath, "native");
    }
  });
}

els.viewFit.addEventListener("click", () => {
  if (state.excelWorkbook) {
    state.view.rotation = 0;
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-rotate", value: 0 }, "*");
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-fit" }, "*");
    return;
  }
  if (state.viewMode === "medium") {
    state.thumbScale = 1;
    els.pdfThumbs.style.setProperty("--thumb-scale", "1");
    return;
  }
  state.view.rotation = 0;
  state.view.userZoomed = false;
  fitPdfPage();
});
els.viewRotate.addEventListener("click", () => {
  // Во 2-м режиме поворот полностью заблокирован: поворот работает только для большого окна в 1-м и 3-м режимах
  if (state.viewMode === "medium") return;

  if (state.excelWorkbook) {
    state.view.rotation = (state.view.rotation + 90) % 360;
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-rotate", value: state.view.rotation }, "*");
    els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-fit" }, "*");
    return;
  }

  // В 1-м и 3-м режимах вращается ТОЛЬКО главное окно большого просмотра (миниатюры не вращаются вообще)
  state.view.rotation = (state.view.rotation + 90) % 360;
  fitPdfPage();
});
els.viewPanMode.addEventListener("click", () => {
  state.view.panMode = !state.view.panMode;
  updateViewTransform();
});
if (els.viewStandardMode) els.viewStandardMode.addEventListener("click", () => setViewerMode("standard"));
if (els.viewFullMode) els.viewFullMode.addEventListener("click", () => setViewerMode("full"));

// Блокируем масштабирование всего окна браузера (Ctrl + / Ctrl - / Ctrl + wheel в пустых местах),
// чтобы каркас окон, рамки и разметка никогда не «плыли»
window.addEventListener("wheel", (event) => {
  if (event.ctrlKey) {
    event.preventDefault();
  }
}, { passive: false });

window.addEventListener("keydown", (event) => {
  if (event.ctrlKey && (event.key === "+" || event.key === "=" || event.key === "-" || event.key === "_" || event.key === "0")) {
    event.preventDefault();
  }
});

els.pdfPageImage.addEventListener("dblclick", (e) => {
  e.stopPropagation();
  setViewerMode(state.viewMode === "full" ? "standard" : "full");
});

if (els.pdfStage) {
  els.pdfStage.addEventListener("dblclick", (e) => {
    if (e.target.closest(".viewer-controls, .excel-tabs-bar, button, a, input")) return;
    setViewerMode(state.viewMode === "full" ? "standard" : "full");
  });
  els.pdfStage.addEventListener("contextmenu", (e) => {
    // ПКМ в полном экране: меню исходного файла. По пустому полю тоже
    // работает, если сейчас отображается документ.
    e.preventDefault();
    e.stopPropagation();
    hideFileContextMenu();
    const info = getActiveSourceInfo();
    if (!info.path) return;
    showFileContextMenu(e.clientX, e.clientY, { path: info.path, isDir: false, ext: info.ext });
  });
}

if (els.excelBookTitle) {
  els.excelBookTitle.addEventListener("dblclick", (e) => {
    e.stopPropagation();
    setViewerMode(state.viewMode === "full" ? "standard" : "full");
  });
}

if (els.pdfPageImage) {
  els.pdfPageImage.addEventListener("contextmenu", (event) => {
    if (!state.activeNativePath) return;
    event.preventDefault();
    event.stopPropagation();
    showFileContextMenu(event.clientX, event.clientY, {
      path: state.activeNativePath,
      isDir: false,
      ext: extOfPath(state.activeNativePath),
    });
  });
}

if (els.excelViewer) {
  els.excelViewer.addEventListener("contextmenu", (event) => {
    const wbPath = state.excelWorkbook?.path || "";
    if (!wbPath) return;
    event.preventDefault();
    event.stopPropagation();
    showFileContextMenu(event.clientX, event.clientY, {
      path: wbPath,
      isDir: false,
      ext: state.excelWorkbook?.extension || extOfPath(wbPath),
    });
  });
}

window.addEventListener("message", (event) => {
  if (event.data?.type === "launcher-toggle-full-view") {
    setViewerMode(state.viewMode === "full" ? "standard" : "full");
  } else if (event.data?.type === "launcher-escape") {
    handleGlobalEscape();
  }
});

window.addEventListener("resize", () => {
  if (state.viewMode === "full" || els.pdfViewer?.classList.contains("stage-active")) {
    fitPdfPage();
  }
});

// Масштабирование только содержимого дерева (колёсико над левой частью)
if (els.sidebar) {
  els.sidebar.addEventListener("wheel", (event) => {
    if (!event.ctrlKey || !event.deltaY) return;
    event.preventDefault();
    event.stopPropagation();
    zoomTree(event.deltaY < 0 ? 1.08 : 0.92);
  }, { passive: false });
}
if (els.scaleWidget) {
  els.scaleWidget.addEventListener("wheel", (event) => {
    if (!event.deltaY) return;
    event.preventDefault();
    event.stopPropagation();
    const factor = event.deltaY < 0 ? 1.08 : 0.92;
    zoomTree(factor);
    zoomThumbs(factor);
  }, { passive: false });
}

// Масштабирование только картинки на сцене (в полноэкранном 3-м режиме)
els.pdfStage.addEventListener("wheel", (event) => {
  if (event.target.closest(".viewer-controls")) return;
  if (!event.ctrlKey || (els.pdfPageImage.hidden && !state.excelWorkbook)) return;
  event.preventDefault();
  event.stopPropagation();
  zoomPdf(event.deltaY < 0 ? 1.12 : 0.89);
}, { passive: false });

// Масштабирование только миниатюр (Ctrl + колёсико над правой лентой)
els.pdfThumbs.addEventListener("wheel", (event) => {
  if (!event.ctrlKey || !event.deltaY) return;
  event.preventDefault();
  event.stopPropagation();
  zoomThumbs(event.deltaY < 0 ? 1.15 : 0.87);
}, { passive: false });

// Панель управления масштабом (кнопки −, 100%, + в верхней панели)
// 100% работает везде во всех окнах: в левом поле, во втором поле, а при просмотре чертежа дублирует «Вписать»
if (els.scaleResetBtn) {
  els.scaleResetBtn.addEventListener("click", () => {
    // 1. Сброс масштаба левого поля (дерево и список объектов)
    state.treeScale = 1;
    document.documentElement.style.setProperty("--tree-scale", "1");

    // 2. Сброс масштаба второго поля (лента миниатюр)
    state.thumbScale = 1;
    els.pdfThumbs.style.setProperty("--thumb-scale", "1");

    try {
      localStorage.setItem("launcher_tree_scale", "1");
      localStorage.setItem("launcher_thumb_scale", "1");
    } catch {}

    // 3. Единый чертёж / документ (дублирует действие «Вписать» / сброс)
    if (state.excelWorkbook) {
      state.view.rotation = 0;
      els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-rotate", value: 0 }, "*");
      els.excelSheetFrame.contentWindow?.postMessage({ type: "launcher-sheet-fit" }, "*");
    } else if (!els.pdfPageImage.hidden) {
      state.view.rotation = 0;
      state.view.userZoomed = false;
      fitPdfPage();
    }

    updateScaleIndicator();
    showToast("Масштаб: 100% (исходный)");
  });
}

if (els.scaleMinusBtn) {
  els.scaleMinusBtn.addEventListener("click", () => {
    const isStageViewing = state.viewMode === "full" || els.pdfViewer?.classList.contains("stage-active");
    if (isStageViewing && (!els.pdfPageImage.hidden || state.excelWorkbook)) {
      zoomPdf(0.89);
    } else {
      zoomTree(0.9);
      zoomThumbs(0.9);
    }
  });
}

if (els.scalePlusBtn) {
  els.scalePlusBtn.addEventListener("click", () => {
    const isStageViewing = state.viewMode === "full" || els.pdfViewer?.classList.contains("stage-active");
    if (isStageViewing && (!els.pdfPageImage.hidden || state.excelWorkbook)) {
      zoomPdf(1.12);
    } else {
      zoomTree(1.1);
      zoomThumbs(1.1);
    }
  });
}

// Восстановление сохранённых масштабов при старте
try {
  const savedTreeScale = parseFloat(localStorage.getItem("launcher_tree_scale") || "1");
  if (savedTreeScale && savedTreeScale >= 0.75 && savedTreeScale <= 2.5) {
    state.treeScale = savedTreeScale;
    document.documentElement.style.setProperty("--tree-scale", String(state.treeScale));
  }
  const savedThumbScale = parseFloat(localStorage.getItem("launcher_thumb_scale") || "1");
  if (savedThumbScale && savedThumbScale >= 0.4 && savedThumbScale <= 5) {
    state.thumbScale = savedThumbScale;
    els.pdfThumbs.style.setProperty("--thumb-scale", String(state.thumbScale));
  }
  updateScaleIndicator();
} catch {}

els.pdfStage.addEventListener("pointerdown", (event) => {
  if (event.target.closest(".viewer-controls")) return;
  if (!state.view.panMode || els.pdfPageImage.hidden || event.button !== 0) return;
  event.preventDefault();
  state.view.dragging = true;
  state.view.dragStartX = event.clientX;
  state.view.dragStartY = event.clientY;
  state.view.dragPanX = state.view.panX;
  state.view.dragPanY = state.view.panY;
  els.pdfStage.setPointerCapture(event.pointerId);
  updateViewTransform();
});

els.pdfStage.addEventListener("pointermove", (event) => {
  if (!state.view.dragging) return;
  state.view.panX = state.view.dragPanX + (event.clientX - state.view.dragStartX);
  state.view.panY = state.view.dragPanY + (event.clientY - state.view.dragStartY);
  updateViewTransform();
});

function endPdfDrag(event) {
  if (!state.view.dragging) return;
  state.view.dragging = false;
  try {
    els.pdfStage.releasePointerCapture(event.pointerId);
  } catch {
    // Pointer capture may already be released by the browser.
  }
  updateViewTransform();
}

els.pdfStage.addEventListener("pointerup", endPdfDrag);
els.pdfStage.addEventListener("pointercancel", endPdfDrag);

function handleGlobalEscape() {
  if (els.settingsModal && !els.settingsModal.hidden) {
    closeSettingsModal();
    return;
  }
  if (!els.pdfPageImage.hidden && state.view.dragging) {
    state.view.dragging = false;
    updateViewTransform();
    return;
  }
  // 1. Из полноэкранного режима возвращаемся в стандартный режим с двумя полями
  if (state.viewMode === "full") {
    setViewerMode("standard");
    if (state.excelWorkbooks?.length > 1 && els.excelViewer) {
      setStageActive(false);
      els.excelViewer.hidden = true;
      if (els.wordViewer) els.wordViewer.hidden = true;
      els.pdfViewer?.classList.remove("empty");
      renderExcelWorkbookRail();
    }
    return;
  }
  // 1.1. Если открыта сцена отдельного документа/таблицы из ленты — возвращаемся к ленте миниатюр
  if (els.pdfViewer?.classList.contains("stage-active") && (state.renderedPages?.length > 1 || state.excelWorkbooks?.length > 1)) {
    setStageActive(false);
    if (els.excelViewer) els.excelViewer.hidden = true;
    if (els.wordViewer) els.wordViewer.hidden = true;
    els.pdfViewer?.classList.remove("empty");
    renderExcelWorkbookRail();
    return;
  }
  // 2. В режиме дерева:
  if (inTreeMode()) {
    // 2.1. Если введён поисковый запрос — сбрасываем поиск
    if (els.treeSearch && els.treeSearch.value.trim()) {
      els.treeSearch.value = "";
      renderTree();
      return;
    }
    // 2.2. Если включён фильтр форматов или изменений — сбрасываем фильтры
    if (state.activeFilters.size > 0 || state.diffFilter) {
      state.activeFilters.clear();
      state.diffFilter = false;
      renderFormats();
      renderTree();
      return;
    }
    // 2.3. Если выделены узлы или строка в дереве — снимаем выделение
    if (state.selectedPaths.size > 0 || state.revealedPath) {
      state.selectedPaths.clear();
      state.revealedPath = "";
      renderTree();
      return;
    }
    // 2.4. Если в дереве ничего не выбрано — выходим назад к списку объектов!
    stopDiffPolling();
    setMode("objects");
    return;
  }
  // 3. В режиме списка объектов: если выбран объект, снимаем выбор и возвращаемся к исходному виду
  if (state.selectedObjectId !== null) {
    state.selectedObjectId = null;
    updateObjectListSelection();
    updateObjectButtons();
    updateObjectStats();
    updateViewerBanner('Политика "Локальный компьютер"');
    return;
  }
}

function stepThumbnails(delta) {
  const pages = state.renderedPages;
  if (!pages?.length) return;
  let currentIdx = pages.findIndex((p) => pageKey(p) === state.activePageKey);
  if (currentIdx < 0) currentIdx = 0;
  let nextIdx = currentIdx + delta;
  if (nextIdx < 0) nextIdx = pages.length - 1;
  else if (nextIdx >= pages.length) nextIdx = 0;

  const nextPage = pages[nextIdx];
  if (nextPage) {
    showPdfPage(nextPage);
    const targetDocPath = thumbPathForPage(nextPage);
    if (targetDocPath) {
      revealPathInTree(targetDocPath, { updateSelection: true, skipScroll: false });
    }
    const key = pageKey(nextPage);
    const thumbEl = Array.from(els.pdfThumbs?.querySelectorAll(".pdf-thumb") || []).find(
      (el) => el.dataset.pageKey === key
    );
    if (thumbEl) {
      thumbEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
    }
  }
}

window.addEventListener("keydown", (event) => {
  const activeTag = document.activeElement?.tagName;
  const isInput = activeTag === "INPUT" || activeTag === "TEXTAREA";

  if (event.key === "Escape") {
    event.preventDefault();
    handleGlobalEscape();
    return;
  }

  if (isInput) return;

  // Навигация клавишами Вверх / Вниз:
  // Если мышь/фокус в правой ленте миниатюр (Зона 2) и есть страницы:
  // ArrowUp / ArrowDown переключают миниатюры в ленте!
  if (event.key === "ArrowUp" || event.key === "ArrowDown") {
    const isThumbsActive = state.viewMode === "full" || state.activeNavZone === "thumbs" || els.pdfViewer?.contains(document.activeElement);
    if (state.renderedPages?.length && isThumbsActive) {
      event.preventDefault();
      stepThumbnails(event.key === "ArrowDown" ? 1 : -1);
      return;
    }
    if (state.excelWorkbooks?.length > 1 && isThumbsActive) {
      event.preventDefault();
      const books = state.excelWorkbooks;
      let currentIdx = state.excelWorkbookIndex;
      if (event.key === "ArrowDown") {
        currentIdx = currentIdx >= 0 && currentIdx < books.length - 1 ? currentIdx + 1 : 0;
      } else {
        currentIdx = currentIdx > 0 ? currentIdx - 1 : books.length - 1;
      }
      const isStage = Boolean(els.pdfViewer?.classList.contains("stage-active"));
      if (isStage) {
        activateExcelWorkbook(currentIdx).catch(showOperationError);
      } else {
        state.excelWorkbookIndex = currentIdx;
        renderExcelWorkbookRail();
      }
      const nextBook = books[currentIdx];
      if (nextBook?.path) {
        revealPathInTree(nextBook.path, { skipScroll: false });
      }
      const wbThumb = els.pdfThumbs?.querySelector(`.excel-book-thumb[data-workbook-index="${currentIdx}"]`);
      if (wbThumb) {
        wbThumb.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "center" });
      }
      return;
    }
    // Иначе (Зона 1 — дерево слева): навигация по списку файлов
    if (inTreeMode()) {
      const files = state.visibleRows.filter((r) => r.type === "file");
      if (!files.length) return;
      event.preventDefault();
      const currentPath = Array.from(state.selectedPaths)[0] || state.revealedPath;
      let idx = files.findIndex((f) => f.path === currentPath);
      if (event.key === "ArrowDown") {
        idx = idx >= 0 && idx < files.length - 1 ? idx + 1 : 0;
      } else {
        idx = idx > 0 ? idx - 1 : files.length - 1;
      }
      const nextFile = files[idx];
      if (nextFile) {
        selectNode(nextFile);
        const row = els.objectTree?.querySelector(`.tree-row[data-path="${CSS.escape(nextFile.path)}"]`);
        if (row) row.scrollIntoView({ block: "nearest", behavior: "smooth" });
      }
      return;
    }
  }

  // Навигация клавишами Влево / Вправо по книгам Excel
  if (state.excelWorkbooks?.length > 1 && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
    event.preventDefault();
    const books = state.excelWorkbooks;
    let currentIdx = state.excelWorkbookIndex;
    if (event.key === "ArrowRight") {
      currentIdx = currentIdx >= 0 && currentIdx < books.length - 1 ? currentIdx + 1 : 0;
    } else {
      currentIdx = currentIdx > 0 ? currentIdx - 1 : books.length - 1;
    }
    const isStage = Boolean(els.pdfViewer?.classList.contains("stage-active"));
    if (isStage) {
      activateExcelWorkbook(currentIdx).catch(showOperationError);
    } else {
      state.excelWorkbookIndex = currentIdx;
      renderExcelWorkbookRail();
    }
    const nextBook = books[currentIdx];
    if (nextBook?.path) {
      revealPathInTree(nextBook.path, { skipScroll: false });
    }
    return;
  }

  // Навигация клавишами Влево / Вправо по страницам / миниатюрам
  if (state.renderedPages?.length && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
    event.preventDefault();
    stepThumbnails(event.key === "ArrowRight" ? 1 : -1);
    return;
  }
});

// Отслеживание двух активных зон навигации:
// Зона 1: Дерево слева (папки, подпапки, файлы)
// Зона 2: Лента миниатюр справа (вертикальная лента страниц)
if (els.pdfViewer) {
  els.pdfViewer.addEventListener("pointerdown", () => { state.activeNavZone = "thumbs"; });
  els.pdfViewer.addEventListener("mouseenter", () => { state.activeNavZone = "thumbs"; });
  els.pdfViewer.addEventListener("focusin", () => { state.activeNavZone = "thumbs"; });
}
if (els.pdfThumbs) {
  els.pdfThumbs.addEventListener("pointerdown", () => { state.activeNavZone = "thumbs"; });
  els.pdfThumbs.addEventListener("mouseenter", () => { state.activeNavZone = "thumbs"; });
  els.pdfThumbs.addEventListener("focusin", () => { state.activeNavZone = "thumbs"; });
  els.pdfThumbs.addEventListener("click", (event) => {
    // Клик по пустому месту ленты (мимо карточек): снять всё выделение.
    // Клики по карточкам сюда не доходят (stopPropagation в карточке).
    // Превью при этом не трогаем.
    if (event.target.closest(".pdf-thumb")) return;
    if (!state.selectedPaths.size && !state.revealedPath) return;
    state.activeNavZone = "thumbs";
    state.selectedPaths.clear();
    state.revealedPath = "";
    state.lastSelectedIndex = null;
    setActiveNativePath("");
    updateTreeSelectionHighlight();
  });
}
if (els.objectTree) {
  els.objectTree.addEventListener("pointerdown", () => { state.activeNavZone = "tree"; });
  els.objectTree.addEventListener("mouseenter", () => { state.activeNavZone = "tree"; });
  els.objectTree.addEventListener("focusin", () => { state.activeNavZone = "tree"; });
  els.objectTree.addEventListener("click", (event) => {
    // Клик по пустому месту дерева (мимо строк): снять всё выделение.
    // Клики по строкам сюда не доходят (stopPropagation в строке).
    if (event.target.closest(".tree-row")) return;
    if (!state.selectedPaths.size && !state.revealedPath) return;
    state.activeNavZone = "tree";
    state.selectedPaths.clear();
    state.revealedPath = "";
    state.lastSelectedIndex = null;
    setActiveNativePath("");
    updateTreeSelectionHighlight();
  });
}
const treeShellContainer = document.getElementById("treeShell") || document.querySelector(".tree-panel");
if (treeShellContainer) {
  treeShellContainer.addEventListener("pointerdown", () => { state.activeNavZone = "tree"; });
  treeShellContainer.addEventListener("mouseenter", () => { state.activeNavZone = "tree"; });
}

// Отслеживание позиции курсора мыши для точной привязки к листу под курсором при прокрутке колесом
window.addEventListener("mousemove", (e) => {
  state._lastMouseX = e.clientX;
  state._lastMouseY = e.clientY;
}, { passive: true });

let _ribbonSyncRaf = null;
function syncRibbonScrollToTree() {
  if (_ribbonSyncRaf) return;
  _ribbonSyncRaf = requestAnimationFrame(() => {
    _ribbonSyncRaf = null;
    if (state.viewMode === "full" || !els.pdfThumbs) return;
    if (!state.renderedPages?.length && !state.excelWorkbooks?.length) return;

    let targetThumb = null;
    // Если курсор находится над лентой миниатюр — берём миниатюру точно под курсором
    if (state._lastMouseX !== undefined && state._lastMouseY !== undefined) {
      const el = document.elementFromPoint(state._lastMouseX, state._lastMouseY);
      targetThumb = el?.closest(".pdf-thumb");
    }
    // Если курсор не над миниатюрой (или прокрутка колесом/скроллбаром вне неё) — берём миниатюру по центральной ватерлинии
    if (!targetThumb) {
      const rect = els.pdfThumbs.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        const midY = rect.top + rect.height / 2;
        const midX = rect.left + rect.width / 2;
        const el = document.elementFromPoint(midX, midY);
        targetThumb = el?.closest(".pdf-thumb");
      }
    }
    if (!targetThumb) return;

    // 1. Для Excel-книг
    if (targetThumb.classList.contains("excel-book-thumb") && targetThumb.dataset.workbookIndex !== undefined) {
      const idx = Number(targetThumb.dataset.workbookIndex);
      const wb = state.excelWorkbooks?.[idx];
      if (wb?.path && state.revealedPath !== wb.path) {
        revealPathInTree(wb.path, { updateSelection: false, skipScroll: false });
      }
      return;
    }

    // 2. Для PDF-страниц
    if (!targetThumb.dataset?.pageKey) return;
    const key = targetThumb.dataset.pageKey;
    const page = state.renderedPages.find((p) => pageKey(p) === key);
    if (!page) return;
    const docPath = thumbPathForPage(page);
    if (docPath && state.revealedPath !== docPath) {
      revealPathInTree(docPath, { updateSelection: false, skipScroll: false });
    }
  });
}
if (els.pdfThumbs) {
  els.pdfThumbs.addEventListener("scroll", syncRibbonScrollToTree, { passive: true });
}

document.addEventListener("click", (event) => {
  if (els.objectListState.classList.contains("active") && !event.target.closest(".object-row, button, textarea")) {
    state.selectedObjectId = null;
    renderObjectList();
  }
  // В режиме дерева: сбрасываем выделение ТОЛЬКО если кликнули по пустому фону самого дерева
  // Клик в просмотрщик, ленту миниатюр, скроллбар или тулбар никогда не сбрасывает дерево!
  if (inTreeMode() && els.objectTree && els.objectTree === event.target) {
    state.selectedPaths.clear();
    state.revealedPath = "";
    updateTreeSelectionHighlight();
  }
});

setMode("objects");
setViewerMode("standard");
fetch('/api/version')
  .then(r => r.json())
  .then(data => {
    const el = document.getElementById('app-version-badge');
    if (el) el.textContent = 'v' + data.version;
  })
  .catch(() => {
    const el = document.getElementById('app-version-badge');
    if (el) el.textContent = 'v3.1.0';
  });

// Автообновление (Спринт 2): проверка через 3 секунды после загрузки.
function showUpdateBanner(data) {
  const badge = document.getElementById('app-version-badge');
  if (!badge || document.getElementById('update-pill')) return;
  const pill = document.createElement('button');
  pill.type = 'button';
  pill.id = 'update-pill';
  pill.title = data.changelog || 'Доступно обновление лаунчера';
  pill.textContent = 'Доступно обновление v' + data.latest_version + ' [Обновить]';
  pill.style.cssText = 'margin-left:8px;padding:2px 12px;border-radius:999px;border:1px solid #93c5fd;background:#dbeafe;color:#1d4ed8;font-size:12px;cursor:pointer;white-space:nowrap;';
  pill.addEventListener('click', () => {
    pill.disabled = true;
    pill.textContent = 'Загрузка обновления...';
    fetch('/api/updates/apply', { method: 'POST' })
      .then(r => r.json())
      .then(() => {
        showToast('Идет загрузка и установка обновления... Лаунчер перезапустится через несколько секунд.');
      })
      .catch(() => {
        showToast('Идет загрузка и установка обновления... Лаунчер перезапустится через несколько секунд.');
      });
  });
  badge.after(pill);
}

function checkUpdates() {
  fetch('/api/updates/check')
    .then(r => r.json())
    .then(data => {
      if (data && data.has_update) {
        showUpdateBanner(data);
      }
    }).catch(() => {});
}
setTimeout(checkUpdates, 3000);
loadObjectSummaries().catch((error) => {
  els.objectList.innerHTML = `<div class="empty-note">Ошибка загрузки списка: ${error.message}</div>`;
});

// Кэш настроек из /api/config/apps: действие по умолчанию,
// переключатель нативных программ и пути exe по расширениям.
async function loadAppSettings() {
  try {
    const res = await fetch("/api/config/apps");
    if (res.ok) {
      state.appSettings = await res.json();
    }
  } catch (err) {
  }
  return state.appSettings;
}

function extractAppPath(val) {
  if (!val) return "";
  if (typeof val === "object" && val !== null) {
    return String(val.path || "").trim();
  }
  return String(val || "").trim();
}

function settingsCustomExe(extension) {
  if (!extension) return "";
  const raw = String(extension).trim().toLowerCase();
  const withDot = raw.startsWith(".") ? raw : "." + raw;
  const withoutDot = raw.startsWith(".") ? raw.slice(1) : raw;
  const val =
    state.appSettings?.[withDot] ||
    state.appSettings?.[withoutDot] ||
    state.appSettings?.[withDot.toUpperCase()] ||
    state.appSettings?.[withoutDot.toUpperCase()] ||
    (withDot === ".dwg" || withDot === ".dxf" ? state.appSettings?.settingDwgExe : "") ||
    (withDot === ".pdf" ? state.appSettings?.settingPdfExe : "") ||
    ([".doc", ".docx", ".rtf", ".odt"].includes(withDot) ? state.appSettings?.settingWordExe : "") ||
    ([".xls", ".xlsx", ".xlsm", ".csv", ".ods", ".xlsb"].includes(withDot) ? state.appSettings?.settingExcelExe : "") ||
    ([".ppt", ".pptx", ".odp"].includes(withDot) ? state.appSettings?.settingPptExe : "") ||
    ([".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".ico", ".svg"].includes(withDot) ? state.appSettings?.settingImgExe : "") ||
    ([".zip", ".rar", ".7z", ".tar", ".gz"].includes(withDot) ? state.appSettings?.settingArchExe : "") ||
    ([".txt", ".log", ".ini", ".cfg", ".json", ".xml", ".yaml", ".yml"].includes(withDot) ? state.appSettings?.settingTxtExe : "") ||
    ([".mp4", ".avi", ".mov", ".mkv", ".mp3", ".wav"].includes(withDot) ? state.appSettings?.settingMediaExe : "") ||
    "";
  return extractAppPath(val);
}

const NATIVE_APP_GROUPS = [
  { id: "settingDwgExe", label: "AutoCAD / DWG (.dwg, .dxf)", exts: [".dwg", ".dxf"], ph: "Например: C:\\Program Files\\Autodesk\\AutoCAD 2024\\acad.exe" },
  { id: "settingPdfExe", label: "PDF-просмотрщик (.pdf)", exts: [".pdf"], ph: "Например: C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe" },
  { id: "settingWordExe", label: "Microsoft Word (.doc, .docx, .rtf, .odt)", exts: [".doc", ".docx", ".rtf", ".odt"], ph: "Например: C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE" },
  { id: "settingExcelExe", label: "Microsoft Excel (.xls, .xlsx, .xlsm, .csv, .ods)", exts: [".xls", ".xlsx", ".xlsm", ".csv", ".ods"], ph: "Например: C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" },
  { id: "settingPptExe", label: "Microsoft PowerPoint (.ppt, .pptx, .odp)", exts: [".ppt", ".pptx", ".odp"], ph: "Например: C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE" },
  { id: "settingImgExe", label: "Просмотр изображений (.jpg, .png, .gif и др.)", exts: [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".ico", ".svg"], ph: "Например: C:\\Windows\\System32\\mspaint.exe" },
  { id: "settingArchExe", label: "Архиватор (.zip, .rar, .7z и др.)", exts: [".zip", ".rar", ".7z", ".tar", ".gz"], ph: "Например: C:\\Program Files\\7-Zip\\7zFM.exe" },
  { id: "settingTxtExe", label: "Текстовые файлы (.txt, .log и др.)", exts: [".txt", ".log", ".ini", ".cfg", ".json", ".xml", ".yaml", ".yml"], ph: "Например: C:\\Windows\\System32\\notepad.exe" },
  { id: "settingMediaExe", label: "Видео и музыка (.mp4, .mp3 и др.)", exts: [".mp4", ".avi", ".mov", ".mkv", ".mp3", ".wav"], ph: "Например: C:\\Program Files\\VideoLAN\\VLC\\vlc.exe" },
];

function renderSettingsAppRows(cfg) {
  const form = document.getElementById("settingsAppsForm");
  if (!form) return;
  form.replaceChildren();
  for (const group of NATIVE_APP_GROUPS) {
    const current = group.exts.map((e) => extractAppPath(cfg?.[e])).find(Boolean) || "";
    const field = document.createElement("div");
    field.className = "settings-field";
    const label = document.createElement("label");
    label.className = "settings-label";
    label.setAttribute("for", group.id);
    label.textContent = group.label + ":";
    const row = document.createElement("div");
    row.className = "settings-input-row";
    const input = document.createElement("input");
    input.type = "text";
    input.id = group.id;
    input.className = "settings-input";
    input.placeholder = group.ph;
    input.spellcheck = false;
    input.value = current;
    input.dataset.exts = group.exts.join(",");
    const browse = document.createElement("button");
    browse.type = "button";
    browse.className = "settings-browse-btn";
    browse.dataset.target = group.id;
    browse.title = "Выбрать файл .exe через проводник";
    browse.textContent = "Обзор…";
    browse.addEventListener("click", () => browseExeForSetting(group.id));
    const clear = document.createElement("button");
    clear.type = "button";
    clear.className = "settings-clear-btn";
    clear.title = "Очистить";
    clear.textContent = "✕";
    clear.addEventListener("click", () => {
      input.value = "";
      updateDwgHintVisibility();
    });
    input.addEventListener("input", () => {
      if (group.id === "settingDwgExe") updateDwgHintVisibility();
    });
    row.append(input, browse, clear);
    field.append(label, row);
    form.append(field);
  }
}

// Контекстное меню файла/папки (правая кнопка мыши).
// Пункты: Скопировать путь, Открыть в проводнике, Открыть в программе.
function hideFileContextMenu() {
  if (els.contextMenu) els.contextMenu.hidden = true;
}

function showFileContextMenu(clientX, clientY, target) {
  const menu = els.contextMenu;
  if (!menu || !target?.path) return;
  const items = [];
  if (!target.isDir) {
    items.push({
      label: "Открыть в программе по умолчанию",
      run: () => openFileByPath(target.path, "native"),
    });
    items.push({ sep: true });
  }
  items.push({
    label: "Открыть в проводнике",
    run: () => openFileByPath(target.path, "explorer"),
  });
  items.push({ sep: true });
  items.push({
    label: "Скопировать путь",
    hint: "",
    run: () => copyPathToClipboard(target.path),
  });

  menu.replaceChildren();
  for (const it of items) {
    if (it.sep) {
      const d = document.createElement("div");
      d.className = "context-menu-separator";
      menu.append(d);
      continue;
    }
    const b = document.createElement("button");
    b.type = "button";
    b.className = "context-menu-item";
    b.title = it.label;
    const check = document.createElement("span");
    check.className = "context-menu-check";
    check.textContent = it.hint || "";
    const lab = document.createElement("span");
    lab.textContent = it.label;
    b.append(check, lab);
    b.addEventListener("click", (event) => {
      event.stopPropagation();
      hideFileContextMenu();
      it.run();
    });
    menu.append(b);
  }
  menu.style.left = clientX + "px";
  menu.style.top = clientY + "px";
  menu.hidden = false;
  const w = menu.offsetWidth;
  const h = menu.offsetHeight;
  menu.style.left = Math.max(4, Math.min(clientX, window.innerWidth - w - 4)) + "px";
  menu.style.top = Math.max(4, Math.min(clientY, window.innerHeight - h - 4)) + "px";
}

document.addEventListener("click", () => hideFileContextMenu());
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && els.contextMenu && !els.contextMenu.hidden) {
    event.preventDefault();
    event.stopPropagation();
    hideFileContextMenu();
  }
}, true);
window.addEventListener("resize", () => hideFileContextMenu());

function setDwgHint(message) {
  const hint = document.getElementById("settingsDwgHint");
  if (!hint) return;
  if (message) {
    hint.textContent = message;
    hint.hidden = false;
  } else {
    hint.textContent = "";
    hint.hidden = true;
  }
}

function updateDwgHintVisibility() {
  const dwgInput = document.getElementById("settingDwgExe");
  if (!dwgInput || (dwgInput.value && dwgInput.value.trim())) {
    setDwgHint("");
  }
}

// Управление модальным окном "Параметры"
async function openSettingsModal(highlightExt) {
  setDwgHint("");
  try {
    const res = await fetch("/api/config/apps");
    if (res.ok) {
      const cfg = await res.json();
      renderSettingsAppRows(cfg);
    }
  } catch (err) {
  }
  if (els.settingsModal) {
    els.settingsModal.hidden = false;
    if (highlightExt) {
      const norm = String(highlightExt).toLowerCase();
      const dot = norm.startsWith(".") ? norm : "." + norm;
      const targetInput = Array.from(document.querySelectorAll("#settingsAppsForm .settings-input")).find(
        (inp) => (inp.dataset.exts || "").split(",").map((e) => e.trim().toLowerCase()).includes(dot)
      );
      if (targetInput) {
        setTimeout(() => {
          targetInput.focus();
          targetInput.select?.();
          targetInput.classList.add("input-updated");
          setTimeout(() => targetInput.classList.remove("input-updated"), 2000);
        }, 100);
      }
    }
  }
}

function closeSettingsModal() {
  if (els.settingsModal) els.settingsModal.hidden = true;
}

function applyImportedToInputs(imported, onlyIfEmpty = false) {
  if (!imported) return 0;
  let count = 0;
  const inputs = Array.from(document.querySelectorAll("#settingsAppsForm .settings-input"));
  inputs.forEach((input) => {
    if (onlyIfEmpty && input.value && input.value.trim()) {
      return;
    }
    const exts = String(input.dataset.exts || "").split(",").map((e) => e.trim().toLowerCase()).filter(Boolean);
    for (const ext of exts) {
      const val = imported[ext];
      const p = extractAppPath(val);
      if (p) {
        input.value = p;
        input.classList.add("input-updated");
        setTimeout(() => input.classList.remove("input-updated"), 1500);
        count++;
        break;
      }
    }
  });
  return count;
}

async function parseApiResponse(res) {
  let data = null;
  try {
    data = await res.json();
  } catch (parseErr) {
    data = null;
  }
  if (!res.ok) {
    if (data && (data.message || data.error || data.error_code)) {
      const code = data.error_code ? ` [${data.error_code}]` : "";
      const msg = data.message || data.error || "Неизвестная ошибка сервера";
      throw new Error(`Ошибка сервера (${res.status})${code}: ${msg}`);
    }
    throw new Error(`Ошибка сервера HTTP ${res.status}. Проверьте состояние процесса Launcher.`);
  }
  return data || {};
}

function handleApiFetchError(err, prefix = "Не удалось определить программы Windows") {
  const isNetwork =
    err instanceof TypeError ||
    String(err.message || "").toLowerCase().includes("failed to fetch") ||
    String(err.message || "").toLowerCase().includes("networkerror") ||
    String(err.message || "").toLowerCase().includes("load failed");
  if (isNetwork) {
    showToast(
      "Нет соединения с локальным backend Launcher.\nПолностью перезапустите Launcher.\nПодробности доступны в журнале Launcher."
    );
  } else {
    showToast(`${prefix}: ${err.message || err}.\nПодробности записаны в журнал Launcher.`);
  }
}

async function fillEmptyFromWindows() {
  const btn = document.getElementById("settingsFillEmptyBtn");
  const prevText = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
    btn.textContent = "⏳ Поиск в Windows...";
  }
  try {
    const res = await fetch("/api/config/apps/import-windows", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      body: JSON.stringify({ mode: "fill_empty" })
    });
    const data = await parseApiResponse(res);
    const count = applyImportedToInputs(data.imported || {}, true);
    if (data.skipped?.[".dwg"] === "no_usable_windows_association" || data.report?.dwgHint) {
      const dwgInput = document.getElementById("settingDwgExe");
      if (!dwgInput || !dwgInput.value) {
        setDwgHint("Windows вернула системный CAD launcher, а не фактическое приложение. Выберите DWG Viewer или AutoCAD вручную.");
      }
    } else {
      updateDwgHintVisibility();
    }

    if (count > 0) {
      showToast(`Заполнено пустых полей: ${count}. Не забудьте нажать «Сохранить»!`);
    } else {
      showToast("Все поля уже заполнены или Windows не вернула подходящих ассоциаций.");
    }
  } catch (err) {
    handleApiFetchError(err, "Не удалось определить программы Windows");
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = prevText;
    }
  }
}

async function replaceFromWindows() {
  const confirmed = confirm(
    "Заменить сохранённые приложения программами по умолчанию Windows?\n\nСуществующие настройки для поддерживаемых типов файлов будут перезаписаны ассоциациями из Windows (при наличии пригодного приложения)."
  );
  if (!confirmed) {
    return;
  }

  const btn = document.getElementById("settingsReplaceBtn");
  const prevText = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
    btn.textContent = "⏳ Замена из Windows...";
  }
  try {
    const res = await fetch("/api/config/apps/import-windows", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      body: JSON.stringify({ mode: "replace_confirmed" })
    });
    const data = await parseApiResponse(res);
    const count = applyImportedToInputs(data.imported || {}, false);
    if (data.skipped?.[".dwg"] === "no_usable_windows_association" || data.report?.dwgHint) {
      setDwgHint("Windows вернула системный CAD launcher, а не фактическое приложение. Выберите DWG Viewer или AutoCAD вручную.");
    } else {
      updateDwgHintVisibility();
    }

    if (count > 0) {
      showToast(`Обновлено категорий из Windows: ${count}. Не забудьте нажать «Сохранить»!`);
    } else {
      showToast("Подходящих ассоциаций Windows не найдено.");
    }
  } catch (err) {
    handleApiFetchError(err, "Не удалось заменить программы Windows");
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = prevText;
    }
  }
}

async function openWindowsSettings() {
  const btn = document.getElementById("settingsOpenWinSettingsBtn");
  const prevText = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
  }
  try {
    const res = await fetch("/api/config/apps/open-windows-settings", {
      method: "POST",
      headers: { "Accept": "application/json" }
    });
    await parseApiResponse(res);
    showToast("Открываем параметры приложений Windows…");
  } catch (err) {
    handleApiFetchError(err, "Не удалось открыть настройки Windows");
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = prevText;
    }
  }
}

async function saveSettings() {
  const cfg = {
    useNativeApps: state.appSettings?.useNativeApps !== false,
  };
  document.querySelectorAll("#settingsAppsForm .settings-input").forEach((input) => {
    const value = (input.value || "").trim().replace(/^["']+|["']+$/g, "").trim();
    String(input.dataset.exts || "").split(",").map((e) => e.trim()).filter(Boolean).forEach((ext) => {
      const lower = ext.toLowerCase();
      const dot = lower.startsWith(".") ? lower : "." + lower;
      const noDot = dot.slice(1);
      cfg[dot] = value;
      cfg[noDot] = value;
      cfg[dot.toUpperCase()] = value;
      cfg[noDot.toUpperCase()] = value;
    });
    if (input.id) {
      cfg[input.id] = value;
    }
  });
  try {
    const res = await fetch("/api/config/apps", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(cfg),
    });
    if (!res.ok) {
      const data = await res.json();
      throw new Error(data.error || "Не удалось сохранить настройки");
    }
    closeSettingsModal();
    await loadAppSettings().catch(() => {});
    showToast("Параметры успешно сохранены");
  } catch (err) {
    showOperationError(err);
  }
}

async function browseExeForSetting(inputId) {
  const input = document.getElementById(inputId);
  if (!input) return;
  const btn = document.querySelector(`.settings-browse-btn[data-target="${inputId}"]`);
  if (btn) btn.disabled = true;
  try {
    const currentPath = (input.value || "").trim();
    const res = await fetch("/api/choose-exe", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ current: currentPath }),
    });
    if (!res.ok) {
      throw new Error(`Ошибка сервера: HTTP ${res.status}`);
    }
    const data = await res.json();
    if (data.path) {
      input.value = data.path;
      input.classList.add("input-updated");
      setTimeout(() => input.classList.remove("input-updated"), 1500);
      updateDwgHintVisibility();
    } else if (data.error && !data.cancelled && !data.timedOut) {
      showOperationError(new Error(data.error));
    }
  } catch (err) {
    showOperationError(err);
  } finally {
    if (btn) btn.disabled = false;
  }
}

if (els.btnSettings) els.btnSettings.addEventListener("click", openSettingsModal);
if (els.settingsCloseBtn) els.settingsCloseBtn.addEventListener("click", closeSettingsModal);
if (els.settingsCancelBtn) els.settingsCancelBtn.addEventListener("click", closeSettingsModal);
if (els.settingsFillEmptyBtn) els.settingsFillEmptyBtn.addEventListener("click", fillEmptyFromWindows);
if (els.settingsReplaceBtn) els.settingsReplaceBtn.addEventListener("click", replaceFromWindows);
if (els.settingsOpenWinSettingsBtn) els.settingsOpenWinSettingsBtn.addEventListener("click", openWindowsSettings);
if (els.settingsSaveBtn) els.settingsSaveBtn.addEventListener("click", saveSettings);
if (els.settingsModal) {
  els.settingsModal.addEventListener("click", (event) => {
    if (event.target === els.settingsModal) closeSettingsModal();
  });
}
document.querySelectorAll(".settings-browse-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    const targetId = btn.dataset.target;
    if (targetId) browseExeForSetting(targetId);
  });
});
document.querySelectorAll(".settings-clear-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    const targetId = btn.dataset.target;
    const input = document.getElementById(targetId);
    if (input) {
      input.value = "";
      updateDwgHintVisibility();
    }
  });
});



// Настройки подгружаем сразу, чтобы клики знали действие по умолчанию.
loadAppSettings().catch(() => {});

// Версия кода в заголовке окна: видно, какой сборкой пользуемся.
// Если заголовок без версии — окно старое, его надо закрыть и открыть заново.
try {
  const vm = String((typeof document !== "undefined" && document.currentScript && document.currentScript.src) || "").match(/[?&]v=([^&]+)/);
  if (vm) document.title = `F-Engineering Launcher v3 · ${decodeURIComponent(vm[1])}`;
} catch (e) {}
