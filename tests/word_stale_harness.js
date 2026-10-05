/* Поведенческий тест защиты от устаревшего Word-просмотра.
 * Исполняет НАСТОЯЩИЙ app/frontend/app.js в node с минимальными стабами
 * браузера и управляемым fetch. Никаких новых зависимостей.
 * Печатает JSON-вердикт; exit 0 — всё зелёное.
 * Запуск: node tests/word_stale_harness.js (из корня проекта)
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const APP = path.join(__dirname, "..", "app", "frontend", "app.js");

/* ---------- стабы браузера ---------- */
function mockEl() {
  const store = new Map();
  const fn = function () { return mockEl(); };
  return new Proxy(fn, {
    get(t, p) {
      if (p === Symbol.toPrimitive) return () => "";
      if (p === "then") return undefined;
      if (store.has(p)) return store.get(p);
      if (p === "querySelector") return () => null;
      if (p === "querySelectorAll") return () => [];
      if (p === "dataset") return {};
      if (p === "classList") return { add() {}, remove() {}, toggle() {}, contains() { return false; } };
      if (p === "style") return {};
      return (...a) => mockEl();
    },
    set(t, p, v) { store.set(p, v); return true; },
  });
}
const documentStub = {
  querySelector: () => mockEl(),
  querySelectorAll: () => [],
  getElementById: () => mockEl(),
  createElement: () => mockEl(),
  createDocumentFragment: () => mockEl(),
  addEventListener() {},
  body: mockEl(),
  activeElement: null,
};

/* ---------- управляемый fetch ---------- */
const fetchCalls = [];
const gates = new Map();
const autoRoutes = new Map();
function fkey(endpoint, body) {
  let f = "";
  try {
    const b = JSON.parse(body || "{}");
    f = b.file || (b.files || []).join("|");
  } catch (_) {}
  return endpoint + "|" + f;
}
function stubFetch(url, opts) {
  const o = opts || {};
  const body = typeof o.body === "string" ? o.body : "";
  const key = fkey(url, body);
  fetchCalls.push({ url, body });
  if (autoRoutes.has(key)) {
    const r = autoRoutes.get(key);
    if (r && r.ok === false) {
      return Promise.resolve({ ok: false, status: 400, json: async () => ({ error: r.error || "err" }) });
    }
    return Promise.resolve({ ok: true, status: 200, json: async () => r });
  }
  return new Promise((resolve) => {
    if (!gates.has(key)) gates.set(key, []);
    gates.get(key).push(resolve);
  });
}
function respond(endpoint, file, payload) {
  const key = endpoint + "|" + file;
  const list = gates.get(key) || [];
  gates.delete(key);
  for (const resolve of list) {
    resolve({ ok: true, status: 200, json: async () => payload });
  }
  return list.length;
}
function respondError(endpoint, file, message) {
  const key = endpoint + "|" + file;
  const list = gates.get(key) || [];
  gates.delete(key);
  for (const resolve of list) {
    resolve({ ok: false, status: 400, json: async () => ({ error: message }) });
  }
  return list.length;
}
function wordRenderPayload(file, urlPrefix) {
  return {
    dpi: 150,
    documents: [{
      sourcePath: file, sourceName: file.split(/[\\/]/).pop(), path: file + ".pdf",
      sourceType: "DOCX", pages: 2, renderedPages: 2, cacheHit: false,
      items: [
        { page: 1, url: urlPrefix + "1.png", bytes: 10 },
        { page: 2, url: urlPrefix + "2.png", bytes: 10 },
      ],
    }],
    totalPages: 2, renderedPages: 2, errors: [],
  };
}
function pdfRenderPayload(file, urlPrefix) {
  return {
    dpi: 150,
    documents: [{
      sourcePath: file, path: file, sourceType: "PDF",
      pages: 1, renderedPages: 1, cacheHit: false,
      items: [{ page: 1, url: urlPrefix + "1.png", bytes: 10 }],
    }],
    totalPages: 1, renderedPages: 1, errors: [],
  };
}

const DRIVER = `
;(async () => {
  const tick = async (n) => { for (let i = 0; i < (n || 8); i++) await new Promise((r) => setTimeout(r, 0)); };
  const norm = (s) => String(s || "").replace(/\\//g, "\\\\").toLowerCase();
  const out = {};
  const shot = () => ({
    src: els.pdfPageImage.src,
    detail: els.progressDetail.textContent,
    pages: (state.renderedPages || []).map((p) => p.sourcePath || p.documentPath || ""),
    wordFetches: fetchCalls.filter((c) => c.url === "/api/word/preview").length,
    renderFetches: (f) => fetchCalls.filter((c) => c.url === "/api/word/render" && (c.body || "").includes(f)).length,
  });

  // T1: Word A задержан, пользователь выбирает Word B.
  state.selectedPaths = new Set(["C:\\\\t\\\\A.docx"]);
  const pA = showWordPreviewPaginated({ type: "file", name: "A.docx", path: "C:\\\\t\\\\A.docx" }, {});
  await tick();
  state.selectedPaths = new Set(["C:\\\\t\\\\B.docx"]);
  const pB = showWordPreviewPaginated({ type: "file", name: "B.docx", path: "C:\\\\t\\\\B.docx" }, {});
  await tick();
  respond("/api/word/render", "C:\\\\t\\\\B.docx", wordRenderPayload("C:\\\\t\\\\B.docx", "/B"));
  await pB; await tick(12);
  const afterB = shot();
  respond("/api/word/render", "C:\\\\t\\\\A.docx", wordRenderPayload("C:\\\\t\\\\A.docx", "/A"));
  await pA; await tick(12);
  const afterLateA = shot();
  out.t1_shownB = String(afterB.src).endsWith("/B1.png");
  out.t1_stillB = afterLateA.src === afterB.src;
  out.t1_noApages = afterLateA.pages.every((p) => norm(p) !== norm("C:\\\\t\\\\A.docx"));
  out.t1_noHtmlFallback = afterLateA.wordFetches === 0;
  out.t1_detailStable = afterLateA.detail === afterB.detail;

  // T2: Word A2 задержан, пользователь выбирает PDF.
  state.selectedPaths = new Set(["C:\\\\t\\\\A2.docx"]);
  const pA2 = showWordPreviewPaginated({ type: "file", name: "A2.docx", path: "C:\\\\t\\\\A2.docx" }, {});
  await tick();
  state.selectedPaths = new Set(["C:\\\\t\\\\C.pdf"]);
  const pC = previewFileDirectly({ type: "file", name: "C.pdf", path: "C:\\\\t\\\\C.pdf", extension: "PDF" }, {});
  await tick();
  respond("/api/pdf/render", "C:\\\\t\\\\C.pdf", pdfRenderPayload("C:\\\\t\\\\C.pdf", "/C"));
  await pC; await tick(12);
  const afterC = shot();
  respond("/api/word/render", "C:\\\\t\\\\A2.docx", wordRenderPayload("C:\\\\t\\\\A2.docx", "/A2"));
  await tick(12);
  const afterLateA2 = shot();
  out.t2_pdfShown = String(afterC.src).endsWith("/C1.png");
  out.t2_stillPdf = afterLateA2.src === afterC.src;
  out.t2_noA2pages = afterLateA2.pages.every((p) => norm(p) !== norm("C:\\\\t\\\\A2.docx"));
  try { await pA2; } catch (_) {}

  // T3: Word A3 задержан, пользователь выбирает изображение.
  state.selectedPaths = new Set(["C:\\\\t\\\\A3.docx"]);
  const pA3 = showWordPreviewPaginated({ type: "file", name: "A3.docx", path: "C:\\\\t\\\\A3.docx" }, {});
  await tick();
  state.selectedPaths = new Set(["C:\\\\t\\\\I.png"]);
  const pI = previewFileDirectly({ type: "file", name: "I.png", path: "C:\\\\t\\\\I.png", extension: "PNG" }, {});
  await pI; await tick(12);
  const afterI = shot();
  respond("/api/word/render", "C:\\\\t\\\\A3.docx", wordRenderPayload("C:\\\\t\\\\A3.docx", "/A3"));
  await tick(12);
  const afterLateA3 = shot();
  out.t3_imgShown = String(afterI.src).includes("I.png");
  out.t3_stillImg = afterLateA3.src === afterI.src;
  out.t3_noA3pages = afterLateA3.pages.every((p) => norm(p) !== norm("C:\\\\t\\\\A3.docx"));
  try { await pA3; } catch (_) {}

  // T4: COM упал -> ровно один HTML-фолбэк, дальше тишина (без зацикливания).
  const beforeHtml = shot().wordFetches;
  state.selectedPaths = new Set(["C:\\\\t\\\\F.docx"]);
  const pF = showWordPreviewPaginated({ type: "file", name: "F.docx", path: "C:\\\\t\\\\F.docx" }, {});
  await tick();
  respondError("/api/word/render", "C:\\\\t\\\\F.docx", "Word COM broken");
  autoRoutes.set("/api/word/preview|C:\\\\t\\\\F.docx", { url: "/html/F", paragraphs: 1, tables: 0, bytes: 10 });
  await pF; await tick(12);
  const rendersF = shot().renderFetches("F.docx");
  out.t4_oneHtmlFallback = (shot().wordFetches - beforeHtml) === 1;
  out.t4_noRenderLoop = rendersF === 1;
  out.t4_frameShown = String(els.wordDocFrame.src).includes("/html/F");

  // A: Word A задержан через обычную точку входа, выбран B, ошибка A приходит поздно.
  state.selectedPaths = new Set(["C:\\\\t\\\\G.docx"]);
  const pG = previewFileDirectly({ type: "file", name: "G.docx", path: "C:\\\\t\\\\G.docx", extension: "DOCX" }, {});
  await tick();
  state.selectedPaths = new Set(["C:\\\\t\\\\H.docx"]);
  const pH = previewFileDirectly({ type: "file", name: "H.docx", path: "C:\\\\t\\\\H.docx", extension: "DOCX" }, {});
  await tick();
  respond("/api/word/render", "C:\\\\t\\\\H.docx", wordRenderPayload("C:\\\\t\\\\H.docx", "/H"));
  await pH; await tick(12);
  const afterH = shot();
  respondError("/api/word/render", "C:\\\\t\\\\G.docx", "Word COM broken late");
  await tick(12);
  try { await pG; } catch (_) {}
  const afterLateGerr = shot();
  out.a_stillB = afterLateGerr.src === afterH.src;
  out.a_noGpages = afterLateGerr.pages.every((p) => norm(p) !== norm("C:\\\\t\\\\G.docx"));
  out.a_noHtmlFallback = afterLateGerr.wordFetches === afterH.wordFetches;
  out.a_noHtmlForG = !fetchCalls.some((c) => c.url === "/api/word/preview"
    && (c.body || "").includes("G.docx"));
  out.a_detailStable = afterLateGerr.detail === afterH.detail;

  // B: PDF-путь A упал -> начался HTML-запрос A -> выбран B -> поздний HTML-ответ A.
  state.selectedPaths = new Set(["C:\\\\t\\\\J.docx"]);
  const pJ = previewFileDirectly({ type: "file", name: "J.docx", path: "C:\\\\t\\\\J.docx", extension: "DOCX" }, {});
  await tick();
  respondError("/api/word/render", "C:\\\\t\\\\J.docx", "Word COM broken early");
  await tick(12);
  out.b_htmlFiredBeforeSwitch = fetchCalls.some((c) => c.url === "/api/word/preview"
    && (c.body || "").includes("J.docx"));
  state.selectedPaths = new Set(["C:\\\\t\\\\K.docx"]);
  const pK = previewFileDirectly({ type: "file", name: "K.docx", path: "C:\\\\t\\\\K.docx", extension: "DOCX" }, {});
  await tick();
  respond("/api/word/render", "C:\\\\t\\\\K.docx", wordRenderPayload("C:\\\\t\\\\K.docx", "/K"));
  await pK; await tick(12);
  const afterK = shot();
  const frameBeforeLateJ = String(els.wordDocFrame.src);
  respond("/api/word/preview", "C:\\\\t\\\\J.docx", { url: "/html/J", paragraphs: 1, tables: 0, bytes: 10 });
  await tick(12);
  const afterLateJhtml = shot();
  out.b_stillB = afterLateJhtml.src === afterK.src;
  out.b_frameStable = String(els.wordDocFrame.src) === frameBeforeLateJ;
  out.b_detailStable = afterLateJhtml.detail === afterK.detail;
  out.b_noJpages = afterLateJhtml.pages.every((p) => norm(p) !== norm("C:\\\\t\\\\J.docx"));
  try { await pJ; } catch (_) {}

  // C: оба маршрута Word упали — ограниченное число запросов, без зацикливания.
  state.selectedPaths = new Set(["C:\\\\t\\\\Q.docx"]);
  const qBefore = shot();
  const pQ = previewFileDirectly({ type: "file", name: "Q.docx", path: "C:\\\\t\\\\Q.docx", extension: "DOCX" }, {});
  await tick();
  respondError("/api/word/render", "C:\\\\t\\\\Q.docx", "Word COM broken Q");
  autoRoutes.set("/api/word/preview|C:\\\\t\\\\Q.docx", { ok: false, error: "HTML broken Q" });
  for (let i = 0; i < 40 && fetchCalls.filter((c) => c.url === "/api/word/render" && (c.body || "").includes("Q.docx")).length < 2; i++) await tick();
  respondError("/api/word/render", "C:\\\\t\\\\Q.docx", "Word COM broken Q again");
  await pQ; await tick(12);
  await tick(12);
  const qAfter = shot();
  const rendersQ = qAfter.renderFetches("Q.docx");
  out.c_renderBounded = rendersQ === 2;
  out.c_oneHtmlAttempt = (qAfter.wordFetches - qBefore.wordFetches) === 1;
  await tick(12);
  const qLater = shot();
  out.c_quiescent = (qLater.wordFetches === qAfter.wordFetches)
    && (qLater.renderFetches("Q.docx") === rendersQ);

  return out;
})();
`;

async function main() {
  const code = fs.readFileSync(APP, "utf8");
  const sandbox = {
    console, setTimeout, clearTimeout, setInterval, clearInterval,
    performance, AbortController, fetch: stubFetch,
    document: documentStub,
    localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    requestAnimationFrame: (cb) => setTimeout(cb, 0),
    CSS: { escape: (s) => s },
    window: { addEventListener() {}, setTimeout, innerWidth: 1280, innerHeight: 768, isSecureContext: false },
    navigator: {},
    confirm: () => false,
    IntersectionObserver: class { constructor() {} observe() {} disconnect() {} },
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  // Пробрасываем respond/autoRoutes внутрь контекста.
  sandbox.respond = respond;
  sandbox.respondError = respondError;
  sandbox.autoRoutes = autoRoutes;
  sandbox.fetchCalls = fetchCalls;
  sandbox.wordRenderPayload = wordRenderPayload;
  sandbox.pdfRenderPayload = pdfRenderPayload;
  let verdict;
  try {
    verdict = await vm.runInContext(code + "\n" + DRIVER, sandbox, { filename: "app-under-test.js" });
  } catch (e) {
    console.log(JSON.stringify({ harness_error: String((e && e.stack) || e) }));
    process.exit(2);
  }
  const results = [];
  for (const k of Object.keys(verdict)) {
    results.push({ name: k, pass: verdict[k] === true, actual: String(verdict[k]) });
  }
  const failed = results.filter((r) => !r.pass);
  console.log(JSON.stringify({ results, failed: failed.length }));
  process.exit(failed.length ? 1 : 0);
}

main().catch((e) => { console.log(JSON.stringify({ harness_error: String(e) })); process.exit(2); });
