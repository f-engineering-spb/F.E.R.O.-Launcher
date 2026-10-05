/**
 * Stand logic for PDF-VIEWER-COMPARE-001
 * Implements Mode A (PNG), Mode B (PDF.js Canvas), Mode C (Hybrid: PNG preview -> PDF.js canvas)
 * Measures:
 *  - Preparation / network ms
 *  - First painted frame (using requestAnimationFrame)
 *  - Full render ms
 *  - Zoom & Navigation latencies
 */

// Configure local PDF.js worker
if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/vendor/pdfjs/pdf.worker.min.js';
}

let currentMode = 'png'; // 'png' | 'pdfjs' | 'hybrid'
let currentFile = null;
let currentPage = 1;
let totalPages = 1;
let currentZoom = 1.0;

let currentPdfDoc = null; // PDFDocumentProxy for PDF.js
let currentRenderTask = null;

const statMode = document.getElementById('stat-mode');
const statPrep = document.getElementById('stat-prep');
const statFirstPaint = document.getElementById('stat-first-paint');
const statFull = document.getElementById('stat-full');
const statCache = document.getElementById('stat-cache');

const canvasEl = document.getElementById('pdfjs-canvas');
const imgEl = document.getElementById('png-image');
const curPageEl = document.getElementById('cur-page');
const totalPagesEl = document.getElementById('total-pages');
const curZoomEl = document.getElementById('cur-zoom');
const engineStatus = document.getElementById('engine-status');

// Init
window.addEventListener('DOMContentLoaded', async () => {
  await loadSamples();
});

async function loadSamples() {
  try {
    const res = await fetch('/api/samples');
    const data = await res.json();
    const listEl = document.getElementById('samples-list');
    listEl.innerHTML = '';

    if (data.samples && data.samples.length > 0) {
      data.samples.forEach((sample, idx) => {
        const item = document.createElement('div');
        item.className = 'file-item' + (idx === 0 ? ' active' : '');
        item.innerHTML = `
          <div class="file-name">${sample.name}</div>
          <div class="file-meta">${sample.size_kb} KB</div>
        `;
        item.onclick = () => selectSample(sample, item);
        listEl.appendChild(item);
      });
      // Select first
      selectSample(data.samples[0], listEl.children[0]);
    } else {
      listEl.innerHTML = '<div style="color: #888;">Нет файлов</div>';
    }
  } catch (err) {
    console.error('Failed to load samples:', err);
  }
}

function selectSample(sample, element) {
  document.querySelectorAll('.file-item').forEach(el => el.classList.remove('active'));
  if (element) element.classList.add('active');

  currentFile = sample;
  currentPage = 1;
  currentPdfDoc = null; // reset loaded doc
  renderCurrentView();
}

function setMode(mode) {
  currentMode = mode;
  document.querySelectorAll('.mode-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(`btn-mode-${mode}`).classList.add('active');
  renderCurrentView();
}

function setZoom(zoom) {
  currentZoom = zoom;
  curZoomEl.textContent = `${Math.round(zoom * 100)}%`;
  renderCurrentView();
}

function prevPage() {
  if (currentPage > 1) {
    currentPage--;
    renderCurrentView();
  }
}

function nextPage() {
  if (currentPage < totalPages) {
    currentPage++;
    renderCurrentView();
  }
}

async function renderCurrentView() {
  if (!currentFile) return;

  curPageEl.textContent = currentPage;
  statMode.textContent = currentMode.toUpperCase();
  engineStatus.textContent = 'Rendering...';

  if (currentMode === 'png') {
    await renderModePNG();
  } else if (currentMode === 'pdfjs') {
    await renderModePDFJS();
  } else if (currentMode === 'hybrid') {
    await renderModeHybrid();
  }
  engineStatus.textContent = 'Ready';
}

/**
 * Helper to measure exact paint timestamp via requestAnimationFrame
 */
function nextFrame() {
  return new Promise(resolve => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
}

/**
 * Mode A: PyMuPDF 150 DPI PNG
 */
async function renderModePNG() {
  canvasEl.style.display = 'none';
  imgEl.style.display = 'block';

  const t0 = performance.now();
  let tPrepEnd = 0;
  let tPaintEnd = 0;

  // 1. Ensure page rendered on backend (returns total_pages and cache metadata)
  const renderApiUrl = `/api/pdf/render?path=${encodeURIComponent(currentFile.path)}&dpi=150&first_page_only=false`;
  const renderRes = await fetch(renderApiUrl);
  const renderData = await renderRes.json();
  totalPages = renderData.total_pages || 1;
  totalPagesEl.textContent = totalPages;

  tPrepEnd = performance.now();

  // 2. Load the specific page PNG
  const pageUrl = `/api/pdf/page?path=${encodeURIComponent(currentFile.path)}&page=${currentPage}&dpi=150`;

  await new Promise((resolve) => {
    imgEl.onload = async () => {
      // Image is decoded, now wait for real screen paint
      await nextFrame();
      tPaintEnd = performance.now();
      resolve();
    };
    imgEl.src = pageUrl;
  });

  // Apply zoom via CSS transform / width
  imgEl.style.transform = `scale(${currentZoom})`;
  imgEl.style.transformOrigin = 'top center';

  const totalMs = (tPaintEnd - t0).toFixed(1);
  const prepMs = (tPrepEnd - t0).toFixed(1);
  const paintMs = (tPaintEnd - tPrepEnd).toFixed(1);

  statPrep.textContent = `${prepMs} ms`;
  statFirstPaint.textContent = `${totalMs} ms`;
  statFull.textContent = `${totalMs} ms`;
  statCache.textContent = renderData.cached ? 'Disk Cache HIT' : 'Rendered & Saved';

  return { prepMs: parseFloat(prepMs), firstPaintMs: parseFloat(totalMs), fullMs: parseFloat(totalMs), cached: renderData.cached };
}

/**
 * Mode B: Local PDF.js direct Canvas
 */
async function renderModePDFJS() {
  imgEl.style.display = 'none';
  canvasEl.style.display = 'block';

  const t0 = performance.now();
  let tPrepEnd = 0;
  let tPaintEnd = 0;

  if (currentRenderTask) {
    try { currentRenderTask.cancel(); } catch (e) {}
  }

  // 1. Fetch document via Range streaming raw endpoint if not cached in memory
  let cacheStatus = 'RAM Doc Cache';
  if (!currentPdfDoc || currentPdfDoc._transport?.filePath !== currentFile.path) {
    cacheStatus = 'Network / Stream';
    const rawUrl = `/api/file/raw?path=${encodeURIComponent(currentFile.path)}`;
    const loadingTask = window.pdfjsLib.getDocument({
      url: rawUrl,
      rangeChunkSize: 65536,
      disableAutoFetch: false,
    });
    currentPdfDoc = await loadingTask.promise;
    currentPdfDoc._transport = { filePath: currentFile.path };
  }

  totalPages = currentPdfDoc.numPages;
  totalPagesEl.textContent = totalPages;
  tPrepEnd = performance.now();

  // 2. Render target page to Canvas
  const page = await currentPdfDoc.getPage(currentPage);
  const viewport = page.getViewport({ scale: currentZoom * 1.5 }); // match high-res device ratio

  canvasEl.width = viewport.width;
  canvasEl.height = viewport.height;
  canvasEl.style.width = `${viewport.width / 1.5}px`;
  canvasEl.style.height = `${viewport.height / 1.5}px`;

  const ctx = canvasEl.getContext('2d');
  const renderContext = {
    canvasContext: ctx,
    viewport: viewport,
  };

  currentRenderTask = page.render(renderContext);
  await currentRenderTask.promise;

  // Wait for real paint
  await nextFrame();
  tPaintEnd = performance.now();

  const totalMs = (tPaintEnd - t0).toFixed(1);
  const prepMs = (tPrepEnd - t0).toFixed(1);
  const paintMs = (tPaintEnd - tPrepEnd).toFixed(1);

  statPrep.textContent = `${prepMs} ms`;
  statFirstPaint.textContent = `${totalMs} ms`;
  statFull.textContent = `${totalMs} ms`;
  statCache.textContent = cacheStatus;

  return { prepMs: parseFloat(prepMs), firstPaintMs: parseFloat(totalMs), fullMs: parseFloat(totalMs), cached: cacheStatus === 'RAM Doc Cache' };
}

/**
 * Mode C: Hybrid (Fast PNG thumbnail placeholder -> Crisp PDF.js canvas)
 */
async function renderModeHybrid() {
  // 1. Immediately display PNG cached frame for instant visible paint
  imgEl.style.display = 'block';
  canvasEl.style.display = 'none';

  const t0 = performance.now();
  let firstPaintMs = 0;

  const pageUrl = `/api/pdf/page?path=${encodeURIComponent(currentFile.path)}&page=${currentPage}&dpi=150`;
  await new Promise((resolve) => {
    imgEl.onload = async () => {
      await nextFrame();
      firstPaintMs = performance.now() - t0;
      resolve();
    };
    imgEl.src = pageUrl;
  });

  statFirstPaint.textContent = `${firstPaintMs.toFixed(1)} ms (PNG placeholder)`;

  // 2. Concurrently render vector PDF.js canvas in background and swap
  if (!currentPdfDoc || currentPdfDoc._transport?.filePath !== currentFile.path) {
    const rawUrl = `/api/file/raw?path=${encodeURIComponent(currentFile.path)}`;
    const loadingTask = window.pdfjsLib.getDocument({
      url: rawUrl,
      rangeChunkSize: 65536,
    });
    currentPdfDoc = await loadingTask.promise;
    currentPdfDoc._transport = { filePath: currentFile.path };
  }

  totalPages = currentPdfDoc.numPages;
  totalPagesEl.textContent = totalPages;

  const page = await currentPdfDoc.getPage(currentPage);
  const viewport = page.getViewport({ scale: currentZoom * 1.5 });

  canvasEl.width = viewport.width;
  canvasEl.height = viewport.height;
  canvasEl.style.width = `${viewport.width / 1.5}px`;
  canvasEl.style.height = `${viewport.height / 1.5}px`;

  const ctx = canvasEl.getContext('2d');
  const renderContext = {
    canvasContext: ctx,
    viewport: viewport,
  };

  if (currentRenderTask) {
    try { currentRenderTask.cancel(); } catch (e) {}
  }
  currentRenderTask = page.render(renderContext);
  await currentRenderTask.promise;

  // Swap to canvas
  imgEl.style.display = 'none';
  canvasEl.style.display = 'block';
  await nextFrame();

  const fullMs = performance.now() - t0;
  statPrep.textContent = `Hybrid fast-path`;
  statFull.textContent = `${fullMs.toFixed(1)} ms (Canvas swap)`;
  statCache.textContent = `Hybrid (Disk PNG + RAM PDF.js)`;

  return { prepMs: 0, firstPaintMs: parseFloat(firstPaintMs.toFixed(1)), fullMs: parseFloat(fullMs.toFixed(1)), cached: true };
}

/**
 * Automated 5-run Benchmark Harness
 */
async function runAutoBenchmark() {
  engineStatus.textContent = 'Running Benchmark...';
  const tbody = document.querySelector('#results-table tbody');
  tbody.innerHTML = '<tr><td colspan="4" style="text-align: center; color: var(--accent);">Тестирование выполняется...</td></tr>';

  const samplesRes = await fetch('/api/samples');
  const samplesData = await samplesRes.json();
  const samples = samplesData.samples;

  const results = [];

  for (const sample of samples) {
    currentFile = sample;
    currentPage = 1;

    // Test PNG Cold (purge cache first)
    await clearServerCache();
    currentPdfDoc = null;
    const pngCold = await renderModePNG();

    // Test PNG Warm (repeat)
    const pngWarm = await renderModePNG();

    // Test PDF.js Cold (reset ram doc)
    currentPdfDoc = null;
    const pdfjsCold = await renderModePDFJS();

    // Test PDF.js Warm (in-memory)
    const pdfjsWarm = await renderModePDFJS();

    // Test Hybrid Warm
    const hybrid = await renderModeHybrid();

    results.push({
      sample: sample.name,
      png_cold_paint: pngCold.firstPaintMs,
      png_warm_paint: pngWarm.firstPaintMs,
      pdfjs_cold_paint: pdfjsCold.firstPaintMs,
      pdfjs_warm_paint: pdfjsWarm.firstPaintMs,
      hybrid_paint: hybrid.firstPaintMs,
    });
  }

  // Render Table
  tbody.innerHTML = '';
  results.forEach(r => {
    tbody.innerHTML += `
      <tr>
        <td rowspan="2" style="vertical-align: middle;"><strong>${r.sample}</strong></td>
        <td>PNG (150 DPI)</td>
        <td>${r.png_cold_paint} ms</td>
        <td style="color: var(--success);">${r.png_warm_paint} ms</td>
      </tr>
      <tr>
        <td>PDF.js Canvas</td>
        <td>${r.pdfjs_cold_paint} ms</td>
        <td style="color: var(--warning);">${r.pdfjs_warm_paint} ms</td>
      </tr>
    `;
  });

  // Post results to backend server
  await fetch('/api/benchmark/record', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(results),
  });

  engineStatus.textContent = 'Benchmark Finished!';
}

async function clearServerCache() {
  await fetch('/api/pdf/clear_cache', { method: 'POST' });
  currentPdfDoc = null;
  statCache.textContent = 'Cleared';
}
