"""
Headless Benchmark Runner for PDF-VIEWER-COMPARE-001.
Simulates and executes end-to-end performance benchmarks for:
- Mode A: PyMuPDF 150 DPI PNG rasterization (Cold render & Warm persistent disk cache hit)
- Mode B: PDF.js engine simulation using local PDF.js in Node.js (V8) environment:
          fetching, parsing xref, font building, operator list construction.
- Mode C: Hybrid model (Persistent disk image instant paint + progressive vector stage).

Outputs structured statistical metrics across 5 iterations per sample file.
"""

from __future__ import annotations

import json
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
STAND_DIR = PROJECT_ROOT / "stand"
SAMPLES_DIR = STAND_DIR / "samples"
VENDOR_DIR = STAND_DIR / "vendor"

sys.path.insert(0, str(PROJECT_ROOT))
import app.backend.server as app_server

NODE_BENCH_SCRIPT = STAND_DIR / "node_pdfjs_bench.js"

# Create helper Node script to execute local Mozilla PDF.js v3.11.174
NODE_SCRIPT_CONTENT = """
const fs = require('fs');
const path = require('path');

const pdfjs = require('./vendor/pdfjs/pdf.min.js');

async function benchmarkPdfjs(filePath, iterations) {
  const data = new Uint8Array(fs.readFileSync(filePath));
  const results = [];

  for (let i = 0; i < iterations; i++) {
    const t0 = performance.now();
    // Cold load (fresh document proxy)
    const doc = await pdfjs.getDocument({
      data: data.slice(0),
      useWorkerFetch: false,
      isEvalSupported: false,
      useSystemFonts: true
    }).promise;
    const tDocLoaded = performance.now();
    const page = await doc.getPage(1);
    const tPageLoaded = performance.now();
    const ops = await page.getOperatorList();
    const tOpsReady = performance.now();

    results.push({
      run: i + 1,
      doc_parse_ms: tDocLoaded - t0,
      page_load_ms: tPageLoaded - tDocLoaded,
      operator_list_ms: tOpsReady - tPageLoaded,
      total_pipeline_ms: tOpsReady - t0
    });
    await doc.destroy();
  }

  // Warm run on already loaded doc
  const docWarm = await pdfjs.getDocument({
    data: data.slice(0),
    useWorkerFetch: false,
    isEvalSupported: false,
    useSystemFonts: true
  }).promise;
  const tWarmStart = performance.now();
  const pageWarm = await docWarm.getPage(1);
  const opsWarm = await pageWarm.getOperatorList();
  const warmTotal = performance.now() - tWarmStart;
  await docWarm.destroy();

  console.log(JSON.stringify({ runs: results, warm_doc_page_ms: warmTotal }));
}

const targetFile = process.argv[2];
const iters = parseInt(process.argv[3] || '5', 10);
benchmarkPdfjs(targetFile, iters).catch(err => {
  console.error(err);
  process.exit(1);
});
"""

def setup_node_bench():
    with open(NODE_BENCH_SCRIPT, "w", encoding="utf-8") as f:
        f.write(NODE_SCRIPT_CONTENT)

def run_pdfjs_benchmark(file_path: Path, iterations: int = 5) -> dict:
    cmd = ["node", str(NODE_BENCH_SCRIPT), str(file_path), str(iterations)]
    proc = subprocess.run(cmd, cwd=str(STAND_DIR), capture_output=True, text=True, check=True)
    return json.loads(proc.stdout.strip())

def run_pymupdf_benchmark(file_path: Path, iterations: int = 5) -> dict:
    temp_dir = tempfile.mkdtemp(prefix="bench_cache_")
    cache_dir = Path(temp_dir)
    orig_cache = app_server.PDF_CACHE_DIR
    app_server.PDF_CACHE_DIR = cache_dir

    try:
        cold_runs = []
        warm_runs = []

        # Iteration loop
        for i in range(iterations):
            # 1. Measure Cold: clean cache first
            if cache_dir.exists():
                shutil.rmtree(cache_dir, ignore_errors=True)
            cache_dir.mkdir(parents=True, exist_ok=True)

            t0 = time.perf_counter()
            meta_cold = app_server.render_pdf(file_path, dpi=150, first_page_only=True)
            key = app_server.pdf_cache_key(file_path, 150)
            page_png = cache_dir / key / "page-1.png"
            png_bytes = page_png.read_bytes() if page_png.exists() else b""
            t_cold = (time.perf_counter() - t0) * 1000.0

            cold_runs.append({
                "run": i + 1,
                "total_ms": t_cold,
                "render_ms": meta_cold.get("render_time_ms", 0),
                "cached": meta_cold.get("cacheHit", False),
                "bytes_len": len(png_bytes)
            })

            # 2. Measure Warm (Persistent disk cache hit)
            t0_warm = time.perf_counter()
            meta_warm = app_server.render_pdf(file_path, dpi=150, first_page_only=True)
            page_png_warm = cache_dir / key / "page-1.png"
            png_bytes_warm = page_png_warm.read_bytes() if page_png_warm.exists() else b""
            t_warm = (time.perf_counter() - t0_warm) * 1000.0

            warm_runs.append({
                "run": i + 1,
                "total_ms": t_warm,
                "cached": meta_warm.get("cacheHit", False)
            })

        return {
            "cold_runs": cold_runs,
            "warm_runs": warm_runs
        }
    finally:
        app_server.PDF_CACHE_DIR = orig_cache
        shutil.rmtree(cache_dir, ignore_errors=True)

def main():
    print("=" * 70)
    print("RUNNING BENCHMARK: PDF-VIEWER-COMPARE-001")
    print("=" * 70)

    setup_node_bench()

    test_files = sorted(list(SAMPLES_DIR.glob("*.pdf")))
    eng_pdf = PROJECT_ROOT / "КР-724-Р-КМ1.1.pdf"
    if eng_pdf.exists():
        test_files.append(eng_pdf)

    all_metrics = {}

    for pdf_file in test_files:
        print(f"\nBenchmarking {pdf_file.name} ({round(pdf_file.stat().st_size / 1024, 1)} KB)...")

        # 1. Mode A: PyMuPDF PNG
        png_data = run_pymupdf_benchmark(pdf_file, iterations=5)
        cold_totals = [r["total_ms"] for r in png_data["cold_runs"]]
        warm_totals = [r["total_ms"] for r in png_data["warm_runs"]]

        png_cold_mean = statistics.mean(cold_totals)
        png_cold_stdev = statistics.stdev(cold_totals) if len(cold_totals) > 1 else 0.0
        png_warm_mean = statistics.mean(warm_totals)
        png_warm_stdev = statistics.stdev(warm_totals) if len(warm_totals) > 1 else 0.0

        print(f"  [Mode A: PNG (PyMuPDF)]")
        print(f"    Cold 1st page paint: {png_cold_mean:.2f} ms (±{png_cold_stdev:.2f} ms)")
        print(f"    Warm persistent hit: {png_warm_mean:.2f} ms (±{png_warm_stdev:.2f} ms)")

        # 2. Mode B: PDF.js
        pdfjs_data = run_pdfjs_benchmark(pdf_file, iterations=5)
        pdfjs_cold_totals = [r["total_pipeline_ms"] for r in pdfjs_data["runs"]]
        pdfjs_cold_mean = statistics.mean(pdfjs_cold_totals)
        pdfjs_cold_stdev = statistics.stdev(pdfjs_cold_totals) if len(pdfjs_cold_totals) > 1 else 0.0
        pdfjs_warm_mean = pdfjs_data["warm_doc_page_ms"]

        print(f"  [Mode B: PDF.js Canvas Engine]")
        print(f"    Cold 1st page parse+ops: {pdfjs_cold_mean:.2f} ms (±{pdfjs_cold_stdev:.2f} ms)")
        print(f"    Warm RAM doc cached:     {pdfjs_warm_mean:.2f} ms")

        # 3. Mode C: Hybrid
        hybrid_first_paint = png_warm_mean
        print(f"  [Mode C: Hybrid Architecture]")
        print(f"    First visible frame (PNG placeholder): {hybrid_first_paint:.2f} ms")
        print(f"    Seamless canvas vector swap:           {pdfjs_cold_mean:.2f} ms")

        all_metrics[pdf_file.name] = {
            "size_kb": round(pdf_file.stat().st_size / 1024, 1),
            "png": {
                "cold_mean": round(png_cold_mean, 2),
                "cold_stdev": round(png_cold_stdev, 2),
                "warm_mean": round(png_warm_mean, 2),
                "warm_stdev": round(png_warm_stdev, 2)
            },
            "pdfjs": {
                "cold_mean": round(pdfjs_cold_mean, 2),
                "cold_stdev": round(pdfjs_cold_stdev, 2),
                "warm_mean": round(pdfjs_warm_mean, 2)
            },
            "hybrid": {
                "first_paint_mean": round(hybrid_first_paint, 2),
                "vector_settle_mean": round(pdfjs_cold_mean, 2)
            }
        }

    # Save benchmark data to json
    results_path = STAND_DIR / "benchmark_summary.json"
    with open(results_path, "w", encoding="utf-8") as f:
        json.dump(all_metrics, f, ensure_ascii=False, indent=2)
    print(f"\nAll benchmark results saved to: {results_path}")

if __name__ == "__main__":
    main()
