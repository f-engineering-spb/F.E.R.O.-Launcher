
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
