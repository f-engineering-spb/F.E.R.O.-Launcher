"""
Isolated Stand Server for PDF-VIEWER-COMPARE-001.
Runs on port 8999 (default).
Provides:
1. Static files from stand/ directory (vendor/pdfjs, samples, etc.).
2. HTTP 206 Partial Content (Range request) support for raw PDF files (/api/file/raw?path=...).
3. S03 PyMuPDF rendering endpoint (/api/pdf/render) with persistent cache using app.backend.server functions.
4. Benchmark metrics recording endpoint (/api/benchmark/record).
"""

from __future__ import annotations

import json
import mimetypes
import os
import re
import sys
from http import HTTPStatus
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

# Ensure app root is in sys.path
STAND_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = STAND_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import app.backend.server as app_server

SAMPLES_DIR = STAND_DIR / "samples"
VENDOR_DIR = STAND_DIR / "vendor"
BENCHMARK_RESULTS_FILE = STAND_DIR / "benchmark_results.json"

class StandRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STAND_DIR), **kwargs)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        if path == "/api/samples":
            self._handle_samples_list()
        elif path == "/api/file/raw":
            self._handle_raw_file(query)
        elif path == "/api/pdf/render":
            self._handle_pdf_render(query)
        elif path == "/api/pdf/page":
            self._handle_pdf_page(query)
        elif path == "/api/benchmark/results":
            self._handle_get_benchmark_results()
        else:
            super().do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/pdf/clear_cache":
            self._handle_clear_cache()
        elif path == "/api/benchmark/record":
            self._handle_record_benchmark()
        else:
            self.send_error(HTTPStatus.NOT_FOUND, "Endpoint not found")

    def _send_json(self, status: int, data: dict | list) -> None:
        body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _handle_samples_list(self) -> None:
        samples = []
        if SAMPLES_DIR.exists():
            for f in sorted(SAMPLES_DIR.glob("*.pdf")):
                samples.append({
                    "name": f.name,
                    "path": str(f.resolve()),
                    "size_bytes": f.stat().st_size,
                    "size_kb": round(f.stat().st_size / 1024, 1),
                })
        for f in sorted(PROJECT_ROOT.glob("*.pdf")):
            if f.name.startswith("sample_"):
                continue
            samples.append({
                "name": f.name,
                "path": str(f.resolve()),
                "size_bytes": f.stat().st_size,
                "size_kb": round(f.stat().st_size / 1024, 1),
            })
        self._send_json(HTTPStatus.OK, {"samples": samples})

    def _handle_raw_file(self, query: dict) -> None:
        target_path_str = query.get("path", [None])[0]
        if not target_path_str:
            self.send_error(HTTPStatus.BAD_REQUEST, "Missing path parameter")
            return

        file_path = Path(unquote(target_path_str)).resolve()
        if not file_path.exists() or not file_path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, f"File not found: {file_path}")
            return

        # Handle HTTP 206 Partial Content (Range request) for PDF.js streaming
        file_size = file_path.stat().st_size
        range_header = self.headers.get("Range")
        content_type = mimetypes.guess_type(str(file_path))[0] or "application/pdf"

        if range_header:
            range_match = re.match(r"bytes=(\d+)-(\d*)", range_header)
            if range_match:
                start = int(range_match.group(1))
                end = int(range_match.group(2)) if range_match.group(2) else file_size - 1
                if start >= file_size:
                    self.send_response(HTTPStatus.REQUESTED_RANGE_NOT_SATISFIABLE)
                    self.send_header("Content-Range", f"bytes */{file_size}")
                    self.end_headers()
                    return

                end = min(end, file_size - 1)
                length = end - start + 1

                self.send_response(HTTPStatus.PARTIAL_CONTENT)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Range", f"bytes {start}-{end}/{file_size}")
                self.send_header("Content-Length", str(length))
                self.send_header("Accept-Ranges", "bytes")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()

                with open(file_path, "rb") as f:
                    f.seek(start)
                    self.wfile.write(f.read(length))
                return

        # Full file response
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(file_size))
        self.send_header("Accept-Ranges", "bytes")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

        with open(file_path, "rb") as f:
            while chunk := f.read(64 * 1024):
                self.wfile.write(chunk)

    def _handle_pdf_render(self, query: dict) -> None:
        target_path_str = query.get("path", [None])[0]
        first_page_only = query.get("first_page_only", ["false"])[0].lower() == "true"
        dpi_str = query.get("dpi", ["150"])[0]
        try:
            dpi = int(dpi_str)
        except ValueError:
            dpi = 150

        if not target_path_str:
            self.send_error(HTTPStatus.BAD_REQUEST, "Missing path parameter")
            return

        file_path = Path(unquote(target_path_str)).resolve()
        if not file_path.exists() or not file_path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, f"File not found: {file_path}")
            return

        try:
            result = app_server.render_pdf(
                file_path,
                dpi=dpi,
                first_page_only=first_page_only,
            )
            self._send_json(HTTPStatus.OK, result)
        except Exception as e:
            self.send_error(HTTPStatus.INTERNAL_SERVER_ERROR, f"PDF Render error: {e}")

    def _handle_pdf_page(self, query: dict) -> None:
        target_path_str = query.get("path", [None])[0]
        page_num_str = query.get("page", ["1"])[0]
        dpi_str = query.get("dpi", ["150"])[0]
        try:
            page_num = int(page_num_str)
            dpi = int(dpi_str)
        except ValueError:
            self.send_error(HTTPStatus.BAD_REQUEST, "Invalid page or dpi parameter")
            return

        if not target_path_str:
            self.send_error(HTTPStatus.BAD_REQUEST, "Missing path parameter")
            return

        file_path = Path(unquote(target_path_str)).resolve()
        key = app_server.pdf_cache_key(file_path, dpi)
        cache_entry = app_server.PDF_CACHE_DIR / key
        page_png = cache_entry / f"page-{page_num}.png"

        if not page_png.exists():
            # Lazy render first page if missing
            app_server.render_pdf(file_path, dpi=dpi, first_page_only=False)

        if not page_png.exists():
            self.send_error(HTTPStatus.NOT_FOUND, f"Page {page_num} not found")
            return

        png_bytes = page_png.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "image/png")
        self.send_header("Content-Length", str(len(png_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(png_bytes)

    def _handle_clear_cache(self) -> None:
        import shutil
        count = 0
        if app_server.PDF_CACHE_DIR.exists():
            for item in app_server.PDF_CACHE_DIR.iterdir():
                if item.is_dir():
                    shutil.rmtree(item, ignore_errors=True)
                    count += 1
        self._send_json(HTTPStatus.OK, {"cleared_count": count, "status": "cache purged"})

    def _handle_record_benchmark(self) -> None:
        length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(length)
        try:
            payload = json.loads(raw_body.decode("utf-8"))
            existing = []
            if BENCHMARK_RESULTS_FILE.exists():
                try:
                    with open(BENCHMARK_RESULTS_FILE, "r", encoding="utf-8") as f:
                        existing = json.load(f)
                except Exception:
                    existing = []
            if isinstance(payload, list):
                existing.extend(payload)
            else:
                existing.append(payload)

            with open(BENCHMARK_RESULTS_FILE, "w", encoding="utf-8") as f:
                json.dump(existing, f, ensure_ascii=False, indent=2)

            self._send_json(HTTPStatus.OK, {"status": "saved", "total_records": len(existing)})
        except Exception as e:
            self.send_error(HTTPStatus.BAD_REQUEST, f"Invalid JSON payload: {e}")

    def _handle_get_benchmark_results(self) -> None:
        if BENCHMARK_RESULTS_FILE.exists():
            with open(BENCHMARK_RESULTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._send_json(HTTPStatus.OK, data)
        else:
            self._send_json(HTTPStatus.OK, [])

def run_server(port: int = 8999):
    server = HTTPServer(("127.0.0.1", port), StandRequestHandler)
    print(f"Stand Server running at http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down stand server.")
    finally:
        server.server_close()

if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8999
    run_server(port)
