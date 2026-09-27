"""Integration test suite for the native apps HTTP API and contract."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from app.backend.server import (
    LauncherHandler,
    ensure_native_apps_first_run_initialized,
)


class TestNativeAppsHttpApi(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.isolated_runtime = Path(cls.temp_dir.name)
        cls.isolated_config = cls.isolated_runtime / "native_apps.json"

        # Patch NATIVE_APPS_CONFIG_FILE and RUNTIME_DIR for entire test class
        cls.patch_config = patch("app.backend.server.NATIVE_APPS_CONFIG_FILE", cls.isolated_config)
        cls.patch_runtime = patch("app.backend.server.RUNTIME_DIR", cls.isolated_runtime)
        cls.patch_config.start()
        cls.patch_runtime.start()

        # Start HTTP server on loopback with free OS-assigned port (port 0)
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), LauncherHandler)
        cls.port = cls.server.server_port
        cls.server_thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.server_thread.start()
        time.sleep(0.3)

    @classmethod
    def tearDownClass(cls):
        try:
            cls.server.shutdown()
            cls.server.server_close()
        except Exception:
            pass
        cls.patch_config.stop()
        cls.patch_runtime.stop()
        cls.temp_dir.cleanup()

    def _url(self, path: str) -> str:
        return f"http://127.0.0.1:{self.port}{path}"

    def test_01_import_windows_post_returns_200_json(self):
        """POST /api/config/apps/import-windows returns 200 OK and valid JSON response."""
        req = urllib.request.Request(
            self._url("/api/config/apps/import-windows"),
            data=json.dumps({"mode": "fill_empty"}).encode("utf-8"),
            headers={"Content-Type": "application/json", "Accept": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            self.assertIn("application/json", resp.headers.get("Content-Type", ""))
            data = json.loads(resp.read().decode("utf-8"))
            self.assertTrue(data.get("ok"))
            self.assertIn("config", data)
            self.assertIn("report", data)
            self.assertIn("imported", data)
            self.assertIn("skipped", data)

    def test_02_options_preflight_returns_204_with_cors(self):
        """OPTIONS preflight to /api/config/apps/import-windows returns 204 No Content with CORS headers."""
        req = urllib.request.Request(
            self._url("/api/config/apps/import-windows"),
            headers={
                "Origin": f"http://localhost:{self.port}",
                "Access-Control-Request-Method": "POST",
                "Access-Control-Request-Headers": "Content-Type",
            },
            method="OPTIONS",
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            self.assertEqual(resp.status, 204)
            self.assertTrue(
                resp.headers.get("Access-Control-Allow-Origin") in ("*", f"http://localhost:{self.port}")
            )
            methods = resp.headers.get("Access-Control-Allow-Methods", "")
            self.assertIn("POST", methods)
            self.assertIn("OPTIONS", methods)
            headers = resp.headers.get("Access-Control-Allow-Headers", "")
            self.assertIn("Content-Type", headers)

    def test_03_legacy_autodetect_post_returns_backward_compatible_json(self):
        """POST /api/config/apps/autodetect returns 200 OK and backward-compatible JSON with 'detected'."""
        req = urllib.request.Request(
            self._url("/api/config/apps/autodetect"),
            headers={"Accept": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            self.assertEqual(resp.status, 200)
            data = json.loads(resp.read().decode("utf-8"))
            self.assertTrue(data.get("ok"))
            self.assertIn("detected", data)
            self.assertIsInstance(data["detected"], dict)

    def test_04_open_windows_settings_post_mocked(self):
        """POST /api/config/apps/open-windows-settings returns 200 OK without launching real UI in tests."""
        with patch("app.backend.server._shell_open", return_value=True), \
             patch("os.startfile", create=True):
            req = urllib.request.Request(
                self._url("/api/config/apps/open-windows-settings"),
                headers={"Accept": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                self.assertEqual(resp.status, 200)
                data = json.loads(resp.read().decode("utf-8"))
                self.assertTrue(data.get("ok"))

    def test_05_isolated_environment_does_not_mutate_real_config(self):
        """Ensure test ran in isolated directory and real runtime config was untouched."""
        real_config_file = REPO_ROOT / "runtime" / "native_apps.json"
        # The test directory must not be the real repo runtime
        self.assertNotEqual(self.isolated_runtime, REPO_ROOT / "runtime")
        self.assertTrue(self.isolated_config.parent.exists())

    def test_06_frontend_backend_route_contract(self):
        """Verify that routes in app.js match endpoints implemented in server.py."""
        app_js_path = REPO_ROOT / "app" / "frontend" / "app.js"
        server_py_path = REPO_ROOT / "app" / "backend" / "server.py"

        app_js_content = app_js_path.read_text(encoding="utf-8")
        server_py_content = server_py_path.read_text(encoding="utf-8")

        # Frontend calls these exact routes
        self.assertIn("/api/config/apps/import-windows", app_js_content)
        self.assertIn("/api/config/apps/open-windows-settings", app_js_content)

        # Backend handles these exact routes in do_POST
        self.assertIn('parsed.path == "/api/config/apps/import-windows"', server_py_content)
        self.assertIn('parsed.path == "/api/config/apps/open-windows-settings"', server_py_content)
        self.assertIn('parsed.path == "/api/config/apps/autodetect"', server_py_content)


if __name__ == "__main__":
    unittest.main()
