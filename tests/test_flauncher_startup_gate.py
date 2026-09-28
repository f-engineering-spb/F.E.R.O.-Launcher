"""Startup readiness gate tests for app/flauncher.pyw (Windows-only).

Covers: hidden window creation, show() exactly once after all gates,
no main-window show on health/objects failure, fail-closed single instance.
No AutoCAD, no real GUI: pywebview, sockets and HTTP are faked (stdlib only).
"""

from __future__ import annotations

import importlib.util
import json
import sys
import time
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parent.parent
FLAUNCHER_PATH = REPO_ROOT / "app" / "flauncher.pyw"


class _FakeEvent:
    def __init__(self):
        self.handlers = []

    def __iadd__(self, handler):
        self.handlers.append(handler)
        return self

    def fire(self, *args, **kwargs):
        for handler in list(self.handlers):
            handler(*args, **kwargs)


class _FakeEvents:
    def __init__(self):
        self.loaded = _FakeEvent()
        self.closed = _FakeEvent()


class _FakeWindow:
    instances: list = []

    def __init__(self, title, url=None, html=None, **kwargs):
        self.title = title
        self.url = url
        self.html = html
        self.kwargs = kwargs
        self.events = _FakeEvents()
        self.show_calls = 0
        self.maximize_calls = 0
        self.loaded_html = None
        self.size = None
        _FakeWindow.instances.append(self)

    def show(self):
        self.show_calls += 1

    def maximize(self):
        self.maximize_calls += 1

    def resize(self, width, height):
        self.size = (width, height)

    def load_html(self, html):
        self.loaded_html = html


class _FakeWebview:
    def __init__(self):
        self.start_calls = 0

    def create_window(self, *args, **kwargs):
        return _FakeWindow(*args, **kwargs)

    def start(self, *args, **kwargs):
        self.start_calls += 1


class _FakeResponse:
    def __init__(self, status=200, payload=None):
        self.status = status
        self.payload = payload if payload is not None else {"items": []}

    def read(self):
        return json.dumps(self.payload).encode("utf-8")

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


def load_flauncher():
    """Exec app/flauncher.pyw. No side effects (webview is imported lazily)."""
    spec = importlib.util.spec_from_file_location("flauncher_gate_under_test", FLAUNCHER_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def wait_until(predicate, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if predicate():
            return True
        time.sleep(0.05)
    return predicate()


class TestStartupGate(unittest.TestCase):
    def setUp(self):
        self.fl = load_flauncher()
        self.fl.READINESS_TIMEOUT_SECONDS = 2
        # Fake pywebview stays installed for the whole test (main imports it lazily).
        _FakeWindow.instances = []
        self.fake_webview = _FakeWebview()
        self._old_webview = sys.modules.get("webview")
        sys.modules["webview"] = self.fake_webview
        # Never touch the real system: record shutdown/browser/error calls.
        self.shutdown_calls = []
        self.browser_calls = []
        self.error_calls = []
        fl = self.fl
        self._orig_shutdown = fl.shutdown_server
        self._orig_browser = fl.open_browser_window
        self._orig_error = fl.show_startup_error
        fl.shutdown_server = lambda *a, **k: self.shutdown_calls.append((a, k))
        fl.open_browser_window = lambda *a, **k: self.browser_calls.append((a, k))
        fl.show_startup_error = lambda *a, **k: self.error_calls.append((a, k))
        self.addCleanup(self._restore)

    def _restore(self):
        self.fl.shutdown_server = self._orig_shutdown
        self.fl.open_browser_window = self._orig_browser
        self.fl.show_startup_error = self._orig_error
        if self._old_webview is None:
            sys.modules.pop("webview", None)
        else:
            sys.modules["webview"] = self._old_webview

    def _ok_urlopen(self, *args, **kwargs):
        return _FakeResponse(200, {"items": [{"id": "x"}]})

    def _run_main_healthy(self, fire_loaded=True):
        fl = self.fl
        fl.find_or_start_server = lambda: (8780, None)
        with mock.patch("urllib.request.urlopen", side_effect=self._ok_urlopen):
            with mock.patch("socket.socket") as mock_socket:
                mock_socket.return_value.bind.return_value = None
                thread = __import__("threading").Thread(target=fl.main, daemon=True)
                thread.start()
                main_window = None
                for _ in range(100):
                    time.sleep(0.05)
                    wins = [w for w in _FakeWindow.instances if w.url is not None]
                    if wins:
                        main_window = wins[0]
                        break
                self.assertIsNotNone(main_window, "hidden main window was not created")
                if fire_loaded:
                    main_window.events.loaded.fire()
                thread.join(timeout=10)
        return main_window

    def test_window_created_hidden_without_fullscreen(self):
        window = self._run_main_healthy()
        self.assertTrue(window.kwargs.get("hidden"), "window must be created hidden")
        self.assertNotIn("fullscreen", window.kwargs)
        self.assertIsNot(window.kwargs.get("maximized"), True)

    def test_show_called_once_after_all_gates(self):
        window = self._run_main_healthy()
        self.assertTrue(
            wait_until(lambda: window.show_calls == 1, timeout=5),
            "window.show() must be called exactly once after gates",
        )
        self.assertEqual(window.show_calls, 1)
        self.assertEqual(window.maximize_calls, 1)

    def test_no_show_before_frontend_loaded(self):
        fl = self.fl
        fl.find_or_start_server = lambda: (8780, None)
        with mock.patch("urllib.request.urlopen", side_effect=self._ok_urlopen):
            with mock.patch("socket.socket") as mock_socket:
                mock_socket.return_value.bind.return_value = None
                thread = __import__("threading").Thread(target=fl.main, daemon=True)
                thread.start()
                time.sleep(1.0)
                wins = [w for w in _FakeWindow.instances if w.url is not None]
                # Loaded never fired: main window must stay hidden.
                for w in wins:
                    self.assertEqual(w.show_calls, 0)
                thread.join(timeout=10)

    def test_objects_failure_shows_no_main_window(self):
        fl = self.fl
        fl.find_or_start_server = lambda: (8780, None)

        def failing_urlopen(*args, **kwargs):
            raise ConnectionError("backend down")

        with mock.patch("urllib.request.urlopen", side_effect=failing_urlopen):
            with mock.patch("socket.socket") as mock_socket:
                mock_socket.return_value.bind.return_value = None
                fl.main()
        mains = [w for w in _FakeWindow.instances if w.url is not None]
        for w in mains:
            self.assertEqual(w.show_calls, 0, "main window must not show on objects failure")
        self.assertTrue(self.error_calls, "startup error must be reported on objects failure")

    def test_mutex_held_creates_no_ui(self):
        fl = self.fl
        started = []
        fl.find_or_start_server = lambda: started.append(True) or (8780, None)
        focus_calls = []
        fl.focus_existing_window = lambda: focus_calls.append(True) or False
        with mock.patch("socket.socket") as mock_socket:
            mock_socket.return_value.bind.side_effect = OSError("busy")
            fl.main()
        self.assertEqual(_FakeWindow.instances, [], "no window when mutex is held")
        self.assertEqual(started, [], "no backend start when mutex is held")
        self.assertEqual(len(focus_calls), 1)

    def test_wait_for_objects_ready_success_and_timeout(self):
        fl = self.fl
        with mock.patch("urllib.request.urlopen", side_effect=self._ok_urlopen):
            ok, detail = fl.wait_for_objects_ready("http://127.0.0.1:9/", 2)
        self.assertTrue(ok)
        self.assertIn("objects ready", detail)

        def bad_urlopen(*args, **kwargs):
            raise ConnectionError("down")

        with mock.patch("urllib.request.urlopen", side_effect=bad_urlopen):
            ok, detail = fl.wait_for_objects_ready("http://127.0.0.1:9/", 0.3)
        self.assertFalse(ok)
        self.assertTrue(detail)

    def test_error_window_is_small_with_reason(self):
        fl = self.fl
        fl.show_startup_error = self._orig_error  # exercise the real one with fake webview
        fake = _FakeWebview()
        sys.modules["webview"] = fake
        try:
            fl.show_startup_error("objects endpoint failure: HTTP 500")
        finally:
            sys.modules.pop("webview", None)
        self.assertEqual(len(_FakeWindow.instances), 1)
        err_window = _FakeWindow.instances[0]
        self.assertEqual(err_window.kwargs.get("width"), 620)
        self.assertIn("objects endpoint failure", err_window.html)


if __name__ == "__main__":
    unittest.main()
