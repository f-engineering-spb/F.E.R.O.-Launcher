"""Поведенческий тест защиты от устаревшего Word-просмотра.

Исполняет настоящий app/frontend/app.js в node (движок доступен в среде,
новых зависимостей нет) и моделирует: задержанный запрос A, выбор B,
готовность B, поздний ответ A. Проверяет действия и состояние:
экран, состояние, отсутствие фолбэка и стабильность статуса.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HARNESS = ROOT / "tests" / "word_stale_harness.js"


class WordStalePreviewBehavior(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if shutil.which("node") is None:
            raise unittest.SkipTest("node недоступен в среде")

    def test_stale_word_preview_never_clobbers(self):
        proc = subprocess.run(
            ["node", str(HARNESS)],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
        try:
            verdict = json.loads(proc.stdout[proc.stdout.index("{"):proc.stdout.rindex("}") + 1])
        except ValueError:
            self.fail(f"harness без JSON-вердикта (exit={proc.returncode}): {proc.stdout[-500:]} {proc.stderr[-500:]}")
        self.assertEqual(proc.returncode, 0, f"harness exit != 0: {verdict}")
        failed = verdict.get("failed", -1)
        names = [r["name"] for r in verdict.get("results", []) if not r.get("pass")]
        self.assertEqual(failed, 0, f"провалены поведенческие проверки: {names}")


if __name__ == "__main__":
    unittest.main()
