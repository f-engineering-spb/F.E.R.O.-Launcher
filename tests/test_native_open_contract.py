"""Automated test suite for the native file opening priority contract.

Validates:
1. Normalization of file extensions (.DWG, dwg, .dwg -> .dwg).
2. Priority of configured DWG Viewer over Windows fallback.
3. Priority of configured AutoCAD over Windows fallback.
4. Fallback to Windows default when setting is empty.
5. Error raised (configured_app_missing) when configured EXE does not exist.
6. Safe argument passing (separate list arguments, no cmd/shell concatenation).
7. Original source path invariant (no preview/cache path substitution).
8. Protection of manual user choices against silent autodetect overwrite.
9. No invocation of accoreconsole or render pipelines during native open.
10. Preservation of contract across other formats (PDF, DOCX, XLSX).
11. Rejection of AcLauncher.exe / ZwLauncher.exe in autodetect for DWG.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from app.backend.server import (
    FORBIDDEN_AUTODETECT_EXES,
    NativeOpenError,
    detect_windows_app_for_ext,
    get_configured_app_for_extension,
    launch_custom_app,
    launch_native_file,
    launch_system_default,
    normalize_file_extension,
    validate_configured_executable,
)


class TestNativeOpenContract(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.base_dir = Path(self.temp_dir.name)

        # Create dummy executables
        self.fake_viewer_exe = self.base_dir / "DWGTrueView.exe"
        self.fake_viewer_exe.write_text("dummy viewer exe")

        self.fake_acad_exe = self.base_dir / "acad.exe"
        self.fake_acad_exe.write_text("dummy acad exe")

        # Create dummy test files
        self.test_dwg = self.base_dir / "чертёж [объект 1].dwg"
        self.test_dwg.write_bytes(b"AC1032 dummy dwg")

        self.test_pdf = self.base_dir / "документ с пробелами.pdf"
        self.test_pdf.write_bytes(b"%PDF-1.4 dummy pdf")

        self.test_xlsx = self.base_dir / "таблица (v1).xlsx"
        self.test_xlsx.write_bytes(b"dummy xlsx")

    def test_01_extension_normalization(self):
        """1. .DWG, dwg, .dwg normalize to .dwg."""
        self.assertEqual(normalize_file_extension(".DWG"), ".dwg")
        self.assertEqual(normalize_file_extension("DWG"), ".dwg")
        self.assertEqual(normalize_file_extension("dwg"), ".dwg")
        self.assertEqual(normalize_file_extension(".dwg"), ".dwg")
        self.assertEqual(normalize_file_extension(Path("test.DWG")), ".dwg")
        self.assertEqual(normalize_file_extension(Path("folder/file.DXF")), ".dxf")
        self.assertEqual(normalize_file_extension(""), "")

    def test_02_configured_dwg_viewer_priority(self):
        """2. Valid configured DWG Viewer has priority over Windows fallback and AutoCAD."""
        config = {
            "useNativeApps": True,
            ".dwg": str(self.fake_viewer_exe),
        }
        with patch("app.backend.server.load_native_apps_config", return_value=config), \
             patch("app.backend.server.launch_custom_app", return_value="custom-app:DWGTrueView.exe") as mock_custom, \
             patch("app.backend.server.launch_system_default") as mock_system:
            res = launch_native_file(self.test_dwg)
            self.assertEqual(res, "custom-app:DWGTrueView.exe")
            mock_custom.assert_called_once_with(str(self.fake_viewer_exe), self.test_dwg)
            mock_system.assert_not_called()

    def test_03_configured_autocad_priority(self):
        """3. Valid configured AutoCAD has priority over Windows fallback."""
        config = {
            "useNativeApps": True,
            ".dwg": str(self.fake_acad_exe),
        }
        with patch("app.backend.server.load_native_apps_config", return_value=config), \
             patch("app.backend.server.launch_custom_app", return_value="custom-app:acad.exe") as mock_custom, \
             patch("app.backend.server.launch_system_default") as mock_system:
            res = launch_native_file(self.test_dwg)
            self.assertEqual(res, "custom-app:acad.exe")
            mock_custom.assert_called_once_with(str(self.fake_acad_exe), self.test_dwg)
            mock_system.assert_not_called()

    def test_04_empty_setting_uses_windows_fallback(self):
        """4. Empty configured app uses Windows fallback (launch_system_default)."""
        config = {
            "useNativeApps": True,
            ".dwg": "",
            "settingDwgExe": "",
        }
        with patch("app.backend.server.load_native_apps_config", return_value=config), \
             patch("app.backend.server.launch_custom_app") as mock_custom, \
             patch("app.backend.server.launch_system_default", return_value="system-default") as mock_system:
            res = launch_native_file(self.test_dwg)
            self.assertEqual(res, "system-default")
            mock_custom.assert_not_called()
            mock_system.assert_called_once_with(self.test_dwg)

    def test_05_missing_configured_exe_raises_structured_error(self):
        """5. Missing configured EXE returns configured_app_missing without fallback."""
        non_existent_exe = self.base_dir / "non_existent_viewer.exe"
        config = {
            "useNativeApps": True,
            ".dwg": str(non_existent_exe),
        }
        with patch("app.backend.server.load_native_apps_config", return_value=config), \
             patch("app.backend.server.launch_custom_app") as mock_custom, \
             patch("app.backend.server.launch_system_default") as mock_system:
            with self.assertRaises(NativeOpenError) as ctx:
                launch_native_file(self.test_dwg)
            self.assertEqual(ctx.exception.error_code, "configured_app_missing")
            self.assertIn(".dwg", ctx.exception.details.get("extension", ""))
            mock_custom.assert_not_called()
            mock_system.assert_not_called()

    def test_06_safe_argument_passing(self):
        """6. Paths with spaces, Cyrillic, brackets passed as separate list arguments, never string concat."""
        with patch("app.backend.server._shell_open", return_value=False), \
             patch("subprocess.Popen") as mock_popen, \
             patch("app.backend.server.bring_native_window_to_front"), \
             patch("app.backend.server._bring_window_to_front"):
            mock_popen.return_value.pid = 12345
            res = launch_custom_app(str(self.fake_viewer_exe), self.test_dwg)
            self.assertEqual(res, "custom-app:DWGTrueView.exe")
            # Verify Popen called with exact separate list items [exe, file]
            call_args = mock_popen.call_args
            self.assertIsNotNone(call_args)
            args_list = call_args[0][0]
            self.assertIsInstance(args_list, list)
            self.assertEqual(len(args_list), 2)
            self.assertEqual(args_list[0], os.path.normpath(str(self.fake_viewer_exe)))
            self.assertEqual(args_list[1], os.path.normpath(str(self.test_dwg)))
            # Verify shell=True is NOT passed
            self.assertFalse(call_args[1].get("shell", False))

    def test_07_original_source_path_invariant(self):
        """7. Native opening receives original source path, not preview or cache path."""
        config = {
            "useNativeApps": True,
            ".dwg": str(self.fake_viewer_exe),
        }
        with patch("app.backend.server.load_native_apps_config", return_value=config), \
             patch("app.backend.server.launch_custom_app") as mock_custom:
            mock_custom.return_value = "custom-app:DWGTrueView.exe"
            launch_native_file(self.test_dwg)
            called_path = mock_custom.call_args[0][1]
            self.assertEqual(called_path, self.test_dwg)
            self.assertNotIn("cache", str(called_path).casefold())

    def test_08_manual_user_choice_not_overwritten(self):
        """8. Config loader preserves user-configured app for .dwg."""
        config = {
            "useNativeApps": True,
            ".dwg": str(self.fake_viewer_exe),
            "settingDwgExe": str(self.fake_viewer_exe),
        }
        found = get_configured_app_for_extension(".dwg", config)
        self.assertEqual(found, str(self.fake_viewer_exe))

    def test_09_no_accoreconsole_or_render_pipeline_coupling(self):
        """9. Native DWG opening does not invoke accoreconsole or render pipelines."""
        config = {
            "useNativeApps": True,
            ".dwg": str(self.fake_viewer_exe),
        }
        with patch("app.backend.server.load_native_apps_config", return_value=config), \
             patch("app.backend.server.launch_custom_app", return_value="custom-app:DWGTrueView.exe"), \
             patch("app.backend.server.render_dwg_model") as mock_render_dwg, \
             patch("app.backend.server.dwg_to_model_pdf") as mock_dwg_to_pdf:
            res = launch_native_file(self.test_dwg)
            self.assertEqual(res, "custom-app:DWGTrueView.exe")
            mock_render_dwg.assert_not_called()
            mock_dwg_to_pdf.assert_not_called()

    def test_10_other_formats_contract_integrity(self):
        """10. PDF and XLSX adhere to the exact same contract without regressions."""
        # PDF with custom app
        fake_pdf_app = self.base_dir / "Acrobat.exe"
        fake_pdf_app.write_text("dummy acrobat")
        config_pdf = {"useNativeApps": True, ".pdf": str(fake_pdf_app)}
        with patch("app.backend.server.load_native_apps_config", return_value=config_pdf), \
             patch("app.backend.server.launch_custom_app", return_value="custom-app:Acrobat.exe") as mock_pdf:
            self.assertEqual(launch_native_file(self.test_pdf), "custom-app:Acrobat.exe")
            mock_pdf.assert_called_once_with(str(fake_pdf_app), self.test_pdf)

        # XLSX with empty app uses system default
        config_empty = {"useNativeApps": True, ".xlsx": ""}
        with patch("app.backend.server.load_native_apps_config", return_value=config_empty), \
             patch("app.backend.server.launch_system_default", return_value="system-default") as mock_sys:
            self.assertEqual(launch_native_file(self.test_xlsx), "system-default")
            mock_sys.assert_called_once_with(self.test_xlsx)

    def test_11_autodetect_rejects_aclauncher(self):
        """11. Autodetect rejects AcLauncher.exe and ZwLauncher.exe for DWG/DXF."""
        self.assertIn("aclauncher.exe", FORBIDDEN_AUTODETECT_EXES)
        self.assertIn("zwlauncher.exe", FORBIDDEN_AUTODETECT_EXES)

        # Mock winreg returning AcLauncher command line
        mock_k = MagicMock()
        mock_k.__enter__.return_value = mock_k
        mock_k.__exit__.return_value = None
        with patch("app.backend.server.winreg") as mock_winreg:
            mock_winreg.OpenKey.return_value = mock_k
            mock_winreg.QueryValueEx.return_value = ('"C:\\Program Files\\Common Files\\Autodesk Shared\\AcShellEx\\AcLauncher.exe" /O "%1"', 1)
            mock_winreg.EnumValue.side_effect = OSError
            detected = detect_windows_app_for_ext(".dwg")
            self.assertEqual(detected, "", "AcLauncher.exe must be rejected as an autodetect candidate!")

    def test_12_validate_configured_executable(self):
        """12. validate_configured_executable properly handles valid, missing, and non-exe files."""
        # Valid executable file
        is_val, clean = validate_configured_executable(str(self.fake_viewer_exe))
        self.assertTrue(is_val)
        self.assertEqual(clean, str(self.fake_viewer_exe))

        # Quoted executable path
        is_val, clean = validate_configured_executable(f'"{self.fake_viewer_exe}"')
        self.assertTrue(is_val)
        self.assertEqual(clean, str(self.fake_viewer_exe))

        # Missing file
        is_val, clean = validate_configured_executable("C:\\NonExistent\\SomeViewer.exe")
        self.assertFalse(is_val)

        # Empty string
        is_val, clean = validate_configured_executable("")
        self.assertFalse(is_val)


if __name__ == "__main__":
    unittest.main()
