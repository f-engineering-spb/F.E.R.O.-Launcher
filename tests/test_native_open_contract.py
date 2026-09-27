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
    ALL_SUPPORTED_EXTENSIONS,
    FORBIDDEN_AUTODETECT_EXES,
    NativeOpenError,
    detect_windows_app_for_ext,
    ensure_native_apps_first_run_initialized,
    get_configured_app_for_extension,
    import_windows_default_app_mappings,
    launch_custom_app,
    launch_native_file,
    launch_system_default,
    load_native_apps_config,
    migrate_and_normalize_apps_config,
    normalize_file_extension,
    save_native_apps_config,
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

    def test_13_first_run_missing_config_initializes_from_windows(self):
        """13. When config file does not exist, ensure_native_apps_first_run_initialized imports Windows defaults safely."""
        config_path = self.base_dir / "native_apps.json"

        def mock_detect(ext):
            if ext == ".pdf":
                return str(self.fake_viewer_exe)
            if ext == ".dwg":
                return "C:\\Autodesk\\AcLauncher.exe"  # Must be rejected!
            return ""

        with patch("app.backend.server.NATIVE_APPS_CONFIG_FILE", config_path), \
             patch("app.backend.server.detect_windows_app_for_ext", side_effect=mock_detect):
            res = ensure_native_apps_first_run_initialized()
            self.assertIsInstance(res, dict)
            self.assertTrue(config_path.exists())

            loaded = load_native_apps_config()
            self.assertTrue(loaded.get("firstRunCompleted"))
            self.assertEqual(loaded.get("version"), 2)
            # PDF was imported
            pdf_entry = loaded.get(".pdf")
            self.assertIsInstance(pdf_entry, dict)
            self.assertEqual(pdf_entry.get("path"), str(self.fake_viewer_exe))
            self.assertEqual(pdf_entry.get("source"), "windows_import")
            # DWG launcher was rejected, path remains empty
            self.assertEqual(loaded.get(".dwg", {}).get("path", ""), "")

    def test_14_second_run_preserves_manual_config(self):
        """14. When firstRunCompleted is True, startup does not re-import and retains existing manual config."""
        config_path = self.base_dir / "native_apps.json"
        initial_cfg = {
            "version": 2,
            "firstRunCompleted": True,
            "useNativeApps": True,
            ".dwg": {
                "path": str(self.fake_viewer_exe),
                "source": "manual",
                "updated_at": "2026-01-01T00:00:00Z",
                "valid_at_save": True,
            },
        }
        import json
        config_path.write_text(json.dumps(initial_cfg), encoding="utf-8")

        with patch("app.backend.server.NATIVE_APPS_CONFIG_FILE", config_path), \
             patch("app.backend.server.detect_windows_app_for_ext") as mock_detect:
            res = ensure_native_apps_first_run_initialized()
            self.assertIsInstance(res, dict)
            mock_detect.assert_not_called()

            loaded = load_native_apps_config()
            self.assertEqual(get_configured_app_for_extension(".dwg", loaded), str(self.fake_viewer_exe))

    def test_15_legacy_string_migration_is_pure(self):
        """15. migrate_and_normalize_apps_config migrates flat string configs to dicts in-memory without side effects."""
        raw_legacy = {
            "useNativeApps": True,
            ".dwg": str(self.fake_viewer_exe),
            ".pdf": "C:\\Program Files\\Adobe\\Acrobat.exe",
        }
        migrated, had_legacy = migrate_and_normalize_apps_config(raw_legacy)
        self.assertTrue(had_legacy)
        self.assertIsInstance(migrated[".dwg"], dict)
        self.assertEqual(migrated[".dwg"]["path"], str(self.fake_viewer_exe))
        self.assertEqual(migrated[".dwg"]["source"], "manual")
        self.assertTrue(migrated[".dwg"]["valid_at_save"])

        self.assertIsInstance(migrated[".pdf"], dict)
        self.assertEqual(migrated[".pdf"]["path"], "C:\\Program Files\\Adobe\\Acrobat.exe")
        self.assertEqual(migrated[".pdf"]["source"], "manual")

        # Original dict should not be mutated
        self.assertIsInstance(raw_legacy[".dwg"], str)

    def test_16_fill_empty_mode_preserves_existing_values(self):
        """16. fill_empty mode populates only empty slots and leaves non-empty user settings untouched."""
        existing_cfg = {
            "version": 2,
            "firstRunCompleted": True,
            "useNativeApps": True,
            ".dwg": {
                "path": str(self.fake_viewer_exe),
                "source": "manual",
                "updated_at": "2026-01-01T00:00:00Z",
                "valid_at_save": True,
            },
        }

        def mock_detect(ext):
            if ext == ".dwg":
                return str(self.fake_acad_exe)  # Should NOT overwrite fake_viewer_exe!
            if ext == ".pdf":
                return str(self.fake_acad_exe)  # Empty slot, should be filled
            return ""

        with patch("app.backend.server.detect_windows_app_for_ext", side_effect=mock_detect):
            updated, report = import_windows_default_app_mappings("fill_empty", existing_cfg, save=False)
            # DWG was skipped because not empty
            self.assertEqual(report["details"][".dwg"]["reason"], "already_configured")
            self.assertEqual(updated[".dwg"]["path"], str(self.fake_viewer_exe))
            # PDF was filled
            self.assertIn(".pdf", report["updated"])
            self.assertEqual(updated[".pdf"]["path"], str(self.fake_acad_exe))
            self.assertEqual(updated[".pdf"]["source"], "windows_import")

    def test_17_replace_confirmed_mode_rejects_aclauncher_and_preserves_dwg(self):
        """17. replace_confirmed mode replaces existing valid values, BUT if Windows returns AcLauncher for DWG, it is NOT cleared."""
        existing_cfg = {
            "version": 2,
            "firstRunCompleted": True,
            "useNativeApps": True,
            ".dwg": {
                "path": str(self.fake_viewer_exe),
                "source": "manual",
                "updated_at": "2026-01-01T00:00:00Z",
                "valid_at_save": True,
            },
            ".pdf": {
                "path": str(self.fake_viewer_exe),
                "source": "manual",
                "updated_at": "2026-01-01T00:00:00Z",
                "valid_at_save": True,
            }
        }

        def mock_detect(ext):
            if ext == ".dwg":
                return "C:\\Autodesk\\AcLauncher.exe"  # Unusable launcher
            if ext == ".pdf":
                return str(self.fake_acad_exe)  # Usable replacement
            return ""

        with patch("app.backend.server.detect_windows_app_for_ext", side_effect=mock_detect):
            updated, report = import_windows_default_app_mappings("replace_confirmed", existing_cfg, save=False)
            # DWG launcher is rejected, existing DWG Viewer is kept!
            self.assertEqual(report["details"][".dwg"]["reason"], "no_usable_windows_association")
            self.assertEqual(updated[".dwg"]["path"], str(self.fake_viewer_exe))
            # PDF is replaced
            self.assertIn(".pdf", report["updated"])
            self.assertEqual(updated[".pdf"]["path"], str(self.fake_acad_exe))

    def test_18_pure_load_config_has_no_side_effects(self):
        """18. load_native_apps_config is pure read/migration: no autodetect, no file writes."""
        config_path = self.base_dir / "native_apps.json"
        config_path.write_text('{"useNativeApps": true, ".dwg": "C:\\\\acad.exe"}', encoding="utf-8")
        stat_before = config_path.stat().st_mtime_ns

        with patch("app.backend.server.NATIVE_APPS_CONFIG_FILE", config_path), \
             patch("app.backend.server.detect_windows_app_for_ext") as mock_detect, \
             patch("app.backend.server.save_native_apps_config") as mock_save:
            cfg = load_native_apps_config()
            mock_detect.assert_not_called()
            mock_save.assert_not_called()
            self.assertEqual(config_path.stat().st_mtime_ns, stat_before)
            self.assertIsInstance(cfg.get(".dwg"), dict)
            self.assertEqual(cfg[".dwg"]["path"], "C:\\acad.exe")


if __name__ == "__main__":
    unittest.main()

