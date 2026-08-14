from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import launcher


class LauncherTests(unittest.TestCase):
    def test_self_test_validates_operation_catalog(self) -> None:
        self.assertEqual(launcher.self_test(), 0)

    def test_main_routes_self_test_without_opening_gui(self) -> None:
        with patch.object(
            launcher, "self_test", return_value=0
        ) as self_test, patch.object(launcher, "launch") as launch:
            self.assertEqual(launcher.main(["--self-test"]), 0)
        self_test.assert_called_once_with()
        launch.assert_not_called()

    def test_main_records_self_test_failure_without_windowed_error_dialog(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            error_file = Path(folder) / "self-test-error.log"
            with patch.object(
                launcher, "self_test", side_effect=RuntimeError("diagnostic failure")
            ), patch.dict(
                os.environ,
                {launcher._SELF_TEST_ERROR_FILE_ENV: str(error_file)},
                clear=False,
            ):
                self.assertEqual(launcher.main(["--self-test"]), 87)
            self.assertIn("diagnostic failure", error_file.read_text(encoding="utf-8"))

    def test_multiprocessing_runtime_probe(self) -> None:
        launcher._verify_multiprocessing_runtime()

    def test_frozen_self_test_checks_tk_runtime(self) -> None:
        with patch.object(launcher.sys, "frozen", True, create=True), patch.object(
            launcher, "_verify_tk_runtime"
        ) as verify_tk, patch.object(
            launcher, "_verify_multiprocessing_runtime"
        ) as verify_multiprocessing:
            self.assertEqual(launcher.self_test(), 0)
        verify_tk.assert_called_once_with()
        verify_multiprocessing.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
