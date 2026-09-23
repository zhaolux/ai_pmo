from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path

from ai_pmo.current_workbooks import resolve_current


class CurrentWorkbookTests(unittest.TestCase):
    def _register(self, suite: Path, values: dict[str, str]) -> None:
        config = suite / "automation" / "config"
        config.mkdir(parents=True, exist_ok=True)
        (config / "current_workbooks.json").write_text(
            json.dumps(values, ensure_ascii=False), encoding="utf-8",
        )

    def test_registered_version_wins_even_if_old_version_has_newer_mtime(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            directory = suite / "02_计划与进度管理"
            directory.mkdir()
            current = directory / "计划与进度管理-20260920-V2.xlsx"
            old = directory / "计划与进度管理-20260918-V1.xlsx"
            current.write_bytes(b"current")
            old.write_bytes(b"old")
            os.utime(old, (current.stat().st_mtime + 100, current.stat().st_mtime + 100))
            self._register(suite, {"plan": "02_计划与进度管理/计划与进度管理-20260920-V2.xlsx"})

            self.assertEqual(resolve_current(suite, "plan"), current)

    def test_missing_registration_does_not_fall_back_to_latest_file(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            directory = suite / "02_计划与进度管理"
            directory.mkdir()
            (directory / "计划与进度管理-20260920-V1.xlsx").write_bytes(b"present")

            with self.assertRaises(FileNotFoundError):
                resolve_current(suite, "plan")

    def test_rejects_archive_and_path_escape(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            archived = suite / "02_计划与进度管理" / "_archive"
            archived.mkdir(parents=True)
            (archived / "计划与进度管理-20260918-V1.xlsx").write_bytes(b"old")
            for path in (
                "02_计划与进度管理/_archive/计划与进度管理-20260918-V1.xlsx",
                "../elsewhere.xlsx",
            ):
                self._register(suite, {"plan": path})
                with self.assertRaises(ValueError):
                    resolve_current(suite, "plan")

    def test_rejects_wrong_module_prefix_and_missing_registered_file(self):
        with tempfile.TemporaryDirectory() as folder:
            suite = Path(folder)
            wrong = suite / "02_计划与进度管理" / "风险管理-20260920-V1.xlsx"
            wrong.parent.mkdir()
            wrong.write_bytes(b"wrong")
            self._register(suite, {"plan": "02_计划与进度管理/风险管理-20260920-V1.xlsx"})
            with self.assertRaises(ValueError):
                resolve_current(suite, "plan")
            self._register(suite, {"plan": "02_计划与进度管理/计划与进度管理-20260920-V1.xlsx"})
            with self.assertRaises(FileNotFoundError):
                resolve_current(suite, "plan")


if __name__ == "__main__":
    unittest.main()
