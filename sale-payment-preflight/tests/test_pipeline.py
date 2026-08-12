#!/usr/bin/env python3
"""End-to-end acceptance checks for the three synthetic case bundles."""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]
CASES_ROOT = SKILL_ROOT / "assets" / "demo-cases"
RUNNER = SKILL_ROOT / "scripts" / "run_preflight.py"


def folder_hashes(folder: Path) -> dict[str, str]:
    return {
        path.relative_to(folder).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in sorted(folder.rglob("*"))
        if path.is_file()
    }


class PipelineAcceptanceTests(unittest.TestCase):
    maxDiff = None

    @classmethod
    def setUpClass(cls):
        if not CASES_ROOT.exists():
            subprocess.run([sys.executable, str(SKILL_ROOT / "tests" / "generate_demo_cases.py")], check=True)

    def test_each_synthetic_case_matches_expected_controls(self):
        cases = sorted(path for path in CASES_ROOT.iterdir() if path.is_dir())
        self.assertEqual(3, len(cases))
        with tempfile.TemporaryDirectory(prefix="sale-preflight-tests-") as temp_root:
            for case in cases:
                with self.subTest(case=case.name):
                    materials = case / "materials"
                    expected = json.loads((case / "expected.json").read_text(encoding="utf-8"))
                    before = folder_hashes(materials)
                    output = Path(temp_root) / case.name
                    completed = subprocess.run(
                        [
                            sys.executable,
                            "-X",
                            "utf8",
                            str(RUNNER),
                            "--case-dir",
                            str(materials),
                            "--config",
                            str(case / "case-config.json"),
                            "--output-dir",
                            str(output),
                        ],
                        check=False,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                    )
                    self.assertEqual(0, completed.returncode, completed.stderr)
                    analysis = json.loads((output / "analysis.json").read_text(encoding="utf-8"))
                    titles = {item["title"] for item in analysis["issues"]}

                    self.assertEqual(expected["outstanding_principal"], analysis["ledger"]["outstanding_principal"])
                    self.assertGreaterEqual(len(analysis["materials"]), expected["minimum_materials"])
                    for title in expected.get("required_issue_titles", []):
                        self.assertIn(title, titles)
                    for title in expected.get("required_issue_titles_absent", []):
                        self.assertNotIn(title, titles)
                    self.assertGreaterEqual(
                        analysis["audit"]["unsupported_or_blocked_files"],
                        expected.get("minimum_blocked_or_ocr", 0),
                    )
                    self.assertTrue(analysis["audit"]["originals_unchanged"])
                    self.assertTrue(analysis["audit"]["source_traceability_complete"])
                    self.assertEqual(0, analysis["audit"]["fabricated_fact_count"])
                    self.assertEqual(before, folder_hashes(materials))

    def test_every_fact_has_precise_material_location(self):
        case = CASES_ROOT / "01_complete_performance"
        with tempfile.TemporaryDirectory(prefix="sale-preflight-trace-") as temp_root:
            output = Path(temp_root)
            subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    str(RUNNER),
                    "--case-dir",
                    str(case / "materials"),
                    "--config",
                    str(case / "case-config.json"),
                    "--output-dir",
                    str(output),
                ],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            analysis = json.loads((output / "analysis.json").read_text(encoding="utf-8"))
            self.assertGreater(len(analysis["facts"]), 0)
            timeline_dates = {event["date"] for event in analysis["timeline"]}
            self.assertTrue({"2025-11-15", "2025-12-10", "2025-12-20", "2026-01-20"}.issubset(timeline_dates))
            for fact in analysis["facts"]:
                self.assertRegex(fact["source"]["material_id"], r"^M\d{3}$")
                self.assertTrue(fact["source"]["location"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
