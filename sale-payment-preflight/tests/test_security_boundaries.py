#!/usr/bin/env python3
"""Security and scope-gate acceptance tests for sale-payment-preflight."""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]
RUNNER = SKILL_ROOT / "scripts" / "run_preflight.py"
BASE_CASE = SKILL_ROOT / "assets" / "demo-cases" / "01_complete_performance"
ANOMALY_CASE = SKILL_ROOT / "assets" / "demo-cases" / "03_conflict_missing"


def folder_hashes(folder: Path) -> dict[str, str]:
    return {
        path.relative_to(folder).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in sorted(folder.rglob("*"))
        if path.is_file()
    }


def run_preflight(materials: Path, config: Path, output: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            sys.executable,
            "-X",
            "utf8",
            str(RUNNER),
            "--case-dir",
            str(materials),
            "--config",
            str(config),
            "--output-dir",
            str(output),
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


class SecurityBoundaryTests(unittest.TestCase):
    def test_scope_gate_rejects_unauthorized_buyer_and_non_cny_config(self):
        base_config = json.loads((BASE_CASE / "case-config.json").read_text(encoding="utf-8"))
        mutations = (
            ("unauthorized", {"authorized": False}, "authorized=true"),
            ("buyer-side", {"client_side": "buyer"}, "仅支持卖方"),
            ("non-cny", {"currency": "USD"}, "仅支持单币种 CNY"),
        )
        with tempfile.TemporaryDirectory(prefix="sale-preflight-scope-") as temp_root:
            root = Path(temp_root)
            for name, mutation, expected_error in mutations:
                with self.subTest(name=name):
                    config = dict(base_config)
                    config.update(mutation)
                    config_path = root / f"{name}.json"
                    config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
                    output = root / f"output-{name}"
                    completed = run_preflight(BASE_CASE / "materials", config_path, output)
                    self.assertEqual(2, completed.returncode)
                    self.assertIn(expected_error, completed.stderr)
                    self.assertFalse((output / "analysis.json").exists())

    def test_output_inside_materials_is_rejected_without_writing(self):
        with tempfile.TemporaryDirectory(prefix="sale-preflight-output-boundary-") as temp_root:
            root = Path(temp_root)
            materials = root / "materials"
            shutil.copytree(BASE_CASE / "materials", materials)
            before = folder_hashes(materials)
            output = materials / "generated-output"
            completed = run_preflight(materials, BASE_CASE / "case-config.json", output)
            self.assertEqual(2, completed.returncode)
            self.assertIn("输出目录必须位于案件材料目录之外", completed.stderr)
            self.assertFalse(output.exists())
            self.assertEqual(before, folder_hashes(materials))

    def test_injection_ocr_and_non_cny_ledger_are_quarantined(self):
        before = folder_hashes(ANOMALY_CASE / "materials")
        with tempfile.TemporaryDirectory(prefix="sale-preflight-anomaly-") as temp_root:
            output = Path(temp_root)
            completed = run_preflight(
                ANOMALY_CASE / "materials",
                ANOMALY_CASE / "case-config.json",
                output,
            )
            self.assertEqual(0, completed.returncode, completed.stderr)
            analysis = json.loads((output / "analysis.json").read_text(encoding="utf-8"))
            titles = {item["title"] for item in analysis["issues"]}
            statuses = {item["status"] for item in analysis["materials"]}

            self.assertIn("材料含疑似提示词注入文本", titles)
            self.assertIn("发现非 CNY 流水", titles)
            self.assertIn("needs_ocr", statuses)
            self.assertTrue(analysis["audit"]["embedded_instruction_materials"])
            self.assertFalse(analysis["audit"]["network_access_used"])
            self.assertTrue(analysis["audit"]["originals_unchanged"])
            self.assertEqual("100000.00", analysis["ledger"]["outstanding_principal"])
            self.assertEqual(before, folder_hashes(ANOMALY_CASE / "materials"))

    def test_encrypted_and_unsupported_files_are_inventory_only(self):
        try:
            from pypdf import PdfWriter
        except ImportError as exc:  # pragma: no cover - the main PDF path has the same dependency
            self.skipTest(f"pypdf unavailable: {exc}")

        with tempfile.TemporaryDirectory(prefix="sale-preflight-blocked-") as temp_root:
            root = Path(temp_root)
            materials = root / "materials"
            shutil.copytree(BASE_CASE / "materials", materials)

            writer = PdfWriter()
            writer.add_blank_page(width=72, height=72)
            writer.encrypt("synthetic-password")
            encrypted = materials / "90_encrypted.pdf"
            with encrypted.open("wb") as stream:
                writer.write(stream)
            (materials / "91_unsupported.bin").write_bytes(b"synthetic unsupported evidence")

            before = folder_hashes(materials)
            output = root / "output"
            completed = run_preflight(materials, BASE_CASE / "case-config.json", output)
            self.assertEqual(0, completed.returncode, completed.stderr)
            analysis = json.loads((output / "analysis.json").read_text(encoding="utf-8"))
            by_path = {item["relative_path"]: item for item in analysis["materials"]}

            self.assertEqual("blocked_encrypted", by_path["90_encrypted.pdf"]["status"])
            self.assertEqual("unsupported", by_path["91_unsupported.bin"]["status"])
            self.assertEqual(0, by_path["90_encrypted.pdf"]["locations_extracted"])
            self.assertEqual(0, by_path["91_unsupported.bin"]["locations_extracted"])
            self.assertGreaterEqual(analysis["audit"]["unsupported_or_blocked_files"], 2)
            self.assertEqual(before, folder_hashes(materials))


if __name__ == "__main__":
    unittest.main(verbosity=2)
