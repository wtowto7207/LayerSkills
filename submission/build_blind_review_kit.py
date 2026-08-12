#!/usr/bin/env python3
"""Build answer-free lawyer blind-review and separately sealed answer packages."""

from __future__ import annotations

import hashlib
import io
import json
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "sale-payment-preflight"
CASES = SKILL / "assets" / "demo-cases"
OUTPUTS = ROOT / "outputs" / "sale-payment-preflight-20260812"
RELEASE = ROOT / "submission" / "release"
FIXED_TIME = (2026, 8, 12, 0, 0, 0)


def info(name: str) -> zipfile.ZipInfo:
    item = zipfile.ZipInfo(name, FIXED_TIME)
    item.compress_type = zipfile.ZIP_DEFLATED
    item.external_attr = 0o644 << 16
    return item


def add_bytes(archive: zipfile.ZipFile, name: str, data: bytes, manifest: list[dict]) -> None:
    archive.writestr(info(name), data)
    manifest.append({"path": name, "size": len(data), "sha256": hashlib.sha256(data).hexdigest()})


def add_file(archive: zipfile.ZipFile, source: Path, name: str, manifest: list[dict]) -> None:
    add_bytes(archive, name, source.read_bytes(), manifest)


def build_blind_skill() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for path in sorted(SKILL.rglob("*")):
            if not path.is_file():
                continue
            relative = path.relative_to(SKILL).as_posix()
            parts = set(path.relative_to(SKILL).parts)
            if "__pycache__" in parts or path.suffix == ".pyc" or "tests" in parts:
                continue
            if path.name == "expected.json" and "demo-cases" in parts:
                continue
            archive.writestr(info(f"sale-payment-preflight/{relative}"), path.read_bytes())
    return buffer.getvalue()


def build_reviewer_package() -> Path:
    target = RELEASE / "blind-review-kit.zip"
    manifest: list[dict] = []
    with zipfile.ZipFile(target, "w") as archive:
        add_bytes(
            archive,
            "01_Skill包/sale-payment-preflight-blind.zip",
            build_blind_skill(),
            manifest,
        )
        for case in sorted(path for path in CASES.iterdir() if path.is_dir()):
            for path in sorted((case / "materials").rglob("*")):
                if path.is_file():
                    add_file(
                        archive,
                        path,
                        f"02_盲测案卷/{case.name}/materials/{path.relative_to(case / 'materials').as_posix()}",
                        manifest,
                    )
            add_file(
                archive,
                case / "case-config.json",
                f"02_盲测案卷/{case.name}/case-config.json",
                manifest,
            )
        add_file(
            archive,
            ROOT / "submission" / "律师盲测与验收记录.xlsx",
            "03_记录表/律师盲测与验收记录.xlsx",
            manifest,
        )
        add_file(
            archive,
            ROOT / "submission" / "blind-review-instructions.md",
            "盲测操作说明.md",
            manifest,
        )
        add_bytes(
            archive,
            "manifest.json",
            json.dumps({"schema_version": "1.0", "role": "blind_reviewer", "files": manifest}, ensure_ascii=False, indent=2).encode("utf-8"),
            [],
        )
    return target


def build_answer_package() -> Path:
    target = RELEASE / "blind-review-answer-key.zip"
    manifest: list[dict] = []
    references: dict[str, dict] = {}
    with zipfile.ZipFile(target, "w") as archive:
        add_file(
            archive,
            ROOT / "submission" / "blind-answer-key-note.md",
            "答案使用说明.md",
            manifest,
        )
        for case in sorted(path for path in CASES.iterdir() if path.is_dir()):
            expected = case / "expected.json"
            add_file(archive, expected, f"expected/{case.name}/expected.json", manifest)
            analysis_path = OUTPUTS / case.name / "analysis.json"
            analysis = json.loads(analysis_path.read_text(encoding="utf-8"))
            references[case.name] = {
                "materials": len(analysis["materials"]),
                "facts": len(analysis["facts"]),
                "issues": len(analysis["issues"]),
                "issue_titles": [item["title"] for item in analysis["issues"]],
                "outstanding_principal": analysis["ledger"]["outstanding_principal"],
                "unsupported_or_blocked_files": analysis["audit"]["unsupported_or_blocked_files"],
                "source_traceability_complete": analysis["audit"]["source_traceability_complete"],
                "fabricated_fact_count": analysis["audit"]["fabricated_fact_count"],
                "network_access_used": analysis["audit"]["network_access_used"],
                "lawyer_review_status": analysis["audit"]["lawyer_review_status"],
            }
        add_bytes(
            archive,
            "technical-reference-results.json",
            json.dumps(
                {
                    "schema_version": "1.0",
                    "warning": "技术参考结果，不构成法律专业答案，不替代律师独立判断。",
                    "cases": references,
                },
                ensure_ascii=False,
                indent=2,
            ).encode("utf-8"),
            manifest,
        )
        add_bytes(
            archive,
            "manifest.json",
            json.dumps({"schema_version": "1.0", "role": "answer_custodian", "files": manifest}, ensure_ascii=False, indent=2).encode("utf-8"),
            [],
        )
    return target


def main() -> int:
    RELEASE.mkdir(parents=True, exist_ok=True)
    reviewer = build_reviewer_package()
    answers = build_answer_package()
    print(json.dumps({
        "status": "ok",
        "reviewer_package": str(reviewer),
        "reviewer_sha256": hashlib.sha256(reviewer.read_bytes()).hexdigest(),
        "answer_package": str(answers),
        "answer_sha256": hashlib.sha256(answers.read_bytes()).hexdigest(),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
