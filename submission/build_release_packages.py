from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path, PurePosixPath


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "sale-payment-preflight"
SUBMISSION = ROOT / "submission"
RELEASE = SUBMISSION / "release"
OUTPUTS = ROOT / "outputs" / "sale-payment-preflight-20260812" / "01_complete_performance"

SKILL_ZIP = RELEASE / "sale-payment-preflight.zip"
ENTRY_ZIP = RELEASE / "sale-payment-preflight-entry-package.zip"
BLIND_KIT = RELEASE / "blind-review-kit.zip"
BLIND_ANSWER = RELEASE / "blind-review-answer-key.zip"
MANIFEST = RELEASE / "release-manifest.json"
SUMS = RELEASE / "SHA256SUMS.txt"

FIXED_TIME = (2026, 8, 12, 0, 0, 0)
EXCLUDED_PARTS = {"__pycache__", "node_modules", ".git", ".pytest_cache"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_member(zf: zipfile.ZipFile, source: Path, arcname: str) -> None:
    info = zipfile.ZipInfo(PurePosixPath(arcname).as_posix(), FIXED_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o100644 << 16
    zf.writestr(info, source.read_bytes())


def skill_files() -> list[Path]:
    files: list[Path] = []
    for path in SKILL.rglob("*"):
        if not path.is_file():
            continue
        relative = path.relative_to(SKILL)
        if any(part in EXCLUDED_PARTS for part in relative.parts):
            continue
        if path.suffix.lower() in EXCLUDED_SUFFIXES:
            continue
        files.append(path)
    return sorted(files, key=lambda item: item.relative_to(SKILL).as_posix())


def build_skill_zip() -> None:
    with zipfile.ZipFile(SKILL_ZIP, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for source in skill_files():
            arcname = PurePosixPath(SKILL.name, source.relative_to(SKILL).as_posix()).as_posix()
            write_member(zf, source, arcname)


def entry_items() -> list[tuple[Path, str]]:
    fixed_outputs = [
        "01_案件材料清单.xlsx",
        "02_履约事实时间轴.xlsx",
        "03_要件证据矩阵.xlsx",
        "04_货款核对表.xlsx",
        "05_矛盾与补证清单.docx",
        "06_诉前案卷体检报告.docx",
        "audit.json",
        "validation.json",
    ]
    items: list[tuple[Path, str]] = [
        (SKILL_ZIP, "01_Skill包/sale-payment-preflight.zip"),
        (SUBMISSION / "compatibility-report.json", "02_说明与验收/compatibility-report.json"),
        (SUBMISSION / "second-platform-attempt.md", "02_说明与验收/second-platform-attempt.md"),
        (SUBMISSION / "README.md", "02_说明与验收/README.md"),
        (SUBMISSION / "参赛作品技术说明书.docx", "02_说明与验收/参赛作品技术说明书.docx"),
        (SUBMISSION / "完成审计.md", "02_说明与验收/完成审计.md"),
        (SUBMISSION / "律师盲测与验收记录.xlsx", "02_说明与验收/律师盲测与验收记录.xlsx"),
        (SUBMISSION / "报名信息预填稿_非官方模板.docx", "02_说明与验收/报名信息预填稿_非官方模板.docx"),
        (BLIND_KIT, "04_盲测辅助/blind-review-kit.zip"),
        (SUBMISSION / "blind-review-instructions.md", "04_盲测辅助/盲测评审说明.md"),
        (SUBMISSION / "sale-payment-preflight_4分钟演示.pptx", "05_演示/sale-payment-preflight_4分钟演示.pptx"),
        (SUBMISSION / "sale-payment-preflight_4分钟演示.mp4", "05_演示/sale-payment-preflight_4分钟演示.mp4"),
    ]
    items.extend((OUTPUTS / filename, f"03_完整履约案输出/{filename}") for filename in fixed_outputs)
    return items


def build_entry_zip() -> None:
    items = entry_items()
    missing = [str(source) for source, _ in items if not source.is_file()]
    if missing:
        raise FileNotFoundError("entry package files missing:\n" + "\n".join(missing))
    with zipfile.ZipFile(ENTRY_ZIP, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for source, arcname in items:
            write_member(zf, source, arcname)


def verify_zip(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"corrupt member in {path.name}: {bad}")
        names = zf.namelist()
    forbidden = [
        name
        for name in names
        if any(part in EXCLUDED_PARTS for part in PurePosixPath(name).parts)
        or PurePosixPath(name).suffix.lower() in EXCLUDED_SUFFIXES
    ]
    if forbidden:
        raise RuntimeError(f"forbidden members in {path.name}: {forbidden}")
    return {"entries": len(names), "bytes": path.stat().st_size}


def verify_embedded_skill() -> None:
    expected = SKILL_ZIP.read_bytes()
    with zipfile.ZipFile(ENTRY_ZIP) as zf:
        actual = zf.read("01_Skill包/sale-payment-preflight.zip")
    if actual != expected:
        raise RuntimeError("embedded Skill zip differs from release Skill zip")


def build_manifest(package_stats: dict[str, dict[str, int]]) -> None:
    artifacts = []
    for path in [SKILL_ZIP, ENTRY_ZIP, BLIND_KIT, BLIND_ANSWER]:
        artifacts.append(
            {
                "file": path.name,
                "sha256": sha256(path),
                "bytes": path.stat().st_size,
                **package_stats[path.name],
            }
        )
    payload = {
        "schema": "sale-payment-preflight.release/v1",
        "built_at": "2026-08-12T00:00:00+08:00",
        "video": {
            "file": "../sale-payment-preflight_4分钟演示.mp4",
            "duration_seconds": 240,
            "slides": 10,
            "synthetic_materials_only": True,
        },
        "artifacts": artifacts,
        "external_gates": [
            "official_registration_form_and_signatures",
            "prc_lawyer_rule_approval",
            "two_rounds_of_blind_review",
            "second_live_skill_agent_run",
        ],
    }
    MANIFEST.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def build_sums() -> None:
    files = [SKILL_ZIP, ENTRY_ZIP, BLIND_KIT, BLIND_ANSWER, MANIFEST]
    lines = [f"{sha256(path)}  {path.name}" for path in files]
    SUMS.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    RELEASE.mkdir(parents=True, exist_ok=True)
    for required in [BLIND_KIT, BLIND_ANSWER]:
        if not required.is_file():
            raise FileNotFoundError(required)
    build_skill_zip()
    build_entry_zip()
    package_stats = {
        path.name: verify_zip(path)
        for path in [SKILL_ZIP, ENTRY_ZIP, BLIND_KIT, BLIND_ANSWER]
    }
    verify_embedded_skill()
    build_manifest(package_stats)
    build_sums()
    print(
        json.dumps(
            {
                "status": "ok",
                "packages": package_stats,
                "manifest": str(MANIFEST),
                "sha256sums": str(SUMS),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
