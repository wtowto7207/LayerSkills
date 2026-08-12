#!/usr/bin/env python3
"""Validate Agent Skills packaging and deterministic core execution in two layouts."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


NAME_PATTERN = re.compile(r"^[a-z0-9-]{1,64}$")
ALLOWED_FRONTMATTER = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def parse_frontmatter(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        raise ValueError("SKILL.md 缺少 YAML frontmatter 起始标记")
    try:
        end = lines.index("---", 1)
    except ValueError as exc:
        raise ValueError("SKILL.md 缺少 YAML frontmatter 结束标记") from exc
    data: dict[str, str] = {}
    for line in lines[1:end]:
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if ":" not in line:
            raise ValueError(f"无法解析 frontmatter 行：{line}")
        key, value = line.split(":", 1)
        data[key.strip()] = value.strip().strip("'\"")
    return data


def validate_skill_contract(skill_dir: Path) -> dict[str, Any]:
    skill_file = skill_dir / "SKILL.md"
    if not skill_file.is_file():
        raise ValueError("缺少 SKILL.md")
    frontmatter = parse_frontmatter(skill_file)
    name = frontmatter.get("name", "")
    description = frontmatter.get("description", "")
    unexpected = sorted(set(frontmatter) - ALLOWED_FRONTMATTER)
    text = skill_file.read_text(encoding="utf-8")
    body_lines = len(text.splitlines())
    errors: list[str] = []
    if not NAME_PATTERN.fullmatch(name):
        errors.append("name 不符合 Agent Skills 命名约束")
    if not description or len(description) > 1024:
        errors.append("description 为空或超过 1024 字符")
    if unexpected:
        errors.append(f"存在开放规范之外的 frontmatter 字段：{unexpected}")
    if body_lines >= 500:
        errors.append("SKILL.md 达到或超过 500 行")
    if re.search(r"[A-Za-z]:\\\\", text):
        errors.append("SKILL.md 含 Windows 绝对路径")
    required_dirs = ["scripts", "references", "assets"]
    missing_dirs = [name for name in required_dirs if not (skill_dir / name).is_dir()]
    if missing_dirs:
        errors.append(f"缺少目录：{missing_dirs}")
    return {
        "status": "passed" if not errors else "failed",
        "errors": errors,
        "name": name,
        "description_length": len(description),
        "skill_md_lines": body_lines,
        "frontmatter_fields": sorted(frontmatter),
        "required_directories": {name: (skill_dir / name).is_dir() for name in required_dirs},
    }


def run_staged(staged_skill: Path, output: Path) -> dict[str, Any]:
    case = staged_skill / "assets" / "demo-cases" / "01_complete_performance"
    runner = staged_skill / "scripts" / "run_preflight.py"
    completed = subprocess.run(
        [
            sys.executable,
            "-X",
            "utf8",
            str(runner),
            "--case-dir",
            str(case / "materials"),
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
    if completed.returncode != 0:
        return {
            "status": "failed",
            "returncode": completed.returncode,
            "stderr": completed.stderr.strip(),
        }
    analysis = json.loads((output / "analysis.json").read_text(encoding="utf-8"))
    invariants = {
        "materials": len(analysis["materials"]),
        "facts": len(analysis["facts"]),
        "issues": len(analysis["issues"]),
        "outstanding_principal": analysis["ledger"]["outstanding_principal"],
        "unsupported_or_blocked_files": analysis["audit"]["unsupported_or_blocked_files"],
        "originals_unchanged": analysis["audit"]["originals_unchanged"],
        "source_traceability_complete": analysis["audit"]["source_traceability_complete"],
        "fabricated_fact_count": analysis["audit"]["fabricated_fact_count"],
        "network_access_used": analysis["audit"]["network_access_used"],
    }
    complete = (
        invariants["materials"] == 6
        and invariants["outstanding_principal"] == "100000.00"
        and invariants["unsupported_or_blocked_files"] == 0
        and invariants["originals_unchanged"] is True
        and invariants["source_traceability_complete"] is True
        and invariants["fabricated_fact_count"] == 0
        and invariants["network_access_used"] is False
    )
    return {
        "status": "passed" if complete else "failed",
        "returncode": completed.returncode,
        "invariants": invariants,
        "analysis_sha256": sha256(output / "analysis.json"),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skill-dir", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    skill_dir = args.skill_dir.resolve()
    contract = validate_skill_contract(skill_dir)

    layouts = {
        "codex_personal_skill_layout": Path("codex-home") / "skills" / skill_dir.name,
        "claude_code_project_skill_layout": Path("claude-project") / ".claude" / "skills" / skill_dir.name,
    }
    runs: dict[str, Any] = {}
    with tempfile.TemporaryDirectory(prefix="sale-preflight-portability-") as temp_root:
        root = Path(temp_root)
        for layout_name, relative in layouts.items():
            staged = root / relative
            shutil.copytree(
                skill_dir,
                staged,
                ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
            )
            output = root / "outputs" / layout_name
            runs[layout_name] = {
                "relative_entrypoint": (relative / "SKILL.md").as_posix(),
                "contract": validate_skill_contract(staged),
                "core_run": run_staged(staged, output),
            }

    invariant_sets = [item.get("core_run", {}).get("invariants") for item in runs.values()]
    parity = bool(invariant_sets) and all(item == invariant_sets[0] for item in invariant_sets)
    passed = contract["status"] == "passed" and parity and all(
        item["contract"]["status"] == "passed" and item["core_run"]["status"] == "passed"
        for item in runs.values()
    )
    report = {
        "schema_version": "1.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "status": "passed" if passed else "failed",
        "claim_scope": "开放 Agent Skills 包结构与确定性核心脚本在两种发现目录布局中的兼容性校验",
        "live_agent_run": False,
        "live_agent_note": "本报告不等同于第二个智能体产品的模型触发、界面或人工复跑证据。",
        "source_contract": contract,
        "layout_runs": runs,
        "invariant_parity": parity,
        "references": [
            "https://code.claude.com/docs/en/skills",
            "https://platform.claude.com/docs/es/agents-and-tools/agent-skills/best-practices",
        ],
    }
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
