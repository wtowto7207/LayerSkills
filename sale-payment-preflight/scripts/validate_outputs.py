#!/usr/bin/env python3
"""Validate the fixed sale-payment-preflight output contract."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from decimal import Decimal
from pathlib import Path


REQUIRED = [
    "analysis.json",
    "audit.json",
    "01_案件材料清单.xlsx",
    "02_履约事实时间轴.xlsx",
    "03_要件证据矩阵.xlsx",
    "04_货款核对表.xlsx",
    "05_矛盾与补证清单.docx",
    "06_诉前案卷体检报告.docx",
]
FORBIDDEN_ASSERTIONS = ("胜诉率", "必然胜诉", "必然败诉", "法院一定支持", "已过诉讼时效")


def zip_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return "\n".join(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in archive.namelist()
            if name.endswith(".xml")
        )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    output_dir = args.output_dir.resolve()
    errors: list[str] = []

    for name in REQUIRED:
        path = output_dir / name
        if not path.is_file() or path.stat().st_size == 0:
            errors.append(f"缺少或为空：{name}")

    if errors:
        print(json.dumps({"status": "failed", "errors": errors}, ensure_ascii=False, indent=2))
        return 1

    analysis = json.loads((output_dir / "analysis.json").read_text(encoding="utf-8"))
    audit = json.loads((output_dir / "audit.json").read_text(encoding="utf-8"))
    if analysis.get("schema_version") != "1.0":
        errors.append("analysis schema_version 不是 1.0")
    if analysis.get("audit", {}).get("network_access_used") is not False:
        errors.append("联网审计状态异常")
    if analysis.get("audit", {}).get("originals_unchanged") is not True:
        errors.append("原件完整性检查未通过")
    if analysis.get("audit", {}).get("fabricated_fact_count") != 0:
        errors.append("虚构事实计数不为 0")
    if analysis.get("audit", {}).get("lawyer_review_status") != "pending":
        errors.append("律师签核前 review status 必须为 pending")
    if audit.get("audit", {}).get("source_traceability_complete") is not True:
        errors.append("事实溯源不完整")

    material_ids = {item["id"] for item in analysis.get("materials", [])}
    fact_ids = {item["id"] for item in analysis.get("facts", [])}
    for fact in analysis.get("facts", []):
        source = fact.get("source") or {}
        if source.get("material_id") not in material_ids or not source.get("location"):
            errors.append(f"事实 {fact.get('id')} 缺少有效来源")
    for item in analysis.get("evidence_matrix", []):
        unknown = set(item.get("evidence_fact_ids", [])) - fact_ids
        if unknown:
            errors.append(f"规则 {item.get('rule_id')} 引用了未知事实 {sorted(unknown)}")

    totals = analysis["ledger"]["totals"]
    calculated = Decimal(totals["应收"]) - Decimal(totals["付款"]) - Decimal(totals["退货"]) - Decimal(totals["折让"])
    if calculated != Decimal(analysis["ledger"]["outstanding_principal"]):
        errors.append("未付本金与分项合计不一致")

    for name in REQUIRED[2:]:
        path = output_dir / name
        try:
            with zipfile.ZipFile(path) as archive:
                if archive.testzip() is not None:
                    errors.append(f"压缩结构损坏：{name}")
        except zipfile.BadZipFile:
            errors.append(f"不是有效的 OOXML 文件：{name}")

    for name in ("05_矛盾与补证清单.docx", "06_诉前案卷体检报告.docx"):
        text = zip_text(output_dir / name)
        if "律师复核前草稿" not in text:
            errors.append(f"{name} 缺少律师复核声明")
        for phrase in FORBIDDEN_ASSERTIONS:
            if phrase in text:
                errors.append(f"{name} 出现禁止的确定性表述：{phrase}")

    for name in REQUIRED[2:6]:
        text = zip_text(output_dir / name)
        if "律师复核前草稿" not in text:
            errors.append(f"{name} 缺少律师复核声明")
        if name == "04_货款核对表.xlsx" and not re.search(r"<(?:\w+:)?f[^>]*>.*?(?:SUMIF|SUMIFS).*?</(?:\w+:)?f>", text, re.S):
            errors.append("货款核对表缺少公式驱动汇总")

    result = {
        "status": "passed" if not errors else "failed",
        "errors": errors,
        "checks": {
            "required_files": len(REQUIRED),
            "materials": len(analysis.get("materials", [])),
            "facts": len(analysis.get("facts", [])),
            "issues": len(analysis.get("issues", [])),
            "outstanding_principal": analysis["ledger"]["outstanding_principal"],
            "lawyer_review_status": analysis["audit"]["lawyer_review_status"],
        },
    }
    (output_dir / "validation.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
