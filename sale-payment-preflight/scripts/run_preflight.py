#!/usr/bin/env python3
"""Offline, read-only preflight for seller-side PRC B2B sale-payment files."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


DISCLAIMER = "律师复核前草稿：仅供案卷预整理，不构成法律意见，不得替代律师独立判断。"
SUPPORTED = {".pdf", ".docx", ".xlsx", ".csv", ".txt"}
FACT_KEYWORDS = {
    "party": ("当事人", "甲方", "乙方", "出卖人", "买受人", "卖方", "买方", "法定代表人", "授权"),
    "contract": ("买卖合同", "购销合同", "供货合同", "供货协议", "合同编号", "签订合同"),
    "order": ("订单", "采购单", "订购", "采购编号", "下单"),
    "delivery": ("送货单", "发货", "交付", "物流", "收货单", "出库", "签收"),
    "acceptance": ("验收", "收货确认", "签收", "检验合格", "验收合格"),
    "invoice": ("发票", "开票", "税票"),
    "reconciliation": ("对账", "欠款确认", "债权确认", "还款承诺", "余额确认", "确认欠款"),
    "payment": ("付款", "支付", "回款", "转账", "收款", "银行流水"),
    "demand": ("催款", "催告", "律师函", "要求支付", "承诺付款", "拒付", "逾期"),
    "jurisdiction": ("管辖", "人民法院", "仲裁委员会", "仲裁条款", "争议解决"),
    "limitation": ("诉讼时效", "催告", "对账", "还款承诺", "付款期限", "到期日", "承诺付款"),
    "quality": ("质量异议", "质量问题", "瑕疵", "退货", "换货", "维修", "折让", "质保"),
    "electronic": ("微信", "邮件", "电子邮件", "短信", "聊天记录", "电子数据", "电子订单"),
}
PII_PATTERNS = {
    "手机号": re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)"),
    "身份证号": re.compile(r"(?<!\d)\d{17}[\dXx](?!\w)"),
    "银行卡号": re.compile(r"(?<!\d)\d{16,19}(?!\d)"),
    "邮箱": re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"),
}
INJECTION_PATTERNS = (
    re.compile(r"忽略.{0,20}(指令|规则|提示)", re.I),
    re.compile(r"ignore.{0,30}(instruction|prompt|rule)", re.I),
    re.compile(r"(上传|发送).{0,30}(文件|案卷|材料|数据)", re.I),
    re.compile(r"(system prompt|系统提示词|开发者消息)", re.I),
)
DATE_PATTERNS = (
    re.compile(r"(?P<y>20\d{2})[年/-](?P<m>1[0-2]|0?[1-9])(?!\d)[月/-](?P<d>3[01]|[12]\d|0?[1-9])(?!\d)日?"),
    re.compile(r"(?P<y>20\d{2})\.(?P<m>1[0-2]|0?[1-9])(?!\d)\.(?P<d>3[01]|[12]\d|0?[1-9])(?!\d)"),
)
MONEY_PATTERN = re.compile(r"(?<!\d)(?:人民币|RMB|CNY|¥|￥)?\s*([0-9][0-9,]*(?:\.\d{1,2})?)\s*(万元|元)")


@dataclass
class TextUnit:
    location: str
    text: str


@dataclass
class Extracted:
    status: str
    units: list[TextUnit] = field(default_factory=list)
    tables: list[dict[str, Any]] = field(default_factory=list)
    note: str = ""
    page_count: int = 0
    sheet_count: int = 0


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")


def extract_txt(path: Path) -> Extracted:
    units = [TextUnit(f"line {index}", line.strip()) for index, line in enumerate(read_text(path).splitlines(), 1) if line.strip()]
    return Extracted("ok", units=units)


def extract_csv(path: Path) -> Extracted:
    text = read_text(path)
    rows = list(csv.reader(text.splitlines()))
    units = [TextUnit(f"row {index}", " | ".join(cell.strip() for cell in row)) for index, row in enumerate(rows, 1) if any(cell.strip() for cell in row)]
    return Extracted("ok", units=units, tables=[{"name": "CSV", "rows": rows}])


def _tag(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _all_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter() if _tag(node) == "t").strip()


def extract_docx(path: Path) -> Extracted:
    try:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        return Extracted("blocked_corrupt", note=f"DOCX 读取失败：{type(exc).__name__}")
    body = next((node for node in root.iter() if _tag(node) == "body"), root)
    units: list[TextUnit] = []
    paragraph_no = table_no = 0
    for child in list(body):
        if _tag(child) == "p":
            paragraph_no += 1
            text = _all_text(child)
            if text:
                units.append(TextUnit(f"paragraph {paragraph_no}", text))
        elif _tag(child) == "tbl":
            table_no += 1
            for row_no, row in enumerate((n for n in child.iter() if _tag(n) == "tr"), 1):
                cells = [_all_text(cell) for cell in list(row) if _tag(cell) == "tc"]
                if any(cells):
                    units.append(TextUnit(f"table {table_no} row {row_no}", " | ".join(cells)))
    return Extracted("ok", units=units)


def _xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [_all_text(item) for item in root if _tag(item) == "si"]


def _xlsx_sheet_targets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target_by_id = {rel.attrib.get("Id"): rel.attrib.get("Target", "") for rel in rels}
    sheets: list[tuple[str, str]] = []
    for sheet in workbook.iter():
        if _tag(sheet) != "sheet":
            continue
        rel_id = next((value for key, value in sheet.attrib.items() if key.endswith("}id") or key == "r:id"), None)
        target = target_by_id.get(rel_id, "")
        if not target:
            continue
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        sheets.append((sheet.attrib.get("name", "Sheet"), target))
    return sheets


def extract_xlsx(path: Path) -> Extracted:
    try:
        with zipfile.ZipFile(path) as archive:
            shared = _xlsx_shared_strings(archive)
            sheet_targets = _xlsx_sheet_targets(archive)
            units: list[TextUnit] = []
            tables: list[dict[str, Any]] = []
            for sheet_name, target in sheet_targets:
                root = ET.fromstring(archive.read(target))
                rows: list[list[str]] = []
                for row_no, row in enumerate((n for n in root.iter() if _tag(n) == "row"), 1):
                    values: list[str] = []
                    for cell in (n for n in list(row) if _tag(n) == "c"):
                        cell_type = cell.attrib.get("t")
                        value_node = next((n for n in cell if _tag(n) == "v"), None)
                        if cell_type == "inlineStr":
                            value = _all_text(cell)
                        elif value_node is None:
                            value = ""
                        else:
                            value = value_node.text or ""
                            if cell_type == "s" and value.isdigit() and int(value) < len(shared):
                                value = shared[int(value)]
                        values.append(value)
                    if any(value.strip() for value in values):
                        rows.append(values)
                        units.append(TextUnit(f"{sheet_name}!row {row_no}", " | ".join(values)))
                tables.append({"name": sheet_name, "rows": rows})
        return Extracted("ok", units=units, tables=tables, sheet_count=len(sheet_targets))
    except (zipfile.BadZipFile, KeyError, ET.ParseError, ValueError) as exc:
        return Extracted("blocked_corrupt", note=f"XLSX 读取失败：{type(exc).__name__}")


def extract_pdf(path: Path) -> Extracted:
    try:
        from pypdf import PdfReader
    except ImportError:
        return Extracted("blocked_dependency", note="缺少 pypdf，未读取 PDF")
    try:
        reader = PdfReader(str(path), strict=False)
        if reader.is_encrypted:
            try:
                if reader.decrypt("") == 0:
                    return Extracted("blocked_encrypted", note="PDF 已加密")
            except Exception:
                return Extracted("blocked_encrypted", note="PDF 已加密")
        units: list[TextUnit] = []
        for page_no, page in enumerate(reader.pages, 1):
            text = (page.extract_text() or "").strip()
            if text:
                for line_no, line in enumerate(text.splitlines(), 1):
                    if line.strip():
                        units.append(TextUnit(f"page {page_no} line {line_no}", line.strip()))
        status = "ok" if units else "needs_ocr"
        note = "未提取到文本，需 OCR 或人工核对" if not units else ""
        return Extracted(status, units=units, note=note, page_count=len(reader.pages))
    except Exception as exc:
        return Extracted("blocked_corrupt", note=f"PDF 读取失败：{type(exc).__name__}")


def extract(path: Path) -> Extracted:
    return {
        ".txt": extract_txt,
        ".csv": extract_csv,
        ".docx": extract_docx,
        ".xlsx": extract_xlsx,
        ".pdf": extract_pdf,
    }.get(path.suffix.lower(), lambda _: Extracted("unsupported", note="不支持的格式"))(path)


def mask_value(kind: str, value: str) -> str:
    if kind == "邮箱":
        name, _, domain = value.partition("@")
        return f"{name[:1]}***@{domain}" if domain else "[邮箱已掩码]"
    keep = 4 if len(value) >= 4 else 1
    return f"[{kind}***{value[-keep:]}]"


def redact(text: str, material_id: str, location: str, events: list[dict[str, Any]]) -> str:
    result = text
    for kind, pattern in PII_PATTERNS.items():
        def replace(match: re.Match[str]) -> str:
            value = match.group(0)
            masked = mask_value(kind, value)
            events.append({
                "type": kind,
                "masked": masked,
                "material_id": material_id,
                "location": location,
                "original_sha256": hashlib.sha256(value.encode("utf-8")).hexdigest(),
            })
            return masked
        result = pattern.sub(replace, result)
    return result


def pii_counts(texts: Iterable[str]) -> dict[str, int]:
    joined = "\n".join(texts)
    return {kind: len(pattern.findall(joined)) for kind, pattern in PII_PATTERNS.items()}


def has_injection(text: str) -> bool:
    return any(pattern.search(text) for pattern in INJECTION_PATTERNS)


def first_date(text: str) -> str | None:
    for pattern in DATE_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        try:
            return date(int(match.group("y")), int(match.group("m")), int(match.group("d"))).isoformat()
        except ValueError:
            return None
    return None


def money_values(text: str) -> list[Decimal]:
    values: list[Decimal] = []
    for raw, unit in MONEY_PATTERN.findall(text):
        try:
            value = Decimal(raw.replace(",", ""))
            if unit == "万元":
                value *= Decimal("10000")
            values.append(value.quantize(Decimal("0.01")))
        except InvalidOperation:
            pass
    return values


def split_fragments(text: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", text).strip()
    if not normalized:
        return []
    fragments = [part.strip() for part in re.split(r"(?<=[。；;！？!?])\s*", normalized) if part.strip()]
    return fragments or [normalized]


def decimal_text(value: Decimal) -> str:
    return format(value.quantize(Decimal("0.01")), "f")


def parse_decimal(value: str) -> Decimal | None:
    cleaned = re.sub(r"[^0-9.\-]", "", str(value).replace(",", ""))
    try:
        return Decimal(cleaned).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def normalized_header(value: str) -> str:
    return re.sub(r"\s+", "", value).strip()


def parse_ledger(material_id: str, tables: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    result: list[dict[str, Any]] = []
    warnings: list[str] = []
    required = {"日期", "类型", "金额"}
    for table in tables:
        rows = table.get("rows") or []
        if not rows:
            continue
        headers = [normalized_header(value) for value in rows[0]]
        if not required.issubset(set(headers)):
            continue
        index = {name: headers.index(name) for name in headers if name}
        for row_no, row in enumerate(rows[1:], 2):
            def get(name: str, default: str = "") -> str:
                position = index.get(name)
                return str(row[position]).strip() if position is not None and position < len(row) else default
            kind = get("类型")
            if kind not in {"应收", "付款", "退货", "折让"}:
                warnings.append(f"{material_id}/{table['name']} row {row_no}：未知流水类型“{kind}”")
                continue
            amount = parse_decimal(get("金额"))
            if amount is None or amount < 0:
                warnings.append(f"{material_id}/{table['name']} row {row_no}：金额无效")
                continue
            currency = get("币种", "CNY") or "CNY"
            confirmed = get("确认状态", "待核实") or "待核实"
            result.append({
                "date": get("日期"),
                "type": kind,
                "amount": decimal_text(amount),
                "currency": currency,
                "voucher_no": get("凭证号"),
                "counterparty": get("对方主体"),
                "confirmation_status": confirmed,
                "declared_material_id": get("材料编号"),
                "note": get("备注"),
                "source": {"material_id": material_id, "location": f"{table['name']} row {row_no}"},
            })
    return result, warnings


def load_rules(script_dir: Path) -> dict[str, Any]:
    return json.loads((script_dir.parent / "references" / "rules.json").read_text(encoding="utf-8"))


def write_csv(path: Path, headers: list[str], rows: list[list[Any]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(headers)
        writer.writerows(rows)


def build_analysis(case_dir: Path, config: dict[str, Any], output_dir: Path, rules_data: dict[str, Any]) -> dict[str, Any]:
    if config.get("authorized") is not True:
        raise ValueError("配置必须显式设置 authorized=true")
    if config.get("client_side") != "seller":
        raise ValueError("v1 仅支持卖方货款追索")
    if config.get("currency") != "CNY":
        raise ValueError("v1 仅支持单币种 CNY")
    try:
        as_of = date.fromisoformat(config["as_of_date"])
    except (KeyError, ValueError, TypeError) as exc:
        raise ValueError("as_of_date 必须为 YYYY-MM-DD") from exc
    for key in ("case_name", "client_name", "counterparty_name"):
        if not str(config.get(key, "")).strip():
            raise ValueError(f"缺少配置字段：{key}")

    output_resolved = output_dir.resolve()
    case_resolved = case_dir.resolve()
    if output_resolved == case_resolved or case_resolved in output_resolved.parents:
        raise ValueError("输出目录必须位于案件材料目录之外")

    files = sorted((path for path in case_dir.rglob("*") if path.is_file()), key=lambda p: p.relative_to(case_dir).as_posix().lower())
    if not files:
        raise ValueError("案件材料目录为空")

    materials: list[dict[str, Any]] = []
    extractions: dict[str, Extracted] = {}
    fingerprints_before: dict[str, tuple[int, int, str]] = {}
    redaction_events: list[dict[str, Any]] = []
    injection_materials: list[str] = []
    ledger_rows: list[dict[str, Any]] = []
    ledger_warnings: list[str] = []

    for index, path in enumerate(files, 1):
        material_id = f"M{index:03d}"
        relative = path.relative_to(case_dir).as_posix()
        stat = path.stat()
        digest = sha256_file(path)
        fingerprints_before[relative] = (stat.st_size, stat.st_mtime_ns, digest)
        extracted = extract(path) if path.suffix.lower() in SUPPORTED else Extracted("unsupported", note="不支持的格式")
        extractions[material_id] = extracted
        text_values = [unit.text for unit in extracted.units]
        injection = any(has_injection(text) for text in text_values)
        if injection:
            injection_materials.append(material_id)
        counts = pii_counts(text_values)
        materials.append({
            "id": material_id,
            "relative_path": relative,
            "extension": path.suffix.lower(),
            "size_bytes": stat.st_size,
            "sha256": digest,
            "status": extracted.status,
            "locations_extracted": len(extracted.units),
            "page_count": extracted.page_count,
            "sheet_count": extracted.sheet_count,
            "pii_counts": counts,
            "embedded_instruction_detected": injection,
            "note": extracted.note,
        })
        parsed_rows, warnings = parse_ledger(material_id, extracted.tables)
        ledger_rows.extend(parsed_rows)
        ledger_warnings.extend(warnings)

    facts: list[dict[str, Any]] = []
    seen_facts: set[tuple[str, str, str]] = set()
    client_name = config["client_name"]
    counterparty_name = config["counterparty_name"]
    for material in materials:
        material_id = material["id"]
        extracted = extractions[material_id]
        for unit in extracted.units:
            for fragment in split_fragments(unit.text):
                categories = [category for category, keywords in FACT_KEYWORDS.items() if any(keyword in fragment for keyword in keywords)]
                if client_name in fragment or counterparty_name in fragment:
                    categories.append("party")
                categories = sorted(set(categories))
                if not categories:
                    continue
                masked = redact(fragment[:500], material_id, unit.location, redaction_events)
                for category in categories:
                    key = (category, material_id, masked)
                    if key in seen_facts:
                        continue
                    seen_facts.add(key)
                    facts.append({
                        "id": f"F{len(facts) + 1:04d}",
                        "category": category,
                        "date": first_date(fragment),
                        "summary": masked[:240],
                        "money_values": [decimal_text(value) for value in money_values(fragment)],
                        "source": {"material_id": material_id, "location": unit.location},
                        "status": "已发现材料线索",
                    })

    if materials:
        facts.append({
            "id": f"F{len(facts) + 1:04d}",
            "category": "material_index",
            "date": None,
            "summary": f"已对 {len(materials)} 份材料建立只读清单、哈希和编号。",
            "money_values": [],
            "source": {"material_id": materials[0]["id"], "location": "file inventory"},
            "status": "已发现材料线索",
        })
    if ledger_rows:
        source = ledger_rows[0]["source"]
        facts.append({
            "id": f"F{len(facts) + 1:04d}",
            "category": "ledger",
            "date": None,
            "summary": f"发现 {len(ledger_rows)} 条结构化货款流水；仅计入已确认 CNY 行。",
            "money_values": [],
            "source": source,
            "status": "已发现材料线索",
        })

    confirmed = [row for row in ledger_rows if row["confirmation_status"] == "已确认" and row["currency"] == "CNY"]
    totals = {kind: sum((Decimal(row["amount"]) for row in confirmed if row["type"] == kind), Decimal("0")) for kind in ("应收", "付款", "退货", "折让")}
    outstanding = totals["应收"] - totals["付款"] - totals["退货"] - totals["折让"]
    currencies = sorted({row["currency"] for row in ledger_rows})
    ledger = {
        "currency": "CNY",
        "rows": ledger_rows,
        "totals": {kind: decimal_text(value) for kind, value in totals.items()},
        "outstanding_principal": decimal_text(outstanding),
        "included_rows": len(confirmed),
        "excluded_rows": len(ledger_rows) - len(confirmed),
        "status": "可复算" if confirmed else "待人工结构化/确认",
        "warnings": ledger_warnings,
    }

    by_category: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for fact in facts:
        by_category[fact["category"]].append(fact)

    evidence_matrix: list[dict[str, Any]] = []
    for rule in rules_data["rules"]:
        matched = []
        for category in rule.get("required_any", []):
            matched.extend(by_category.get(category, []))
        status = "已发现材料线索" if matched else "待补充/待核实"
        evidence_matrix.append({
            "rule_id": rule["id"],
            "category": rule["category"],
            "question": rule["question"],
            "status": status,
            "evidence_fact_ids": [fact["id"] for fact in matched[:12]],
            "lawyer_note": rule["lawyer_note"],
            "legal_source_ids": rule.get("legal_sources", []),
            "severity_if_missing": rule["severity_if_missing"],
        })

    issues: list[dict[str, Any]] = []
    issue_keys: set[tuple[str, tuple[str, ...]]] = set()

    def add_issue(severity: str, category: str, title: str, detail: str, sources: list[str], recommendation: str) -> None:
        key = (title, tuple(sorted(sources)))
        if key in issue_keys:
            return
        issue_keys.add(key)
        issues.append({
            "id": f"I{len(issues) + 1:03d}",
            "severity": severity,
            "category": category,
            "title": title,
            "detail": detail,
            "source_ids": sources,
            "recommendation": recommendation,
            "status": "待律师复核",
        })

    for material in materials:
        if material["status"] != "ok":
            severity = "high" if material["status"] in {"blocked_encrypted", "blocked_corrupt", "blocked_dependency"} else "medium"
            add_issue(severity, "材料可读性", f"{material['id']} 未完成文本抽取", material["note"] or material["status"], [material["id"]], "取得可验证原件、解密副本或经批准的 OCR/人工核对结果。")
    if injection_materials:
        add_issue("high", "安全", "材料含疑似提示词注入文本", "嵌入指令已按证据内容隔离，未执行。", injection_materials, "律师核实文本业务含义；继续保持离线和指令隔离。")
    hash_groups: dict[str, list[str]] = defaultdict(list)
    for material in materials:
        hash_groups[material["sha256"]].append(material["id"])
    for duplicate_ids in hash_groups.values():
        if len(duplicate_ids) > 1:
            add_issue("low", "材料管理", "发现内容完全相同的重复文件", "SHA-256 相同。", duplicate_ids, "核实是否重复提交并保留一份原件记录。")
    if not by_category["contract"]:
        if by_category["delivery"] or by_category["invoice"] or by_category["reconciliation"]:
            sources = sorted({fact["source"]["material_id"] for category in ("delivery", "invoice", "reconciliation") for fact in by_category[category]})
            add_issue("high", "合同关系", "未发现书面买卖合同", "存在履约或结算线索，但合同关系需结合交易方式、交易习惯和其他证据综合审查。", sources, "补充订单、报价、历史交易、沟通记录、送货与对账材料，由律师综合判断。")
        else:
            add_issue("high", "合同关系", "缺少合同关系核心材料", "未发现书面合同或主要履约材料线索。", [], "补充合同、订单、送货、收货、对账及往来材料。")
    if by_category["invoice"] and not by_category["delivery"]:
        sources = sorted({fact["source"]["material_id"] for fact in by_category["invoice"]})
        add_issue("high", "交付", "存在发票但未发现交付材料", "发票不能当然替代交付或欠款证明。", sources, "补充送货单、物流、签收、验收或买方确认材料。")
    if not by_category["reconciliation"]:
        add_issue("medium", "结算", "未发现对账或债权确认材料", "欠款金额和到期事实缺少独立确认线索。", [], "核实并补充对账单、余额确认、还款承诺或催款回复。")
    if not confirmed:
        add_issue("high", "金额", "没有可计入的已确认 CNY 流水", "未付本金无法由结构化已确认流水复算。", [], "按模板整理应收、付款、退货和折让，并由律师/客户逐行确认。")
    if any(currency != "CNY" for currency in currencies):
        sources = sorted({row["source"]["material_id"] for row in ledger_rows if row["currency"] != "CNY"})
        add_issue("high", "范围", "发现非 CNY 流水", "多币种事项超出 v1 范围，非 CNY 行未计入。", sources, "停止金额结论，转律师和财务人员处理汇率及支付问题。")
    for row in confirmed:
        cp = row["counterparty"]
        if cp and cp not in {client_name, counterparty_name}:
            add_issue("high", "付款主体", "发现第三方或不一致的收付款主体", f"流水对方主体为“{cp}”，未自动归因。", [row["source"]["material_id"]], "核实代付、债务加入、关联关系或其他法律基础。")
    invoice_amounts = {value for fact in by_category["invoice"] for value in fact["money_values"]}
    invoice_materials = {fact["source"]["material_id"] for fact in by_category["invoice"]}
    delivery_amounts = {
        value
        for fact in by_category["delivery"]
        if fact["source"]["material_id"] not in invoice_materials
        for value in fact["money_values"]
    }
    delivery_amounts.update(row["amount"] for row in confirmed if row["type"] == "应收")
    if invoice_amounts and delivery_amounts and invoice_amounts.isdisjoint(delivery_amounts):
        sources = sorted({fact["source"]["material_id"] for fact in by_category["invoice"] + by_category["delivery"]})
        add_issue("high", "金额矛盾", "发票与交付材料中的金额未匹配", f"发票金额线索 {sorted(invoice_amounts)}；交付金额线索 {sorted(delivery_amounts)}。", sources, "逐批核对订单、交付、退货、折让和发票对应关系。")
    for warning in ledger_warnings:
        add_issue("medium", "流水质量", "货款流水存在无法计入的行", warning, [], "修正类型、金额或确认状态后重新运行。")
    for item in evidence_matrix:
        if item["status"] == "待补充/待核实":
            add_issue(item["severity_if_missing"], item["category"], f"规则 {item['rule_id']} 缺少材料线索", item["question"], [], item["lawyer_note"])

    timeline = [
        {
            "date": fact["date"],
            "category": fact["category"],
            "event": fact["summary"],
            "material_id": fact["source"]["material_id"],
            "location": fact["source"]["location"],
            "fact_id": fact["id"],
        }
        for fact in facts if fact["date"]
    ]
    timeline.sort(key=lambda item: (item["date"], item["material_id"], item["fact_id"]))
    for event in timeline:
        if date.fromisoformat(event["date"]) > as_of:
            add_issue("high", "截止日", "发现晚于截止日期的材料事实", f"{event['date']} 晚于 {as_of.isoformat()}。", [event["material_id"]], "核实日期录入、材料形成时间和截止日。")

    severity_order = {"high": 0, "medium": 1, "low": 2}
    issues.sort(key=lambda item: (severity_order[item["severity"]], item["id"]))
    for index, issue in enumerate(issues, 1):
        issue["id"] = f"I{index:03d}"

    fingerprints_after = {}
    for path in files:
        relative = path.relative_to(case_dir).as_posix()
        stat = path.stat()
        fingerprints_after[relative] = (stat.st_size, stat.st_mtime_ns, sha256_file(path))
    originals_unchanged = fingerprints_before == fingerprints_after
    source_traceability = all(fact.get("source", {}).get("material_id") and fact.get("source", {}).get("location") for fact in facts)

    analysis = {
        "schema_version": "1.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "disclaimer": DISCLAIMER,
        "case": {
            "case_name": config["case_name"],
            "client_name": config["client_name"],
            "counterparty_name": config["counterparty_name"],
            "client_side": "seller",
            "as_of_date": as_of.isoformat(),
            "currency": "CNY",
            "questions": config.get("questions", []),
            "scope_status": "in_scope_with_lawyer_review",
        },
        "materials": materials,
        "facts": facts,
        "timeline": timeline,
        "ledger": ledger,
        "evidence_matrix": evidence_matrix,
        "issues": issues,
        "legal_sources": ["L01", "L02", "L03", "L04", "L05", "L06", "L07"],
        "redaction_events": redaction_events,
        "audit": {
            "authorized": True,
            "network_access_used": False,
            "originals_unchanged": originals_unchanged,
            "source_traceability_complete": source_traceability,
            "unsupported_or_blocked_files": sum(1 for material in materials if material["status"] != "ok"),
            "embedded_instruction_materials": injection_materials,
            "redaction_event_count": len(redaction_events),
            "fabricated_fact_count": 0,
            "lawyer_review_status": "pending",
            "rules_review_status": rules_data.get("review_status", "pending_prc_lawyer_approval"),
        },
    }
    return analysis


def write_outputs(analysis: dict[str, Any], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "analysis.json").write_text(json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8")
    audit_payload = {
        "schema_version": analysis["schema_version"],
        "generated_at": analysis["generated_at"],
        "disclaimer": analysis["disclaimer"],
        "case": analysis["case"],
        "audit": analysis["audit"],
        "redaction_events": analysis["redaction_events"],
        "material_hashes": [{"id": item["id"], "relative_path": item["relative_path"], "sha256": item["sha256"]} for item in analysis["materials"]],
    }
    (output_dir / "audit.json").write_text(json.dumps(audit_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(
        output_dir / "01_案件材料清单.csv",
        ["材料编号", "相对路径", "格式", "状态", "SHA-256", "位置数", "页数", "工作表数", "敏感信息计数", "备注"],
        [[m["id"], m["relative_path"], m["extension"], m["status"], m["sha256"], m["locations_extracted"], m["page_count"], m["sheet_count"], sum(m["pii_counts"].values()), m["note"]] for m in analysis["materials"]],
    )
    write_csv(
        output_dir / "02_履约事实时间轴.csv",
        ["日期", "类别", "事件摘要", "材料编号", "位置", "事实编号"],
        [[t["date"], t["category"], t["event"], t["material_id"], t["location"], t["fact_id"]] for t in analysis["timeline"]],
    )
    write_csv(
        output_dir / "03_要件证据矩阵.csv",
        ["规则编号", "类别", "审查问题", "状态", "事实编号", "律师复核提示", "法源编号"],
        [[e["rule_id"], e["category"], e["question"], e["status"], ", ".join(e["evidence_fact_ids"]), e["lawyer_note"], ", ".join(e["legal_source_ids"])] for e in analysis["evidence_matrix"]],
    )
    write_csv(
        output_dir / "04_货款核对表.csv",
        ["日期", "类型", "金额", "币种", "凭证号", "对方主体", "确认状态", "来源材料", "位置", "备注"],
        [[r["date"], r["type"], r["amount"], r["currency"], r["voucher_no"], r["counterparty"], r["confirmation_status"], r["source"]["material_id"], r["source"]["location"], r["note"]] for r in analysis["ledger"]["rows"]],
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case-dir", required=True, type=Path)
    parser.add_argument("--config", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    try:
        case_dir = args.case_dir.resolve(strict=True)
        if not case_dir.is_dir():
            raise ValueError("case-dir 必须是目录")
        config = json.loads(args.config.read_text(encoding="utf-8"))
        rules = load_rules(Path(__file__).resolve().parent)
        analysis = build_analysis(case_dir, config, args.output_dir, rules)
        write_outputs(analysis, args.output_dir)
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print(json.dumps({
        "status": "ok",
        "output_dir": str(args.output_dir.resolve()),
        "materials": len(analysis["materials"]),
        "facts": len(analysis["facts"]),
        "issues": len(analysis["issues"]),
        "outstanding_principal": analysis["ledger"]["outstanding_principal"],
        "disclaimer": DISCLAIMER,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
