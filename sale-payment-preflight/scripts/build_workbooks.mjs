#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const DISCLAIMER = "律师复核前草稿：仅供案卷预整理，不构成法律意见，不得替代律师独立判断。";
const COLORS = {
  navy: "#0B2545",
  blue: "#2E74B5",
  teal: "#0F766E",
  lightBlue: "#E8EEF5",
  lightGray: "#F2F4F7",
  border: "#D9DEE7",
  red: "#FCE8E6",
  gold: "#FFF4CE",
  muted: "#667085",
};

function arg(name) {
  const index = process.argv.indexOf(name);
  if (index < 0 || index + 1 >= process.argv.length) throw new Error(`Missing ${name}`);
  return process.argv[index + 1];
}

const analysisPath = path.resolve(arg("--analysis"));
const outputDir = path.resolve(arg("--output-dir"));
const previewDir = process.argv.includes("--preview-dir") ? path.resolve(arg("--preview-dir")) : null;
const analysis = JSON.parse(await fs.readFile(analysisPath, "utf8"));
await fs.mkdir(outputDir, { recursive: true });
if (previewDir) await fs.mkdir(previewDir, { recursive: true });

function baseWorkbook(sheetName, title, columns) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(sheetName);
  sheet.showGridLines = false;
  const last = String.fromCharCode(64 + columns);
  const titleRange = sheet.getRange(`A1:${last}1`);
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format = {
    fill: COLORS.navy,
    font: { bold: true, color: "#FFFFFF", size: 18 },
    verticalAlignment: "center",
  };
  titleRange.format.rowHeight = 32;
  const noteRange = sheet.getRange(`A2:${last}2`);
  noteRange.merge();
  noteRange.values = [[DISCLAIMER]];
  noteRange.format = {
    fill: COLORS.lightBlue,
    font: { color: COLORS.navy, italic: true, size: 10 },
    wrapText: true,
    verticalAlignment: "center",
  };
  noteRange.format.rowHeight = 28;
  return { workbook, sheet, last };
}

function styleHeader(range) {
  range.format = {
    fill: COLORS.blue,
    font: { bold: true, color: "#FFFFFF", size: 10 },
    wrapText: true,
    verticalAlignment: "center",
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
  range.format.rowHeight = 30;
}

function styleBody(range) {
  range.format = {
    font: { color: "#1F2937", size: 10 },
    wrapText: true,
    verticalAlignment: "center",
    borders: {
      insideHorizontal: { style: "thin", color: COLORS.border },
      bottom: { style: "thin", color: COLORS.border },
    },
  };
}

function setWidths(sheet, widths, maxRow) {
  widths.forEach((width, index) => {
    const col = String.fromCharCode(65 + index);
    sheet.getRange(`${col}1:${col}${maxRow}`).format.columnWidth = width;
  });
}

async function exportAndVerify(workbook, sheetName, filename, previewName, keyRange) {
  const table = await workbook.inspect({
    kind: "table",
    range: `${sheetName}!${keyRange}`,
    include: "values,formulas",
    tableMaxRows: 16,
    tableMaxCols: 12,
    maxChars: 5000,
  });
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: `formula errors in ${filename}`,
    maxChars: 2000,
  });
  if (errors.ndjson && /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(errors.ndjson)) {
    throw new Error(`Formula error detected in ${filename}: ${errors.ndjson}`);
  }
  const blob = await SpreadsheetFile.exportXlsx(workbook);
  await blob.save(path.join(outputDir, filename));
  if (previewDir) {
    const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1.5, format: "png" });
    await fs.writeFile(path.join(previewDir, previewName), new Uint8Array(await preview.arrayBuffer()));
  }
  return table.ndjson?.slice(0, 600) ?? "";
}

async function buildMaterials() {
  const { workbook, sheet } = baseWorkbook("材料清单", "01 案件材料清单", 10);
  sheet.getRange("A3:D3").values = [["材料总数", null, "未完整抽取", null]];
  sheet.getRange("B3").formulas = [["=COUNTA(A6:A205)"]];
  sheet.getRange("D3").formulas = [["=COUNTIFS(A6:A205,\"<>\",D6:D205,\"<>ok\")"]];
  sheet.getRange("A3:D3").format = { fill: COLORS.lightGray, font: { bold: true, color: COLORS.navy }, borders: { preset: "outside", style: "thin", color: COLORS.border } };
  const headers = ["材料编号", "相对路径", "格式", "状态", "SHA-256", "位置数", "页数", "工作表数", "敏感字段计数", "备注"];
  sheet.getRange("A5:J5").values = [headers];
  styleHeader(sheet.getRange("A5:J5"));
  const rows = analysis.materials.map((m) => [m.id, m.relative_path, m.extension, m.status, m.sha256, m.locations_extracted, m.page_count, m.sheet_count, Object.values(m.pii_counts).reduce((a, b) => a + b, 0), m.note || (m.embedded_instruction_detected ? "疑似嵌入指令已隔离" : "")]);
  if (rows.length) {
    sheet.getRange(`A6:J${5 + rows.length}`).values = rows;
    styleBody(sheet.getRange(`A6:J${5 + rows.length}`));
    sheet.getRange(`F6:I${5 + rows.length}`).format.horizontalAlignment = "center";
  }
  sheet.getRange("D6:D205").conditionalFormats.add("containsText", { text: "blocked", format: { fill: COLORS.red, font: { color: "#9B1C1C", bold: true } } });
  sheet.getRange("D6:D205").conditionalFormats.add("containsText", { text: "needs_ocr", format: { fill: COLORS.gold, font: { color: "#7A5A00", bold: true } } });
  setWidths(sheet, [11, 36, 9, 16, 30, 10, 8, 10, 13, 30], Math.max(205, 5 + rows.length));
  sheet.freezePanes.freezeRows(5);
  return exportAndVerify(workbook, "材料清单", "01_案件材料清单.xlsx", "01_案件材料清单.png", `A1:J${Math.max(6, 5 + rows.length)}`);
}

async function buildTimeline() {
  const { workbook, sheet } = baseWorkbook("履约时间轴", "02 履约事实时间轴", 6);
  sheet.getRange("A3:B3").values = [["有日期事实数", null]];
  sheet.getRange("B3").formulas = [["=COUNTA(A6:A205)"]];
  sheet.getRange("A3:B3").format = { fill: COLORS.lightGray, font: { bold: true, color: COLORS.navy }, borders: { preset: "outside", style: "thin", color: COLORS.border } };
  sheet.getRange("A5:F5").values = [["日期", "类别", "事件摘要", "材料编号", "具体位置", "事实编号"]];
  styleHeader(sheet.getRange("A5:F5"));
  const rows = analysis.timeline.map((t) => [new Date(`${t.date}T00:00:00`), t.category, t.event, t.material_id, t.location, t.fact_id]);
  if (rows.length) {
    sheet.getRange(`A6:F${5 + rows.length}`).values = rows;
    styleBody(sheet.getRange(`A6:F${5 + rows.length}`));
    sheet.getRange(`A6:A${5 + rows.length}`).format.numberFormat = "yyyy-mm-dd";
  }
  setWidths(sheet, [14, 16, 54, 12, 30, 11], Math.max(205, 5 + rows.length));
  sheet.freezePanes.freezeRows(5);
  return exportAndVerify(workbook, "履约时间轴", "02_履约事实时间轴.xlsx", "02_履约事实时间轴.png", `A1:F${Math.max(6, 5 + rows.length)}`);
}

async function buildMatrix() {
  const { workbook, sheet } = baseWorkbook("证据矩阵", "03 要件—证据矩阵", 7);
  sheet.getRange("A3:D3").values = [["规则数", null, "待补充/待核实", null]];
  sheet.getRange("B3").formulas = [["=COUNTA(A6:A105)"]];
  sheet.getRange("D3").formulas = [["=COUNTIF(D6:D105,\"待补充/待核实\")"]];
  sheet.getRange("A3:D3").format = { fill: COLORS.lightGray, font: { bold: true, color: COLORS.navy }, borders: { preset: "outside", style: "thin", color: COLORS.border } };
  sheet.getRange("A5:G5").values = [["规则编号", "类别", "审查问题", "状态", "事实编号", "律师复核提示", "法源编号"]];
  styleHeader(sheet.getRange("A5:G5"));
  const rows = analysis.evidence_matrix.map((e) => [e.rule_id, e.category, e.question, e.status, e.evidence_fact_ids.join(", "), e.lawyer_note, e.legal_source_ids.join(", ")]);
  if (rows.length) {
    sheet.getRange(`A6:G${5 + rows.length}`).values = rows;
    styleBody(sheet.getRange(`A6:G${5 + rows.length}`));
  }
  sheet.getRange("D6:D105").conditionalFormats.add("containsText", { text: "待补充", format: { fill: COLORS.gold, font: { color: "#7A5A00", bold: true } } });
  sheet.getRange("D6:D105").conditionalFormats.add("containsText", { text: "已发现", format: { fill: "#E6F4EA", font: { color: "#137333", bold: true } } });
  setWidths(sheet, [11, 16, 48, 18, 24, 52, 14], Math.max(105, 5 + rows.length));
  sheet.freezePanes.freezeRows(5);
  return exportAndVerify(workbook, "证据矩阵", "03_要件证据矩阵.xlsx", "03_要件证据矩阵.png", `A1:G${Math.max(6, 5 + rows.length)}`);
}

async function buildLedger() {
  const { workbook, sheet } = baseWorkbook("货款核对", "04 货款核对表", 10);
  sheet.getRange("A3:J3").values = [["已确认应收", null, "已确认付款", null, "退货+折让", null, "未付本金", null, "计入行数", null]];
  sheet.getRange("B3").formulas = [["=SUMIFS(C7:C206,B7:B206,\"应收\",F7:F206,\"已确认\",D7:D206,\"CNY\")"]];
  sheet.getRange("D3").formulas = [["=SUMIFS(C7:C206,B7:B206,\"付款\",F7:F206,\"已确认\",D7:D206,\"CNY\")"]];
  sheet.getRange("F3").formulas = [["=SUMIFS(C7:C206,B7:B206,\"退货\",F7:F206,\"已确认\",D7:D206,\"CNY\")+SUMIFS(C7:C206,B7:B206,\"折让\",F7:F206,\"已确认\",D7:D206,\"CNY\")"]];
  sheet.getRange("H3").formulas = [["=B3-D3-F3"]];
  sheet.getRange("J3").formulas = [["=COUNTIFS(F7:F206,\"已确认\",D7:D206,\"CNY\")"]];
  sheet.getRange("A3:J3").format = { fill: COLORS.lightGray, font: { bold: true, color: COLORS.navy }, borders: { preset: "outside", style: "thin", color: COLORS.border } };
  for (const cell of ["B3", "D3", "F3", "H3"]) {
    sheet.getRange(cell).format.numberFormat = "¥#,##0.00";
  }
  sheet.getRange("A5:J5").merge();
  sheet.getRange("A5:J5").values = [["口径：仅计入确认状态为“已确认”、币种为 CNY 的流水；未付本金 = 应收 - 付款 - 退货 - 折让。"]];
  sheet.getRange("A5:J5").format = { fill: COLORS.gold, font: { color: "#7A5A00", italic: true, size: 10 }, wrapText: true };
  sheet.getRange("A6:J6").values = [["日期", "类型", "金额", "币种", "凭证号", "确认状态", "对方主体", "来源材料", "具体位置", "备注"]];
  styleHeader(sheet.getRange("A6:J6"));
  const rows = analysis.ledger.rows.map((r) => [r.date ? new Date(`${r.date}T00:00:00`) : null, r.type, Number(r.amount), r.currency, r.voucher_no, r.confirmation_status, r.counterparty, r.source.material_id, r.source.location, r.note]);
  if (rows.length) {
    sheet.getRange(`A7:J${6 + rows.length}`).values = rows;
    styleBody(sheet.getRange(`A7:J${6 + rows.length}`));
    sheet.getRange(`A7:A${6 + rows.length}`).format.numberFormat = "yyyy-mm-dd";
    sheet.getRange(`C7:C${6 + rows.length}`).format.numberFormat = "¥#,##0.00";
  }
  sheet.getRange("F7:F206").conditionalFormats.add("containsText", { text: "待核实", format: { fill: COLORS.gold, font: { color: "#7A5A00", bold: true } } });
  setWidths(sheet, [14, 11, 16, 10, 18, 15, 32, 12, 26, 34], Math.max(206, 6 + rows.length));
  sheet.freezePanes.freezeRows(6);
  return exportAndVerify(workbook, "货款核对", "04_货款核对表.xlsx", "04_货款核对表.png", `A1:J${Math.max(7, 6 + rows.length)}`);
}

const summaries = [];
summaries.push(await buildMaterials());
summaries.push(await buildTimeline());
summaries.push(await buildMatrix());
summaries.push(await buildLedger());
process.stdout.write(JSON.stringify({ status: "ok", workbooks: 4, outputDir, previews: Boolean(previewDir), checks: summaries }, null, 2));
