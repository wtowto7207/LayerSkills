#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(process.argv[process.argv.indexOf("--root") + 1]);
const output = path.join(root, "submission", "律师盲测与验收记录.xlsx");
const previewDir = path.join(root, "tmp", "submission-workbook-previews");
await fs.mkdir(path.dirname(output), { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const colors = { navy: "#0B2545", blue: "#2E74B5", light: "#E8EEF5", gray: "#F2F4F7", border: "#D9DEE7", gold: "#FFF4CE", red: "#FCE8E6", green: "#E6F4EA" };
function title(sheet, address, value) {
  const range = sheet.getRange(address);
  range.merge();
  range.values = [[value]];
  range.format = { fill: colors.navy, font: { bold: true, color: "#FFFFFF", size: 17 }, verticalAlignment: "center" };
  range.format.rowHeight = 30;
}
function header(range) {
  range.format = { fill: colors.blue, font: { bold: true, color: "#FFFFFF", size: 10 }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "all", style: "thin", color: colors.border } };
  range.format.rowHeight = 34;
}
function body(range) {
  range.format = { font: { color: "#1F2937", size: 10 }, wrapText: true, verticalAlignment: "center", borders: { preset: "all", style: "thin", color: colors.border } };
}
function widths(sheet, map, lastRow) {
  for (const [column, width] of Object.entries(map)) sheet.getRange(`${column}1:${column}${lastRow}`).format.columnWidth = width;
}

const workbook = Workbook.create();
const records = workbook.worksheets.add("盲测记录");
records.showGridLines = false;
title(records, "A1:L1", "律师盲测与验收记录");
records.getRange("A2:L2").merge();
records.getRange("A2:L2").values = [["填写规则：三套合成案卷各做两轮；金额、来源、虚构事实和安全门禁为一票否决项。"]];
records.getRange("A2:L2").format = { fill: colors.light, font: { italic: true, color: colors.navy, size: 10 }, wrapText: true };
records.getRange("A4:L4").values = [["轮次", "案卷", "复核人/日期", "本金一致", "预设关键缺证数", "正确识别数", "缺证识别率", "虚构事实数", "来源覆盖率", "输出确定性结论", "安全异常通过", "结论"]];
header(records.getRange("A4:L4"));
const cases = [
  ["第一轮", "完整履约案", "", "", 1, "", null, "", "", "", "", null],
  ["第一轮", "无书面合同案", "", "", 1, "", null, "", "", "", "", null],
  ["第一轮", "冲突缺证案", "", "", 3, "", null, "", "", "", "", null],
  ["第二轮", "完整履约案", "", "", 1, "", null, "", "", "", "", null],
  ["第二轮", "无书面合同案", "", "", 1, "", null, "", "", "", "", null],
  ["第二轮", "冲突缺证案", "", "", 3, "", null, "", "", "", "", null],
];
records.getRange("A5:L10").values = cases;
for (let row = 5; row <= 10; row += 1) {
  records.getRange(`G${row}`).formulas = [[`=IF(F${row}="","",IFERROR(F${row}/E${row},0))`]];
  records.getRange(`L${row}`).formulas = [[`=IF(OR(D${row}="",F${row}="",H${row}="",I${row}="",J${row}="",K${row}=""),"待填写",IF(AND(D${row}="是",G${row}>=0.9,H${row}=0,I${row}=1,J${row}="否",K${row}="是"),"通过","待整改"))`]];
}
body(records.getRange("A5:L10"));
records.getRange("G5:G10").format.numberFormat = "0%";
records.getRange("I5:I10").format.numberFormat = "0%";
records.getRange("L5:L10").conditionalFormats.add("containsText", { text: "通过", format: { fill: colors.green, font: { color: "#137333", bold: true } } });
records.getRange("L5:L10").conditionalFormats.add("containsText", { text: "待整改", format: { fill: colors.red, font: { color: "#9B1C1C", bold: true } } });
records.getRange("A12:L12").merge();
records.getRange("A12:L12").values = [["复核签记：第一轮重点找漏项；修订后第二轮确认金额、来源与门禁。发现真实客户信息时立即停止并更换合成材料。"]];
records.getRange("A12:L12").format = { fill: colors.gold, font: { color: "#7A5A00", bold: true, size: 10 }, wrapText: true };
widths(records, { A: 12, B: 20, C: 24, D: 13, E: 16, F: 14, G: 15, H: 14, I: 15, J: 18, K: 17, L: 14 }, 12);
records.freezePanes.freezeRows(4);

const dashboard = workbook.worksheets.add("验收看板");
dashboard.showGridLines = false;
title(dashboard, "A1:F1", "参赛作品验收看板");
dashboard.getRange("A2:F2").merge();
dashboard.getRange("A2:F2").values = [["绿色为已由自动测试或结构校验确认；黄色为必须由申报团队在投稿前补齐。"]];
dashboard.getRange("A2:F2").format = { fill: colors.light, font: { italic: true, color: colors.navy, size: 10 } };
dashboard.getRange("A4:F4").values = [["指标", "目标", "当前值", "状态", "证据", "责任人"]];
header(dashboard.getRange("A4:F4"));
dashboard.getRange("A5:F16").values = [
  ["本金核算准确率", 1, 1, null, "3/3 合成案卷均为 100,000.00", "技术负责人"],
  ["来源定位覆盖率", 1, 1, null, "audit.source_traceability_complete=true", "技术负责人"],
  ["虚构事实数", 0, 0, null, "fabricated_fact_count=0", "技术负责人"],
  ["关键缺证识别率", 0.9, 1, null, "4/4 预设标题命中", "技术负责人"],
  ["Skill 结构校验", 1, 1, null, "Skill is valid", "技术负责人"],
  ["端到端与安全测试", 6, 6, null, "6 项：三案、来源、边界拒绝、注入/OCR、加密/不支持格式", "技术负责人"],
  ["Skill 目录布局兼容数", 2, 2, null, "Codex/Claude Code 发现目录下核心业务不变量一致", "技术负责人"],
  ["真实智能体环境数", 2, 1, null, "目录兼容报告不替代第二智能体模型/界面实跑", "技术负责人"],
  ["律师盲测轮数", 2, 0, null, "待填写“盲测记录”", "复核律师"],
  ["规则矩阵批准", 1, 0, null, "rules.json 仍为 pending", "承办律师"],
  ["盲测评审包", 1, 1, null, "blind-review-kit.zip 与答案包严格分离", "技术负责人"],
  ["演示视频时长（分钟）", 3, 4, null, "00:04:00；10 页；1080p；仅用合成材料", "技术负责人"],
];
for (let row = 5; row <= 16; row += 1) {
  const formula = row === 7
    ? `=IF(C${row}=B${row},"通过","待完成")`
    : row === 16
      ? `=IF(AND(C${row}>=3,C${row}<=5),"通过","待完成")`
      : `=IF(C${row}>=B${row},"通过","待完成")`;
  dashboard.getRange(`D${row}`).formulas = [[formula]];
}
body(dashboard.getRange("A5:F16"));
dashboard.getRange("B5:C5").format.numberFormat = "0%";
dashboard.getRange("B6:C6").format.numberFormat = "0%";
dashboard.getRange("B8:C8").format.numberFormat = "0%";
dashboard.getRange("D5:D16").conditionalFormats.add("containsText", { text: "通过", format: { fill: colors.green, font: { color: "#137333", bold: true } } });
dashboard.getRange("D5:D16").conditionalFormats.add("containsText", { text: "待完成", format: { fill: colors.gold, font: { color: "#7A5A00", bold: true } } });
dashboard.getRange("A18:F18").merge();
dashboard.getRange("A18:F18").values = [["投稿门禁：只有“真实智能体环境数、律师盲测轮数、规则矩阵批准”全部转为通过，才建议提交优秀入库评审。"]];
dashboard.getRange("A18:F18").format = { fill: colors.red, font: { color: "#9B1C1C", bold: true, size: 10 }, wrapText: true };
widths(dashboard, { A: 28, B: 14, C: 14, D: 15, E: 52, F: 18 }, 18);
dashboard.freezePanes.freezeRows(4);

const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, maxChars: 3000 });
if (errors.ndjson && /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(errors.ndjson)) throw new Error(errors.ndjson);
const blob = await SpreadsheetFile.exportXlsx(workbook);
await blob.save(output);
for (const sheetName of ["盲测记录", "验收看板"]) {
  const image = await workbook.render({ sheetName, autoCrop: "all", scale: 1.35, format: "png" });
  await fs.writeFile(path.join(previewDir, `${sheetName}.png`), new Uint8Array(await image.arrayBuffer()));
}
process.stdout.write(JSON.stringify({ status: "ok", output, previewDir }, null, 2));
