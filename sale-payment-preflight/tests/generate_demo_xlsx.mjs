#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const rootArg = process.argv.indexOf("--root");
if (rootArg < 0 || rootArg + 1 >= process.argv.length) throw new Error("Missing --root");
const root = path.resolve(process.argv[rootArg + 1]);
const output = path.join(root, "01_complete_performance", "materials", "06_联系人核验.xlsx");

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("核验记录");
sheet.showGridLines = false;
sheet.getRange("A1:D1").merge();
sheet.getRange("A1:D1").values = [["联系人核验记录（合成演示）"]];
sheet.getRange("A1:D1").format = {
  fill: "#0B2545",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  verticalAlignment: "center",
};
sheet.getRange("A2:D2").merge();
sheet.getRange("A2:D2").values = [["全部名称与信息均属虚构；本表用于验证 XLSX 离线抽取能力。"]];
sheet.getRange("A2:D2").format = { fill: "#E8EEF5", font: { italic: true, color: "#0B2545", size: 10 } };
sheet.getRange("A4:D4").values = [["主体", "联系人", "角色", "核验状态"]];
sheet.getRange("A4:D4").format = { fill: "#2E74B5", font: { bold: true, color: "#FFFFFF", size: 10 } };
sheet.getRange("A5:D6").values = [
  ["浙江甲辰设备销售（演示）有限公司", "赵某", "销售经办", "待律师复核"],
  ["杭州乙木制造（演示）有限公司", "王某", "收货联系人", "待律师复核"],
];
sheet.getRange("A4:D6").format.wrapText = true;
for (const [column, width] of [["A", 38], ["B", 16], ["C", 20], ["D", 18]]) {
  sheet.getRange(`${column}1:${column}8`).format.columnWidth = width;
}
sheet.freezePanes.freezeRows(4);

await fs.mkdir(path.dirname(output), { recursive: true });
const blob = await SpreadsheetFile.exportXlsx(workbook);
await blob.save(output);
process.stdout.write(JSON.stringify({ status: "ok", output }, null, 2));
