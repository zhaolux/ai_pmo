import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const [sourcePath, outputPath] = process.argv.slice(2);
if (!sourcePath || !outputPath) {
  throw new Error("用法: node build_agent_results_excel.mjs <report.json> <output.xlsx>");
}
const outputDir = outputPath.slice(0, outputPath.lastIndexOf("/"));
const report = JSON.parse(await fs.readFile(sourcePath, "utf8"));

const wb = Workbook.create();
const font = "Arial";
const navy = "#1F4E78";
const blue = "#5B9BD5";
const lightBlue = "#D9EAF7";
const border = "#D9E1F2";
const text = "#1F2937";
const amber = "#FFF2CC";

function baseSheet(name) {
  const sheet = wb.worksheets.add(name);
  sheet.showGridLines = false;
  return sheet;
}

function title(sheet, range, value, subtitle, endColumn = "H") {
  sheet.getRange(range).values = [[value]];
  sheet.getRange(range).format.font = { name: font, size: 15, bold: true, color: navy };
  const titleRow = Number(range.match(/\d+/)[0]);
  const row = titleRow + 1;
  sheet.mergeCells(`A${titleRow}:${endColumn}${titleRow}`);
  sheet.mergeCells(`A${row}:${endColumn}${row}`);
  sheet.getRange(`A${row}`).values = [[subtitle]];
  sheet.getRange(`A${row}`).format.font = { name: font, size: 10, italic: true, color: "#667085" };
}

function header(range) {
  range.format.fill = navy;
  range.format.font = { name: font, size: 10, bold: true, color: "#FFFFFF" };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.borders = { preset: "all", style: "thin", color: "#FFFFFF" };
}

function body(range) {
  range.format.font = { name: font, size: 10, color: text };
  range.format.verticalAlignment = "center";
  range.format.borders = { preset: "all", style: "thin", color: border };
}

const overview = baseSheet("Agent总览");
overview.tabColor = navy;
title(overview, "A2", "AI PMO Agent运行总览", `数据日期 ${report.data_date}｜运行ID ${report.run_id}`);
overview.getRange("A5:H5").values = [["项目健康度", "发现总数", "重大", "高", "中", "需人工审批", "待确认", "运行策略"]];
header(overview.getRange("A5:H5"));
overview.getRange("A6:H6").formulas = [[
  `="${report.health}"`,
  "=COUNTA('发现清单'!$A$5:$A$204)",
  "=COUNTIF('发现清单'!$C$5:$C$204,\"重大\")",
  "=COUNTIF('发现清单'!$C$5:$C$204,\"高\")",
  "=COUNTIF('发现清单'!$C$5:$C$204,\"中\")",
  "=COUNTIF('发现清单'!$I$5:$I$204,\"是\")",
  "=COUNTIF('发现清单'!$K$5:$K$204,\"待确认\")",
  "=\"只读建议\"",
]];
body(overview.getRange("A6:H6"));
overview.getRange("A6:H6").format.font = { name: font, size: 12, bold: true, color: text };
overview.getRange("A6:H6").format.horizontalAlignment = "center";
overview.getRange("A6").format.fill = report.health === "红" ? "#F4CCCC" : (report.health === "黄" ? amber : "#D9EAD3");

overview.getRange("A9:D9").values = [["Agent", "职责", "发现数量", "运行摘要"]];
header(overview.getRange("A9:D9"));
const agentLabels = {
  schedule_resource: ["进度资源Agent", "识别延期、近期到期和责任信息"],
  risk_issue: ["风险问题Agent", "识别风险、问题、变更和待决策事项"],
  cost_contract: ["成本合同Agent", "识别预算完整性、采购估算和付款证据风险"],
  quality_acceptance: ["质量验收Agent", "识别质量门禁、严重缺陷、交付物和UAT证据风险"],
  communication_report: ["沟通报告Agent", "识别逾期行动、待决策和升级事项"],
};
const summaryRows = report.agent_results.map(item => [
  (agentLabels[item.agent] || [item.agent, "识别项目管理关注事项"])[0],
  (agentLabels[item.agent] || [item.agent, "识别项目管理关注事项"])[1],
  item.finding_count,
  item.summary,
]);
overview.getRange(`A10:D${9 + summaryRows.length}`).values = summaryRows;
body(overview.getRange(`A10:D${9 + summaryRows.length}`));
const boundaryRow = 11 + summaryRows.length;
overview.getRange(`A${boundaryRow}:H${boundaryRow}`).values = [["人工确认边界", report.approval_boundary, "", "", "", "", "", ""]];
overview.mergeCells(`B${boundaryRow}:H${boundaryRow}`);
overview.getRange(`A${boundaryRow}:H${boundaryRow}`).format.fill = amber;
overview.getRange(`A${boundaryRow}:H${boundaryRow}`).format.font = { name: font, size: 10, bold: true, color: text };
overview.getRange(`A${boundaryRow}:H${boundaryRow}`).format.borders = { preset: "all", style: "thin", color: "#D6B656" };
overview.getRange(`B${boundaryRow}`).format.wrapText = true;
overview.getRange("A:A").format.columnWidth = 18;
overview.getRange("B:B").format.columnWidth = 22;
overview.getRange("C:G").format.columnWidth = 14;
overview.getRange("H:H").format.columnWidth = 28;
overview.getRange("D:D").format.columnWidth = 38;
overview.getRange("5:6").format.rowHeight = 28;
overview.getRange(`${boundaryRow}:${boundaryRow}`).format.rowHeight = 36;

const management = baseSheet("管理层摘要");
management.tabColor = "#70AD47";
title(management, "A2", "AI PMO管理层摘要", `数据日期 ${report.data_date}｜仅基于已验证的台账事实`);
management.getRange("A5:F5").values = [["状态判断", "发现总数", "重大", "高", "中", "需审批"]];
header(management.getRange("A5:F5"));
management.getRange("A6:F6").values = [[
  report.management_summary.status_judgment,
  report.management_summary.finding_count,
  report.management_summary.major_count,
  report.management_summary.high_count,
  report.management_summary.medium_count,
  report.management_summary.approval_count,
]];
body(management.getRange("A6:F6"));
management.getRange("A6:F6").format.font = { name: font, size: 12, bold: true, color: text };
management.getRange("A6:F6").format.horizontalAlignment = "center";
management.getRange("A6").format.fill = report.health === "红" ? "#F4CCCC" : (report.health === "黄" ? amber : "#D9EAD3");
management.getRange("A9:E9").values = [["严重度", "关注事项", "对象ID", "责任人", "建议动作"]];
header(management.getRange("A9:E9"));
const priorityRows = report.management_summary.top_priorities.map(item => [item.severity, item.title, item.object_id, item.owner, item.recommendation]);
if (priorityRows.length) {
  management.getRange(`A10:E${9 + priorityRows.length}`).values = priorityRows;
  body(management.getRange(`A10:E${9 + priorityRows.length}`));
  management.getRange(`B10:E${9 + priorityRows.length}`).format.wrapText = true;
}
const actionRow = 12 + Math.max(priorityRows.length, 1);
management.getRange(`A${actionRow}:E${actionRow}`).values = [["建议次序", report.management_summary.next_action, "", "", ""]];
management.mergeCells(`B${actionRow}:E${actionRow}`);
management.getRange(`A${actionRow}:E${actionRow}`).format.fill = amber;
management.getRange(`A${actionRow}:E${actionRow}`).format.borders = { preset: "all", style: "thin", color: "#D6B656" };
management.getRange(`A${actionRow}:E${actionRow}`).format.font = { name: font, size: 10, bold: true, color: text };
management.getRange("A:A").format.columnWidth = 14;
management.getRange("B:B").format.columnWidth = 28;
management.getRange("C:C").format.columnWidth = 20;
management.getRange("D:D").format.columnWidth = 18;
management.getRange("E:E").format.columnWidth = 58;
management.getRange("F:F").format.columnWidth = 14;

const findings = baseSheet("发现清单");
findings.tabColor = blue;
title(findings, "A2", "Agent发现清单", "黄色列由项目经理或授权人确认；源项目台账不会被本文件自动修改。", "L");
const findingHeaders = ["发现ID", "Agent", "严重度", "对象ID", "标题", "详细说明", "建议动作", "责任人", "需审批", "数据日期", "确认状态", "人工结论"];
findings.getRange("A4:L4").values = [findingHeaders];
header(findings.getRange("A4:L4"));
const findingRows = report.findings.map(item => [
  item.finding_id,
  (agentLabels[item.agent] || [item.agent])[0],
  item.severity,
  item.object_id,
  item.title,
  item.detail,
  item.recommendation,
  item.owner,
  item.requires_approval ? "是" : "否",
  new Date(`${report.data_date}T00:00:00`),
  "待确认",
  "",
]);
const lastFindingRow = 4 + findingRows.length;
findings.getRange(`A5:L${lastFindingRow}`).values = findingRows;
body(findings.getRange(`A5:L${lastFindingRow}`));
findings.getRange(`F5:G${lastFindingRow}`).format.wrapText = true;
findings.getRange(`J5:J${lastFindingRow}`).setNumberFormat("yyyy-mm-dd");
findings.getRange(`K5:L${lastFindingRow}`).format.fill = amber;
findings.getRange(`K5:K${lastFindingRow}`).dataValidation = { rule: { type: "list", values: ["待确认", "已接受", "需调整", "已关闭"] } };
for (let row = 5; row <= lastFindingRow; row += 1) {
  const severity = findings.getRange(`C${row}`);
  const value = findingRows[row - 5][2];
  if (value === "重大") severity.format.fill = "#F4CCCC";
  if (value === "高") severity.format.fill = "#FCE4D6";
  if (value === "中") severity.format.fill = amber;
}
findings.freezePanes.freezeRows(4);
findings.freezePanes.freezeColumns(4);
const widths = [22, 16, 10, 16, 22, 46, 44, 16, 10, 13, 13, 36];
widths.forEach((width, i) => findings.getRangeByIndexes(0, i, 1, 1).format.columnWidth = width);
findings.getRange(`5:${lastFindingRow}`).format.rowHeight = 42;

const evidence = baseSheet("证据明细");
title(evidence, "A2", "Agent证据明细", "每项发现对应到来源文件、工作表、记录ID和源行号。请依据证据确认AI建议。", "E");
evidence.getRange("A4:E4").values = [["发现ID", "来源文件", "工作表", "记录ID", "源行号"]];
header(evidence.getRange("A4:E4"));
const evidenceRows = [];
for (const item of report.findings) {
  for (const source of item.evidence) {
    evidenceRows.push([item.finding_id, source.file, source.sheet, source.record_id, source.row ?? ""]);
  }
}
const lastEvidenceRow = 4 + evidenceRows.length;
evidence.getRange(`A5:E${lastEvidenceRow}`).values = evidenceRows;
body(evidence.getRange(`A5:E${lastEvidenceRow}`));
evidence.freezePanes.freezeRows(4);
evidence.getRange("A:A").format.columnWidth = 24;
evidence.getRange("B:B").format.columnWidth = 38;
evidence.getRange("C:C").format.columnWidth = 22;
evidence.getRange("D:D").format.columnWidth = 18;
evidence.getRange("E:E").format.columnWidth = 12;

const instructions = baseSheet("使用说明");
instructions.tabColor = "#A6A6A6";
title(instructions, "A2", "Agent运行结果使用说明", "本工作簿是Agent分析快照；确认结果应回写源项目台账后重新运行Agent。 ", "B");
instructions.getRange("A5:B12").values = [
  ["项目", "能源行业合同全生命周期管理系统"],
  ["数据日期", report.data_date],
  ["来源报告", sourcePath],
  ["运行命令", "automation/run.sh agents --date YYYY-MM-DD"],
  ["更新方式", "先更新源Excel，再运行Agent；命令会自动刷新JSON、数据库审计和本工作簿"],
  ["黄色字段", "由项目经理或授权人填写确认状态和人工结论"],
  ["审批边界", "Agent只提出建议，不直接修改基线、预算、状态、风险等级或关闭结论"],
  ["证据要求", "正式结论须能追溯到来源文件、工作表、记录ID和源行号"],
];
body(instructions.getRange("A5:B12"));
instructions.getRange("A5:A12").format.fill = lightBlue;
instructions.getRange("A5:A12").format.font = { name: font, size: 10, bold: true, color: navy };
instructions.getRange("B5:B12").format.wrapText = true;
instructions.getRange("A:A").format.columnWidth = 18;
instructions.getRange("B:B").format.columnWidth = 90;
instructions.getRange("5:12").format.rowHeight = 30;

wb.recalculate();
await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
