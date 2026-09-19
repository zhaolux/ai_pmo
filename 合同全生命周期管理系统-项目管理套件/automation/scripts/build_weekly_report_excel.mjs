import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const [payloadPath, outputPath] = process.argv.slice(2);
if (!payloadPath || !outputPath) throw new Error("用法: node build_weekly_report_excel.mjs <payload.json> <output.xlsx>");
const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));
const wb = Workbook.create();
const navy = "#1F4E78", light = "#D9EAF7", border = "#D9E1F2", amber = "#FFF2CC", text = "#1F2937";
const font = "Arial";
function base(name) { const s = wb.worksheets.add(name); s.showGridLines = false; return s; }
function title(s, value, subtitle, end = "H") {
  s.getRange("A2").values = [[value]]; s.mergeCells(`A2:${end}2`);
  s.getRange("A2").format.font = { name: font, size: 15, bold: true, color: navy };
  s.getRange("A3").values = [[subtitle]]; s.mergeCells(`A3:${end}3`);
  s.getRange("A3").format.font = { name: font, size: 10, italic: true, color: "#667085" };
}
function header(r) { r.format.fill = navy; r.format.font = { name: font, size: 10, bold: true, color: "#FFFFFF" }; r.format.horizontalAlignment = "center"; r.format.borders = { preset: "all", style: "thin", color: "#FFFFFF" }; }
function body(r) { r.format.font = { name: font, size: 10, color: text }; r.format.verticalAlignment = "center"; r.format.borders = { preset: "all", style: "thin", color: border }; }

const summary = base("周报摘要");
title(summary, "项目周报", `${payload.project_name}｜数据截止日 ${payload.data_date}`);
summary.getRange("A5:F5").values = [["项目状态", "发现总数", "重大", "高", "中", "需人工确认"]]; header(summary.getRange("A5:F5"));
summary.getRange("A6:F6").values = [[payload.status, payload.metrics.finding_count, payload.metrics.major_count, payload.metrics.high_count, payload.metrics.medium_count, payload.metrics.approval_count]]; body(summary.getRange("A6:F6"));
summary.getRange("A6:F6").format.font = { name: font, size: 12, bold: true, color: text }; summary.getRange("A6:F6").format.horizontalAlignment = "center";
summary.getRange("A6").format.fill = payload.status === "红" ? "#F4CCCC" : payload.status === "黄" ? amber : "#D9EAD3";
summary.getRange("A9:H9").values = [["总体判断", payload.conclusion, "", "", "", "", "", ""]]; summary.mergeCells("B9:H9"); body(summary.getRange("A9:H9")); summary.getRange("A9:H9").format.fill = light; summary.getRange("B9").format.wrapText = true;
summary.getRange("A12:E12").values = [["严重度", "重点事项", "对象ID", "责任人", "建议动作"]]; header(summary.getRange("A12:E12"));
const priorityRows = payload.priorities.map(x => [x.severity, x.title, x.object_id, x.owner, x.recommendation]);
if (priorityRows.length) { summary.getRange(`A13:E${12 + priorityRows.length}`).values = priorityRows; body(summary.getRange(`A13:E${12 + priorityRows.length}`)); summary.getRange(`B13:E${12 + priorityRows.length}`).format.wrapText = true; }
[14, 18, 18, 18, 52, 14, 14, 14].forEach((w, i) => summary.getRangeByIndexes(0, i, 1, 1).format.columnWidth = w);

const decisions = base("决策与支持");
title(decisions, "需决策与支持事项", "管理层按严重度和时限确认。", "D");
decisions.getRange("A5:D5").values = [["严重度", "事项", "对象ID", "决策或责任人"]]; header(decisions.getRange("A5:D5"));
const decisionRows = payload.decisions.map(x => [x.severity, x.title, x.object_id, x.owner]);
if (decisionRows.length) { decisions.getRange(`A6:D${5 + decisionRows.length}`).values = decisionRows; body(decisions.getRange(`A6:D${5 + decisionRows.length}`)); }
decisions.getRange("A:A").format.columnWidth = 14; decisions.getRange("B:B").format.columnWidth = 38; decisions.getRange("C:C").format.columnWidth = 24; decisions.getRange("D:D").format.columnWidth = 28;

const agents = base("Agent运行情况");
title(agents, "Agent运行情况", "各专业Agent的本周扫描结果。", "C");
agents.getRange("A5:C5").values = [["Agent", "发现数", "运行摘要"]]; header(agents.getRange("A5:C5"));
const agentRows = payload.agent_status.map(x => [x.agent, x.finding_count, x.summary]);
agents.getRange(`A6:C${5 + agentRows.length}`).values = agentRows; body(agents.getRange(`A6:C${5 + agentRows.length}`));
agents.getRange("A:A").format.columnWidth = 26; agents.getRange("B:B").format.columnWidth = 14; agents.getRange("C:C").format.columnWidth = 56;

const instructions = base("使用说明");
title(instructions, "项目周报使用说明", "本周报由AI PMO Agent已验证事实自动生成。", "B");
instructions.getRange("A5:B9").values = [["数据截止日", payload.data_date], ["更新命令", "automation/run.sh weekly-report --date YYYY-MM-DD"], ["数据来源", "AI PMO Agent JSON报告"], ["人工确认", "状态、基线、预算、风险等级和关闭结论须由授权人确认"], ["下周重点", payload.next_action]]; body(instructions.getRange("A5:B9")); instructions.getRange("A5:A9").format.fill = light; instructions.getRange("A5:A9").format.font = { name: font, size: 10, bold: true, color: navy }; instructions.getRange("B5:B9").format.wrapText = true; instructions.getRange("A:A").format.columnWidth = 18; instructions.getRange("B:B").format.columnWidth = 90;

wb.recalculate();
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
