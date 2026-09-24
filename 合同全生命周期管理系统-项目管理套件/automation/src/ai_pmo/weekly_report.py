from __future__ import annotations

import json
import os
import subprocess
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook
from .current_workbooks import resolve_current
from .output_guard import ensure_outputs_available


DEFAULT_NODE = Path(
    "/Users/zhaolu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"
)
DOC_FONT = "Arial Unicode MS"


def extract_cost_overview(suite: Path) -> dict | None:
    """从当前成本工作簿读取周报成本口径；缺簿或缺合计行时返回 None（周报不因此失败）。

    数字全部为格式化字符串，与 G15 材料口径一致（人天千分位、万元两位、单价两位）。
    """
    try:
        cost = resolve_current(suite, "cost")
        workbook = load_workbook(cost, data_only=True, read_only=True)
    except Exception:
        return None
    try:
        budget = workbook["人工成本预算"]
        total_row = None
        for row in budget.iter_rows(min_col=1, max_col=1):
            if row[0].value == "合计":
                total_row = row[0].row
                break
        if total_row is None:
            return None
        person_days = float(budget.cell(total_row, 4).value or 0)
        total_cost = float(budget.cell(total_row, 6).value or 0)
        if person_days <= 0:
            return None
        detail = workbook["月度投入明细"]
        actual = float(detail["AQ15"].value or 0)
        return {
            "person_days": person_days,
            "total_cost_wan": f"{total_cost / 10000:.2f}",
            "average_day_rate": f"{total_cost / person_days:,.2f}",
            "actual_person_days": actual,
            "completion_rate": (actual / person_days) if person_days else None,
        }
    except Exception:
        return None
    finally:
        workbook.close()


def _font(run, size: int | None = None, bold: bool | None = None) -> None:
    run.font.name = DOC_FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), DOC_FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def weekly_report_paths(suite: Path, as_of: date) -> tuple[Path, Path]:
    directory = suite / "08_沟通会议与报告" / "周报"
    stem = f"项目周报-{as_of:%Y%m%d}-V1"
    return directory / f"{stem}.xlsx", directory / f"{stem}.docx"


def weekly_report_payload(report: dict, project_name: str, cost_overview: dict | None = None) -> dict:
    summary = report["management_summary"]
    decisions = [
        {
            "severity": item["severity"], "title": item["title"],
            "object_id": item["object_id"], "owner": item["owner"],
            "recommendation": item.get("recommendation", ""),
        }
        for item in report.get("findings", [])
        if item.get("object_id", "").startswith("DEC-")
        or item.get("title") == "待决策事项"
    ]
    if not decisions:
        decisions = summary["decisions_needed"]
    decisions.sort(key=lambda item: (not item["object_id"].startswith("DEC-"), item["object_id"]))
    payload = {
        "project_name": project_name,
        "data_date": report["data_date"],
        "status": report["health"],
        "conclusion": (
            f"项目当前健康度为{report['health']}。"
            f"共识别{summary['finding_count']}项关注事项，其中"
            f"重大{summary['major_count']}项、高{summary['high_count']}项、"
            f"中{summary['medium_count']}项，{summary['approval_count']}项需人工确认。"
        ),
        "metrics": {
            "finding_count": summary["finding_count"],
            "major_count": summary["major_count"],
            "high_count": summary["high_count"],
            "medium_count": summary["medium_count"],
            "approval_count": summary["approval_count"],
        },
        "agent_status": report["agent_results"],
        "priorities": summary["top_priorities"],
        "decisions": decisions[:10],
        "next_action": summary["next_action"],
    }
    if cost_overview:
        payload["cost"] = {
            "person_days": f"{cost_overview['person_days']:,.0f}",
            "total_cost_wan": cost_overview["total_cost_wan"],
            "average_day_rate": cost_overview["average_day_rate"],
            "actual_person_days": f"{cost_overview['actual_person_days']:,.0f}",
            "completion_rate": (
                f"{cost_overview['completion_rate'] * 100:.1f}%"
                if cost_overview.get("completion_rate") is not None else ""
            ),
        }
    return payload


def _shade(cell, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:fill"), fill)
    properties.append(shading)


def _borders(cell, color: str = "D9D9D9") -> None:
    properties = cell._tc.get_or_add_tcPr()
    borders = properties.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        shading = properties.first_child_found_in("w:shd")
        if shading is None:
            properties.append(borders)
        else:
            properties.insert(properties.index(shading), borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:color"), color)
        borders.append(node)


def _table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float] | None = None):
    table = document.add_table(rows=1, cols=len(headers))
    table.autofit = False
    for index, value in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = str(value)
        _shade(cell, "1F4E78")
        _borders(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for run in cell.paragraphs[0].runs:
            _font(run, 9, True)
            run.font.color.rgb = RGBColor(255, 255, 255)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = "" if value is None else str(value)
            _borders(cells[index])
            if row_index % 2:
                _shade(cells[index], "F3F6FA")
            cells[index].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in cells[index].paragraphs:
                for run in paragraph.runs:
                    _font(run, 9)
    if widths:
        for row in table.rows:
            for index, width in enumerate(widths):
                row.cells[index].width = Cm(width)
    return table


def build_weekly_docx(payload: dict, path: Path) -> Path:
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    styles = document.styles
    styles["Normal"].font.name = DOC_FONT
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)
    styles["Normal"].font.size = Pt(10)
    for name, size in (("Title", 20), ("Heading 1", 14), ("Heading 2", 11)):
        styles[name].font.name = DOC_FONT
        styles[name]._element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)
        styles[name].font.size = Pt(size)
        styles[name].font.color.rgb = RGBColor(0, 0, 0)

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("项目周报")
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(f"{payload['project_name']}｜数据截止日 {payload['data_date']}")

    document.add_heading("一 项目状态判断", level=1)
    document.add_paragraph(payload["conclusion"])
    metrics = payload["metrics"]
    _table(document, ["健康度", "发现总数", "重大", "高", "中", "需人工确认"], [[
        payload["status"], metrics["finding_count"], metrics["major_count"], metrics["high_count"],
        metrics["medium_count"], metrics["approval_count"],
    ]], [2.2, 2.4, 1.8, 1.8, 1.8, 2.8])

    document.add_heading("二 专业Agent运行情况", level=1)
    _table(document, ["Agent", "发现数", "运行摘要"], [
        [item["agent"], item["finding_count"], item["summary"]] for item in payload["agent_status"]
    ], [4.2, 2.2, 10.0])

    document.add_heading("三 本周重点关注事项", level=1)
    _table(document, ["严重度", "关注事项", "对象ID", "责任人", "建议动作"], [
        [item["severity"], item["title"], item["object_id"], item["owner"], item["recommendation"]]
        for item in payload["priorities"]
    ], [1.6, 3.5, 2.5, 2.5, 6.5])

    document.add_heading("四 需决策与支持事项", level=1)
    _table(document, ["严重度", "事项", "对象ID", "决策或责任人"], [
        [item["severity"], item["title"], item["object_id"], item["owner"]]
        for item in payload["decisions"]
    ], [2.0, 6.0, 3.0, 5.5])

    if payload.get("cost"):
        cost = payload["cost"]
        document.add_heading("五 成本投入概况", level=1)
        document.add_paragraph(
            "成本口径取自当前成本工作簿《人工成本预算》与《月度投入明细》，"
            "与驾驶舱及 G15 会前材料同一口径。"
        )
        _table(document, ["计划人天", "人工成本(万元)", "平均人日单价(元)", "实际投入(人天)", "完成率"], [[
            cost["person_days"], cost["total_cost_wan"], cost["average_day_rate"],
            cost["actual_person_days"], cost["completion_rate"],
        ]], [2.4, 3.0, 3.4, 3.0, 2.2])

    document.add_heading("六 下周管理重点", level=1)
    document.add_paragraph(payload["next_action"])
    document.add_paragraph("重点完成CA、电子文档和BI报表产品POC证据汇总，为9月30日选型决策提供事实、方案、影响和建议。")
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    return path


def build_weekly_reports(
    report_path: Path, project_name: str, suite: Path, as_of: date, script_path: Path,
    *, replace_generated: bool = False,
) -> tuple[Path, Path]:
    xlsx_path, docx_path = weekly_report_paths(suite, as_of)
    ensure_outputs_available((xlsx_path, docx_path), replace_generated=replace_generated)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    payload = weekly_report_payload(report, project_name, extract_cost_overview(suite))
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    payload_path = xlsx_path.parent / f".weekly-report-{as_of:%Y%m%d}.json"
    payload_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    node = Path(os.environ.get("AI_PMO_NODE", str(DEFAULT_NODE)))
    try:
        subprocess.run([str(node), str(script_path), str(payload_path), str(xlsx_path)], check=True)
        build_weekly_docx(payload, docx_path)
    finally:
        payload_path.unlink(missing_ok=True)
    return xlsx_path, docx_path
