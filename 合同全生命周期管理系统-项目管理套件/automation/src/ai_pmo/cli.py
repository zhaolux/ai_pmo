from __future__ import annotations

import argparse
import json
from datetime import date, datetime, timezone

from .builder import build, build_output_paths
from .config import get_paths, load_json
from .database import connect, initialize, seed_baseline
from .scanner import scan
from .validator import validate
from .orchestrator import run_agents
from .excel_export import agent_excel_path, build_agent_excel
from .weekly_report import build_weekly_reports, weekly_report_paths
from .decision_brief import build_decision_brief, decision_brief_paths
from .output_guard import ensure_outputs_available
from .finding_review import import_reviews, review_summary


def _log(command: str, status: str, message: str) -> None:
    paths = get_paths()
    paths.logs.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc).isoformat()
    with connect(paths.database) as connection:
        connection.execute(
            "INSERT INTO run_log(started_at, finished_at, command, status, message) VALUES(?,?,?,?,?)",
            (now, now, command, status, message),
        )
    with (paths.logs / "ai_pmo.jsonl").open("a", encoding="utf-8") as stream:
        stream.write(json.dumps({"time": now, "command": command, "status": status,
                                 "message": message}, ensure_ascii=False) + "\n")


def _preview_targets(
    command: str, paths, report_date: date, *, replace_generated: bool = False,
) -> None:
    if command in {"build", "run"}:
        targets = list(build_output_paths(paths).values())
    elif command == "agents":
        targets = [
            paths.outputs / "agents" / f"agent-report-{report_date:%Y%m%d}.json",
            agent_excel_path(paths.suite, report_date),
        ]
    elif command == "weekly-report":
        targets = list(weekly_report_paths(paths.suite, report_date))
    elif command == "decision-brief":
        targets = list(decision_brief_paths(paths.suite, report_date))
    else:
        raise ValueError(f"{command} 不生成文件，无需预览目标")
    print(f"{command} 目标清单（只读预览）：")
    for target in targets:
        status = "将新建"
        if target.exists():
            status = "已存在，显式允许替换" if replace_generated else "已存在，默认拒绝覆盖"
        print(f"- {status}：{target}")
    print("未创建或修改文件；正式执行时仍会重新检查目标。")


def execute(
    command: str, as_of: date | None = None, *,
    replace_generated: bool = False, preview_targets: bool = False,
    review_file=None, reviewer: str | None = None, run_id: str | None = None,
) -> int:
    paths = get_paths()
    project = load_json("project.json")
    try:
        if replace_generated and command not in {"agents", "weekly-report", "decision-brief"}:
            raise ValueError("--replace-generated 仅适用于 agents、weekly-report、decision-brief")
        if preview_targets:
            _preview_targets(
                command, paths, as_of or date.today(),
                replace_generated=replace_generated,
            )
            return 0
        if command in {"build", "run"}:
            ensure_outputs_available(build_output_paths(paths).values())
        if command == "agents":
            report_date = as_of or date.today()
            ensure_outputs_available(
                (
                    paths.outputs / "agents" / f"agent-report-{report_date:%Y%m%d}.json",
                    agent_excel_path(paths.suite, report_date),
                ),
                replace_generated=replace_generated,
            )
        if command in {"init", "seed", "run"}:
            initialize(paths.database, project)
            print(f"已初始化中央数据库：{paths.database}")
        if command in {"seed", "run"}:
            phases = load_json("phases.json")
            resources = load_json("resources.json")
            seed_baseline(paths.database, project, phases, resources)
            print(f"已写入 {len(phases)} 个阶段、{sum(x['headcount'] for x in resources)} 人的角色编制")
        if command in {"scan", "run"}:
            result = scan(paths)
            print(f"已扫描 {result['files']} 个文件、{result['sheets']} 个Sheet")
        if command in {"validate", "run"}:
            errors = validate(paths)
            if errors:
                print("校验发现以下事项：")
                for error in errors:
                    print(f"- {error}")
            else:
                print("数据与配置校验通过")
        if command in {"build", "run"}:
            outputs = build(paths)
            print(f"已生成 {len(outputs)} 个模块工作簿")
        if command == "agents":
            initialize(paths.database, project)
            report = run_agents(
                paths.suite, project["project_id"], as_of or date.today(),
                paths.outputs / "agents", paths.database,
                replace_generated=replace_generated,
            )
            print(f"已生成AI PMO Agent报告：{report}")
            workbook = build_agent_excel(
                report,
                agent_excel_path(paths.suite, as_of or date.today()),
                paths.automation / "scripts" / "build_agent_results_excel.mjs",
                replace_generated=replace_generated,
            )
            print(f"已更新AI PMO Agent结果工作簿：{workbook}")
        if command == "weekly-report":
            report_date = as_of or date.today()
            report_path = paths.outputs / "agents" / f"agent-report-{report_date:%Y%m%d}.json"
            if not report_path.exists():
                raise FileNotFoundError(f"请先运行 agents 命令生成报告：{report_path}")
            xlsx, docx = build_weekly_reports(
                report_path, project["project_name"], paths.suite, report_date,
                paths.automation / "scripts" / "build_weekly_report_excel.mjs",
                replace_generated=replace_generated,
            )
            print(f"已生成项目周报Excel：{xlsx}")
            print(f"已生成项目周报Word：{docx}")
        if command == "decision-brief":
            report_date = as_of or date.today()
            mapping = load_json("poc_decision_mapping.json")
            report_path = paths.outputs / "agents" / f"agent-report-{report_date:%Y%m%d}.json"
            agent_report = json.loads(report_path.read_text(encoding="utf-8")) if report_path.exists() else None
            json_file, docx_file = build_decision_brief(
                paths.suite, mapping, report_date, agent_report,
                replace_generated=replace_generated,
            )
            print(f"已生成POC决策事实包：{json_file}")
            print(f"已生成POC选型决策简报：{docx_file}")
        if command == "review-import":
            if review_file is None or not reviewer:
                raise ValueError("review-import需要--file和--reviewer")
            initialize(paths.database, project)
            result = import_reviews(paths.database, review_file, reviewer)
            print(f"已导入{result['imported']}条人工确认：运行ID {result['run_id']}")
        if command == "review-summary":
            if not run_id:
                raise ValueError("review-summary需要--run-id")
            initialize(paths.database, project)
            print(json.dumps(review_summary(paths.database, run_id), ensure_ascii=False))
        _log(command, "SUCCESS", "命令执行完成")
        return 0
    except FileExistsError as exc:
        print(f"执行已停止：{exc}")
        return 1
    except Exception as exc:
        try:
            initialize(paths.database, project)
            _log(command, "FAILED", str(exc))
        finally:
            print(f"执行失败：{exc}")
        return 1


def main() -> None:
    parser = argparse.ArgumentParser(description="能源行业合同管理AI PMO生成器")
    parser.add_argument("command", choices=["init", "seed", "scan", "validate", "build", "run", "agents", "weekly-report", "decision-brief", "review-import", "review-summary"])
    parser.add_argument("--date", type=date.fromisoformat, default=None, help="Agent分析的数据日期，格式YYYY-MM-DD")
    parser.add_argument("--replace-generated", action="store_true", help="显式替换同日期派生报告，不适用于业务工作簿")
    parser.add_argument("--preview-targets", action="store_true", help="只读列出生成文件及同名冲突，不执行命令")
    parser.add_argument("--file", type=__import__('pathlib').Path, default=None, help="填写了人工确认结果的Agent Excel")
    parser.add_argument("--reviewer", default=None, help="人工确认人")
    parser.add_argument("--run-id", default=None, help="要查询的Agent运行ID")
    args = parser.parse_args()
    raise SystemExit(execute(
        args.command, args.date, replace_generated=args.replace_generated,
        preview_targets=args.preview_targets,
        review_file=args.file, reviewer=args.reviewer, run_id=args.run_id,
    ))
