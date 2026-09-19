from __future__ import annotations

import json
from datetime import date

from .config import Paths, load_json
from .database import connect


PLANNING_SCHEMA = """
CREATE TABLE IF NOT EXISTS phase_resource_plan (
  phase_id TEXT NOT NULL REFERENCES project_phase(phase_id),
  role_code TEXT NOT NULL REFERENCES resource_role(role_code),
  planned_headcount INTEGER NOT NULL CHECK(planned_headcount >= 0),
  planning_status TEXT NOT NULL DEFAULT '规划假设',
  PRIMARY KEY (phase_id, role_code)
);
CREATE TABLE IF NOT EXISTS wbs_item (
  wbs_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  phase_id TEXT NOT NULL REFERENCES project_phase(phase_id),
  task_name TEXT NOT NULL,
  owner_role TEXT NOT NULL REFERENCES resource_role(role_code),
  planned_start TEXT NOT NULL,
  planned_finish TEXT NOT NULL,
  dependency TEXT,
  deliverable TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT '未开始'
);
CREATE TABLE IF NOT EXISTS baseline_milestone (
  milestone_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  phase_id TEXT NOT NULL REFERENCES project_phase(phase_id),
  milestone_name TEXT NOT NULL,
  planned_date TEXT NOT NULL,
  planned_start TEXT,
  planned_finish TEXT,
  actual_start TEXT,
  actual_finish TEXT,
  deliverable TEXT NOT NULL,
  exit_criteria TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT '未开始'
);
"""


def seed_planning(paths: Paths) -> dict[str, int]:
    project = load_json("project.json")
    resources = load_json("phase_resources.json")
    wbs = load_json("wbs.json")
    milestones = load_json("milestones.json")
    with connect(paths.database) as connection:
        connection.executescript(PLANNING_SCHEMA)
        existing = {row[1] for row in connection.execute("PRAGMA table_info(baseline_milestone)")}
        for column in ("planned_start", "planned_finish", "actual_start", "actual_finish"):
            if column not in existing:
                connection.execute(f"ALTER TABLE baseline_milestone ADD COLUMN {column} TEXT")
        connection.execute("DELETE FROM phase_resource_plan")
        connection.execute("DELETE FROM wbs_item WHERE project_id=?", (project["project_id"],))
        connection.execute("DELETE FROM baseline_milestone WHERE project_id=?", (project["project_id"],))
        for item in resources:
            connection.execute(
                "INSERT INTO phase_resource_plan(phase_id,role_code,planned_headcount) VALUES(?,?,?)",
                (item["phase_id"], item["role_code"], item["headcount"]),
            )
        for item in wbs:
            connection.execute(
                """INSERT INTO wbs_item
                   (wbs_id,project_id,phase_id,task_name,owner_role,planned_start,
                    planned_finish,dependency,deliverable,status)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (item["wbs_id"], project["project_id"], item["phase_id"],
                 item["task_name"], item["owner_role"], item["planned_start"],
                 item["planned_finish"], item["dependency"], item["deliverable"],
                 "计划完成日期已到，待确认" if item["planned_finish"] < "2026-09-16" else
                 "进行中（待确认）" if item["planned_start"] <= "2026-09-16" <= item["planned_finish"] else
                 "未开始"),
            )
        for item in milestones:
            connection.execute(
                """INSERT INTO baseline_milestone
                   (milestone_id,project_id,phase_id,milestone_name,planned_date,
                    planned_start,planned_finish,actual_start,actual_finish,
                    deliverable,exit_criteria,status)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (item["milestone_id"], project["project_id"], item["phase_id"],
                 item["milestone_name"], item["planned_finish"], item["planned_start"],
                 item["planned_finish"], None, None, item["deliverable"],
                 item["exit_criteria"],
                 "计划日期已到，待确认" if item["planned_finish"] <= "2026-09-16" else "未开始"),
            )
    return {"phase_resources": len(resources), "wbs": len(wbs), "milestones": len(milestones)}


def validate_planning(paths: Paths) -> list[str]:
    project = load_json("project.json")
    errors: list[str] = []
    with connect(paths.database) as connection:
        phase_rows = connection.execute(
            "SELECT phase_id, planned_start, planned_finish FROM project_phase ORDER BY planned_start"
        ).fetchall()
        for row in phase_rows:
            if row["planned_finish"] and row["planned_start"] > row["planned_finish"]:
                errors.append(f"阶段日期倒置：{row['phase_id']}")
        pool = {row["role_code"]: row["planned_headcount"] for row in
                connection.execute("SELECT role_code, planned_headcount FROM resource_role")}
        for row in connection.execute(
            "SELECT phase_id,role_code,planned_headcount FROM phase_resource_plan"
        ):
            if row["planned_headcount"] > pool[row["role_code"]]:
                errors.append(f"阶段资源超过编制：{row['phase_id']} {row['role_code']}")
        peaks = connection.execute(
            """SELECT phase_id,SUM(planned_headcount) AS people
               FROM phase_resource_plan GROUP BY phase_id ORDER BY phase_id"""
        ).fetchall()
        if max((row["people"] for row in peaks), default=0) != project["planned_headcount"]:
            errors.append("资源投入峰值没有达到计划30人")
        phase_bounds = {row["phase_id"]: (row["planned_start"], row["planned_finish"])
                        for row in phase_rows}
        for row in connection.execute(
            "SELECT wbs_id,phase_id,planned_start,planned_finish,dependency FROM wbs_item"
        ):
            start, finish = phase_bounds[row["phase_id"]]
            if row["planned_start"] < start or (finish and row["planned_finish"] > finish):
                errors.append(f"WBS超出所属阶段：{row['wbs_id']}")
            if row["planned_start"] > row["planned_finish"]:
                errors.append(f"WBS日期倒置：{row['wbs_id']}")
        for row in connection.execute(
            "SELECT milestone_id,phase_id,planned_start,planned_finish FROM baseline_milestone"
        ):
            start, finish = phase_bounds[row["phase_id"]]
            if row["planned_start"] < start or (finish and row["planned_finish"] > finish):
                errors.append(f"里程碑超出所属阶段：{row['milestone_id']}")
            if row["planned_start"] > row["planned_finish"]:
                errors.append(f"里程碑日期倒置：{row['milestone_id']}")
    return errors


def export_planning_json(paths: Paths) -> None:
    target = paths.data_center / "exports"
    target.mkdir(parents=True, exist_ok=True)
    with connect(paths.database) as connection:
        for table in ("project_phase", "resource_role", "phase_resource_plan",
                      "wbs_item", "baseline_milestone"):
            rows = [dict(row) for row in connection.execute(f"SELECT * FROM {table}")]
            (target / f"{table}.json").write_text(
                json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8"
            )
