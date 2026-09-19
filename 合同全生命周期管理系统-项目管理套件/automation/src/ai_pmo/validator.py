from __future__ import annotations

from .config import Paths, load_json
from .database import connect


def validate(paths: Paths) -> list[str]:
    project = load_json("project.json")
    errors: list[str] = []
    if project["hourly_rate"] != project["person_day_rate"] / project["hours_per_day"]:
        errors.append("小时单价与人天单价不一致")
    if project["hours_per_week"] != 40:
        errors.append("每人每周可用工时不是40小时")
    with connect(paths.database) as connection:
        role_total = connection.execute(
            "SELECT COALESCE(SUM(planned_headcount), 0) FROM resource_role"
        ).fetchone()[0]
        if role_total and role_total != project["planned_headcount"]:
            errors.append(
                f"角色编制合计{role_total}人，与计划总人数{project['planned_headcount']}人不一致"
            )
        duplicated = connection.execute(
            """SELECT person_id, week_start, SUM(planned_hours) AS hours
               FROM task_assignment GROUP BY person_id, week_start
               HAVING SUM(planned_hours) > ?""",
            (project["hours_per_week"],),
        ).fetchall()
        for row in duplicated:
            errors.append(f"资源超载：{row['person_id']} {row['week_start']} {row['hours']}小时")
        unmapped = connection.execute(
            "SELECT sheet_name FROM source_sheet WHERE module_code IS NULL"
        ).fetchall()
        for row in unmapped:
            errors.append(f"V2 Sheet尚未映射：{row['sheet_name']}")
    return errors
