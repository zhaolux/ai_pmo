from __future__ import annotations

import sqlite3
from pathlib import Path


SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS project (
  project_id TEXT PRIMARY KEY,
  project_name TEXT NOT NULL,
  industry TEXT NOT NULL,
  as_of_date TEXT
);
CREATE TABLE IF NOT EXISTS project_baseline (
  project_id TEXT PRIMARY KEY REFERENCES project(project_id),
  preparation_start TEXT NOT NULL,
  formal_start_date TEXT NOT NULL,
  planned_go_live_date TEXT NOT NULL,
  trial_operation_start TEXT NOT NULL,
  planned_headcount INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS project_phase (
  phase_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  phase_name TEXT NOT NULL,
  planned_start TEXT NOT NULL,
  planned_finish TEXT,
  phase_type TEXT NOT NULL,
  primary_deliverables TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS resource_role (
  role_code TEXT PRIMARY KEY,
  role_name TEXT NOT NULL,
  planned_headcount INTEGER NOT NULL CHECK(planned_headcount >= 0),
  hours_per_week REAL NOT NULL,
  hourly_rate REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS source_file (
  source_id INTEGER PRIMARY KEY AUTOINCREMENT,
  relative_path TEXT NOT NULL UNIQUE,
  file_type TEXT NOT NULL,
  size_bytes INTEGER NOT NULL,
  modified_at TEXT NOT NULL,
  sha256 TEXT NOT NULL,
  scanned_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS source_sheet (
  source_id INTEGER NOT NULL REFERENCES source_file(source_id) ON DELETE CASCADE,
  sheet_name TEXT NOT NULL,
  sheet_order INTEGER NOT NULL,
  max_row INTEGER,
  max_column INTEGER,
  module_code TEXT,
  PRIMARY KEY (source_id, sheet_name)
);
CREATE TABLE IF NOT EXISTS person (
  person_id TEXT PRIMARY KEY,
  person_name TEXT NOT NULL,
  active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS task (
  task_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  task_name TEXT NOT NULL,
  status TEXT NOT NULL,
  planned_start TEXT,
  planned_finish TEXT,
  progress REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS task_assignment (
  task_id TEXT NOT NULL REFERENCES task(task_id),
  person_id TEXT NOT NULL REFERENCES person(person_id),
  week_start TEXT NOT NULL,
  planned_hours REAL NOT NULL,
  PRIMARY KEY (task_id, person_id, week_start)
);
CREATE TABLE IF NOT EXISTS milestone (
  milestone_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  milestone_name TEXT NOT NULL,
  planned_date TEXT,
  status TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS requirement (
  requirement_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  requirement_name TEXT NOT NULL,
  priority TEXT,
  status TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS risk (
  risk_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  risk_name TEXT NOT NULL,
  probability INTEGER,
  impact INTEGER,
  owner_person_id TEXT REFERENCES person(person_id),
  status TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS issue (
  issue_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  issue_name TEXT NOT NULL,
  status TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS change_request (
  change_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  change_name TEXT NOT NULL,
  status TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS deliverable (
  deliverable_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  deliverable_name TEXT NOT NULL,
  status TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS run_log (
  run_id INTEGER PRIMARY KEY AUTOINCREMENT,
  started_at TEXT NOT NULL,
  finished_at TEXT,
  command TEXT NOT NULL,
  status TEXT NOT NULL,
  message TEXT
);
CREATE TABLE IF NOT EXISTS agent_run (
  run_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES project(project_id),
  data_date TEXT NOT NULL,
  generated_at TEXT NOT NULL,
  health TEXT NOT NULL,
  write_policy TEXT NOT NULL,
  report_path TEXT
);
CREATE TABLE IF NOT EXISTS agent_finding (
  finding_id TEXT NOT NULL,
  run_id TEXT NOT NULL REFERENCES agent_run(run_id) ON DELETE CASCADE,
  agent TEXT NOT NULL,
  object_id TEXT NOT NULL,
  severity TEXT NOT NULL,
  title TEXT NOT NULL,
  detail TEXT NOT NULL,
  recommendation TEXT NOT NULL,
  owner TEXT NOT NULL,
  requires_approval INTEGER NOT NULL,
  evidence_json TEXT NOT NULL,
  PRIMARY KEY (run_id, finding_id)
);
"""


def connect(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize(path: Path, project: dict) -> None:
    with connect(path) as connection:
        connection.executescript(SCHEMA)
        connection.execute(
            """INSERT INTO project(project_id, project_name, industry)
               VALUES(?, ?, ?)
               ON CONFLICT(project_id) DO UPDATE SET
                 project_name=excluded.project_name,
                 industry=excluded.industry""",
            (project["project_id"], project["project_name"], project["industry"]),
        )


def seed_baseline(path: Path, project: dict, phases: list[dict], resources: list[dict]) -> None:
    import json

    initialize(path, project)
    with connect(path) as connection:
        connection.execute(
            """INSERT INTO project_baseline
               (project_id, preparation_start, formal_start_date, planned_go_live_date,
                trial_operation_start, planned_headcount)
               VALUES(?, ?, ?, ?, ?, ?)
               ON CONFLICT(project_id) DO UPDATE SET
                 preparation_start=excluded.preparation_start,
                 formal_start_date=excluded.formal_start_date,
                 planned_go_live_date=excluded.planned_go_live_date,
                 trial_operation_start=excluded.trial_operation_start,
                 planned_headcount=excluded.planned_headcount""",
            (project["project_id"], project["project_preparation_start"],
             project["formal_start_date"], project["planned_go_live_date"],
             project["trial_operation_start"], project["planned_headcount"]),
        )
        for phase in phases:
            connection.execute(
                """INSERT INTO project_phase
                   (phase_id, project_id, phase_name, planned_start, planned_finish,
                    phase_type, primary_deliverables)
                   VALUES(?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(phase_id) DO UPDATE SET
                     project_id=excluded.project_id,
                     phase_name=excluded.phase_name,
                     planned_start=excluded.planned_start,
                     planned_finish=excluded.planned_finish,
                     phase_type=excluded.phase_type,
                     primary_deliverables=excluded.primary_deliverables""",
                (phase["phase_id"], project["project_id"], phase["phase_name"],
                 phase["planned_start"], phase["planned_finish"], phase["phase_type"],
                 json.dumps(phase["primary_deliverables"], ensure_ascii=False)),
            )
        for resource in resources:
            connection.execute(
                """INSERT INTO resource_role
                   (role_code, role_name, planned_headcount, hours_per_week, hourly_rate)
                   VALUES(?, ?, ?, ?, ?)
                   ON CONFLICT(role_code) DO UPDATE SET
                     role_name=excluded.role_name,
                     planned_headcount=excluded.planned_headcount,
                     hours_per_week=excluded.hours_per_week,
                     hourly_rate=excluded.hourly_rate""",
                (resource["role_code"], resource["role_name"], resource["headcount"],
                 project["hours_per_week"], project["hourly_rate"]),
            )
