from __future__ import annotations

from pathlib import Path
from math import isclose
from openpyxl import load_workbook


def check_milestone_links(milestone_ids, task_milestones):
    milestones = set(milestone_ids)
    referenced = set(task_milestones)
    return {
        'errors': [f'总控任务引用不存在的里程碑 {item}' for item in sorted(referenced - milestones)],
        'warnings': [f'里程碑 {item} 没有总控任务' for item in sorted(milestones - referenced)],
    }


def compare_dashboard_snapshot(expected_sources, actual_sources, expected_metrics, actual_metrics):
    warnings = []
    for key, expected in expected_sources.items():
        actual = actual_sources.get(key)
        if actual != expected:
            warnings.append(f'驾驶舱来源快照过期：{key} 登记为{expected}，快照为{actual or "缺失"}')
    for key, expected in expected_metrics.items():
        if key not in actual_metrics or actual_metrics[key] in (None, ''):
            warnings.append(f'驾驶舱指标待刷新：{key} 源台账={expected}，快照缺失')
        elif (isinstance(expected, (int, float)) and isinstance(actual_metrics[key], (int, float))
              and isclose(float(expected), float(actual_metrics[key]), rel_tol=1e-9, abs_tol=0.005)):
            continue
        elif actual_metrics[key] != expected:
            warnings.append(f'驾驶舱指标不一致：{key} 源台账={expected}，快照={actual_metrics[key]}')
    return {'errors': [], 'warnings': warnings}


def dashboard_metrics_from_cells(values_by_row):
    rows = {
        '待决策事项': 14, '计划投入人天': 17, '计划人工成本': 18,
        '人日单价': 19, '里程碑总数': 36,
    }
    return {name: values_by_row.get(row) for name, row in rows.items()}


def check_migration_progress(interface: dict) -> dict:
    """Report migration source-inventory progress from the migration workbook interface."""
    clues = interface.get('来源线索数')
    inventoried = interface.get('已盘点来源系统数')
    gate_date = interface.get('G15 目标日期')
    if hasattr(gate_date, 'date'):
        gate_date = gate_date.date().isoformat()
    counts: dict = {}
    warnings: list[str] = []
    if clues in (None, ''):
        warnings.append('迁移工作簿缺少_数据接口盘点统计，无法核对来源系统盘点进度')
    else:
        counts = {'迁移来源线索': clues, '已盘点来源': inventoried}
        if inventoried != clues:
            warnings.append(
                f'迁移来源系统盘点未完成：已盘点 {inventoried or 0}/{clues}'
                + (f'，G15 目标日期 {gate_date}' if gate_date else '')
            )
    return {'counts': counts, 'warnings': warnings}


def check_milestone_owner_drift(plan_owners: dict[str, str],
                                dashboard_owners: dict[str, str]) -> dict:
    """里程碑负责人在计划工作簿与驾驶舱快照之间出现差异时提示。

    plan_owners / dashboard_owners: {里程碑名称: 负责人}。差异通常意味着
    里程碑计划刚被调整而驾驶舱尚未刷新，需要确认总控台账对应任务负责人已联动。
    """
    warnings = []
    for name, owner in plan_owners.items():
        snap = dashboard_owners.get(name)
        if snap and snap != owner:
            warnings.append(
                f'里程碑“{name}”负责人已变更：驾驶舱快照为{snap}，里程碑计划为{owner}；'
                '请确认总控台账对应任务负责人已联动，并刷新驾驶舱快照'
            )
    return {'errors': [], 'warnings': warnings}


def check_milestone_owner_alignment(milestones: list[dict],
                                    batch_owners: dict[str, list[str]]) -> dict:
    """里程碑负责人应直接负责其批次内至少一条任务。

    milestones: [{'id', 'name', 'owner', 'status'}]；batch_owners: {MS编号: [任务负责人…]}
    （任务负责人按业务→后端→前端→测试取第一个非空，空字符串表示该任务未落实负责人）。
    仅校验进行中/已完成且批次任务已全部落实负责人的里程碑，避免对未排期工作误报。
    """
    warnings = []
    for m in milestones:
        if m['status'] not in ('进行中', '已完成'):
            continue
        owners = batch_owners.get(m['id'], [])
        if not owners or any(not o for o in owners):
            continue
        if m['owner'] not in owners:
            warnings.append(
                f'里程碑 {m["id"]}“{m["name"]}”（负责人 {m["owner"]}）批次任务已全部落实负责人，'
                f'但无任何任务由 {m["owner"]} 负责，请核对里程碑负责人或任务负责人'
            )
    return {'errors': [], 'warnings': warnings}


def _dashboard_milestone_owners(sheet) -> dict[str, str]:
    """从驾驶舱项目驾驶舱表提取里程碑区块 {名称: 责任人}，表结构变化时返回空。"""
    try:
        rows = list(sheet.iter_rows(values_only=True))
    except Exception:
        return {}
    for i, row in enumerate(rows):
        cells = [str(v) if v is not None else '' for v in row]
        if '里程碑' in cells and '责任人' in cells:
            name_col, owner_col = cells.index('里程碑'), cells.index('责任人')
            owners: dict[str, str] = {}
            for data in rows[i + 1:]:
                name = data[name_col] if name_col < len(data) else None
                owner = data[owner_col] if owner_col < len(data) else None
                if name in (None, ''):
                    break
                owners[str(name)] = str(owner) if owner else ''
            return owners
    return {}


def _ids(sheet, prefix, start_row=1):
    return [str(row[0]) for row in sheet.iter_rows(min_row=start_row, values_only=True)
            if row and isinstance(row[0], str) and row[0].startswith(prefix)]


def _interface(sheet):
    return {str(row[0]): row[1] for row in sheet.iter_rows(values_only=True)
            if len(row) > 1 and row[0] not in (None, '')}


def audit_portfolio(plan_path: Path, risk_path: Path, change_path: Path,
                    quality_path: Path, deliverable_path: Path, cost_path: Path,
                    ai_pmo_path: Path, registered_sources: dict[str, str],
                    migration_path: Path | None = None) -> dict:
    books = [load_workbook(path, read_only=True, data_only=True) for path in
             (plan_path, risk_path, change_path, quality_path, deliverable_path, cost_path, ai_pmo_path)]
    migration_book = load_workbook(migration_path, read_only=True, data_only=True) if migration_path else None
    if migration_book is not None:
        books.append(migration_book)
    plan, risk, change, quality, deliverable, cost, ai_pmo = books[:7]
    try:
        milestone_ids = _ids(plan['里程碑计划'], 'MS', 4)
        task_milestones = [str(row[5]) for row in plan['项目总控台账'].iter_rows(min_row=4, values_only=True)
                           if row and isinstance(row[0], str) and row[0].startswith('T') and row[5]]
        links = check_milestone_links(milestone_ids, task_milestones)
        ms_rows = []
        plan_owners: dict[str, str] = {}
        for row in plan['里程碑计划'].iter_rows(min_row=4, values_only=True):
            if row and isinstance(row[0], str) and row[0].startswith('MS'):
                ms_rows.append({'id': row[0], 'name': row[2], 'owner': row[4],
                                'status': row[10]})
                if row[2]:
                    plan_owners[str(row[2])] = str(row[4]) if row[4] else ''
        batch_owners: dict[str, list[str]] = {}
        for row in plan['项目总控台账'].iter_rows(min_row=4, values_only=True):
            if row and isinstance(row[0], str) and row[0].startswith('T') and row[5]:
                owner = next((str(v) for v in (row[11], row[15], row[19], row[23])
                              if v not in (None, '')), '')
                batch_owners.setdefault(str(row[5]), []).append(owner)
        owner_drift = check_milestone_owner_drift(
            plan_owners, _dashboard_milestone_owners(ai_pmo['项目驾驶舱']))
        owner_align = check_milestone_owner_alignment(ms_rows, batch_owners)
        interface = _interface(ai_pmo['_数据接口'])
        source_snapshot = {'plan': interface.get('里程碑来源'), 'risk': interface.get('风险来源')}
        risk_ids = _ids(risk['风险登记册'], 'RSK-', 9)
        decision_ids = _ids(change['待决策事项'], 'D-', 5)
        task_ids = _ids(plan['项目总控台账'], 'T', 4)
        cost_sheet = cost['人工成本预算']
        expected_metrics = {
            '总任务数': len(task_ids), '风险总数': len(risk_ids),
            '待决策事项': len(decision_ids), '里程碑总数': len(milestone_ids),
            '计划投入人天': cost_sheet['D14'].value,
            '计划人工成本': cost_sheet['F14'].value,
            '人日单价': cost_sheet['E3'].value,
        }
        actual_metrics = {name: interface.get(name) for name in ('总任务数', '风险总数')}
        actual_metrics.update(dashboard_metrics_from_cells({row: ai_pmo['_数据接口'].cell(row, 2).value
                                                             for row in (14, 17, 18, 19, 36)}))
        snapshot = compare_dashboard_snapshot(
            {key: value for key, value in registered_sources.items() if key in source_snapshot},
            source_snapshot, expected_metrics, actual_metrics,
        )
        counts = {
            '里程碑': len(milestone_ids), '总控任务': len(task_ids), '风险': len(risk_ids),
            '待决策': len(decision_ids), '缺陷': len(_ids(quality['缺陷台账'], 'DEF-', 5)),
            'UAT': len(_ids(quality['UAT验收'], 'UAT-', 5)),
            '交付物': len(_ids(deliverable['交付物台账'], 'D-', 5)),
        }
        migration = {'counts': {}, 'warnings': []}
        if migration_book is not None:
            migration = check_migration_progress(_interface(migration_book['_数据接口']))
            counts.update(migration['counts'])
        return {'counts': counts, 'errors': links['errors'] + snapshot['errors'] + owner_drift['errors'] + owner_align['errors'],
                'warnings': links['warnings'] + snapshot['warnings'] + migration['warnings']
                + owner_drift['warnings'] + owner_align['warnings']}
    finally:
        for book in books: book.close()
