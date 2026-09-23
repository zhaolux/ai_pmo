import unittest

from ai_pmo.portfolio_audit import (
    check_migration_progress, check_milestone_links, check_milestone_owner_alignment,
    check_milestone_owner_drift, compare_dashboard_snapshot,
    dashboard_metrics_from_cells,
)


class PortfolioAuditTests(unittest.TestCase):
    def test_milestone_links_report_missing_and_uncovered_ids(self):
        result = check_milestone_links(['MS001', 'MS002'], ['MS001', 'MS999'])
        self.assertEqual(result['errors'], ['总控任务引用不存在的里程碑 MS999'])
        self.assertEqual(result['warnings'], ['里程碑 MS002 没有总控任务'])

    def test_dashboard_snapshot_reports_source_and_metric_drift_as_warnings(self):
        result = compare_dashboard_snapshot(
            expected_sources={'plan': '02/new.xlsx'},
            actual_sources={'plan': '02/old.xlsx'},
            expected_metrics={'总任务数': 135, '风险总数': 12},
            actual_metrics={'总任务数': 135, '风险总数': 11},
        )
        self.assertEqual(result['errors'], [])
        self.assertIn('驾驶舱来源快照过期：plan 登记为02/new.xlsx，快照为02/old.xlsx', result['warnings'])
        self.assertIn('驾驶舱指标不一致：风险总数 源台账=12，快照=11', result['warnings'])

    def test_missing_dashboard_metric_is_pending_refresh_not_zero(self):
        result = compare_dashboard_snapshot({}, {}, {'计划人工成本': 100}, {})
        self.assertEqual(result['warnings'], ['驾驶舱指标待刷新：计划人工成本 源台账=100，快照缺失'])

    def test_ignores_insignificant_floating_point_difference(self):
        result = compare_dashboard_snapshot({}, {}, {'人日单价': 1454.18502202643},
                                            {'人日单价': 1454.1850220264316})
        self.assertEqual(result['warnings'], [])

    def test_dashboard_metrics_use_fixed_interface_cells_when_labels_are_absent(self):
        values = {14: 4, 17: 1135, 18: 1650500, 19: 1454.18, 36: 17}
        self.assertEqual(dashboard_metrics_from_cells(values), {
            '待决策事项': 4, '里程碑总数': 17, '计划投入人天': 1135,
            '计划人工成本': 1650500, '人日单价': 1454.18,
        })

    def test_migration_progress_warns_until_all_sources_inventoried(self):
        result = check_migration_progress(
            {'来源线索数': 6, '已盘点来源系统数': 2, 'G15 目标日期': '2026-10-13'})
        self.assertEqual(result['counts'], {'迁移来源线索': 6, '已盘点来源': 2})
        self.assertEqual(result['warnings'], [
            '迁移来源系统盘点未完成：已盘点 2/6，G15 目标日期 2026-10-13'])

    def test_migration_progress_is_silent_when_inventory_complete(self):
        result = check_migration_progress({'来源线索数': 6, '已盘点来源系统数': 6})
        self.assertEqual(result['counts'], {'迁移来源线索': 6, '已盘点来源': 6})
        self.assertEqual(result['warnings'], [])

    def test_migration_progress_flags_missing_interface_stats(self):
        result = check_migration_progress({})
        self.assertEqual(result['counts'], {})
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('缺少_数据接口', result['warnings'][0])

    def test_milestone_owner_drift_flags_owner_changed_after_dashboard_snapshot(self):
        result = check_milestone_owner_drift(
            {'可行性研究完成': '项目经理', '可研工作启动': '项目经理'},
            {'可行性研究完成': '项目集经理', '可研工作启动': '项目经理'},
        )
        self.assertEqual(result['errors'], [])
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('可行性研究完成', result['warnings'][0])
        self.assertIn('项目集经理', result['warnings'][0])
        self.assertIn('项目经理', result['warnings'][0])

    def test_milestone_owner_drift_ignores_milestones_absent_from_dashboard(self):
        result = check_milestone_owner_drift({'正式启动': '项目集经理'}, {})
        self.assertEqual(result['warnings'], [])

    def test_milestone_owner_alignment_warns_when_owner_own_no_assigned_task(self):
        milestones = [{'id': 'MS001', 'name': '可研工作启动', 'owner': '项目集经理',
                       'status': '已完成'}]
        result = check_milestone_owner_alignment(milestones, {'MS001': ['项目经理', '产品经理']})
        self.assertEqual(len(result['warnings']), 1)
        self.assertIn('MS001', result['warnings'][0])

    def test_milestone_owner_alignment_skips_unassigned_or_not_started_batches(self):
        milestones = [
            {'id': 'MS001', 'name': 'A', 'owner': '项目集经理', 'status': '已完成'},
            {'id': 'MS002', 'name': 'B', 'owner': '技术经理', 'status': '未开始'},
        ]
        result = check_milestone_owner_alignment(
            milestones, {'MS001': ['项目经理', ''], 'MS002': ['技术经理']})
        self.assertEqual(result['warnings'], [])


if __name__ == '__main__':
    unittest.main()
