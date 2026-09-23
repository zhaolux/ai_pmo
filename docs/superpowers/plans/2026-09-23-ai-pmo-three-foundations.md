# AI PMO 三项基础能力实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 建立唯一数据主源审计、六维健康度与历史快照、Agent发现事项人工确认闭环三项可重复运行能力。

**Architecture:** 继续以 `current_workbooks.json` 登记业务主源，新增对象级数据所有权注册表并纳入日常只读巡检。Agent运行时由独立健康度模块计算六维得分并把总分和维度快照写入SQLite。人工确认仍在Agent结果Excel中填写，再由显式导入命令校验运行ID和发现ID后写入审计表，不反向覆盖业务台账。

**Tech Stack:** Python 3.12、SQLite、openpyxl、unittest、现有 `ai_pmo` CLI 与Excel生成脚本。

---

### Task 1 数据主源与联动审计

**Files:**
- Create: `合同全生命周期管理系统-项目管理套件/automation/config/data_ownership.json`
- Create: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/source_ownership.py`
- Create: `合同全生命周期管理系统-项目管理套件/automation/scripts/audit_source_ownership.py`
- Create: `合同全生命周期管理系统-项目管理套件/automation/tests/test_source_ownership.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/daily_run.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/tests/test_daily_run.py`

- [ ] 先写测试：重复对象主源、缺少登记文件、Sheet缺失、ID表头缺失必须失败；合法注册表返回对象计数和零错误。
- [ ] 运行单测并确认因 `source_ownership` 尚不存在而失败。
- [ ] 实现注册表读取、当前工作簿解析、Sheet及主键表头校验，不写任何工作簿。
- [ ] 增加只读命令并把“数据主源审计”放在日常巡检第一步。
- [ ] 运行相关测试和真实套件审计，预期错误0；不自动修复业务数据。

### Task 2 六维健康度与历史快照

**Files:**
- Create: `合同全生命周期管理系统-项目管理套件/automation/config/health_model.json`
- Create: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/health_model.py`
- Create: `合同全生命周期管理系统-项目管理套件/automation/tests/test_health_model.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/agents/schedule_resource.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/tests/test_agents.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/database.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/orchestrator.py`

- [ ] 先写测试：六维权重合计100；严重度扣分可解释；重大事项触发红色硬门槛；分数边界85/70正确；资源超负荷产生 `RES-` 发现。
- [ ] 运行测试并确认健康度模块缺失、旧主控仅按发现数量判断而失败。
- [ ] 实现六维配置、扣分明细、硬门槛和总分/状态计算；未识别严重度不得静默忽略。
- [ ] 扩展进度资源Agent，只读取数据日期所在周的资源计划并报告超负荷，不把低负荷自动判为问题。
- [ ] 新增 `health_snapshot` 和 `health_dimension_snapshot` 表；每次Agent报告持久化时同事务写入快照。
- [ ] 报告JSON新增 `health_score`、`health_dimensions`、`health_rules_version`，保留原 `health` 字段兼容周报和Excel。
- [ ] 运行相关测试并在临时数据库核对总分、六维快照和运行ID一致。

### Task 3 Agent发现事项人工确认闭环

**Files:**
- Create: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/finding_review.py`
- Create: `合同全生命周期管理系统-项目管理套件/automation/tests/test_finding_review.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/database.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/src/ai_pmo/cli.py`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/scripts/build_agent_results_excel.mjs`

- [ ] 先写测试：仅允许待确认、已接受、需调整、已关闭；工作簿运行ID必须与数据库一致；未知发现ID、重复发现ID、空审核人必须拒绝且不部分写入。
- [ ] 运行测试并确认导入能力不存在而失败。
- [ ] 新增 `agent_finding_review` 审计表，主键为运行ID与发现ID，保存状态、结论、审核人和审核时间。
- [ ] 实现Excel导入函数：读取“发现清单”K:L和“Agent总览”运行ID；校验全量后在单事务中写入，不修改源项目台账。
- [ ] CLI新增 `review-import --file ... --reviewer ...` 和 `review-summary --run-id ...`。
- [ ] 在Agent结果工作簿使用说明中增加确认导入命令和边界说明。
- [ ] 用测试工作簿完成一次待确认→已接受导入，再验证重复导入更新同一审计记录而不新增重复行。

### Task 4 文档与全量验收

**Files:**
- Modify: `合同全生命周期管理系统-项目管理套件/automation/README.md`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/docs/DATA_LINEAGE_AND_WRITE_POLICY.md`
- Modify: `合同全生命周期管理系统-项目管理套件/automation/docs/AI_PMO_AGENT_ARCHITECTURE.md`
- Modify: `ROADMAP.md`

- [ ] 更新运行说明、数据血缘、人工审批边界和恢复方式。
- [ ] 运行 `unittest discover` 全量测试。
- [ ] 运行真实套件数据主源审计、数据联动审计、组合审计和日常预览。
- [ ] 在临时输出目录运行Agent，核对JSON六维得分、SQLite快照与Excel确认字段。
- [ ] 执行 `git diff --check`，不提交、不推送、不清理现有未提交文件。

## 自检结论

- 规格覆盖：三项能力分别由Task 1、Task 2、Task 3实现，Task 4统一验收。
- 数据边界：业务台账仍是唯一事实源；健康快照和确认记录是审计数据，不取得审批效力。
- 安全边界：所有新审计默认只读；确认导入只写中央数据库，不写业务工作簿。
- Git边界：依据项目 `AGENTS.md`，本计划不包含自动提交或推送步骤。
