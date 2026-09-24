#!/usr/bin/env python3
"""把成本口径数字（计划人天/人工成本/平均人日单价）联动到 G15 会前材料。

更新对象（按文件名模式定位最新版）：
- 01_项目启动与治理/G15-会前签认跟踪-*.xlsx 的 签认跟踪!H8
- 01_项目启动与治理/G15-未决事项处置建议包-*.md
- 01_项目启动与治理/G15-会前材料包清单与缺口对照-*.md

数字来源：05 成本簿 人工成本预算 合计行动态定位（与 sync_cost_to_ai_pmo 同口径）。
默认仅预览；--apply 时先备份到目标目录 _archive/成本口径联动前-YYYYMMDD/ 再原位更新。
xlsx 走 zipfile 级 sharedStrings 手术（不动公式缓存）；md 为文本正则替换。
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from scripts.sync_cost_to_ai_pmo import CostMetrics, extract_cost_metrics, backup_workbook

G15_DIR = "01_项目启动与治理"
XLSX_PATTERN = "G15-会前签认跟踪-*.xlsx"
MD_PATTERNS = ("G15-未决事项处置建议包-*.md", "G15-会前材料包清单与缺口对照-*.md")

# 人天 / 万元 / 单价 三连（保留原文分隔符与标签写法；单价含小数）
TRIPLE_RE = re.compile(
    r"[\d,]+ 人天(\s*/\s*|/) ?[\d.]+ 万元(\s*/\s*|/) ?(平均人日单价|均价) [\d,]+(?:\.\d+)? 元"
)
WAN_RE = re.compile(r"人工成本 [\d.]+ 万（月度投入明细驱动）")
QUOTE_RE = re.compile(r"\"[\d,]+ 人天/[\d.]+ 万元（月度投入明细驱动）\"")
DASHBOARD_RE = re.compile(
    r"驾驶舱 \d{4}-\d{2}-\d{2} 口径为 [\d,]+ 人天/[\d.]+ 万元/均价 [\d,]+(?:\.\d+)? 元"
)
H8_RE = re.compile(
    r"确认计划人天唯一来源为《月度投入明细》：[\d,]+ 人天 / [\d.]+ 万元 / 均价 [\d,]+(?:\.\d+)? 元（G15 会前完成）"
)


def fmt_person_days(v: float) -> str:
    return f"{v:,.0f}"


def fmt_wan(v: float) -> str:
    return f"{v / 10000:.2f}"


def fmt_rate(v: float) -> str:
    return f"{v:,.2f}"


def replace_md_text(text: str, metrics: CostMetrics, today: str) -> tuple[str, int]:
    pd, wan, rate = fmt_person_days(metrics.person_days), fmt_wan(metrics.total_cost), fmt_rate(metrics.average_day_rate)
    n = 0

    def triple(m: re.Match) -> str:
        nonlocal n
        n += 1
        return f"{pd} 人天{m.group(1)}{wan} 万元{m.group(2)}{m.group(3)} {rate} 元"

    text = TRIPLE_RE.sub(triple, text)
    new = WAN_RE.sub(f"人工成本 {wan} 万（月度投入明细驱动）", text)
    n += len(WAN_RE.findall(text)); text = new
    new = QUOTE_RE.sub(f'"{pd} 人天/{wan} 万元（月度投入明细驱动）"', text)
    n += len(QUOTE_RE.findall(text)); text = new
    new = DASHBOARD_RE.sub(f"驾驶舱 {today} 口径为 {pd} 人天/{wan} 万元/均价 {rate} 元", text)
    n += len(DASHBOARD_RE.findall(text)); text = new
    return text, n


def iter_string_parts(path: Path):
    """产出 (zip条目名, 文本)；优先 sharedStrings（原生 UTF-8），
    其次工作表内联字符串（数字字符引用已解码为真实字符）。"""
    import html
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "xl/sharedStrings.xml" in names:
            yield "xl/sharedStrings.xml", z.read("xl/sharedStrings.xml").decode("utf-8")
        else:
            for name in names:
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    yield name, html.unescape(z.read(name).decode("utf-8"))


def escape_non_ascii(text: str) -> str:
    """按 openpyxl 风格把非 ASCII 字符编码为数字字符引用（ASCII 原样保留）。"""
    return "".join(ch if ord(ch) < 128 else f"&#{ord(ch)};" for ch in text)


def replace_xlsx(path: Path, metrics: CostMetrics) -> int:
    """zipfile 级替换签认跟踪口径句（sharedStrings 或工作表内联字符串）；返回替换次数。"""
    pd, wan, rate = fmt_person_days(metrics.person_days), fmt_wan(metrics.total_cost), fmt_rate(metrics.average_day_rate)
    repl = f"确认计划人天唯一来源为《月度投入明细》：{pd} 人天 / {wan} 万元 / 均价 {rate} 元（G15 会前完成）"
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {name: z.read(name) for name in names}
        infos = {i.filename: i for i in z.infolist()}
    count = 0
    for part, xml in list(iter_string_parts(path)):
        new, n = H8_RE.subn(repl, xml)
        if n:
            # sharedStrings 为原生 UTF-8；工作表内联字符串需恢复数字字符引用
            if part.startswith("xl/worksheets/"):
                new = escape_non_ascii(new)
            data[part] = new.encode("utf-8")
            count += n
    if count:
        tmp = path.with_suffix(path.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for name in names:
                zo.writestr(infos[name], data[name])
        tmp.replace(path)
    return count


def locate_targets(suite: Path) -> dict[str, list[Path]]:
    g15 = suite / G15_DIR
    targets: dict[str, list[Path]] = {"xlsx": [], "md": []}
    if not g15.exists():
        return targets
    xlsx = sorted(g15.glob(XLSX_PATTERN))
    if xlsx:
        targets["xlsx"].append(xlsx[-1])
    for pattern in MD_PATTERNS:
        matches = sorted(g15.glob(pattern))
        if matches:
            targets["md"].append(matches[-1])
    return targets


def preview_sync(suite: Path) -> dict[str, object]:
    from ai_pmo.current_workbooks import resolve_current
    cost = resolve_current(suite, "cost")
    metrics = extract_cost_metrics(cost)
    targets = locate_targets(suite)
    plan: list[dict[str, object]] = []
    pd, wan, rate = fmt_person_days(metrics.person_days), fmt_wan(metrics.total_cost), fmt_rate(metrics.average_day_rate)
    for path in targets["xlsx"]:
        xml = "".join(text for _, text in iter_string_parts(path))
        hits = len(H8_RE.findall(xml))
        # 已是最新口径时给出 0，便于区分"无需更新"
        plan.append({"file": path, "kind": "xlsx", "replacements": hits,
                     "sample": f"{pd} 人天 / {wan} 万元 / 均价 {rate} 元"})
    for path in targets["md"]:
        text = path.read_text(encoding="utf-8")
        hits = (len(TRIPLE_RE.findall(text)) + len(WAN_RE.findall(text))
                + len(QUOTE_RE.findall(text)) + len(DASHBOARD_RE.findall(text)))
        plan.append({"file": path, "kind": "md", "replacements": hits,
                     "sample": f"{pd} 人天 / {wan} 万元 / 均价 {rate} 元"})
    return {"source": cost, "metrics": metrics, "plan": plan}


def sync(suite: Path, *, expected_source: Path | None = None) -> list[Path]:
    from ai_pmo.current_workbooks import resolve_current
    cost = resolve_current(suite, "cost")
    if expected_source is not None and cost != expected_source:
        raise RuntimeError(f"源工作簿在预览后已变化：{expected_source} -> {cost}")
    metrics = extract_cost_metrics(cost)
    today = date.today().isoformat()
    backups: list[Path] = []
    for path in locate_targets(suite)["xlsx"]:
        xml = "".join(text for _, text in iter_string_parts(path))
        if not H8_RE.findall(xml):
            continue
        backups.append(backup_workbook(path))
        replace_xlsx(path, metrics)
    for path in locate_targets(suite)["md"]:
        text = path.read_text(encoding="utf-8")
        new, n = replace_md_text(text, metrics, today)
        if n == 0:
            continue
        backups.append(backup_workbook(path))
        path.write_text(new, encoding="utf-8")
    return backups


def main() -> None:
    default_suite = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description="将成本口径数字联动到 G15 会前材料（签认跟踪与两份 md）")
    parser.add_argument("--suite", type=Path, default=default_suite)
    parser.add_argument("--apply", action="store_true", help="备份目标后执行原位更新；默认仅预览")
    args = parser.parse_args()
    suite = args.suite.resolve()
    preview = preview_sync(suite)
    m = preview["metrics"]
    print(f"读取成本工作簿：{preview['source']}")
    print(f"计划人天={fmt_person_days(m.person_days)}，人工成本={fmt_wan(m.total_cost)} 万元，平均人日单价={fmt_rate(m.average_day_rate)} 元")
    if not preview["plan"]:
        print("未找到 G15 会前材料目标，无需联动。")
        return
    for item in preview["plan"]:
        print(f"将更新：{item['file'].relative_to(suite)}（命中 {item['replacements']} 处）")
    if not args.apply:
        print("仅预览，未修改文件。执行更新请添加 --apply。")
        return
    backups = sync(suite, expected_source=preview["source"])
    for backup in backups:
        print(f"备份：{backup}")


if __name__ == "__main__":
    main()
