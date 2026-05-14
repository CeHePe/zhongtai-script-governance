from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path
from _project_root import find_project_root
from typing import Any

from openpyxl import load_workbook


METRICS = {
    "before": "年度饱和收入_打折前_区域",
    "after": "年度饱和收入_打折后_区域",
}


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(base: Path, predicate) -> Path:
    matches = [
        p
        for p in base.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$") and predicate(p.name)
    ]
    if not matches:
        raise FileNotFoundError("未找到符合条件的工作簿")
    if len(matches) > 1:
        matches.sort(key=lambda p: p.name)
    return matches[0]


def iter_rows(path: Path) -> list[tuple[Any, ...]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    # 部分导出的 xlsx 维度标记错误为 A1，需要重置后才能读到实际单元格。
    ws.reset_dimensions()
    return list(ws.iter_rows(values_only=True))


def read_indicator_rows(base: Path) -> dict[str, dict[str, Any]]:
    path = find_workbook(base, lambda name: name.startswith("JKS_") and "指标清单" in name)
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    headers = [str(v).strip() if v is not None else f"未命名_{i}" for i, v in enumerate(next(ws.iter_rows(values_only=True)), 1)]
    found: dict[str, dict[str, Any]] = {}
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        row_dict = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
        metric_name = row_dict.get("指标名称")
        for key, expected in METRICS.items():
            if metric_name == expected:
                found[key] = {"excel_row": row_num, **row_dict}
    missing = [name for key, name in METRICS.items() if key not in found]
    if missing:
        raise RuntimeError(f"指标清单缺少行：{missing}")
    return found


def as_number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    return float(value)


def calculate(base: Path, month: str | None) -> tuple[str, dict[str, list[float]], dict[str, Any]]:
    ledger = find_workbook(base, lambda name: "新增年度饱和收入台账" in name)
    rows = iter_rows(ledger)
    if not rows:
        raise RuntimeError("新增年度饱和收入台账为空")

    headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    col = {name: idx for idx, name in enumerate(headers)}
    required_cols = ["报表年月", "区域", "年度饱和收入_打折前（元）", "口径认定金额（元）"]
    missing_cols = [name for name in required_cols if name not in col]
    if missing_cols:
        raise RuntimeError(f"新增年度饱和收入台账缺少字段：{missing_cols}")

    months = sorted({str(row[col["报表年月"]]) for row in rows[1:] if row[col["报表年月"]] not in (None, "")})
    if not months:
        raise RuntimeError("新增年度饱和收入台账没有可用报表年月")
    selected_month = month or months[-1]
    if selected_month not in months:
        raise RuntimeError(f"台账没有报表年月 {selected_month}；可用月份：{', '.join(months)}")

    agg: dict[str, list[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])
    for row in rows[1:]:
        if str(row[col["报表年月"]]) != selected_month:
            continue
        region = row[col["区域"]]
        if region in (None, ""):
            region = "(空区域)"
        agg[str(region)][0] += 1
        agg[str(region)][1] += as_number(row[col["年度饱和收入_打折前（元）"]])
        agg[str(region)][2] += as_number(row[col["口径认定金额（元）"]])

    meta = {"ledger": ledger.name, "available_months": months, "headers": headers}
    return selected_month, dict(sorted(agg.items())), meta


def main() -> None:
    configure_stdout()
    parser = argparse.ArgumentParser(description="计算年度饱和收入区域指标的正确值")
    parser.add_argument("--month", help="报表年月，例如 2025-12；不传则使用台账最新月份")
    args = parser.parse_args()

    base = find_project_root(__file__)
    indicator_rows = read_indicator_rows(base)
    month, result, meta = calculate(base, args.month)

    print("指标清单口径：")
    for key in ("before", "after"):
        row = indicator_rows[key]
        print(
            f"- Excel行{row['excel_row']} {row['指标名称']} | "
            f"组织维度={row['组织维度']} | 累计/月度={row['累计/月度']} | "
            f"取数方式={row['取数方式']} | 取数对应表={row['取数对应表']}"
        )
        print(f"  计算逻辑={row['计算逻辑']}")

    print("\n依赖确认：")
    print(f"- 必需台账：{meta['ledger']}；未缺失")
    print("- 来源字段：年度饱和收入_打折前（元）、口径认定金额（元）")
    print("- 累计/月度：指标清单为累计，但逻辑是按区域+报表年月直接取台账，不做1月至报表月逐月累加")
    print("- D类已撤场排除：指标清单未要求")
    print("- 非考核项目排除：指标清单未要求")
    print("- 区域级额外调整：指标清单未要求")
    print("- 高维汇总依赖：不从项目报表汇总，直接按区域+报表年月从手工台账汇总")

    print(f"\n报表年月：{month}")
    print("区域\t记录数\t年度饱和收入_打折前（元）\t年度饱和收入_打折前（万元）\t年度饱和收入_打折后（元）\t年度饱和收入_打折后（万元）")
    for region, (count, before_yuan, after_yuan) in result.items():
        print(
            f"{region}\t{int(count)}\t{before_yuan:.2f}\t{before_yuan / 10000:.4f}\t"
            f"{after_yuan:.2f}\t{after_yuan / 10000:.4f}"
        )


if __name__ == "__main__":
    main()
