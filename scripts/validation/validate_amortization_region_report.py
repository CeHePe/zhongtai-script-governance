from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

from validate_amortization_project_report import (
    PROJECT_FILE,
    REPORT_PREFIX,
    ROOT,
    TOLERANCE_YUAN,
    build_plan_mapping,
    find_latest_plan_file,
    format_output,
    normalize_period,
)


@dataclass(frozen=True)
class FieldSpec:
    position: int
    name: str


FIELDS = [
    FieldSpec(1, "序号"),
    FieldSpec(2, "区域"),
    FieldSpec(3, "分摊类型"),
    FieldSpec(4, "事项类型"),
    FieldSpec(5, "带资/整改金额（含税）"),
    FieldSpec(6, "带资/整改金额（不含税）"),
    FieldSpec(7, "剩余摊销金额（不含税）"),
    FieldSpec(8, "剩余发生金额（不含税）"),
    FieldSpec(9, "往年累计-已摊销金额"),
    FieldSpec(10, "往年累计-已发生金额"),
    FieldSpec(11, "往年累计-还原金额"),
    FieldSpec(12, "当年累计-已摊销金额"),
    FieldSpec(13, "当年累计-已发生金额"),
    FieldSpec(14, "当年累计-还原金额"),
    FieldSpec(15, "一季度-已摊销金额"),
    FieldSpec(16, "一季度-已发生金额"),
    FieldSpec(17, "一季度-还原金额"),
    FieldSpec(18, "二季度-已摊销金额"),
    FieldSpec(19, "二季度-已发生金额"),
    FieldSpec(20, "二季度-还原金额"),
    FieldSpec(21, "三季度-已摊销金额"),
    FieldSpec(22, "三季度-已发生金额"),
    FieldSpec(23, "三季度-还原金额"),
    FieldSpec(24, "四季度-已摊销金额"),
    FieldSpec(25, "四季度-已发生金额"),
    FieldSpec(26, "四季度-还原金额"),
]

PROJECT_VALUE_POSITIONS = [9, 10, 11, 12, *range(16, 34)]
KEY_COLUMNS = ["region", "type", "item"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate every field in the regional amortization report."
    )
    parser.add_argument("--period", default="202512", help="Report period in YYYYMM format.")
    parser.add_argument("--output", type=Path, help="Local-only validation workbook path.")
    return parser.parse_args()


def load_reports(period: str) -> tuple[Path, pd.DataFrame, Path, pd.DataFrame]:
    region_path = ROOT / f"{REPORT_PREFIX}{period}区域.xlsx"
    project_path = ROOT / f"{REPORT_PREFIX}{period}项目.xlsx"
    for path in (region_path, project_path):
        if not path.exists():
            raise FileNotFoundError(f"Report not found: {path.name}")
    region = pd.read_excel(region_path, header=[0, 1])
    project = pd.read_excel(project_path, header=[0, 1])
    if region.shape[1] != len(FIELDS):
        raise ValueError(f"Expected {len(FIELDS)} regional fields, found {region.shape[1]}")
    if project.shape[1] != 33:
        raise ValueError(f"Expected 33 project fields, found {project.shape[1]}")
    return region_path, region, project_path, project


def load_latest_plan() -> tuple[Path, pd.DataFrame, str]:
    path = find_latest_plan_file()
    plan = pd.read_excel(path)
    snapshots = sorted({normalize_period(value) for value in plan["数据年月"].dropna()})
    if not snapshots:
        raise ValueError("Plan ledger has no snapshot month")
    selected = snapshots[-1]
    return path, plan[plan["数据年月"].map(normalize_period).eq(selected)].copy(), selected


def attach_project_codes(
    project_report: pd.DataFrame, plan: pd.DataFrame
) -> tuple[pd.DataFrame, int]:
    mapped, _ = build_plan_mapping(project_report, plan)
    unmatched = int(mapped["code"].isna().sum())
    if unmatched:
        raise ValueError(f"Project report has {unmatched} rows unmatched to the latest plan ledger")
    result = project_report.copy()
    result["_code"] = ""
    result.loc[mapped["report_index"].astype(int), "_code"] = mapped["code"].astype(str).to_numpy()
    return result, unmatched


def exclusion_mask(project_report: pd.DataFrame, report_period: pd.Period) -> pd.Series:
    query = pd.read_excel(ROOT / PROJECT_FILE)
    code_column = "立项编码"
    query[code_column] = query[code_column].astype(str).str.strip()
    lookup = query.drop_duplicates(code_column).set_index(code_column)
    level = project_report["_code"].map(lookup["项目等级"]).astype(str).str.strip()
    status = project_report["_code"].map(lookup["项目状态"]).astype(str).str.strip()
    exit_date = pd.to_datetime(project_report["_code"].map(lookup["已撤场时间"]), errors="coerce")
    return level.eq("D") & status.eq("已撤场") & exit_date.le(report_period.end_time)


def build_groups(
    region: pd.DataFrame, project: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    region_rows = pd.DataFrame(
        {
            "excel_row": region.index + 3,
            "sequence": pd.to_numeric(region.iloc[:, 0], errors="coerce"),
            "region": region.iloc[:, 1].astype(str).str.strip(),
            "type": region.iloc[:, 2].astype(str).str.strip(),
            "item": region.iloc[:, 3].astype(str).str.strip(),
        }
    )
    for position in range(5, 27):
        region_rows[f"value_{position}"] = pd.to_numeric(
            region.iloc[:, position - 1], errors="coerce"
        ).fillna(0.0)

    aggregations: dict[str, object] = {"excel_row": lambda values: ",".join(map(str, values))}
    aggregations.update({f"value_{position}": "sum" for position in range(5, 27)})
    actual_groups = region_rows.groupby(KEY_COLUMNS, as_index=False, dropna=False).agg(aggregations)

    project_keys = [project.columns[index] for index in (1, 6, 7)]
    project_values = [project.columns[position - 1] for position in PROJECT_VALUE_POSITIONS]
    expected_groups = project.groupby(project_keys, as_index=False, dropna=False)[project_values].sum()
    expected_groups.columns = KEY_COLUMNS + [f"expected_{position}" for position in range(5, 27)]
    return region_rows, actual_groups, expected_groups


def equal_number(actual: object, expected: object) -> tuple[bool, float]:
    left = float(pd.to_numeric(pd.Series([actual]), errors="coerce").fillna(0).iloc[0])
    right = float(pd.to_numeric(pd.Series([expected]), errors="coerce").fillna(0).iloc[0])
    diff = left - right
    return abs(diff) < TOLERANCE_YUAN, diff


def validate_keys(region_rows: pd.DataFrame, expected_groups: pd.DataFrame) -> pd.DataFrame:
    expected_keys = set(map(tuple, expected_groups[KEY_COLUMNS].to_numpy()))
    records: list[dict[str, object]] = []
    for index, row in region_rows.iterrows():
        key = tuple(row[column] for column in KEY_COLUMNS)
        key_exists = key in expected_keys
        checks = {
            1: (row["sequence"], index + 1, row["sequence"] == index + 1),
            2: (row["region"], row["region"], key_exists),
            3: (row["type"], row["type"], key_exists),
            4: (row["item"], row["item"], key_exists),
        }
        for position, (actual, expected, passed) in checks.items():
            records.append(
                {
                    "Excel行": int(row["excel_row"]),
                    "字段序号": position,
                    "字段": FIELDS[position - 1].name,
                    "报表值": actual,
                    "期望值": expected,
                    "差异": None,
                    "状态": "通过" if passed else "失败",
                    "说明": "序号连续" if position == 1 else "区域、分摊类型、事项类型组合存在于项目指标底稿",
                }
            )
    return pd.DataFrame(records)


def validate_values(actual_groups: pd.DataFrame, expected_groups: pd.DataFrame) -> pd.DataFrame:
    merged = actual_groups.merge(
        expected_groups, on=KEY_COLUMNS, how="outer", indicator=True, validate="one_to_one"
    )
    records: list[dict[str, object]] = []
    for _, row in merged.iterrows():
        for position in range(5, 27):
            actual = row.get(f"value_{position}")
            expected = row.get(f"expected_{position}")
            passed, diff = equal_number(actual, expected)
            if row["_merge"] != "both":
                passed = False
            records.append(
                {
                    "Excel行": row.get("excel_row", ""),
                    "字段序号": position,
                    "字段": FIELDS[position - 1].name,
                    "报表值": actual,
                    "期望值": expected,
                    "差异": diff,
                    "状态": "通过" if passed else "失败",
                    "说明": "项目报表按区域、分摊类型、事项类型汇总；重复可见组合合并验证",
                }
            )
    return pd.DataFrame(records)


def validate_internal_arithmetic(region: pd.DataFrame) -> pd.DataFrame:
    formulas = {
        6: lambda row: row[4] / 1.06,
        7: lambda row: row[5] - row[8] - row[11],
        8: lambda row: row[5] - row[9] - row[12],
        11: lambda row: row[9] - row[8],
        14: lambda row: row[12] - row[11],
        17: lambda row: row[15] - row[14],
        20: lambda row: row[18] - row[17],
        23: lambda row: row[21] - row[20],
        26: lambda row: row[24] - row[23],
    }
    records: list[dict[str, object]] = []
    numeric = region.apply(pd.to_numeric, errors="coerce").fillna(0.0)
    for index, row in numeric.iterrows():
        for position, formula in formulas.items():
            actual = row.iloc[position - 1]
            expected = formula(row.iloc)
            passed, diff = equal_number(actual, expected)
            records.append(
                {
                    "Excel行": index + 3,
                    "字段序号": position,
                    "字段": FIELDS[position - 1].name,
                    "报表值": actual,
                    "期望值": expected,
                    "差异": diff,
                    "状态": "通过" if passed else "失败",
                    "说明": "报表行内恒等式",
                }
            )
    return pd.DataFrame(records)


def build_summary(detail: pd.DataFrame, internal: pd.DataFrame) -> pd.DataFrame:
    counts = (
        detail.groupby(["字段序号", "字段", "状态"], as_index=False)
        .size()
        .pivot(index=["字段序号", "字段"], columns="状态", values="size")
        .fillna(0)
        .reset_index()
    )
    for column in ("通过", "失败", "阻塞"):
        if column not in counts:
            counts[column] = 0
        counts[column] = counts[column].astype(int)
    internal_failures = internal[internal["状态"].eq("失败")].groupby("字段序号").size()
    counts["内部校验失败"] = counts["字段序号"].map(internal_failures).fillna(0).astype(int)
    counts["测试单元"] = counts[["通过", "失败", "阻塞"]].sum(axis=1)
    counts["字段结论"] = np.select(
        [counts["失败"].gt(0) | counts["内部校验失败"].gt(0), counts["阻塞"].gt(0)],
        ["失败", "阻塞"],
        default="通过",
    )
    return counts[
        ["字段序号", "字段", "测试单元", "通过", "失败", "阻塞", "内部校验失败", "字段结论"]
    ]


def validate(period: str) -> tuple[pd.DataFrame, ...]:
    report_period = pd.Period(f"{period[:4]}-{period[4:6]}", freq="M")
    region_path, region, project_path, project = load_reports(period)
    plan_path, plan, plan_snapshot = load_latest_plan()
    project, unmatched = attach_project_codes(project, plan)
    excluded = exclusion_mask(project, report_period)
    included_project = project.loc[~excluded].copy()
    region_rows, actual_groups, expected_groups = build_groups(region, included_project)
    key_detail = validate_keys(region_rows, expected_groups)
    value_detail = validate_values(actual_groups, expected_groups)
    detail = pd.concat([key_detail, value_detail], ignore_index=True).sort_values(
        ["字段序号", "Excel行"]
    )
    internal = validate_internal_arithmetic(region)
    summary = build_summary(detail, internal)
    sources = pd.DataFrame(
        [
            {"来源": "待测区域报表", "文件": region_path.name, "状态": "已读取", "说明": f"期间={period}"},
            {"来源": "项目指标底稿", "文件": project_path.name, "状态": "已验证", "说明": "按区域、分摊类型、事项类型汇总"},
            {"来源": "计划台账", "文件": plan_path.name, "状态": "已读取", "说明": f"用于项目编码映射；快照={plan_snapshot}"},
            {"来源": "项目主数据", "文件": PROJECT_FILE, "状态": "已读取", "说明": "剔除D类、已撤场且撤场日期不晚于报告期末的项目"},
        ]
    )
    metadata = pd.DataFrame(
        [
            {
                "区域报表行数": len(region),
                "区域可见组合数": len(actual_groups),
                "重复可见组合数": int(region_rows.duplicated(KEY_COLUMNS, keep=False).sum() // 2),
                "项目底稿行数": len(project),
                "剔除项目行数": int(excluded.sum()),
                "未匹配计划行数": unmatched,
                "失败单元": int(detail["状态"].eq("失败").sum())
                + int(internal["状态"].eq("失败").sum()),
                "阻塞单元": int(detail["状态"].eq("阻塞").sum()),
            }
        ]
    )
    return summary, detail, internal, sources, metadata


def main() -> None:
    args = parse_args()
    output = args.output or ROOT / "output" / "spreadsheet" / f"amortization_region_{args.period}_validation.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    summary, detail, internal, sources, metadata = validate(args.period)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="字段汇总", index=False)
        detail.to_excel(writer, sheet_name="逐字段明细", index=False)
        internal.to_excel(writer, sheet_name="内部恒等式", index=False)
        pd.concat(
            [detail[detail["状态"].ne("通过")], internal[internal["状态"].ne("通过")]],
            ignore_index=True,
        ).to_excel(writer, sheet_name="失败及阻塞", index=False)
        sources.to_excel(writer, sheet_name="数据依赖", index=False)
        metadata.to_excel(writer, sheet_name="运行摘要", index=False)
    format_output(output)
    print(f"output={output}")
    print(metadata.to_string(index=False))
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
