from __future__ import annotations

import importlib.util
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "2025-12"
TOLERANCE = 1e-6


def load_base_module():
    """复用半收付归母净利润脚本里的文件查找、编码标准化和项目报表读取函数。"""
    script_path = ROOT / "validate_half_cash_attributable_profit.py"
    spec = importlib.util.spec_from_file_location("half_cash_base", script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def load_indicator_rows(base) -> pd.DataFrame:
    """读取指标清单中“截止性支出_半收付”的项目级指标行。"""
    indicator = pd.read_excel(base.find_workbook(base.INDICATOR), sheet_name=0, dtype=object)
    cols = list(indicator.columns)
    metric = base.u(r"\u622a\u6b62\u6027\u652f\u51fa_\u534a\u6536\u4ed8")
    project = base.u(r"\u9879\u76ee")
    rows = indicator[
        indicator[cols[3]].astype(str).eq(metric)
        & indicator[cols[4]].astype(str).eq(project)
    ][[cols[3], cols[4], cols[5], cols[8], cols[10], cols[12]]].fillna("")
    rows.columns = ["metric_name", "dimension", "cycle", "method", "source_table", "logic"]
    return rows


def find_cutoff_ledger(base) -> Path:
    """本次报表月使用 2025-12 台账，排除历史 202412 台账。"""
    path = ROOT / base.u(r"\u622a\u6b62\u6027\u6536\u652f\u8c03\u6574\u53f0\u8d26.xlsx")
    if not path.exists():
        raise FileNotFoundError(f"Missing cutoff ledger: {path.name}")
    return path


def load_cutoff_ledger(base, query_df: pd.DataFrame) -> pd.DataFrame:
    """读取截止性收支调整台账，并按项目编码补充穿透比例。"""
    raw = pd.read_excel(find_cutoff_ledger(base), header=None, dtype=object)
    code_label = base.u(r"\u7acb\u9879\u7f16\u7801")
    header_idx = next(
        index
        for index in range(min(20, len(raw)))
        if code_label in raw.iloc[index].astype(str).str.strip().tolist()
    )
    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = raw.iloc[header_idx].astype(str).str.strip().tolist()

    type_col = base.u(r"\u7c7b\u578b")
    month_col = base.u(r"\u6570\u636e\u5e74\u6708")
    project_label = base.u(r"\u9879\u76ee")
    cost_col = next(
        column
        for column in df.columns
        if base.u(r"\u5f53\u5e74\u65b0\u589e\u6210\u672c\u91d1\u989d") in str(column)
    )

    source = df[
        df[type_col].astype(str).str.strip().eq(project_label)
        & df[month_col].astype(str).str[:7].eq(REPORT_MONTH)
    ].copy()
    source["code_exact"] = source[code_label].astype(str).str.strip().str.upper()
    source["cost_no_tax"] = pd.to_numeric(source[cost_col], errors="coerce").fillna(0.0)

    query_df = query_df.copy()
    query_df["code_exact"] = query_df["project_code"].astype(str).str.strip().str.upper()
    ratio = query_df[["code_exact", "penetration_ratio"]].drop_duplicates("code_exact")
    source = source.merge(ratio, on="code_exact", how="left")
    source["missing_query_ratio"] = source["penetration_ratio"].isna()
    source["penetration_ratio"] = source["penetration_ratio"].fillna(1.0)

    # 指标清单字面口径：当年新增成本金额（不含税）* 0.81 * 穿透比例。
    source["calc_by_indicator"] = source["cost_no_tax"] * 0.81 * source["penetration_ratio"]
    # 诊断口径：报表数表现为在指标清单字面口径基础上额外 / 1.06。
    source["calc_with_extra_1_06"] = source["calc_by_indicator"] / 1.06
    return source


def build_compare(base, source: pd.DataFrame) -> pd.DataFrame:
    """把台账项目级计算结果和半收付项目报表的截止性支出字段对比。"""
    calc = (
        source.groupby("code_exact", as_index=False)[["calc_by_indicator", "calc_with_extra_1_06"]]
        .sum()
    )
    report = base.load_project_report().copy()
    report = report[report["code_exact"].ne("")].copy()
    compare = report[
        ["region", "line", "project_code", "project_name", "code_exact", "cutoff_expense"]
    ].merge(calc, on="code_exact", how="left")
    compare[["calc_by_indicator", "calc_with_extra_1_06"]] = compare[
        ["calc_by_indicator", "calc_with_extra_1_06"]
    ].fillna(0.0)
    compare["diff_by_indicator"] = compare["calc_by_indicator"] - compare["cutoff_expense"]
    compare["diff_with_extra_1_06"] = compare["calc_with_extra_1_06"] - compare["cutoff_expense"]
    return compare


def load_region_source(base, query_df: pd.DataFrame) -> pd.DataFrame:
    """读取区域级需要的 A 项目累计值和 B 区域累计值明细。"""
    raw = pd.read_excel(find_cutoff_ledger(base), header=None, dtype=object)
    code_label = base.u(r"\u7acb\u9879\u7f16\u7801")
    header_idx = next(
        index
        for index in range(min(20, len(raw)))
        if code_label in raw.iloc[index].astype(str).str.strip().tolist()
    )
    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = raw.iloc[header_idx].astype(str).str.strip().tolist()

    type_col = base.u(r"\u7c7b\u578b")
    month_col = base.u(r"\u6570\u636e\u5e74\u6708")
    region_col = base.u(r"\u533a\u57df")
    project_label = base.u(r"\u9879\u76ee")
    region_label = base.u(r"\u533a\u57df")
    cost_col = next(
        column
        for column in df.columns
        if base.u(r"\u5f53\u5e74\u65b0\u589e\u6210\u672c\u91d1\u989d") in str(column)
    )

    query_df = query_df.copy()
    query_df["code_exact"] = query_df["project_code"].astype(str).str.strip().str.upper()
    ratio = query_df[["code_exact", "region", "penetration_ratio"]].drop_duplicates("code_exact")
    profit_non_assess_codes = base.load_profit_non_assess_codes()

    project = df[
        df[type_col].astype(str).str.strip().eq(project_label)
        & df[month_col].astype(str).str[:7].eq(REPORT_MONTH)
    ].copy()
    project["code_exact"] = project[code_label].astype(str).str.strip().str.upper()
    project["code_norm"] = project[code_label].map(base.normalize_code)
    project["cost_no_tax"] = pd.to_numeric(project[cost_col], errors="coerce").fillna(0.0)
    project = project[~project["code_norm"].isin(profit_non_assess_codes)].copy()
    project = project.merge(ratio, on="code_exact", how="left", suffixes=("", "_query"))
    project["penetration_ratio"] = project["penetration_ratio"].fillna(1.0)
    project["calc_by_indicator"] = project["cost_no_tax"] * 0.81 * project["penetration_ratio"]
    project["calc_with_extra_1_06"] = project["calc_by_indicator"] / 1.06
    project["source_part"] = "A_project"

    region = df[
        df[type_col].astype(str).str.strip().eq(region_label)
        & df[month_col].astype(str).str[:7].eq(REPORT_MONTH)
    ].copy()
    region["region"] = region[region_col].astype(str).str.strip()
    region["cost_no_tax"] = pd.to_numeric(region[cost_col], errors="coerce").fillna(0.0)
    region["penetration_ratio"] = 1.0
    region["calc_by_indicator"] = region["cost_no_tax"] * 0.81
    region["calc_with_extra_1_06"] = region["calc_by_indicator"] / 1.06
    region["source_part"] = "B_region"

    return pd.concat(
        [
            project[["region", "source_part", "cost_no_tax", "penetration_ratio", "calc_by_indicator", "calc_with_extra_1_06"]],
            region[["region", "source_part", "cost_no_tax", "penetration_ratio", "calc_by_indicator", "calc_with_extra_1_06"]],
        ],
        ignore_index=True,
    )


def build_region_compare(base, region_source: pd.DataFrame) -> pd.DataFrame:
    """把区域级 A+B 计算结果与半收付区域报表的截止性支出字段对比。"""
    calc = (
        region_source.groupby(["region", "source_part"], as_index=False)[
            ["calc_by_indicator", "calc_with_extra_1_06"]
        ]
        .sum()
        .pivot(index="region", columns="source_part", values=["calc_by_indicator", "calc_with_extra_1_06"])
        .fillna(0.0)
    )
    calc.columns = [f"{metric}_{part}" for metric, part in calc.columns]
    calc = calc.reset_index()
    for column in [
        "calc_by_indicator_A_project",
        "calc_by_indicator_B_region",
        "calc_with_extra_1_06_A_project",
        "calc_with_extra_1_06_B_region",
    ]:
        if column not in calc.columns:
            calc[column] = 0.0
    calc["calc_by_indicator"] = calc["calc_by_indicator_A_project"] + calc["calc_by_indicator_B_region"]
    calc["calc_with_extra_1_06"] = (
        calc["calc_with_extra_1_06_A_project"] + calc["calc_with_extra_1_06_B_region"]
    )

    report = base.load_region_report()[["region", "cutoff_expense"]].copy()
    compare = report.merge(calc, on="region", how="left").fillna(0.0)
    compare["diff_by_indicator"] = compare["calc_by_indicator"] - compare["cutoff_expense"]
    compare["diff_with_extra_1_06"] = compare["calc_with_extra_1_06"] - compare["cutoff_expense"]
    return compare


def summarize(label: str, compare: pd.DataFrame, calc_col: str, diff_col: str) -> dict:
    """输出总览，方便判断是否通过。"""
    diff = compare[diff_col]
    return {
        "check": label,
        "status": "passed" if diff.abs().le(TOLERANCE).all() else "failed",
        "rows": int(len(compare)),
        "nonzero_calc_projects": int(compare[calc_col].abs().gt(TOLERANCE).sum()),
        "nonzero_report_projects": int(compare["cutoff_expense"].abs().gt(TOLERANCE).sum()),
        "mismatch_rows": int(diff.abs().gt(TOLERANCE).sum()),
        "calc_total": float(compare[calc_col].sum()),
        "report_total": float(compare["cutoff_expense"].sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
    }


def main() -> None:
    base = load_base_module()
    indicator_rows = load_indicator_rows(base)
    query_df = base.load_query()
    source = load_cutoff_ledger(base, query_df)
    compare = build_compare(base, source)
    region_source = load_region_source(base, query_df)
    region_compare = build_region_compare(base, region_source)

    source_nonzero = source[source["calc_by_indicator"].abs().gt(TOLERANCE)].copy()
    report_codes = set(compare["code_exact"])
    missing_report = source_nonzero[~source_nonzero["code_exact"].isin(report_codes)]
    missing_query = source[source["missing_query_ratio"]].copy()

    print("[indicator_rows]")
    print(indicator_rows.to_json(force_ascii=False, orient="records"))
    print("[source_summary]")
    print(
        pd.DataFrame(
            [
                {
                    "ledger": find_cutoff_ledger(base).name,
                    "report_month": REPORT_MONTH,
                    "source_rows": int(len(source)),
                    "source_nonzero_cost_rows": int(source["cost_no_tax"].abs().gt(TOLERANCE).sum()),
                    "missing_query_rows": int(len(missing_query)),
                    "source_nonzero_codes_missing_report": int(missing_report["code_exact"].nunique()),
                }
            ]
        ).to_json(force_ascii=False, orient="records")
    )

    print("[checks]")
    print(
        pd.DataFrame(
            [
                summarize("indicator_literal_cost_no_tax_x_0_81_x_ratio", compare, "calc_by_indicator", "diff_by_indicator"),
                summarize("diagnostic_extra_divide_1_06", compare, "calc_with_extra_1_06", "diff_with_extra_1_06"),
                summarize(
                    "region_indicator_literal_A_plus_B",
                    region_compare,
                    "calc_by_indicator",
                    "diff_by_indicator",
                ),
                summarize(
                    "region_diagnostic_extra_divide_1_06_A_plus_B",
                    region_compare,
                    "calc_with_extra_1_06",
                    "diff_with_extra_1_06",
                ),
            ]
        ).to_json(force_ascii=False, orient="records")
    )

    print("[region_compare_detail]")
    print(
        region_compare[
            [
                "region",
                "cutoff_expense",
                "calc_with_extra_1_06_A_project",
                "calc_with_extra_1_06_B_region",
                "calc_with_extra_1_06",
                "diff_with_extra_1_06",
                "calc_by_indicator",
                "diff_by_indicator",
            ]
        ].to_json(force_ascii=False, orient="records")
    )

    print("[indicator_mismatch_top]")
    mismatch = compare[compare["diff_by_indicator"].abs().gt(TOLERANCE)].copy()
    print(
        mismatch.sort_values("diff_by_indicator", key=lambda series: series.abs(), ascending=False)
        .head(30)[
            [
                "region",
                "line",
                "project_code",
                "project_name",
                "cutoff_expense",
                "calc_by_indicator",
                "diff_by_indicator",
                "calc_with_extra_1_06",
                "diff_with_extra_1_06",
            ]
        ]
        .to_json(force_ascii=False, orient="records")
    )

    print("[blocking_rows]")
    print(
        pd.DataFrame(
            [
                {
                    "missing_query_rows": int(len(missing_query)),
                    "source_nonzero_codes_missing_report": int(missing_report["code_exact"].nunique()),
                }
            ]
        ).to_json(force_ascii=False, orient="records")
    )


if __name__ == "__main__":
    main()
