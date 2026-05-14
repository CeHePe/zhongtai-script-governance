from __future__ import annotations

import importlib.util
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "2025-12"
TOLERANCE = 1e-6


def load_base_module():
    """复用半收付归母净利润校验脚本里的报表读取和编码标准化函数。"""
    script_path = ROOT / "validate_half_cash_attributable_profit.py"
    spec = importlib.util.spec_from_file_location("half_cash_base", script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def month_key(value: object) -> str:
    """把 Excel 日期、202512、2025-12 等月份格式统一成 yyyy-mm。"""
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m")
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    if len(text) == 6 and text.isdigit():
        return f"{text[:4]}-{text[4:]}"
    return text[:7]


def as_number(series: pd.Series) -> pd.Series:
    """把金额列安全转换为数值，空值按 0 处理。"""
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def find_workbook_by_tokens(*tokens: str) -> Path:
    """用 Python 匹配中文文件名，避免 PowerShell 编码影响。"""
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def load_indicator_rows(base) -> pd.DataFrame:
    """读取指标清单中项目和区域维度的精确指标行。"""
    indicator = pd.read_excel(base.find_workbook(base.INDICATOR), sheet_name=0, dtype=object)
    cols = list(indicator.columns)
    metric = base.u(r"\u5f53\u671f\u5e94\u6536\u6c34\u7535\u8d39_\u534a\u6536\u4ed8_\u636e\u5b9e\u5206\u644a")
    target_dims = [base.u(r"\u9879\u76ee"), base.u(r"\u533a\u57df")]
    rows = indicator[
        indicator[cols[3]].astype(str).eq(metric)
        & indicator[cols[4]].astype(str).isin(target_dims)
    ][[cols[3], cols[4], cols[5], cols[8], cols[10], cols[12]]].fillna("")
    rows = rows.copy()
    rows.insert(0, "excel_row", rows.index + 2)
    rows.columns = ["excel_row", "metric_name", "dimension", "cycle", "method", "source_table", "logic"]
    return rows


def load_allocated_water_source(base, query_df: pd.DataFrame) -> pd.DataFrame:
    """读取垫支水电费物业分摊据实分摊台账，计算项目级当期应收据实分摊。"""
    ledger_name = base.u(r"\u57ab\u652f\u6c34\u7535\u8d39\u7269\u4e1a\u5206\u644a\u636e\u5b9e\u5206\u644a")
    raw = pd.read_excel(find_workbook_by_tokens(ledger_name), header=None)
    df = raw.iloc[5:].reset_index(drop=True).copy()
    df.columns = [f"c{i}" for i in range(df.shape[1])]
    df = df[df["c0"].map(month_key).eq(REPORT_MONTH)].copy()
    df["code_norm"] = df["c1"].map(base.normalize_code)

    ratio = query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm")
    df = df.merge(ratio, on="code_norm", how="left")
    df["missing_ratio"] = df["penetration_ratio"].isna()
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)

    # 指标清单第 2103 行口径：取【当期应收】* 穿透比例。
    df["current_receivable_alloc_calc"] = as_number(df["c4"]) * df["penetration_ratio"]
    project = (
        df.groupby("code_norm", as_index=False)["current_receivable_alloc_calc"]
        .sum()
    )
    return df, project


def compare_project(base, project_report: pd.DataFrame, source_project: pd.DataFrame) -> pd.DataFrame:
    """项目维度：源台账计算值与半收付项目报表字段对比。"""
    compare = project_report[
        ["region", "line", "project_code", "project_name", "code_norm", "water_current_receivable_alloc"]
    ].merge(source_project, on="code_norm", how="left")
    compare["current_receivable_alloc_calc"] = compare["current_receivable_alloc_calc"].fillna(0.0)
    compare["diff"] = compare["current_receivable_alloc_calc"] - compare["water_current_receivable_alloc"]
    return compare


def compare_region(
    project_report: pd.DataFrame,
    region_report: pd.DataFrame,
    source_project: pd.DataFrame,
    profit_non_assess_codes: set[str],
) -> pd.DataFrame:
    """区域维度：从项目指标库汇总，排除利润类非考核项目。"""
    source = (
        project_report[["region", "code_norm"]]
        .merge(source_project, on="code_norm", how="left")
        .fillna({"current_receivable_alloc_calc": 0.0})
    )
    source = source[~source["code_norm"].isin(profit_non_assess_codes)].copy()
    rollup = source.groupby("region", as_index=False)["current_receivable_alloc_calc"].sum()
    compare = region_report[["region", "water_current_receivable_alloc"]].merge(rollup, on="region", how="left")
    compare["current_receivable_alloc_calc"] = compare["current_receivable_alloc_calc"].fillna(0.0)
    compare["diff"] = compare["current_receivable_alloc_calc"] - compare["water_current_receivable_alloc"]
    return compare


def summarize(label: str, compare: pd.DataFrame, report_col: str, calc_col: str) -> dict:
    """输出通过情况和关键汇总值。"""
    diff = compare[calc_col] - compare[report_col]
    return {
        "check": label,
        "status": "passed" if diff.abs().le(TOLERANCE).all() else "failed",
        "rows": int(len(compare)),
        "nonzero_report_rows": int(compare[report_col].abs().gt(TOLERANCE).sum()),
        "nonzero_calc_rows": int(compare[calc_col].abs().gt(TOLERANCE).sum()),
        "mismatch_rows": int(diff.abs().gt(TOLERANCE).sum()),
        "calc_total": float(compare[calc_col].sum()),
        "report_total": float(compare[report_col].sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
    }


def main() -> None:
    base = load_base_module()
    indicator_rows = load_indicator_rows(base)
    query_df = base.load_query()
    project_report = base.load_project_report()
    region_report = base.load_region_report()
    region_report = region_report[
        region_report["region"].notna()
        & region_report["region"].astype(str).str.strip().ne("")
        & ~region_report["region"].astype(str).str.strip().isin(["0", "0.0"])
    ].copy()
    profit_non_assess_codes = base.load_profit_non_assess_codes()
    source_detail, source_project = load_allocated_water_source(base, query_df)

    project_compare = compare_project(base, project_report, source_project)
    region_compare = compare_region(project_report, region_report, source_project, profit_non_assess_codes)

    print("[indicator_rows]")
    print(indicator_rows.to_json(force_ascii=False, orient="records"))

    print("[source_summary]")
    print(
        pd.DataFrame(
            [
                {
                    "ledger": find_workbook_by_tokens(
                        base.u(r"\u57ab\u652f\u6c34\u7535\u8d39\u7269\u4e1a\u5206\u644a\u636e\u5b9e\u5206\u644a")
                    ).name,
                    "report_month": REPORT_MONTH,
                    "source_rows": int(len(source_detail)),
                    "source_nonzero_rows": int(source_detail["current_receivable_alloc_calc"].abs().gt(TOLERANCE).sum()),
                    "source_projects": int(source_project["code_norm"].nunique()),
                    "missing_ratio_rows": int(source_detail["missing_ratio"].sum()),
                    "profit_non_assess_codes": int(len(profit_non_assess_codes)),
                }
            ]
        ).to_json(force_ascii=False, orient="records")
    )

    print("[checks]")
    print(
        pd.DataFrame(
            [
                summarize(
                    "project_current_receivable_alloc",
                    project_compare,
                    "water_current_receivable_alloc",
                    "current_receivable_alloc_calc",
                ),
                summarize(
                    "region_current_receivable_alloc",
                    region_compare,
                    "water_current_receivable_alloc",
                    "current_receivable_alloc_calc",
                ),
            ]
        ).to_json(force_ascii=False, orient="records")
    )

    print("[project_mismatch_top]")
    project_mismatch = project_compare[project_compare["diff"].abs().gt(TOLERANCE)].copy()
    print(
        project_mismatch.sort_values("diff", key=lambda series: series.abs(), ascending=False)
        .head(30)
        .to_json(force_ascii=False, orient="records")
    )

    print("[region_detail]")
    print(region_compare.to_json(force_ascii=False, orient="records"))


if __name__ == "__main__":
    main()
