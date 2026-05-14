from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)


def load_base_module():
    """复用半收付归母净利润脚本里的报表映射和计划数计算逻辑。"""
    script_path = ROOT / "validate_half_cash_attributable_profit.py"
    spec = importlib.util.spec_from_file_location("half_cash_base", script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def print_record(label: str, record: dict) -> None:
    """统一输出 JSON，方便后续用脚本解析结果。"""
    print(f"{label} {json.dumps(record, ensure_ascii=False)}")


def summarize_diff(base, label: str, calc: pd.Series, report: pd.Series, tolerance: float | None = None) -> dict:
    """汇总计算值与报表值的差异。"""
    tolerance = base.TOLERANCE if tolerance is None else tolerance
    diff = calc - report
    return {
        "status": "passed" if (diff.abs() <= tolerance).all() else "failed",
        "rows": int(len(diff)),
        "mismatch_rows": int((diff.abs() > tolerance).sum()),
        "calc_total": float(calc.sum()),
        "report_total": float(report.sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "check": label,
        "tolerance": float(tolerance),
    }


def compare_project_plan_metrics(base, project_df: pd.DataFrame, query_df: pd.DataFrame) -> None:
    """项目维度计划数：用计划台账和摊销比例配置复算后对比项目报表。"""
    plan_calc = base.compute_plan_values(base.load_plan_ledger(query_df))
    project_plan = project_df[
        ["region", "project_code", "project_name", "code_norm", "plan_capital", "plan_smart", "plan_quality"]
    ].merge(plan_calc, on="code_norm", how="left", suffixes=("_report", "_calc"))
    project_plan = project_plan.fillna(0.0)

    for component in ["plan_capital", "plan_smart", "plan_quality"]:
        print_record(
            f"PROJECT_{component}",
            summarize_diff(
                base,
                f"project_{component}",
                project_plan[f"{component}_calc"],
                project_plan[f"{component}_report"],
            ),
        )


def compare_region_plan_metrics(base, project_df: pd.DataFrame, region_df: pd.DataFrame) -> None:
    """区域维度计划数：指标清单单月口径有利润考核项目条件，这里排除利润类非考核项目后汇总。"""
    profit_non_assess_codes = base.load_profit_non_assess_codes()
    project_scope = project_df.loc[~project_df["code_norm"].isin(profit_non_assess_codes)].copy()
    rollup = project_scope.groupby("region", as_index=False)[["plan_capital", "plan_smart", "plan_quality"]].sum()
    region_plan = region_df[["region", "plan_capital", "plan_smart", "plan_quality"]].merge(
        rollup, on="region", how="left", suffixes=("_report", "_calc")
    )
    region_plan = region_plan.fillna(0.0)

    for component in ["plan_capital", "plan_smart", "plan_quality"]:
        print_record(
            f"REGION_{component}",
            summarize_diff(
                base,
                f"region_{component}",
                region_plan[f"{component}_calc"],
                region_plan[f"{component}_report"],
            ),
        )


def build_project_minority_from_1_5_2(base, project_df: pd.DataFrame, query_df: pd.DataFrame) -> pd.DataFrame:
    """项目维度少数股东损益：按 1.5.2 半收付净利润和项目穿透比例复算。"""
    source = base.load_1_5_2_project_profit()
    minority = project_df[
        ["region", "project_code", "project_name", "code_exact", "code_norm", "half_net_profit", "minority_loss"]
    ].merge(query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"), on="code_norm", how="left")
    minority["penetration_ratio"] = minority["penetration_ratio"].fillna(1.0)

    minority = minority.merge(source[["code_exact", "half_net_profit_source"]], on="code_exact", how="left")
    fallback = source.groupby("code_norm", as_index=False).agg(
        source_exact_count=("code_exact", "nunique"),
        fallback_half_net_profit_source=("half_net_profit_source", "sum"),
    )
    minority = minority.merge(fallback, on="code_norm", how="left")

    # 报表存在少量无字母前缀编码；只有唯一源编码时才用标准化编码兜底。
    missing_exact = minority["half_net_profit_source"].isna()
    no_prefix_code = ~minority["code_exact"].astype(str).str.match(r"^[A-Z]")
    use_fallback = missing_exact & no_prefix_code & minority["source_exact_count"].eq(1)
    minority.loc[use_fallback, "half_net_profit_source"] = minority.loc[use_fallback, "fallback_half_net_profit_source"]

    minority["half_net_profit_source"] = minority["half_net_profit_source"].fillna(0.0)
    minority["calc_from_1_5_2"] = minority["half_net_profit_source"] * (1 - minority["penetration_ratio"])
    minority["calc_from_report_profit"] = minority["half_net_profit"] * (1 - minority["penetration_ratio"])
    return minority


def compare_project_minority(base, project_df: pd.DataFrame, query_df: pd.DataFrame) -> None:
    """分别输出源表复算和报表半收付净利润反算，便于定位源表差异。"""
    minority = build_project_minority_from_1_5_2(base, project_df, query_df)
    print_record(
        "PROJECT_minority_from_report_profit",
        summarize_diff(base, "project_minority_from_report_profit", minority["calc_from_report_profit"], minority["minority_loss"]),
    )
    print_record(
        "PROJECT_minority_from_1_5_2",
        summarize_diff(base, "project_minority_from_1_5_2", minority["calc_from_1_5_2"], minority["minority_loss"]),
    )

    bad = minority.loc[
        (minority["calc_from_1_5_2"] - minority["minority_loss"]).abs() > base.TOLERANCE,
        [
            "region",
            "project_code",
            "project_name",
            "half_net_profit",
            "half_net_profit_source",
            "penetration_ratio",
            "minority_loss",
            "calc_from_1_5_2",
        ],
    ].copy()
    bad["diff"] = bad["calc_from_1_5_2"] - bad["minority_loss"]
    print_record("PROJECT_minority_from_1_5_2_bad_count", {"rows": int(len(bad))})
    if not bad.empty:
        print(bad.head(20).to_json(orient="records", force_ascii=False))


def compare_region_minority_rollup(base, project_df: pd.DataFrame, region_df: pd.DataFrame) -> None:
    """区域少数股东损益：先做项目报表汇总检查；该指标清单高维口径不是简单项目汇总。"""
    rollup = project_df.groupby("region", as_index=False)["minority_loss"].sum()
    region_minority = region_df[["region", "minority_loss"]].merge(
        rollup, on="region", how="left", suffixes=("_report", "_calc")
    )
    region_minority = region_minority.fillna(0.0)
    print_record(
        "REGION_minority_project_rollup",
        summarize_diff(
            base,
            "region_minority_project_rollup",
            region_minority["minority_loss_calc"],
            region_minority["minority_loss_report"],
        ),
    )


def compare_region_minority_formula(base, region_df: pd.DataFrame) -> None:
    """区域少数股东损益：按用户确认公式，从区域管报归母净利润报表直接取后两项。"""
    mgmt_path = base.find_workbook(base.u(r"\u7ba1\u62a5\u5f52\u6bcd\u51c0\u5229\u6da6"), base.REGION)
    mgmt_df = pd.read_excel(mgmt_path)
    columns = list(mgmt_df.columns)
    mgmt_df = mgmt_df.rename(
        columns={
            columns[1]: "region",
            columns[2]: "management_net_profit",
            columns[5]: "management_minority_loss",
        }
    )
    mgmt_df["management_net_profit"] = pd.to_numeric(mgmt_df["management_net_profit"], errors="coerce").fillna(0.0)
    mgmt_df["management_minority_loss"] = pd.to_numeric(
        mgmt_df["management_minority_loss"], errors="coerce"
    ).fillna(0.0)

    check_df = region_df[["region", "half_net_profit", "minority_loss"]].merge(
        mgmt_df[["region", "management_net_profit", "management_minority_loss"]],
        on="region",
        how="left",
    )
    check_df = check_df.fillna(0.0)

    # 公式：区域级少数股东损益_半收付 = 半收付净利润 * 少数股东损益_管报 / 管报净利润。
    check_df["calc_minority_loss"] = check_df.apply(
        lambda row: 0.0
        if abs(row["management_net_profit"]) < 1e-12
        else row["half_net_profit"] * row["management_minority_loss"] / row["management_net_profit"],
        axis=1,
    )
    print_record(
        "REGION_minority_management_formula",
        summarize_diff(
            base,
            "region_minority_management_formula",
            check_df["calc_minority_loss"],
            check_df["minority_loss"],
            tolerance=1e-4,
        ),
    )


def main() -> None:
    base = load_base_module()
    project_df = base.load_project_report()
    region_df = base.load_region_report()
    query_df = base.load_query()

    compare_project_plan_metrics(base, project_df, query_df)
    compare_region_plan_metrics(base, project_df, region_df)
    compare_project_minority(base, project_df, query_df)
    compare_region_minority_formula(base, region_df)
    compare_region_minority_rollup(base, project_df, region_df)


if __name__ == "__main__":
    main()
