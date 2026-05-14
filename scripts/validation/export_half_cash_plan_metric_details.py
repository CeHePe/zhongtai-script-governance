from __future__ import annotations

import importlib.util
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = find_project_root(__file__)
REPORT_YEAR = 2025
REPORT_MONTH = "2025-12"
OUTPUT_NAME = "三大计划数测试明细.xlsx"


def load_base_module():
    """复用已有半收付归母净利润脚本里的文件查找、编码标准化和报表读取函数。"""
    script_path = SCRIPT_DIR / "validate_half_cash_attributable_profit.py"
    spec = importlib.util.spec_from_file_location("half_cash_base", script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def load_plan_raw(base) -> pd.DataFrame:
    """读取三大计划数台账，并补充项目查询中的穿透比例和区域属性。"""
    plan_path = next(
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and base.PLAN in path.name and base.RATIO not in path.name
    )
    query_df = base.load_query()
    df = pd.read_excel(plan_path)
    df = df.rename(
        columns={
            df.columns[0]: "ledger_no",
            df.columns[1]: "ledger_region",
            df.columns[2]: "project_code",
            df.columns[3]: "project_name",
            df.columns[4]: "plan_type",
            df.columns[5]: "matter_type",
            df.columns[6]: "tax_included_investment",
            df.columns[7]: "years_text",
            df.columns[8]: "start_month",
            df.columns[9]: "end_month",
            df.columns[10]: "meeting_time",
            df.columns[11]: "plan_status",
            df.columns[12]: "early_end_month",
        }
    )
    df["code_norm"] = df["project_code"].map(base.normalize_code)
    df["tax_included_investment"] = pd.to_numeric(df["tax_included_investment"], errors="coerce").fillna(0.0)
    df["years"] = pd.to_numeric(df["years_text"].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    df["start_date"] = pd.to_datetime(df["start_month"].astype(str) + "-01", errors="coerce")
    df["end_date"] = pd.to_datetime(df["end_month"].astype(str) + "-01", errors="coerce")
    df["early_end_date"] = pd.to_datetime(df["early_end_month"].astype(str) + "-01", errors="coerce")

    query_cols = [
        "code_norm",
        "region",
        "business_attr",
        "project_level",
        "project_status",
        "penetration_ratio",
    ]
    df = df.merge(query_df[query_cols].drop_duplicates("code_norm"), on="code_norm", how="left")
    df["penetration_ratio"] = pd.to_numeric(df["penetration_ratio"], errors="coerce").fillna(1.0)
    return df


def build_monthly_detail(base, plan_df: pd.DataFrame) -> pd.DataFrame:
    """把计划台账逐行展开到报表年度各月，并计算每月金额。"""
    ratio_map = base.load_plan_ratio_map()
    capital_actual_df = base.load_capital_actual_raw()
    month_periods = list(pd.period_range(f"{REPORT_YEAR}-01", REPORT_MONTH, freq="M"))
    early_end = base.u(r"\u63d0\u524d\u7ed3\u675f")
    terminated = base.TERMINATED
    capital = base.u(r"\u5e26\u8d44\u644a\u9500")
    smart = base.u(r"\u667a\u80fd\u5316\u6574\u6539")
    quality = base.u(r"\u8d28\u6548\u63d0\u5347")
    component_map = {
        capital: "带资摊销调整计划数",
        smart: "智能化整改摊销调整计划数",
        quality: "质效提升计划数",
    }

    records: list[dict] = []
    for _, row in plan_df.iterrows():
        if row["plan_type"] not in component_map or pd.isna(row["start_date"]) or pd.isna(row["years"]):
            continue

        start_period = base.month_to_period(row["start_date"])
        end_period = base.month_to_period(row["end_date"])
        early_end_period = base.month_to_period(row["early_end_date"])
        effective_end = early_end_period if row["plan_status"] in {early_end, terminated} and early_end_period is not None else end_period
        for month_period in month_periods:
            if start_period is None or month_period < start_period:
                continue
            if effective_end is not None and month_period > effective_end:
                continue

            amort_year_index = base.amort_year_index(month_period, start_period)
            amort_year_no = amort_year_index + 1
            annual_ratio = ""
            divide_tax = ""
            trigger_month = effective_end.strftime("%Y-%m") if effective_end is not None else ""
            prev_actual_base = ""
            prev_plan_base = ""
            current_year_prev_plan_base = ""

            if row["plan_type"] == capital:
                ratio_list = ratio_map.get(int(row["years"]), [])
                annual_ratio_value = ratio_list[amort_year_index] if 0 <= amort_year_index < len(ratio_list) else 0.0
                annual_ratio = annual_ratio_value
                divide_tax = "是"
                if effective_end is not None and month_period == effective_end:
                    prev_year_end = pd.Period(f"{effective_end.year - 1}-12", freq="M")
                    trigger_prev = effective_end - 1
                    year_start = pd.Period(f"{effective_end.year}-01", freq="M")
                    prev_plan_base = sum(
                        base.capital_month_base(row["tax_included_investment"], ratio_map, int(row["years"]), period, start_period)
                        for period in pd.period_range(start_period, min(prev_year_end, end_period), freq="M")
                    ) if end_period is not None and prev_year_end >= start_period else 0.0
                    current_year_prev_plan_base = sum(
                        base.capital_month_base(row["tax_included_investment"], ratio_map, int(row["years"]), period, start_period)
                        for period in pd.period_range(max(year_start, start_period), min(trigger_prev, end_period), freq="M")
                    ) if end_period is not None and trigger_prev >= max(year_start, start_period) else 0.0
                    prev_actual_base = base.capital_prev_actual_base(row["code_norm"], prev_year_end, capital_actual_df)
                    monthly_amount = (
                        (prev_actual_base - prev_plan_base - current_year_prev_plan_base)
                        * 0.81
                        / 1.06
                        * row["penetration_ratio"]
                    )
                    formula = "（累计发生数 - 累计计划数 - 当年1月至触发月上一月计划数） * 0.81 / 1.06 * 项目穿透比例"
                else:
                    monthly_amount = (
                        row["tax_included_investment"]
                        * annual_ratio_value
                        * 0.81
                        / 1.06
                        / 12
                        * row["penetration_ratio"]
                    )
                    formula = "含税计划投资金额 * 年度摊销比例 * 0.81 / 1.06 / 12 * 项目穿透比例"
            elif row["plan_type"] == smart:
                divide_tax = "否"
                monthly_amount = (
                    row["tax_included_investment"] / row["years"] / 12 * row["penetration_ratio"] * 0.81
                )
                formula = "含税计划投资金额 / 分摊年限 / 12 * 项目穿透比例 * 0.81"
            else:
                divide_tax = "是"
                monthly_amount = (
                    row["tax_included_investment"] / row["years"] / 12 * row["penetration_ratio"] * 0.81 / 1.06
                )
                formula = "含税计划投资金额 / 分摊年限 / 12 * 项目穿透比例 * 0.81 / 1.06"

            records.append(
                {
                    "台账序号": row["ledger_no"],
                    "台账区域": row["ledger_region"],
                    "项目查询区域": row["region"],
                    "立项编码": row["project_code"],
                    "项目名称": row["project_name"],
                    "业务属性": row["business_attr"],
                    "项目等级": row["project_level"],
                    "项目状态": row["project_status"],
                    "摊销类型": row["plan_type"],
                    "事项类型": row["matter_type"],
                    "指标": component_map[row["plan_type"]],
                    "计划投资金额_含税": row["tax_included_investment"],
                    "分摊年限": row["years"],
                    "摊销开始日期": row["start_month"],
                    "摊销结束日期": row["end_month"],
                    "计划状态": row["plan_status"],
                    "提前结束年月": row["early_end_month"],
                    "有效结束月份": effective_end.strftime("%Y-%m") if pd.notna(effective_end) else "",
                    "触发月份": trigger_month,
                    "计算月份": month_period.strftime("%Y-%m"),
                    "摊销第几年": amort_year_no,
                    "年度摊销比例": annual_ratio,
                    "累计发生数": prev_actual_base,
                    "累计计划数": prev_plan_base,
                    "当年1月至触发月上一月计划数": current_year_prev_plan_base,
                    "是否除以1.06": divide_tax,
                    "项目穿透比例": row["penetration_ratio"],
                    "计算公式": formula,
                    "月度计算金额": monthly_amount,
                    "code_norm": row["code_norm"],
                }
            )

    return pd.DataFrame(records)


def pivot_metric_amount(df: pd.DataFrame, index_cols: list[str], amount_col: str) -> pd.DataFrame:
    """把三项指标从长表转成宽表。"""
    wide = (
        df.groupby(index_cols + ["指标"], as_index=False)[amount_col]
        .sum()
        .pivot(index=index_cols, columns="指标", values=amount_col)
        .reset_index()
        .fillna(0.0)
    )
    for metric in ["带资摊销调整计划数", "智能化整改摊销调整计划数", "质效提升计划数"]:
        if metric not in wide.columns:
            wide[metric] = 0.0
    return wide


def build_project_compare(base, detail_df: pd.DataFrame, project_report: pd.DataFrame) -> pd.DataFrame:
    """项目层：明细汇总后与项目报表三项计划数对比。"""
    project_calc = pivot_metric_amount(detail_df, ["code_norm"], "月度计算金额").rename(
        columns={
            "带资摊销调整计划数": "带资_计算",
            "智能化整改摊销调整计划数": "智能化_计算",
            "质效提升计划数": "质效_计算",
        }
    )
    compare = project_report[
        ["region", "project_code", "project_name", "code_norm", "plan_capital", "plan_smart", "plan_quality"]
    ].merge(project_calc, on="code_norm", how="left")
    compare = compare.fillna(0.0)
    compare = compare.rename(
        columns={
            "region": "报表区域",
            "project_code": "立项编码",
            "project_name": "项目名称",
            "plan_capital": "带资_报表",
            "plan_smart": "智能化_报表",
            "plan_quality": "质效_报表",
        }
    )
    for prefix in ["带资", "智能化", "质效"]:
        compare[f"{prefix}_差异"] = compare[f"{prefix}_计算"] - compare[f"{prefix}_报表"]
    return compare


def build_region_compare(
    base, detail_df: pd.DataFrame, project_report: pd.DataFrame, region_report: pd.DataFrame
) -> pd.DataFrame:
    """区域层：项目维度验证通过后，按项目报表区域汇总并与区域报表对比。"""
    profit_non_assess_codes = base.load_profit_non_assess_codes()
    project_region = project_report[["code_norm", "region"]].drop_duplicates("code_norm").rename(
        columns={"region": "项目报表区域"}
    )
    region_source = detail_df.merge(project_region, on="code_norm", how="left")
    region_source = region_source.loc[~region_source["code_norm"].isin(profit_non_assess_codes)].copy()
    region_calc = pivot_metric_amount(region_source, ["项目报表区域"], "月度计算金额").rename(
        columns={
            "项目报表区域": "区域",
            "带资摊销调整计划数": "带资_计算",
            "智能化整改摊销调整计划数": "智能化_计算",
            "质效提升计划数": "质效_计算",
        }
    )
    compare = region_report[["region", "plan_capital", "plan_smart", "plan_quality"]].merge(
        region_calc,
        left_on="region",
        right_on="区域",
        how="left",
    )
    compare = compare.fillna(0.0)
    compare = compare.rename(
        columns={
            "region": "报表区域",
            "plan_capital": "带资_报表",
            "plan_smart": "智能化_报表",
            "plan_quality": "质效_报表",
        }
    )
    for prefix in ["带资", "智能化", "质效"]:
        compare[f"{prefix}_差异"] = compare[f"{prefix}_计算"] - compare[f"{prefix}_报表"]
    return compare


def build_query_region_diagnostic(base, detail_df: pd.DataFrame, project_report: pd.DataFrame) -> pd.DataFrame:
    """诊断项：列出项目查询区域和项目报表区域不一致且存在计划数金额的项目。"""
    project_region = project_report[["code_norm", "region", "project_code", "project_name"]].drop_duplicates("code_norm")
    metric_sum = pivot_metric_amount(detail_df, ["code_norm", "项目查询区域"], "月度计算金额").rename(
        columns={
            "带资摊销调整计划数": "带资_计算",
            "智能化整改摊销调整计划数": "智能化_计算",
            "质效提升计划数": "质效_计算",
        }
    )
    diagnostic = metric_sum.merge(project_region, on="code_norm", how="left")
    diagnostic = diagnostic.rename(
        columns={
            "region": "项目报表区域",
            "project_code": "立项编码",
            "project_name": "项目名称",
        }
    )
    amount_cols = ["带资_计算", "智能化_计算", "质效_计算"]
    diagnostic = diagnostic.loc[
        diagnostic["项目查询区域"].astype(str).ne(diagnostic["项目报表区域"].astype(str))
        & diagnostic[amount_cols].abs().sum(axis=1).gt(1e-6)
    ].copy()
    return diagnostic[["立项编码", "项目名称", "项目查询区域", "项目报表区域"] + amount_cols]


def build_region_project_rollup(base, project_report: pd.DataFrame, region_report: pd.DataFrame) -> pd.DataFrame:
    """诊断项：区域报表是否等于项目报表剔除利润类非考核项目后的汇总。"""
    profit_non_assess_codes = base.load_profit_non_assess_codes()
    project_scope = project_report.loc[~project_report["code_norm"].isin(profit_non_assess_codes)].copy()
    rollup = (
        project_scope.groupby("region", as_index=False)[["plan_capital", "plan_smart", "plan_quality"]]
        .sum()
        .rename(
            columns={
                "region": "区域",
                "plan_capital": "带资_项目汇总",
                "plan_smart": "智能化_项目汇总",
                "plan_quality": "质效_项目汇总",
            }
        )
    )
    compare = region_report[["region", "plan_capital", "plan_smart", "plan_quality"]].merge(
        rollup,
        left_on="region",
        right_on="区域",
        how="left",
    )
    compare = compare.fillna(0.0)
    compare = compare.rename(
        columns={
            "region": "报表区域",
            "plan_capital": "带资_报表",
            "plan_smart": "智能化_报表",
            "plan_quality": "质效_报表",
        }
    )
    for prefix in ["带资", "智能化", "质效"]:
        compare[f"{prefix}_差异"] = compare[f"{prefix}_项目汇总"] - compare[f"{prefix}_报表"]
    return compare


def build_scope_note(base) -> pd.DataFrame:
    """输出本次测试口径，便于和明细一起审阅。"""
    return pd.DataFrame(
        [
            {
                "指标": "带资摊销调整计划数",
                "项目口径": "非触发月按 含税计划投资金额 * 年度摊销比例 * 0.81 / 1.06 / 12 * 项目穿透比例；摊销结束月/提前结束月按（累计发生数 - 累计计划数 - 当年1月至触发月上一月计划数）* 0.81 / 1.06 * 项目穿透比例；按2025年逐月累计",
                "区域口径": "项目明细按项目查询区域汇总，并排除利润类非考核项目",
            },
            {
                "指标": "智能化整改摊销调整计划数",
                "项目口径": "含税计划投资金额 / 分摊年限 / 12 * 项目穿透比例 * 0.81，按2025年有效月份累计",
                "区域口径": "项目明细按项目查询区域汇总，并排除利润类非考核项目",
            },
            {
                "指标": "质效提升计划数",
                "项目口径": "含税计划投资金额 / 分摊年限 / 12 * 项目穿透比例 * 0.81 / 1.06，按2025年有效月份累计",
                "区域口径": "项目明细按项目查询区域汇总，并排除利润类非考核项目",
            },
        ]
    )


def main() -> None:
    base = load_base_module()
    plan_df = load_plan_raw(base)
    project_report = base.load_project_report()
    region_report = base.load_region_report()
    detail_df = build_monthly_detail(base, plan_df)
    project_compare = build_project_compare(base, detail_df, project_report)
    region_compare = build_region_compare(base, detail_df, project_report, region_report)
    region_project_rollup = build_region_project_rollup(base, project_report, region_report)
    query_region_diagnostic = build_query_region_diagnostic(base, detail_df, project_report)
    output_path = ROOT / OUTPUT_NAME

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        build_scope_note(base).to_excel(writer, sheet_name="口径说明", index=False)
        detail_df.to_excel(writer, sheet_name="逐月计算明细", index=False)
        project_compare.to_excel(writer, sheet_name="项目汇总对比", index=False)
        region_compare.to_excel(writer, sheet_name="区域源表汇总对比", index=False)
        region_project_rollup.to_excel(writer, sheet_name="区域项目汇总诊断", index=False)
        query_region_diagnostic.to_excel(writer, sheet_name="项目查询区域映射差异", index=False)

    print(output_path)
    print(f"monthly_detail_rows={len(detail_df)}")
    print(f"project_compare_rows={len(project_compare)}")
    print(f"region_compare_rows={len(region_compare)}")


if __name__ == "__main__":
    main()
