from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

import validate_attributable_profit as base
import validate_mgmt_profit_project_region_three_metrics as three


REPORT_YM = "202512"
REPORT_YEAR = 2025
TOLERANCE = 1e-6

LINE_COLUMNS = [
    "idx",
    "region",
    "line",
    "net_profit",
    "d_exit_net_profit",
    "impairment",
    "minority",
    "plan_capital",
    "actual_capital",
    "plan_smart",
    "actual_smart",
    "plan_quality",
    "actual_quality",
    "ebt_cost",
    "other_adjust",
    "discount",
    "cutoff",
    "dev_incentive",
    "jl_jj_restore",
    "interest",
    "satellite",
    "pro_d_exit_sum",
    "prev_region_perf",
    "curr_region_perf",
    "attributable",
]

ROLLUP_METRICS = [
    ("带资摊销计划数", "plan_capital"),
    ("质效提升计划数", "plan_quality"),
    ("智能化整改计划数", "plan_smart"),
    ("带资摊销实际发生数", "actual_capital"),
    ("智能化整改实际发生数", "actual_smart"),
    ("质效提升实际发生数", "actual_quality"),
    ("年度地产关联方EBT成本", "ebt_cost"),
    ("其他考核管报归母净利润", "other_adjust"),
    ("综管折让金额", "discount"),
    ("截止性收支", "cutoff"),
    ("属地前区域业绩调整", "prev_region_perf"),
    ("当前区域业绩调整", "curr_region_perf"),
    ("少数股东损益", "minority"),
]


def jprint(label: str, data: object) -> None:
    print(f"[{label}]")
    print(json.dumps(data, ensure_ascii=False, default=str))


def load_line_report() -> pd.DataFrame:
    df = pd.read_excel(base.find_workbook(base.ATTRIBUTABLE_TOKEN, base.u(r"\u533a\u57df"), base.u(r"\u6761\u7ebf")))
    df.columns = LINE_COLUMNS
    for column in LINE_COLUMNS[3:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def load_flags() -> pd.DataFrame:
    query = three.load_query().copy()
    flags = query[["code_norm", "project_level", "project_status", "exit_date", "penetration_ratio"]].copy()
    flags["penetration_ratio"] = pd.to_numeric(flags["penetration_ratio"], errors="coerce").fillna(1.0)

    exit_dt = pd.to_datetime(flags["exit_date"], errors="coerce")
    exit_ym = exit_dt.dt.strftime("%Y%m").fillna("")
    exit_year = exit_dt.dt.year.fillna(0).astype(int)

    flags["is_d_current_year_exit"] = (
        flags["project_level"].astype(str).str.startswith("D")
        & flags["project_status"].astype(str).eq(base.u(r"\u5df2\u64a4\u573a"))
        & exit_ym.ne("")
        & exit_ym.le(REPORT_YM)
        & exit_year.eq(REPORT_YEAR)
    )
    flags["is_d_any_exit"] = (
        flags["project_level"].astype(str).str.startswith("D")
        & flags["project_status"].astype(str).eq(base.u(r"\u5df2\u64a4\u573a"))
    )
    # 项目查询存在重复立项编码；区域条线逻辑只需要每个编码一条辅助属性。
    return flags.drop_duplicates("code_norm", keep="first")


def summarize(compare: pd.DataFrame, report_col: str, calc_col: str) -> dict:
    diff = compare[calc_col] - compare[report_col]
    return {
        "status": "通过" if diff.abs().le(TOLERANCE).all() else "失败",
        "mismatch_rows": int(diff.abs().gt(TOLERANCE).sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "report_total": float(compare[report_col].sum()),
        "calc_total": float(compare[calc_col].sum()),
        "diff_total": float(diff.sum()),
    }


def rollup_non_d_current_year_exit(project_df: pd.DataFrame, report_col: str) -> pd.DataFrame:
    return (
        project_df.loc[~project_df["is_d_current_year_exit"].fillna(False)]
        .groupby(["region", "line"], as_index=False)[report_col]
        .sum()
        .rename(columns={report_col: "calc"})
    )


def compare_rollup(line_df: pd.DataFrame, project_df: pd.DataFrame, metric_name: str, report_col: str) -> tuple[dict, list[dict]]:
    calc = rollup_non_d_current_year_exit(project_df, report_col)
    compare = line_df[["region", "line", report_col]].merge(calc, on=["region", "line"], how="left")
    compare["calc"] = compare["calc"].fillna(0.0)
    compare["diff"] = compare["calc"] - compare[report_col]
    summary = {"metric": metric_name, "report_col": report_col, **summarize(compare, report_col, "calc")}
    mismatches = compare.loc[
        compare["diff"].abs().gt(TOLERANCE), ["region", "line", report_col, "calc", "diff"]
    ].to_dict(orient="records")
    return summary, mismatches


def compare_d_exit_net_profit(line_df: pd.DataFrame, project_df: pd.DataFrame) -> list[dict]:
    d_exit = project_df.loc[project_df["is_d_current_year_exit"].fillna(False)].copy()
    raw = d_exit.groupby(["region", "line"], as_index=False)["net_profit"].sum().rename(columns={"net_profit": "raw_net"})
    pen = (
        d_exit.assign(pen_net=d_exit["net_profit"] * d_exit["penetration_ratio"])
        .groupby(["region", "line"], as_index=False)["pen_net"]
        .sum()
    )
    compare = line_df[["region", "line", "d_exit_net_profit"]].merge(raw, on=["region", "line"], how="left")
    compare = compare.merge(pen, on=["region", "line"], how="left").fillna(0.0)
    compare["raw_diff"] = compare["raw_net"] - compare["d_exit_net_profit"]
    compare["pen_diff"] = compare["pen_net"] - compare["d_exit_net_profit"]
    return compare.to_dict(orient="records")


def main() -> None:
    project_df = base.load_project_report().copy()
    line_df = load_line_report()
    flags = load_flags()

    project_df = project_df.merge(flags, on="code_norm", how="left")
    project_df["is_d_current_year_exit"] = project_df["is_d_current_year_exit"].fillna(False)
    project_df["is_d_any_exit"] = project_df["is_d_any_exit"].fillna(False)
    project_df["penetration_ratio"] = project_df["penetration_ratio"].fillna(1.0)

    summaries: list[dict] = []
    mismatches: dict[str, list[dict]] = {}
    for metric_name, report_col in ROLLUP_METRICS:
        summary, detail = compare_rollup(line_df, project_df, metric_name, report_col)
        summaries.append(summary)
        if detail:
            mismatches[metric_name] = detail

    d_exit_detail = compare_d_exit_net_profit(line_df, project_df)
    d_exit_failed = [
        row
        for row in d_exit_detail
        if abs(row["raw_diff"]) > TOLERANCE and abs(row["pen_diff"]) > TOLERANCE
    ]
    summaries.append(
        {
            "metric": "D类已撤场项目管报净利润",
            "report_col": "d_exit_net_profit",
            "status": "通过" if not d_exit_failed else "失败",
            "mismatch_rows": int(len(d_exit_failed)),
            "max_abs_diff_raw": float(max(abs(row["raw_diff"]) for row in d_exit_detail) if d_exit_detail else 0.0),
            "max_abs_diff_pen": float(max(abs(row["pen_diff"]) for row in d_exit_detail) if d_exit_detail else 0.0),
        }
    )

    jprint(
        "d_current_year_exit_rule",
        {
            "report_ym": REPORT_YM,
            "report_year": REPORT_YEAR,
            "rule": "项目类型=D 且 项目状态=已撤场 且 报表年月>=实际撤场年月 且 撤场年份=报表年份",
            "matched_project_rows": int(project_df["is_d_current_year_exit"].sum()),
            "matched_regions": project_df.loc[project_df["is_d_current_year_exit"], "region"].value_counts().to_dict(),
        },
    )
    jprint("summary", summaries)
    jprint("mismatches", mismatches)
    jprint("d_exit_compare", d_exit_detail)


if __name__ == "__main__":
    main()
