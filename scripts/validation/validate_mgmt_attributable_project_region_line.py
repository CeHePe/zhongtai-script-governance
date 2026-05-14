from __future__ import annotations

import json
from pathlib import Path
from _project_root import find_project_root

import pandas as pd

import validate_attributable_profit as base
import validate_mgmt_profit_project_region_three_metrics as three


ROOT = find_project_root(__file__)
TOLERANCE = 1e-6

LINE_COLUMNS = [
    "idx",
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


def jprint(label: str, data: object) -> None:
    print(f"[{label}]")
    print(json.dumps(data, ensure_ascii=False, default=str))


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens}, got {len(matches)}")
    return matches[0]


def load_line_report() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(base.ATTRIBUTABLE_TOKEN, base.u(r"\u6761\u7ebf")))
    df.columns = LINE_COLUMNS
    for column in LINE_COLUMNS[2:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def formula_common(df: pd.DataFrame, include_d_exit: bool) -> pd.Series:
    result = (
        df["net_profit"]
        - df["minority"]
        - df["plan_capital"]
        - df["plan_smart"]
        - df["plan_quality"]
        + df["actual_capital"]
        - df["actual_smart"]
        + df["actual_quality"]
        - df["ebt_cost"]
        + df["other_adjust"]
        + df["discount"]
        + df["cutoff"]
        + df["dev_incentive"]
        + df["jl_jj_restore"]
        + df["interest"]
        + df["prev_region_perf"]
        - df["curr_region_perf"]
    )
    if include_d_exit:
        result = result - df["d_exit_net_profit"]
    return result


def summarize_diff(label: str, key_col: str, df: pd.DataFrame, calc_col: str, report_col: str) -> dict:
    diff = df[calc_col] - df[report_col]
    return {
        "check": label,
        "status": "passed" if diff.abs().le(TOLERANCE).all() else "failed",
        "rows": int(len(df)),
        "mismatch_rows": int(diff.abs().gt(TOLERANCE).sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "report_total": float(df[report_col].sum()),
        "calc_total": float(df[calc_col].sum()),
        "diff_total": float(diff.sum()),
        "mismatches": df.loc[
            diff.abs().gt(TOLERANCE),
            [key_col, report_col, calc_col],
        ]
        .assign(diff=diff[diff.abs().gt(TOLERANCE)])
        .to_dict(orient="records"),
    }


def build_correct_cutoff_project(project_df: pd.DataFrame) -> pd.DataFrame:
    flags = three.project_flags(three.load_query())
    cutoff_source, _, _ = three.load_cutoff_source(flags)
    result = project_df[["region", "line", "project_code", "project_name", "code_norm", "cutoff"]].merge(
        cutoff_source[["code_norm", "amount_literal"]].rename(columns={"amount_literal": "correct_cutoff"}),
        on="code_norm",
        how="left",
    )
    result["correct_cutoff"] = result["correct_cutoff"].fillna(0.0)
    result["cutoff_correction"] = result["correct_cutoff"] - result["cutoff"]
    return result


def correction_by_dimension(project_df: pd.DataFrame, region_df: pd.DataFrame, line_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    flags = three.project_flags(three.load_query())
    project_cutoff = build_correct_cutoff_project(project_df)
    project_correct = project_df[["region", "line", "project_code", "project_name", "code_norm", "attributable", "cutoff"]].merge(
        project_cutoff[["code_norm", "correct_cutoff", "cutoff_correction"]],
        on="code_norm",
        how="left",
    )
    project_correct["correct_attributable"] = project_correct["attributable"] + project_correct["cutoff_correction"]

    _, region_source, _ = three.load_cutoff_source(flags)
    region_correct_cutoff = region_source[["region", "amount_literal"]].rename(
        columns={"amount_literal": "correct_cutoff"}
    )
    region_correct = region_df[["region", "attributable", "cutoff"]].merge(region_correct_cutoff, on="region", how="left")
    region_correct["correct_cutoff"] = region_correct["correct_cutoff"].fillna(0.0)
    region_correct["cutoff_correction"] = region_correct["correct_cutoff"] - region_correct["cutoff"]
    region_correct["correct_attributable"] = region_correct["attributable"] + region_correct["cutoff_correction"]

    project_with_flags = project_df.merge(
        flags[["code_norm", "is_d_exit_by_status"]],
        on="code_norm",
        how="left",
    )
    line_correct_cutoff = (
        project_cutoff.merge(project_with_flags[["code_norm", "is_d_exit_by_status"]], on="code_norm", how="left")
        .loc[lambda df: ~df["is_d_exit_by_status"].fillna(False)]
        .groupby("line", as_index=False)["correct_cutoff"]
        .sum()
    )
    line_correct = line_df[["line", "attributable", "cutoff"]].merge(line_correct_cutoff, on="line", how="left")
    line_correct["correct_cutoff"] = line_correct["correct_cutoff"].fillna(0.0)
    line_correct["cutoff_correction"] = line_correct["correct_cutoff"] - line_correct["cutoff"]
    line_correct["correct_attributable"] = line_correct["attributable"] + line_correct["cutoff_correction"]
    return project_correct, region_correct, line_correct


def high_dim_rollup(project_df: pd.DataFrame, region_df: pd.DataFrame, line_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    flags = three.project_flags(three.load_query())
    project_with_flags = project_df.merge(
        flags[["code_norm", "is_d_exit_by_status"]],
        on="code_norm",
        how="left",
    )
    non_d = project_with_flags.loc[~project_with_flags["is_d_exit_by_status"].fillna(False)]
    region_rollup = non_d.groupby("region", as_index=False)["attributable"].sum().rename(
        columns={"attributable": "project_rollup_attributable"}
    )
    region_compare = region_df[["region", "attributable"]].merge(region_rollup, on="region", how="left").fillna(0.0)
    region_compare["rollup_diff"] = region_compare["project_rollup_attributable"] - region_compare["attributable"]

    line_rollup = non_d.groupby("line", as_index=False)["attributable"].sum().rename(
        columns={"attributable": "project_rollup_attributable"}
    )
    line_compare = line_df[["line", "attributable"]].merge(line_rollup, on="line", how="left").fillna(0.0)
    line_compare["rollup_diff"] = line_compare["project_rollup_attributable"] - line_compare["attributable"]
    return region_compare, line_compare


def main() -> None:
    project_df = base.load_project_report()
    region_df = base.load_region_report()
    line_df = load_line_report()

    project_formula = project_df[["project_code", "project_name", "attributable"]].copy()
    project_formula["formula_calc"] = formula_common(project_df, include_d_exit=False)
    region_formula = region_df[["region", "attributable"]].copy()
    region_formula["formula_calc"] = base.region_formula(region_df)
    line_formula = line_df[["line", "attributable"]].copy()
    line_formula["formula_calc"] = formula_common(line_df, include_d_exit=True)

    region_rollup, line_rollup = high_dim_rollup(project_df, region_df, line_df)
    project_correct, region_correct, line_correct = correction_by_dimension(project_df, region_df, line_df)

    jprint(
        "formula_checks",
        [
            summarize_diff("project_current_report_formula", "project_code", project_formula, "formula_calc", "attributable"),
            summarize_diff("region_current_report_formula", "region", region_formula, "formula_calc", "attributable"),
            summarize_diff("line_current_report_formula", "line", line_formula, "formula_calc", "attributable"),
        ],
    )
    jprint(
        "rollup_checks_current_report",
        [
            summarize_diff(
                "region_rollup_from_non_d_project_attributable",
                "region",
                region_rollup,
                "project_rollup_attributable",
                "attributable",
            ),
            summarize_diff(
                "line_rollup_from_non_d_project_attributable",
                "line",
                line_rollup,
                "project_rollup_attributable",
                "attributable",
            ),
        ],
    )
    jprint(
        "corrected_by_cutoff",
        {
            "project_total": {
                "report_attributable": float(project_correct["attributable"].sum()),
                "correct_attributable": float(project_correct["correct_attributable"].sum()),
                "cutoff_correction": float(project_correct["cutoff_correction"].sum()),
            },
            "region": region_correct.to_dict(orient="records"),
            "line": line_correct.to_dict(orient="records"),
            "project_mismatch_top": project_correct.loc[
                project_correct["cutoff_correction"].abs().gt(TOLERANCE),
                [
                    "region",
                    "line",
                    "project_code",
                    "project_name",
                    "attributable",
                    "correct_attributable",
                    "cutoff",
                    "correct_cutoff",
                    "cutoff_correction",
                ],
            ]
            .sort_values("cutoff_correction", key=lambda s: s.abs(), ascending=False)
            .head(20)
            .to_dict(orient="records"),
        },
    )


if __name__ == "__main__":
    main()
