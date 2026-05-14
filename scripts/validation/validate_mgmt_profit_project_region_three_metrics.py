from __future__ import annotations

import json
from pathlib import Path
from _project_root import find_project_root

import pandas as pd

import validate_attributable_profit as base


ROOT = find_project_root(__file__)
REPORT_MONTH = "2025-12"
REPORT_MONTH_COMPACT = "202512"
TOLERANCE = 1e-6


def jprint(label: str, data: object) -> None:
    print(f"[{label}]")
    print(json.dumps(data, ensure_ascii=False, default=str))


def status_from_diff(series: pd.Series) -> str:
    return "passed" if series.abs().le(TOLERANCE).all() else "failed"


def first_col(columns: list[object], token: str) -> object:
    return next(column for column in columns if token in str(column))


def find_attributable_workbook(dimension_token: str) -> Path:
    matches: list[Path] = []
    for path in ROOT.iterdir():
        if path.suffix.lower() != ".xlsx":
            continue
        if base.ATTRIBUTABLE_TOKEN not in path.name or dimension_token not in path.name:
            continue
        if dimension_token == base.REGION_TOKEN and base.u(r"\u6761\u7ebf") in path.name:
            continue
        matches.append(path)
    if len(matches) != 1:
        raise RuntimeError(f"Expected one attributable workbook for {dimension_token}, got {len(matches)}")
    return matches[0]


def find_current_cutoff_workbook() -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx"
        and base.CUTOFF_TOKEN in path.name
        and "202412" not in path.name
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one current cutoff workbook, got {len(matches)}")
    return matches[0]


def clean_text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def project_flags(query_df: pd.DataFrame) -> pd.DataFrame:
    flags = query_df.copy()
    exit_ym = pd.to_datetime(flags["exit_date"], errors="coerce").dt.strftime("%Y%m").fillna("")
    flags["is_d_exit_by_date"] = (
        flags["project_level"].astype(str).str.startswith("D")
        & flags["project_status"].astype(str).eq(base.u(r"\u5df2\u64a4\u573a"))
        & exit_ym.ne("")
        & exit_ym.le(REPORT_MONTH_COMPACT)
    )
    flags["is_d_exit_by_status"] = (
        flags["project_level"].astype(str).str.startswith("D")
        & flags["project_status"].astype(str).eq(base.u(r"\u5df2\u64a4\u573a"))
    )
    return flags[
        [
            "code_norm",
            "region",
            "penetration_ratio",
            "project_level",
            "project_status",
            "exit_date",
            "is_d_exit_by_date",
            "is_d_exit_by_status",
        ]
    ].drop_duplicates("code_norm")


def load_query() -> pd.DataFrame:
    df = pd.read_excel(base.find_workbook(base.QUERY_TOKEN), dtype=object)
    df.columns = [
        "project_code",
        "project_name",
        "region",
        "charge_area",
        "business_attr",
        "project_level",
        "charge_area_size",
        "legal_entity",
        "joint_venture_tag",
        "penetration_ratio",
        "ownership_attr",
        "project_status",
        "entry_date",
        "exit_date",
    ]
    df["code_norm"] = df["project_code"].map(base.normalize_code)
    df["penetration_ratio"] = pd.to_numeric(df["penetration_ratio"], errors="coerce").fillna(1.0)
    return df


def load_project_report() -> pd.DataFrame:
    df = pd.read_excel(find_attributable_workbook(base.PROJECT_TOKEN))
    df.columns = base.PROJECT_COLUMNS
    for column in base.PROJECT_COLUMNS[5:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    df["code_norm"] = df["project_code"].map(base.normalize_code)
    return df


def load_region_report() -> pd.DataFrame:
    df = pd.read_excel(find_attributable_workbook(base.REGION_TOKEN))
    df.columns = base.REGION_COLUMNS
    for column in base.REGION_COLUMNS[2:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def compare_project(report: pd.DataFrame, source: pd.DataFrame, report_col: str, calc_col: str) -> pd.DataFrame:
    result = report[["region", "line", "project_code", "project_name", "code_norm", report_col]].merge(
        source[["code_norm", calc_col]],
        on="code_norm",
        how="left",
    )
    result[calc_col] = result[calc_col].fillna(0.0)
    result["diff"] = result[calc_col] - result[report_col]
    return result


def project_summary(compare: pd.DataFrame, report_col: str, calc_col: str, source: pd.DataFrame) -> dict:
    report_codes = set(compare["code_norm"])
    missing_source = source[source[calc_col].abs().gt(TOLERANCE) & ~source["code_norm"].isin(report_codes)]
    return {
        "status": status_from_diff(compare["diff"]) if missing_source.empty else "failed",
        "report_rows": int(len(compare)),
        "mismatch_rows": int(compare["diff"].abs().gt(TOLERANCE).sum()),
        "max_abs_diff": float(compare["diff"].abs().max() if len(compare) else 0.0),
        "report_total": float(compare[report_col].sum()),
        "calc_total_on_report_codes": float(compare[calc_col].sum()),
        "nonzero_source_codes_missing_report": int(missing_source["code_norm"].nunique()),
        "missing_report_source_total": float(missing_source[calc_col].sum() if len(missing_source) else 0.0),
    }


def region_compare(report: pd.DataFrame, calc: pd.DataFrame, report_col: str, calc_col: str = "amount") -> pd.DataFrame:
    result = report[["region", report_col]].merge(calc[["region", calc_col]], on="region", how="left")
    result[calc_col] = result[calc_col].fillna(0.0)
    result["diff"] = result[calc_col] - result[report_col]
    return result


def region_summary(compare: pd.DataFrame, report_col: str, calc_col: str = "amount") -> dict:
    return {
        "status": status_from_diff(compare["diff"]),
        "regions": int(len(compare)),
        "mismatch_regions": int(compare["diff"].abs().gt(TOLERANCE).sum()),
        "max_abs_diff": float(compare["diff"].abs().max() if len(compare) else 0.0),
        "report_total": float(compare[report_col].sum()),
        "calc_total": float(compare[calc_col].sum()),
        "diff_total": float(compare["diff"].sum()),
    }


def project_report_rollup(
    project_report: pd.DataFrame,
    flags: pd.DataFrame,
    report_col: str,
    exclude_col: str,
    extra_region_amount: pd.DataFrame | None = None,
) -> pd.DataFrame:
    rollup = project_report[["region", "code_norm", report_col]].merge(
        flags[["code_norm", exclude_col]],
        on="code_norm",
        how="left",
    )
    rollup = (
        rollup.loc[~rollup[exclude_col].fillna(False)]
        .groupby("region", as_index=False)[report_col]
        .sum()
        .rename(columns={report_col: "amount"})
    )
    if extra_region_amount is not None:
        rollup = rollup.merge(extra_region_amount, on="region", how="outer").fillna(0.0)
        rollup["amount"] = rollup["amount"] + rollup["extra_amount"]
    return rollup


def load_ebt_source(flags: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    df = pd.read_excel(base.find_workbook(base.EBT_TOKEN), header=1, dtype=object)
    type_col = df.columns[1]
    month_col = df.columns[2]
    code_col = df.columns[4]
    region_col = df.columns[5]
    ebt_col = df.columns[7]
    unpaid_col = df.columns[8]
    df["row_type"] = clean_text(df[type_col])
    df["data_month"] = clean_text(df[month_col])
    df["code_norm"] = df[code_col].map(base.normalize_code)
    df["region"] = clean_text(df[region_col])
    df["amount"] = (
        pd.to_numeric(df[ebt_col], errors="coerce").fillna(0.0)
        + pd.to_numeric(df[unpaid_col], errors="coerce").fillna(0.0)
    ) * 0.81
    df = df[df["data_month"].eq(REPORT_MONTH)].copy()

    project_type = base.u(r"\u9879\u76ee")
    region_type = base.u(r"\u533a\u57df")
    project = df[df["row_type"].eq(project_type)].copy()
    project_calc = project.groupby("code_norm", as_index=False)["amount"].sum()

    project_with_flags = project.merge(flags, on="code_norm", how="left", suffixes=("", "_query"))
    project_with_flags["region_calc"] = project_with_flags["region_query"].fillna(project_with_flags["region"])
    project_region = (
        project_with_flags.loc[~project_with_flags["is_d_exit_by_status"].fillna(False)]
        .groupby("region_calc", as_index=False)["amount"]
        .sum()
        .rename(columns={"region_calc": "region", "amount": "project_non_d_amount"})
    )

    region_level = (
        df[df["row_type"].eq(region_type) & df["code_norm"].eq("")]
        .groupby("region", as_index=False)["amount"]
        .sum()
        .rename(columns={"amount": "region_level_amount"})
    )
    region_calc = region_level.merge(project_region, on="region", how="outer").fillna(0.0)
    region_calc["amount"] = region_calc["region_level_amount"] + region_calc["project_non_d_amount"]

    diagnostics = {
        "source_rows": int(len(df)),
        "type_counts": df["row_type"].value_counts(dropna=False).to_dict(),
        "project_rows": int(len(project)),
        "region_level_rows": int(len(df[df["row_type"].eq(region_type) & df["code_norm"].eq("")])),
    }
    return project_calc, region_calc, diagnostics


def load_discount_source(flags: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    df = pd.read_excel(base.find_workbook(base.DISCOUNT_TOKEN), header=1, dtype=object)
    month_col = df.columns[1]
    code_col = df.columns[3]
    discount_type_col = df.columns[5]
    top_type_col = df.columns[6]
    total_col = df.columns[9]
    related_col = df.columns[11]
    df["data_month"] = clean_text(df[month_col])
    df["code_norm"] = df[code_col].map(base.normalize_code)
    df["discount_type"] = clean_text(df[discount_type_col])
    df["top_type"] = clean_text(df[top_type_col])
    df["total_discount"] = pd.to_numeric(df[total_col], errors="coerce").fillna(0.0)
    df["related_discount"] = pd.to_numeric(df[related_col], errors="coerce").fillna(0.0)
    top_allowed = {base.u(r"\u5229\u6da6"), base.u(r"\u901a\u7528")}
    special_type = base.u(r"\u0032\u0030\u0032\u0033\u5e74\u5e95\u4e4b\u524d\u7684\u5173\u8054\u65b9\u5e94\u6536\u6b3e")
    df = df[df["data_month"].eq(REPORT_MONTH) & df["top_type"].isin(top_allowed)].copy()
    df["base_amount"] = df["total_discount"]
    df.loc[df["discount_type"].eq(special_type), "base_amount"] = df.loc[
        df["discount_type"].eq(special_type), "related_discount"
    ]
    df = df.merge(flags, on="code_norm", how="left")
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)
    df["amount"] = df["base_amount"] * 0.81 / 1.06 * df["penetration_ratio"]
    project_calc = df.groupby("code_norm", as_index=False)["amount"].sum()
    region_calc = (
        df.loc[~df["is_d_exit_by_status"].fillna(False)]
        .groupby("region", as_index=False)["amount"]
        .sum()
    )
    region_calc_by_date = (
        df.loc[~df["is_d_exit_by_date"].fillna(False)]
        .groupby("region", as_index=False)["amount"]
        .sum()
        .rename(columns={"amount": "amount_by_date"})
    )
    region_calc = region_calc.merge(region_calc_by_date, on="region", how="outer").fillna(0.0)
    diagnostics = {
        "filtered_source_rows": int(len(df)),
        "top_type_counts": df["top_type"].value_counts(dropna=False).to_dict(),
        "special_discount_type_rows": int(df["discount_type"].eq(special_type).sum()),
    }
    return project_calc, region_calc, diagnostics


def load_cutoff_source(flags: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    df = pd.read_excel(find_current_cutoff_workbook(), header=1, dtype=object)
    type_col = df.columns[1]
    month_col = df.columns[2]
    code_col = df.columns[4]
    region_col = df.columns[5]
    income_col = first_col(list(df.columns), base.u(r"\u5f53\u5e74\u65b0\u589e\u6536\u5165\u91d1\u989d"))
    cost_col = first_col(list(df.columns), base.u(r"\u5f53\u5e74\u65b0\u589e\u6210\u672c\u91d1\u989d"))
    df["row_type"] = clean_text(df[type_col])
    df["data_month"] = clean_text(df[month_col])
    df["code_norm"] = df[code_col].map(base.normalize_code)
    df["ledger_region"] = clean_text(df[region_col])
    df["base_amount"] = (
        pd.to_numeric(df[income_col], errors="coerce").fillna(0.0)
        + pd.to_numeric(df[cost_col], errors="coerce").fillna(0.0)
    )
    df = df[df["data_month"].eq(REPORT_MONTH)].copy()
    project_type = base.u(r"\u9879\u76ee")
    region_type = base.u(r"\u533a\u57df")

    project = df[df["row_type"].eq(project_type)].merge(flags, on="code_norm", how="left")
    project["penetration_ratio"] = project["penetration_ratio"].fillna(1.0)
    project["region_calc"] = project["region"].fillna(project["ledger_region"])
    project["amount_literal"] = project["base_amount"] * 0.81 * project["penetration_ratio"]
    project["amount"] = project["amount_literal"] / 1.06
    project_calc = project.groupby("code_norm", as_index=False)[["amount", "amount_literal"]].sum()

    project_region_status = (
        project.loc[~project["is_d_exit_by_status"].fillna(False)]
        .groupby("region_calc", as_index=False)[["amount", "amount_literal"]]
        .sum()
        .rename(columns={"region_calc": "region", "amount": "project_amount", "amount_literal": "project_amount_literal"})
    )
    project_region_date = (
        project.loc[~project["is_d_exit_by_date"].fillna(False)]
        .groupby("region_calc", as_index=False)[["amount", "amount_literal"]]
        .sum()
        .rename(
            columns={
                "region_calc": "region",
                "amount": "project_amount_by_date",
                "amount_literal": "project_amount_literal_by_date",
            }
        )
    )
    region_rows = df[df["row_type"].eq(region_type)].copy()
    region_rows["amount_literal"] = region_rows["base_amount"] * 0.81
    region_rows["amount"] = region_rows["amount_literal"] / 1.06
    region_level = (
        region_rows
        .groupby("ledger_region", as_index=False)[["amount", "amount_literal"]]
        .sum()
        .rename(
            columns={
                "ledger_region": "region",
                "amount": "region_level_amount",
                "amount_literal": "region_level_amount_literal",
            }
        )
    )
    region_calc = region_level.merge(project_region_status, on="region", how="outer").merge(
        project_region_date, on="region", how="outer"
    ).fillna(0.0)
    region_calc["amount"] = region_calc["region_level_amount"] + region_calc["project_amount"]
    region_calc["amount_by_date"] = region_calc["region_level_amount"] + region_calc["project_amount_by_date"]
    region_calc["amount_literal"] = (
        region_calc["region_level_amount_literal"] + region_calc["project_amount_literal"]
    )
    region_calc["amount_literal_by_date"] = (
        region_calc["region_level_amount_literal"] + region_calc["project_amount_literal_by_date"]
    )

    diagnostics = {
        "source_rows": int(len(df)),
        "type_counts": df["row_type"].value_counts(dropna=False).to_dict(),
        "project_rows": int(len(project)),
        "region_level_rows": int(len(df[df["row_type"].eq(region_type)])),
        "report_implementation_note": "uses (income + cost) * 0.81 / 1.06; indicator literal omits /1.06",
    }
    return project_calc, region_calc, diagnostics


def main() -> None:
    project_report = load_project_report()
    region_report = load_region_report()
    flags = project_flags(load_query())

    metric_specs = [
        ("ebt", "ebt_cost", load_ebt_source),
        ("discount", "discount", load_discount_source),
        ("cutoff", "cutoff", load_cutoff_source),
    ]

    summaries: list[dict] = []
    details: dict[str, object] = {}

    for metric, report_col, loader in metric_specs:
        project_source, region_source, diagnostics = loader(flags)
        project_cmp = compare_project(project_report, project_source, report_col, "amount")
        region_calc_source = region_source
        region_calc_by_date_source = None
        if metric in {"discount", "cutoff"}:
            extra = None
            if metric == "cutoff":
                extra = region_source[["region", "region_level_amount"]].rename(
                    columns={"region_level_amount": "extra_amount"}
                )
            region_calc_source = project_report_rollup(
                project_report,
                flags,
                report_col,
                "is_d_exit_by_status",
                extra,
            )
            region_calc_by_date_source = project_report_rollup(
                project_report,
                flags,
                report_col,
                "is_d_exit_by_date",
                extra,
            )
        region_cmp = region_compare(region_report, region_calc_source, report_col, "amount")
        summaries.append(
            {
                "metric": metric,
                "project": project_summary(project_cmp, report_col, "amount", project_source),
                "region": region_summary(region_cmp, report_col, "amount"),
                "source": diagnostics,
            }
        )
        details[f"{metric}_project_mismatch_top"] = (
            project_cmp.loc[project_cmp["diff"].abs().gt(TOLERANCE)]
            .sort_values("diff", key=lambda s: s.abs(), ascending=False)
            .head(20)
            .to_dict(orient="records")
        )
        details[f"{metric}_region_compare"] = region_cmp.to_dict(orient="records")
        if region_calc_source is not region_source:
            raw_region_cmp = region_compare(region_report, region_source, report_col, "amount")
            details[f"{metric}_raw_source_region_compare"] = raw_region_cmp.to_dict(orient="records")
        if "amount_by_date" in region_source.columns:
            by_date_source = (
                region_calc_by_date_source
                if region_calc_by_date_source is not None
                else region_source.drop(columns=["amount"], errors="ignore").rename(columns={"amount_by_date": "amount"})
            )
            by_date = region_compare(region_report, by_date_source, report_col, "amount")
            details[f"{metric}_region_by_exit_date_compare"] = by_date.to_dict(orient="records")
            summaries[-1]["region_by_exit_date"] = region_summary(by_date, report_col, "amount")
            merged = region_cmp[["region", "diff"]].rename(columns={"diff": "diff_by_status"}).merge(
                by_date[["region", "diff"]].rename(columns={"diff": "diff_by_exit_date"}),
                on="region",
                how="outer",
            )
            merged["date_minus_status"] = merged["diff_by_exit_date"] - merged["diff_by_status"]
            details[f"{metric}_exit_date_delta"] = merged.to_dict(orient="records")
        if "amount_literal" in project_source.columns:
            project_literal_cmp = compare_project(project_report, project_source, report_col, "amount_literal")
            region_literal_cmp = region_compare(region_report, region_source, report_col, "amount_literal")
            summaries[-1]["project_indicator_literal"] = project_summary(
                project_literal_cmp, report_col, "amount_literal", project_source
            )
            summaries[-1]["region_indicator_literal"] = region_summary(
                region_literal_cmp, report_col, "amount_literal"
            )
            details[f"{metric}_project_indicator_literal_mismatch_top"] = (
                project_literal_cmp.loc[project_literal_cmp["diff"].abs().gt(TOLERANCE)]
                .sort_values("diff", key=lambda s: s.abs(), ascending=False)
                .head(20)
                .to_dict(orient="records")
            )
            details[f"{metric}_region_indicator_literal_compare"] = region_literal_cmp.to_dict(orient="records")
            if "amount_literal_by_date" in region_source.columns:
                region_literal_by_date_source = region_source.drop(columns=["amount"], errors="ignore").rename(
                    columns={"amount_literal_by_date": "amount"}
                )
                region_literal_by_date_cmp = region_compare(
                    region_report,
                    region_literal_by_date_source,
                    report_col,
                    "amount",
                )
                summaries[-1]["region_indicator_literal_by_exit_date"] = region_summary(
                    region_literal_by_date_cmp, report_col, "amount"
                )
                details[f"{metric}_region_indicator_literal_by_exit_date_compare"] = (
                    region_literal_by_date_cmp.to_dict(orient="records")
                )

    jprint("summary", summaries)
    jprint("details", details)


if __name__ == "__main__":
    main()
