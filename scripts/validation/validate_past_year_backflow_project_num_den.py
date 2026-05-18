from __future__ import annotations

import argparse
import json
import sys
import warnings
from pathlib import Path

import pandas as pd

from _project_root import find_project_root


ROOT = find_project_root(__file__)
TOL = 1e-6


def u(text: str) -> str:
    return text.encode("ascii").decode("unicode_escape")


PROJECT = u(r"\u9879\u76ee")
CUMULATIVE = u(r"\u7d2f\u8ba1")
REPORT_MARK = u(r"\u5f80\u5e74\u56de\u6b3e\u7387")
INDICATOR_MARK = u(r"\u6307\u6807\u6e05\u5355")
PROJECT_QUERY_NAME = u(r"\u9879\u76ee\u67e5\u8be2.xlsx")
NON_ASSESS_NAME = u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26.xlsx")

METRIC_NUMERATOR = u(r"\u5f80\u5e74\u56de\u6b3e\u7387_\u5206\u5b50")
METRIC_DENOMINATOR = u(r"\u5f80\u5e74\u56de\u6b3e\u7387_\u5206\u6bcd")
METRIC_RATE = u(r"\u5f80\u5e74\u56de\u6b3e\u7387")

REPORT_COLUMNS = [
    "idx",
    "region",
    "line",
    "project_code",
    "project_name",
    "past_ar_current_recovery",
    "past_water_current_recovery",
    "past_water_current_recovery_alloc",
    "related_water_old_recovery",
    "coin_recovery_past",
    "numerator",
    "past_ar_balance",
    "past_water_balance",
    "past_water_balance_alloc",
    "related_water_old_receivable",
    "past_not_due_prev_year_end",
    "past_coin_balance",
    "denominator",
    "rate",
]

NUMERIC_COLUMNS = REPORT_COLUMNS[5:]


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def normalize_code(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    if text in {"", "NAN", "NONE"}:
        return ""
    if len(text) > 1 and text[0].isalpha() and any(ch.isdigit() for ch in text[1:]):
        return text[1:]
    return text


def clean_value(value: object) -> object:
    if value is None or pd.isna(value):
        return None
    return value


def find_workbook(*tokens: str, required: bool = True) -> Path | None:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and all(token in path.name for token in tokens)
    ]
    if len(matches) == 1:
        return matches[0]
    if required:
        raise RuntimeError(f"Expected one workbook for {tokens!r}, got {len(matches)}: {[p.name for p in matches]}")
    return None


def find_exact_workbook(name: str, required: bool = True) -> Path | None:
    path = ROOT / name
    if path.exists():
        return path
    if required:
        raise RuntimeError(f"Missing workbook: {name}")
    return None


def load_indicator_rows() -> list[dict]:
    indicator = find_workbook("JKS_", INDICATOR_MARK)
    assert indicator is not None
    df = pd.read_excel(indicator, sheet_name=0, dtype=object)

    serial_col = df.columns[0]
    relation_col = df.columns[1]
    component_col = df.columns[2]
    name_col = df.columns[3]
    dimension_col = df.columns[4]
    period_col = df.columns[5]
    method_col = df.columns[8]
    source_table_col = df.columns[10]
    logic_col = df.columns[12]

    metric_names = {METRIC_NUMERATOR, METRIC_DENOMINATOR, METRIC_RATE}
    target = df[
        df[name_col].astype(str).str.strip().isin(metric_names)
        & df[dimension_col].astype(str).str.strip().eq(PROJECT)
        & df[period_col].astype(str).str.strip().eq(CUMULATIVE)
    ].copy()
    if len(target) != 3:
        raise RuntimeError(f"Expected three project cumulative indicator rows, got {len(target)}")

    rows = []
    for idx, row in target.iterrows():
        rows.append(
            {
                "excel_row": int(idx + 2),
                "serial": clean_value(row.get(serial_col)),
                "relation": clean_value(row.get(relation_col)),
                "component": clean_value(row.get(component_col)),
                "metric": clean_value(row.get(name_col)),
                "dimension": clean_value(row.get(dimension_col)),
                "period": clean_value(row.get(period_col)),
                "method": clean_value(row.get(method_col)),
                "source_table": clean_value(row.get(source_table_col)),
                "logic": clean_value(row.get(logic_col)),
            }
        )
    return rows


def load_report(report_month: str) -> pd.DataFrame:
    report = find_workbook(REPORT_MARK, report_month, PROJECT)
    assert report is not None
    raw = pd.read_excel(report, sheet_name=0, header=None, dtype=object)
    data = raw.iloc[2:].reset_index(drop=True).copy()
    data = data[data[3].notna()].copy()
    if data.shape[1] != len(REPORT_COLUMNS):
        raise RuntimeError(f"Unexpected report column count: {data.shape[1]}")

    data.columns = REPORT_COLUMNS
    for column in NUMERIC_COLUMNS:
        data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0.0)
    data["code_norm"] = data["project_code"].map(normalize_code)
    data["report_file"] = report.name
    return data


def calculate_formula(report: pd.DataFrame) -> pd.DataFrame:
    out = report.copy()
    out["calc_numerator"] = (
        out["past_ar_current_recovery"]
        + out["past_water_current_recovery"]
        + out["past_water_current_recovery_alloc"]
        - out["related_water_old_recovery"]
        - out["coin_recovery_past"]
    )
    out["numerator_diff"] = out["numerator"] - out["calc_numerator"]

    out["calc_denominator"] = (
        out["past_ar_balance"]
        + out["past_water_balance"]
        + out["past_water_balance_alloc"]
        - out["related_water_old_receivable"]
        - out["past_not_due_prev_year_end"]
        - out["past_coin_balance"]
    )
    out["denominator_diff"] = out["denominator"] - out["calc_denominator"]
    return out


def summarize_formula(report: pd.DataFrame, actual: str, calculated: str, diff: str) -> dict:
    mismatches = report[report[diff].abs() > TOL]
    return {
        "row_mismatch_count": int(len(mismatches)),
        "actual_total": float(report[actual].sum()),
        "calculated_total": float(report[calculated].sum()),
        "total_diff": float(report[diff].sum()),
        "max_abs_row_diff": float(report[diff].abs().max() if len(report) else 0.0),
    }


def load_scope_check(report: pd.DataFrame) -> dict:
    project_query = find_exact_workbook(PROJECT_QUERY_NAME)
    non_assess = find_exact_workbook(NON_ASSESS_NAME)
    assert project_query is not None and non_assess is not None

    code_col = u(r"\u7acb\u9879\u7f16\u7801")
    level_col = u(r"\u9879\u76ee\u7b49\u7ea7")
    status_col = u(r"\u9879\u76ee\u72b6\u6001")
    exit_status = u(r"\u5df2\u64a4\u573a")

    project_df = pd.read_excel(project_query, dtype=object)
    project_df["code_norm"] = project_df[code_col].map(normalize_code)
    non_assess_df = pd.read_excel(non_assess, dtype=object)
    non_assess_df["code_norm"] = non_assess_df[code_col].map(normalize_code)
    non_assess_codes = set(non_assess_df["code_norm"].dropna())

    merged = report.merge(
        project_df[["code_norm", level_col, status_col]].drop_duplicates("code_norm"),
        on="code_norm",
        how="left",
    )
    is_non_assess = merged["code_norm"].isin(non_assess_codes)
    is_d_exit = merged[level_col].astype(str).str.startswith("D") & merged[status_col].astype(str).eq(exit_status)
    return {
        "project_query_unmatched_rows": int(merged[level_col].isna().sum()),
        "d_exit_rows_in_report": int(is_d_exit.sum()),
        "non_assess_rows_in_report": int(is_non_assess.sum()),
        "d_exit_numerator_total": float(merged.loc[is_d_exit, "numerator"].sum()),
        "d_exit_denominator_total": float(merged.loc[is_d_exit, "denominator"].sum()),
        "non_assess_numerator_total": float(merged.loc[is_non_assess, "numerator"].sum()),
        "non_assess_denominator_total": float(merged.loc[is_non_assess, "denominator"].sum()),
    }


def source_availability() -> dict:
    optional_sources = {
        "formula_addendum": [u(r"\u6307\u6807\u6e05\u5355\u589e\u8865-\u672a\u5230\u8d26\u671f\u4f59\u989d\u53ca\u56de\u6b3e\u516c\u5f0f260211")],
        "business_aging_202412": [u(r"\u4e1a\u52a1\u5e10\u9f84-\u5e74\u5ea6\u5206\u5e03202412")],
        "business_aging_202512": [u(r"\u4e1a\u52a1\u5e10\u9f84-\u5e74\u5ea6\u5206\u5e03202512")],
        "water_detail": [u(r"\u57ab\u652f\u6c34\u7535\u8d39\u660e\u7ec6\u8868")],
        "water_alloc": [u(r"\u57ab\u652f\u6c34\u7535\u8d39\u7269\u4e1a\u5206\u644a\u636e\u5b9e\u5206\u644a")],
        "related_party": [u(r"\u5173\u8054\u65b9\u6c34\u7535\u8d39\u5e94\u6536\u53ca\u5b9e\u6536\u5e74\u5ea6\u5206\u5e03")],
        "not_due_202412": [u(r"\u5e94\u6536\u8d26\u9f84\u53ca\u672a\u5230\u8d26\u671f\u91d1\u989d\u5e74\u5ea6\u5206\u5e03202412")],
        "coin_balance_2024": [u(r"\u91d1\u5e01\u4f59\u989d\u53f0\u8d262024")],
        "coin_recovery": [u(r"\u5c0f\u4e1a\u4e3b\u91d1\u5e01\u56de\u6b3e\u91d1\u989d\u53f0\u8d26")],
    }
    files = [path.name for path in ROOT.iterdir() if path.is_file() and path.suffix.lower() == ".xlsx"]
    result = {}
    for key, tokens in optional_sources.items():
        matches = [name for name in files if all(token in name for token in tokens)]
        result[key] = matches
    return result


def top_mismatches(report: pd.DataFrame, diff: str, actual: str, calculated: str) -> list[dict]:
    cols = ["region", "line", "project_code", "project_name", actual, calculated, diff]
    mismatches = report.loc[report[diff].abs() > TOL, cols].copy()
    if mismatches.empty:
        return []
    mismatches = mismatches.sort_values(diff, key=lambda series: series.abs(), ascending=False).head(20)
    return json.loads(mismatches.to_json(orient="records", force_ascii=False))


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    parser = argparse.ArgumentParser(description="Validate past-year backflow project numerator/denominator formulas.")
    parser.add_argument("--report-month", default="202512")
    args = parser.parse_args()

    indicator_rows = load_indicator_rows()
    report = calculate_formula(load_report(args.report_month))
    numerator = summarize_formula(report, "numerator", "calc_numerator", "numerator_diff")
    denominator = summarize_formula(report, "denominator", "calc_denominator", "denominator_diff")
    status = "passed" if numerator["row_mismatch_count"] == 0 and denominator["row_mismatch_count"] == 0 else "failed"

    summary = {
        "status": status,
        "scope": "project report internal numerator/denominator formula",
        "report_month": args.report_month,
        "report_file": str(report["report_file"].iloc[0]),
        "report_rows": int(len(report)),
        "project_codes": int(report["code_norm"].nunique()),
        "indicator_rows": indicator_rows,
        "source_availability": source_availability(),
        "scope_check": load_scope_check(report),
        "numerator": numerator,
        "denominator": denominator,
        "top_numerator_mismatches": top_mismatches(report, "numerator_diff", "numerator", "calc_numerator"),
        "top_denominator_mismatches": top_mismatches(report, "denominator_diff", "denominator", "calc_denominator"),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
    if status != "passed":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
