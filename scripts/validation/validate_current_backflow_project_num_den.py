from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path

import pandas as pd

from _project_root import find_project_root


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
TOL = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens!r}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


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


def load_indicator_rows() -> list[dict]:
    indicator = find_workbook(u(r"\u6307\u6807\u6e05\u5355"))
    df = pd.read_excel(indicator, sheet_name=0, dtype=object)
    name_col = df.columns[3]
    dim_col = df.columns[4]
    period_col = df.columns[5]
    metrics = {u(r"\u5f53\u671f\u56de\u6b3e\u7387_\u5206\u5b50"), u(r"\u5f53\u671f\u56de\u6b3e\u7387_\u5206\u6bcd")}
    target = df[
        df[name_col].astype(str).str.strip().isin(metrics)
        & df[dim_col].astype(str).str.strip().eq(u(r"\u9879\u76ee"))
        & df[period_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1"))
    ].copy()
    if len(target) != 2:
        raise RuntimeError(f"Expected two indicator rows, got {len(target)}")

    rows = []
    for idx, row in target.iterrows():
        rows.append(
            {
                "excel_row": int(idx + 2),
                "serial": row.get(u(r"\u5e8f\u53f7")),
                "relation": row.get(df.columns[1]),
                "component": row.get(df.columns[2]),
                "metric": row.get(name_col),
                "dimension": row.get(dim_col),
                "period": row.get(period_col),
                "source_method": row.get(df.columns[8]),
                "source_table": row.get(df.columns[10]),
                "logic": row.get(df.columns[12]),
            }
        )
    return rows


def load_report() -> pd.DataFrame:
    report = find_workbook(u(r"\u5f53\u671f\u56de\u6b3e\u7387"), REPORT_MONTH, u(r"\u9879\u76ee"))
    raw = pd.read_excel(report, header=None, dtype=object)
    data = raw.iloc[3:].reset_index(drop=True).copy()
    data = data[data[3].notna()].copy()
    data.columns = [
        "idx",
        "region",
        "line",
        "project_code",
        "project_name",
        "cashflow_total",
        "non_assess_cashflow",
        "related_cashflow",
        "non_assess_related_cashflow",
        "prev_year_ar_current_recovery",
        "old_ar_current_recovery",
        "current_year_ar_recovery",
        "current_water_recovery",
        "current_water_recovery_alloc",
        "related_current_water_recovery",
        "coin_recovery_current_year",
        "prev_year_not_due_current_recovery",
        "numerator",
        "revenue",
        "non_assess_revenue",
        "related_revenue",
        "non_assess_related_revenue",
        "current_water_receivable",
        "current_water_receivable_alloc",
        "related_current_water_receivable",
        "current_coin_balance",
        "prev_year_not_due_balance",
        "current_year_not_due_amount",
        "denominator",
        "rate",
    ]
    for column in data.columns[5:]:
        data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0.0)
    data["code_norm"] = data["project_code"].map(normalize_code)
    data["report_file"] = report.name
    return data


def calculate(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["calc_numerator"] = (
        out["current_year_ar_recovery"]
        + out["current_water_recovery"]
        + out["current_water_recovery_alloc"]
        - out["related_current_water_recovery"]
        - out["coin_recovery_current_year"]
        + out["prev_year_not_due_current_recovery"]
    )
    out["numerator_diff"] = out["calc_numerator"] - out["numerator"]
    out["calc_denominator"] = (
        out["revenue"]
        - out["non_assess_revenue"]
        - out["related_revenue"]
        + out["non_assess_related_revenue"]
        + out["current_water_receivable"]
        + out["current_water_receivable_alloc"]
        - out["related_current_water_receivable"]
        + out["prev_year_not_due_balance"]
        - out["current_year_not_due_amount"]
        - out["current_coin_balance"]
    )
    out["denominator_diff"] = out["calc_denominator"] - out["denominator"]
    return out


def mismatch_summary(df: pd.DataFrame, actual: str, calc: str, diff: str) -> dict:
    mismatches = df[df[diff].abs() > TOL]
    by_code = (
        df.groupby("code_norm", as_index=False)
        .agg(
            rows=("code_norm", "size"),
            project_names=("project_name", lambda values: " / ".join(map(str, values))),
            actual=(actual, "sum"),
            calculated=(calc, "sum"),
        )
        .copy()
    )
    by_code["diff"] = by_code["calculated"] - by_code["actual"]
    code_mismatches = by_code[by_code["diff"].abs() > TOL]
    return {
        "row_mismatch_count": int(len(mismatches)),
        "project_code_mismatch_count": int(len(code_mismatches)),
        "actual_total": float(df[actual].sum()),
        "calculated_total": float(df[calc].sum()),
        "total_diff": float(df[diff].sum()),
        "max_abs_row_diff": float(df[diff].abs().max() if len(df) else 0.0),
    }


def top_mismatches(df: pd.DataFrame, diff: str, calc: str, actual: str) -> pd.DataFrame:
    cols = ["region", "line", "project_code", "project_name", actual, calc, diff]
    return df.loc[df[diff].abs() > TOL, cols].sort_values(diff, key=lambda s: s.abs(), ascending=False).head(30)


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    supplement_hits = [
        path.name
        for path in ROOT.iterdir()
        if path.is_file() and u(r"\u589e\u8865") in path.name and u(r"\u56de\u6b3e") in path.name
    ]
    indicator_rows = load_indicator_rows()
    report = calculate(load_report())
    numerator = mismatch_summary(report, "numerator", "calc_numerator", "numerator_diff")
    denominator = mismatch_summary(report, "denominator", "calc_denominator", "denominator_diff")
    summary = {
        "status": "passed"
        if numerator["project_code_mismatch_count"] == 0 and denominator["project_code_mismatch_count"] == 0
        else "failed",
        "report_month": REPORT_MONTH,
        "report_file": str(report["report_file"].iloc[0]),
        "report_rows": int(len(report)),
        "project_codes": int(report["code_norm"].nunique()),
        "indicator_rows": indicator_rows,
        "supplement_formula_workbook_found": supplement_hits,
        "numerator": numerator,
        "denominator": denominator,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
    print("\nNUMERATOR_MISMATCH_TOP30")
    print(top_mismatches(report, "numerator_diff", "calc_numerator", "numerator").to_string(index=False))
    print("\nDENOMINATOR_MISMATCH_TOP30")
    print(top_mismatches(report, "denominator_diff", "calc_denominator", "denominator").to_string(index=False))


if __name__ == "__main__":
    main()
