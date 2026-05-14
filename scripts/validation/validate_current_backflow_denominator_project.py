from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
TOL = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str, exact_name: str | None = None) -> Path:
    if exact_name is not None:
        matches = [path for path in ROOT.iterdir() if path.is_file() and path.name == exact_name]
    else:
        matches = [
            path
            for path in ROOT.iterdir()
            if path.is_file() and path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
        ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens or exact_name!r}, got {len(matches)}: {[p.name for p in matches]}")
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


def load_indicator_row() -> dict:
    indicator = find_workbook(u(r"\u6307\u6807\u6e05\u5355"))
    df = pd.read_excel(indicator, sheet_name=0, dtype=object)
    name_col = df.columns[3]
    dim_col = df.columns[4]
    period_col = df.columns[5]
    target = df[
        df[name_col].astype(str).str.strip().eq(u(r"\u5f53\u671f\u56de\u6b3e\u7387_\u5206\u6bcd"))
        & df[dim_col].astype(str).str.strip().eq(u(r"\u9879\u76ee"))
        & df[period_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1"))
    ].copy()
    if len(target) != 1:
        raise RuntimeError(f"Expected one indicator row, got {len(target)}")
    row = target.iloc[0]
    return {
        "excel_row": int(target.index[0] + 2),
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


def load_report() -> pd.DataFrame:
    report = find_workbook(u(r"\u5f53\u671f\u56de\u6b3e\u7387"), REPORT_MONTH, u(r"\u9879\u76ee"))
    raw = pd.read_excel(report, header=None, dtype=object)
    df = raw.iloc[3:].reset_index(drop=True).copy()
    df = df[df[3].notna()].copy()
    df.columns = [
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
        "related_current_water_recovery",
        "coin_recovery_current_year",
        "prev_year_not_due_current_recovery",
        "numerator",
        "revenue",
        "non_assess_revenue",
        "related_revenue",
        "non_assess_related_revenue",
        "current_water_receivable",
        "related_current_water_receivable",
        "current_coin_balance",
        "prev_year_not_due_balance",
        "current_year_not_due_amount",
        "denominator",
        "rate",
    ]
    numeric_cols = df.columns[5:]
    for column in numeric_cols:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    df["code_norm"] = df["project_code"].map(normalize_code)
    df["report_file"] = report.name
    return df


def attach_project_flags(df: pd.DataFrame) -> pd.DataFrame:
    query = pd.read_excel(find_workbook(exact_name=u(r"\u9879\u76ee\u67e5\u8be2.xlsx")), dtype=object)
    query["code_norm"] = query[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code)
    query = query[
        [
            "code_norm",
            u(r"\u9879\u76ee\u7b49\u7ea7"),
            u(r"\u9879\u76ee\u72b6\u6001"),
            u(r"\u5df2\u64a4\u573a\u65f6\u95f4"),
        ]
    ].drop_duplicates("code_norm")
    query.columns = ["code_norm", "project_level", "project_status", "exit_date"]

    non_assess = pd.read_excel(find_workbook(exact_name=u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26.xlsx")), dtype=object)
    non_assess_codes = set(non_assess[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code).dropna())

    out = df.merge(query, on="code_norm", how="left")
    out["is_non_assess"] = out["code_norm"].isin(non_assess_codes)
    out["is_d_exit"] = out["project_level"].astype(str).str.startswith("D") & out["project_status"].astype(str).eq(
        u(r"\u5df2\u64a4\u573a")
    )
    return out


def add_calculation(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["calc_denominator"] = (
        out["revenue"]
        - out["non_assess_revenue"]
        - out["related_revenue"]
        + out["non_assess_related_revenue"]
        + out["current_water_receivable"]
        - out["related_current_water_receivable"]
        + out["prev_year_not_due_balance"]
        - out["current_year_not_due_amount"]
        - out["current_coin_balance"]
    )
    out["diff"] = out["calc_denominator"] - out["denominator"]
    return out


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    supplement_hits = [
        path.name
        for path in ROOT.iterdir()
        if path.is_file() and u(r"\u589e\u8865") in path.name and u(r"\u56de\u6b3e") in path.name
    ]
    indicator_row = load_indicator_row()
    report = add_calculation(attach_project_flags(load_report()))

    row_mismatches = report[report["diff"].abs() > TOL].copy()
    by_code = (
        report.groupby("code_norm", as_index=False)
        .agg(
            rows=("code_norm", "size"),
            report_denominator=("denominator", "sum"),
            calc_denominator=("calc_denominator", "sum"),
            project_names=("project_name", lambda values: " / ".join(map(str, values))),
        )
        .copy()
    )
    by_code["diff"] = by_code["calc_denominator"] - by_code["report_denominator"]
    code_mismatches = by_code[by_code["diff"].abs() > TOL].copy()

    summary = {
        "status": "passed_by_project_code" if code_mismatches.empty else "failed",
        "note": "row-level mismatches remain only where the same project code is split across rows"
        if code_mismatches.empty and not row_mismatches.empty
        else "",
        "report_month": REPORT_MONTH,
        "indicator_row": indicator_row,
        "supplement_formula_workbook_found": supplement_hits,
        "report_file": str(report["report_file"].iloc[0]),
        "report_rows": int(len(report)),
        "project_codes": int(report["code_norm"].nunique()),
        "row_mismatch_count": int(len(row_mismatches)),
        "project_code_mismatch_count": int(len(code_mismatches)),
        "report_denominator_total": float(report["denominator"].sum()),
        "calc_denominator_total": float(report["calc_denominator"].sum()),
        "total_diff": float(report["diff"].sum()),
        "non_assessment_rows_flagged_for_reference_only": int(report["is_non_assess"].sum()),
        "d_exit_rows": int(report["is_d_exit"].sum()),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))

    if not row_mismatches.empty:
        print("\nROW_LEVEL_MISMATCHES")
        cols = [
            "region",
            "line",
            "project_code",
            "project_name",
            "denominator",
            "calc_denominator",
            "diff",
            "is_non_assess",
            "is_d_exit",
        ]
        print(row_mismatches[cols].sort_values("diff", key=lambda s: s.abs(), ascending=False).to_string(index=False))

    if not code_mismatches.empty:
        print("\nPROJECT_CODE_MISMATCHES")
        print(code_mismatches.sort_values("diff", key=lambda s: s.abs(), ascending=False).to_string(index=False))


if __name__ == "__main__":
    main()
