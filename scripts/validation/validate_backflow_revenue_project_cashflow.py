from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
TOLERANCE = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str, exact_stem: bool = False) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx"
        and (
            (exact_stem and len(tokens) == 1 and path.stem == tokens[0])
            or (not exact_stem and all(token in path.name for token in tokens))
        )
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    if text in {"", "NAN", "NONE"}:
        return ""
    if len(text) >= 2 and text[0].isalpha() and any(ch.isdigit() for ch in text[1:]):
        return text[1:]
    return text


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def markdown_table(df: pd.DataFrame, floatfmt: str = ".2f") -> str:
    if df.empty:
        return "(empty)"
    headers = list(df.columns)
    rows = []
    for _, row in df.iterrows():
        values = []
        for value in row.tolist():
            if isinstance(value, float):
                values.append(format(value, floatfmt))
            else:
                values.append("" if pd.isna(value) else str(value))
        rows.append(values)
    lines = [
        "| " + " | ".join(str(header) for header in headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    lines.extend("| " + " | ".join(value.replace("\n", " ") for value in row) + " |" for row in rows)
    return "\n".join(lines)


def load_indicator_row() -> dict:
    indicator = find_workbook(u(r"\u6307\u6807\u6e05\u5355"))
    df = pd.read_excel(indicator, sheet_name=0, dtype=object)
    relation_col = df.columns[1]
    component_col = df.columns[2]
    metric_col = df.columns[3]
    dimension_col = df.columns[4]
    period_col = df.columns[5]
    method_col = df.columns[8]
    source_col = df.columns[10]
    logic_col = df.columns[12]

    target = df[
        df[relation_col].astype(str).str.strip().eq(u(r"\u56de\u6b3e\u8425\u6536\u6bd4"))
        & df[component_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41"))
        & df[metric_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41_\u9879\u76ee"))
        & df[dimension_col].astype(str).str.strip().eq(u(r"\u9879\u76ee"))
        & df[period_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1"))
    ].copy()
    if len(target) != 1:
        raise RuntimeError(f"Expected one indicator row, got {len(target)}")
    row = target.iloc[0]
    return {
        "excel_row": int(target.index[0] + 2),
        "serial": row.get(u(r"\u5e8f\u53f7")),
        "report": row.get(relation_col),
        "component": row.get(component_col),
        "metric": row.get(metric_col),
        "dimension": row.get(dimension_col),
        "period": row.get(period_col),
        "source_method": row.get(method_col),
        "source_table": row.get(source_col),
        "logic": row.get(logic_col),
    }


def load_source() -> tuple[pd.DataFrame, dict]:
    path = find_workbook("1.5.2")
    raw = pd.read_excel(path, header=None, dtype=object)
    meta = {
        "source_file": path.name,
        "title": raw.iat[0, 0],
        "filters": [value for value in raw.iloc[1, :8].tolist() if pd.notna(value)],
    }

    data = raw.iloc[5:].reset_index(drop=True).copy()
    data.columns = [f"c{i}" for i in range(data.shape[1])]
    data = data[data["c1"].astype(str).str.contains(r"\d", na=False)].copy()
    data["source_region"] = data["c0"].astype(str).str.strip()
    data["source_code"] = data["c1"].astype(str).str.strip()
    data["source_name"] = data["c2"].astype(str).str.strip()
    data["source_status"] = data["c3"].astype(str).str.strip()
    data["source_legal"] = data["c8"].astype(str).str.strip()
    data["code_norm"] = data["c1"].map(normalize_code)
    data["source_cashflow"] = as_number(data["c10"])
    grouped = (
        data.groupby(["source_code", "code_norm"], as_index=False)
        .agg(
            source_region=("source_region", "first"),
            source_name=("source_name", "first"),
            source_status=("source_status", "first"),
            source_legal=("source_legal", "first"),
            source_cashflow=("source_cashflow", "sum"),
        )
        .rename(columns={"source_code": "project_code"})
    )
    return grouped, meta


def load_report() -> tuple[pd.DataFrame, dict]:
    prefix = u(r"\u56de\u6b3e\u8425\u6536\u6bd4")
    suffix = u(r"\u9879\u76ee")
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and path.stem == f"{prefix}{REPORT_MONTH}{suffix}"
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one report workbook, got {len(matches)}: {[p.name for p in matches]}")
    path = matches[0]
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
    data = raw.iloc[3:].reset_index(drop=True).copy()
    data.columns = [f"c{i}" for i in range(data.shape[1])]
    data = data[data["c3"].astype(str).str.contains(r"\d", na=False)].copy()
    data["region"] = data["c1"].astype(str).str.strip()
    data["line"] = data["c2"].astype(str).str.strip()
    data["project_code"] = data["c3"].astype(str).str.strip()
    data["project_name"] = data["c4"].astype(str).str.strip()
    data["code_norm"] = data["c3"].map(normalize_code)
    data["report_cashflow"] = as_number(data["c5"])
    return data[["region", "line", "project_code", "project_name", "code_norm", "report_cashflow"]], {
        "report_file": path.name
    }


def attach_scope_flags(df: pd.DataFrame) -> pd.DataFrame:
    query_path = find_workbook(u(r"\u9879\u76ee\u67e5\u8be2"))
    query = pd.read_excel(query_path, dtype=object)
    query["code_norm"] = query[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code)
    query_flags = query[
        [
            "code_norm",
            u(r"\u9879\u76ee\u7b49\u7ea7"),
            u(r"\u9879\u76ee\u72b6\u6001"),
            u(r"\u5df2\u64a4\u573a\u65f6\u95f4"),
            u(r"\u7a7f\u900f\u6bd4\u4f8b"),
        ]
    ].drop_duplicates("code_norm")
    query_flags = query_flags.rename(
        columns={
            u(r"\u9879\u76ee\u7b49\u7ea7"): "project_level",
            u(r"\u9879\u76ee\u72b6\u6001"): "project_status",
            u(r"\u5df2\u64a4\u573a\u65f6\u95f4"): "exit_date",
            u(r"\u7a7f\u900f\u6bd4\u4f8b"): "penetration_ratio",
        }
    )

    non_assess_path = find_workbook(u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26"), exact_stem=True)
    non_assess = pd.read_excel(non_assess_path, dtype=object)
    non_assess_codes = set(non_assess[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code).dropna())

    result = df.merge(query_flags, on="code_norm", how="left")
    result["is_d_exit"] = result["project_level"].astype(str).str.startswith("D") & result[
        "project_status"
    ].astype(str).eq(u(r"\u5df2\u64a4\u573a"))
    result["is_non_assess"] = result["code_norm"].isin(non_assess_codes)
    return result


def summarize(merged: pd.DataFrame, extra_source: pd.DataFrame, indicator_row: dict, source_meta: dict, report_meta: dict) -> dict:
    diff = merged["source_cashflow"] - merged["report_cashflow"]
    mismatch = diff.abs() > TOLERANCE
    extra_nonzero = extra_source["source_cashflow"].abs() > TOLERANCE
    mismatch_rows = merged.loc[mismatch].copy()
    return {
        "status": "passed" if not mismatch.any() and not extra_nonzero.any() else "failed",
        "report_month": REPORT_MONTH,
        "indicator_row": indicator_row,
        "source": source_meta,
        "report": report_meta,
        "report_rows": int(len(merged)),
        "report_projects": int(merged["code_norm"].nunique()),
        "mismatch_rows": int(mismatch.sum()),
        "mismatch_d_exit_rows": int(mismatch_rows["is_d_exit"].sum()),
        "mismatch_d_exit_diff_total": float(mismatch_rows.loc[mismatch_rows["is_d_exit"], "diff"].sum()),
        "mismatch_non_assess_rows": int(mismatch_rows["is_non_assess"].sum()),
        "mismatch_non_assess_diff_total": float(mismatch_rows.loc[mismatch_rows["is_non_assess"], "diff"].sum()),
        "report_nonzero_source_zero_rows": int(
            ((merged["source_cashflow"].abs() <= TOLERANCE) & (merged["report_cashflow"].abs() > TOLERANCE)).sum()
        ),
        "source_total_for_report_projects": float(merged["source_cashflow"].sum()),
        "report_total": float(merged["report_cashflow"].sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "extra_source_projects_not_in_report": int(len(extra_source)),
        "extra_source_nonzero_projects_not_in_report": int(extra_nonzero.sum()),
        "extra_source_total_not_in_report": float(extra_source["source_cashflow"].sum()),
    }


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    indicator_row = load_indicator_row()
    source, source_meta = load_source()
    report, report_meta = load_report()

    merged = report.merge(source[["project_code", "source_cashflow"]], on="project_code", how="left")
    fallback = source.groupby("code_norm", as_index=False).agg(
        source_exact_count=("project_code", "nunique"),
        fallback_source_cashflow=("source_cashflow", "sum"),
    )
    merged = merged.merge(fallback, on="code_norm", how="left")
    missing_exact = merged["source_cashflow"].isna()
    no_prefix_code = ~merged["project_code"].str.match(r"^[A-Z]")
    use_fallback = missing_exact & no_prefix_code & merged["source_exact_count"].eq(1)
    merged.loc[use_fallback, "source_cashflow"] = merged.loc[use_fallback, "fallback_source_cashflow"]
    merged["source_cashflow"] = merged["source_cashflow"].fillna(0.0)
    merged["diff"] = merged["source_cashflow"] - merged["report_cashflow"]
    merged = attach_scope_flags(merged)

    extra_source = source[~source["project_code"].isin(set(report["project_code"]))].copy()
    summary = summarize(merged, extra_source, indicator_row, source_meta, report_meta)

    mismatch_detail = merged.loc[
        merged["diff"].abs() > TOLERANCE,
        [
            "region",
            "line",
            "project_code",
            "project_name",
            "report_cashflow",
            "source_cashflow",
            "diff",
            "project_level",
            "project_status",
            "exit_date",
            "penetration_ratio",
            "is_d_exit",
            "is_non_assess",
        ],
    ].sort_values("diff", key=lambda series: series.abs(), ascending=False)

    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
    print("\nMISMATCH_DETAIL")
    display_cols = [
        "region",
        "line",
        "project_code",
        "project_name",
        "report_cashflow",
        "source_cashflow",
        "diff",
        "project_level",
        "project_status",
        "is_d_exit",
        "is_non_assess",
    ]
    print(markdown_table(mismatch_detail[display_cols]))

    extra_nonzero = extra_source[extra_source["source_cashflow"].abs() > TOLERANCE].copy()
    if not extra_nonzero.empty:
        print("\nEXTRA_SOURCE_NOT_IN_REPORT_TOP20")
        print(markdown_table(extra_nonzero.sort_values("source_cashflow", key=lambda series: series.abs(), ascending=False).head(20)))


if __name__ == "__main__":
    main()
