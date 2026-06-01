from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path

import pandas as pd

from _project_root import find_project_root


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
TOLERANCE = 1e-6
RATIO_DENOMINATOR_ZERO_THRESHOLD = 0.01


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str, exact_stem: str | None = None, exact_name: str | None = None) -> Path:
    matches: list[Path]
    if exact_name is not None:
        matches = [path for path in ROOT.iterdir() if path.is_file() and path.name == exact_name]
    elif exact_stem is not None:
        matches = [path for path in ROOT.iterdir() if path.is_file() and path.stem == exact_stem]
    else:
        matches = [
            path
            for path in ROOT.iterdir()
            if path.is_file() and path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
        ]
    if len(matches) != 1:
        target = exact_name or exact_stem or tokens
        raise RuntimeError(f"Expected one workbook for {target!r}, got {len(matches)}: {[p.name for p in matches]}")
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
    rows: list[list[str]] = []
    for _, row in df.iterrows():
        current: list[str] = []
        for value in row.tolist():
            if isinstance(value, float):
                current.append(format(value, floatfmt))
            else:
                current.append("" if pd.isna(value) else str(value))
        rows.append(current)
    lines = [
        "| " + " | ".join(str(header) for header in headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    lines.extend("| " + " | ".join(cell.replace("\n", " ") for cell in row) + " |" for row in rows)
    return "\n".join(lines)


def load_indicator_rows() -> list[dict[str, object]]:
    indicator = find_workbook(u(r"\u6307\u6807\u6e05\u5355"))
    df = pd.read_excel(indicator, sheet_name=0, dtype=object)
    serial_col = df.columns[0]
    relation_col = df.columns[1]
    component_col = df.columns[2]
    metric_col = df.columns[3]
    dimension_col = df.columns[4]
    period_col = df.columns[5]
    method_col = df.columns[8]
    source_col = df.columns[10]
    logic_col = df.columns[12]

    wanted = [
        (u(r"\u56de\u6b3e\u8425\u6536\u6bd4"), u(r"\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41"), u(r"\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41_\u9879\u76ee")),
        ("", u(r"\u56de\u6b3e\u8425\u6536\u6bd4_\u5206\u5b50"), u(r"\u56de\u6b3e\u8425\u6536\u6bd4_\u5206\u5b50_\u9879\u76ee")),
        ("", u(r"\u56de\u6b3e\u8425\u6536\u6bd4_\u5206\u6bcd"), u(r"\u56de\u6b3e\u8425\u6536\u6bd4_\u5206\u6bcd_\u9879\u76ee")),
        ("", u(r"\u56de\u6b3e\u8425\u6536\u6bd4"), u(r"\u56de\u6b3e\u8425\u6536\u6bd4")),
    ]

    rows: list[dict[str, object]] = []
    for relation, component, metric in wanted:
        target = df[
            df[component_col].astype(str).str.strip().eq(component)
            & df[metric_col].astype(str).str.strip().eq(metric)
            & df[dimension_col].astype(str).str.strip().eq(u(r"\u9879\u76ee"))
            & df[period_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1"))
        ].copy()
        if relation:
            target = target[target[relation_col].astype(str).str.strip().eq(relation)]
        if len(target) != 1:
            raise RuntimeError(f"Expected one indicator row for {(relation, component, metric)}, got {len(target)}")
        idx = target.index[0]
        row = target.iloc[0]
        rows.append(
            {
                "excel_row": int(idx + 2),
                "serial": row.get(serial_col),
                "relation": row.get(relation_col),
                "component": row.get(component_col),
                "metric": row.get(metric_col),
                "dimension": row.get(dimension_col),
                "period": row.get(period_col),
                "source_method": row.get(method_col),
                "source_table": row.get(source_col),
                "logic": row.get(logic_col),
            }
        )
    return rows


def load_source_cashflow() -> tuple[pd.DataFrame, dict[str, object]]:
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
    data["project_code"] = data["c1"].astype(str).str.strip()
    data["code_norm"] = data["c1"].map(normalize_code)
    data["source_cashflow"] = as_number(data["c10"])
    grouped = (
        data.groupby(["project_code", "code_norm"], as_index=False)["source_cashflow"]
        .sum()
        .sort_values(["code_norm", "project_code"])
        .reset_index(drop=True)
    )
    return grouped, meta


def load_report() -> tuple[pd.DataFrame, dict[str, object]]:
    path = find_workbook(exact_stem=u(r"\u56de\u6b3e\u8425\u6536\u6bd4") + REPORT_MONTH + u(r"\u9879\u76ee"))
    raw = pd.read_excel(path, header=None, dtype=object)
    data = raw.iloc[3:].reset_index(drop=True).copy()
    data = data[data[3].astype(str).str.contains(r"\d", na=False)].copy()
    data.columns = [
        "idx",
        "region",
        "line",
        "project_code",
        "project_name",
        "cashflow",
        "non_assess_cashflow",
        "related_cashflow",
        "non_assess_related_cashflow",
        "water_recv",
        "non_assess_water_recv",
        "water_recv_alloc",
        "non_assess_water_recv_alloc",
        "related_water_curr_recv",
        "related_water_old_recv",
        "house_offset",
        "coin_recovery",
        "revenue",
        "non_assess_revenue",
        "related_revenue",
        "non_assess_related_revenue",
        "current_year_discount",
        "related_discount_fin_q",
        "total_discount",
        "non_assess_total_discount",
        "related_discount_new_window",
        "water_receivable",
        "water_receivable_alloc",
        "related_water_curr_receivable",
        "not_due_current",
        "not_due_prev_end",
        "coin_current",
        "ratio",
        "numerator",
        "denominator",
        "project_status_report",
    ]
    for column in data.columns[5:35]:
        data[column] = as_number(data[column])
    data["project_code"] = data["project_code"].astype(str).str.strip()
    data["project_name"] = data["project_name"].astype(str).str.strip()
    data["code_norm"] = data["project_code"].map(normalize_code)
    meta = {
        "report_file": path.name,
        "layout_note": u(r"\u62a5\u8868\u5c3e\u90e8\u5b9e\u9645\u6570\u636e\u5217\u987a\u5e8f\u4e3a AG=\u56de\u6b3e\u8425\u6536\u6bd4, AH=\u56de\u6b3e\u8425\u6536\u6bd4\u5206\u5b50, AI=\u56de\u6b3e\u8425\u6536\u6bd4\u5206\u6bcd, AJ=\u9879\u76ee\u72b6\u6001\u3002"),
    }
    return data, meta


def attach_scope_flags(df: pd.DataFrame) -> pd.DataFrame:
    query = pd.read_excel(find_workbook(exact_name=u(r"\u9879\u76ee\u67e5\u8be2.xlsx")), dtype=object)
    query["code_norm"] = query[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code)
    query = query[
        [
            "code_norm",
            u(r"\u9879\u76ee\u7b49\u7ea7"),
            u(r"\u9879\u76ee\u72b6\u6001"),
            u(r"\u5df2\u64a4\u573a\u65f6\u95f4"),
            u(r"\u7a7f\u900f\u6bd4\u4f8b"),
        ]
    ].drop_duplicates("code_norm")
    query.columns = ["code_norm", "project_level", "project_status", "exit_date", "penetration_ratio"]

    non_assess = pd.read_excel(find_workbook(exact_name=u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26.xlsx")), dtype=object)
    non_assess_codes = set(non_assess[u(r"\u7acb\u9879\u7f16\u7801")].map(normalize_code).dropna())

    out = df.merge(query, on="code_norm", how="left")
    out["is_non_assess"] = out["code_norm"].isin(non_assess_codes)
    out["is_d_exit"] = out["project_level"].astype(str).str.startswith("D") & out["project_status"].astype(str).eq(
        u(r"\u5df2\u64a4\u573a")
    )
    return out


def attach_source_cashflow(report: pd.DataFrame, source: pd.DataFrame) -> pd.DataFrame:
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
    return merged


def add_calculations(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["calc_numerator"] = (
        out["cashflow"]
        - out["non_assess_cashflow"]
        - (out["related_cashflow"] - out["non_assess_related_cashflow"])
        + out["water_recv"]
        - out["non_assess_water_recv"]
        + out["water_recv_alloc"]
        - out["non_assess_water_recv_alloc"]
        - (out["related_water_curr_recv"] + out["related_water_old_recv"])
        - out["coin_recovery"]
        + out["house_offset"]
    )
    out["calc_denominator_indicator"] = (
        out["revenue"] * 1.06
        - out["non_assess_revenue"] * 1.06
        - (out["related_revenue"] * 1.06 - out["non_assess_related_revenue"] * 1.06)
        - out["total_discount"]
        + out["current_year_discount"]
        + out["related_discount_fin_q"]
        + out["related_discount_new_window"]
        + out["water_receivable"]
        + out["water_receivable_alloc"]
        - out["related_water_curr_receivable"]
        - out["not_due_current"]
        + out["not_due_prev_end"]
        - out["coin_current"]
    )
    out["calc_denominator_best_visible"] = (
        out["revenue"] * 1.06
        - out["non_assess_revenue"] * 1.06
        - (out["related_revenue"] * 1.06 - out["non_assess_related_revenue"] * 1.06)
        - out["total_discount"]
        + out["current_year_discount"]
        + out["related_discount_fin_q"]
        + out["related_discount_new_window"]
        + out["water_receivable"]
        - out["related_water_curr_receivable"]
        - out["not_due_current"]
        + out["not_due_prev_end"]
        - out["coin_current"]
    )
    out["calc_ratio"] = 0.0
    mask = out["denominator"].abs() >= RATIO_DENOMINATOR_ZERO_THRESHOLD
    out.loc[mask, "calc_ratio"] = out.loc[mask, "numerator"] / out.loc[mask, "denominator"]
    out["cashflow_diff"] = out["source_cashflow"] - out["cashflow"]
    out["numerator_diff"] = out["calc_numerator"] - out["numerator"]
    out["denominator_indicator_diff"] = out["calc_denominator_indicator"] - out["denominator"]
    out["denominator_best_visible_diff"] = out["calc_denominator_best_visible"] - out["denominator"]
    out["ratio_diff"] = out["calc_ratio"] - out["ratio"]
    return out


def summarize(df: pd.DataFrame, label: str, actual: str, calc: str, diff: str) -> dict[str, object]:
    mismatch = df[diff].abs() > TOLERANCE
    return {
        "check": label,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(df)),
        "mismatch_rows": int(mismatch.sum()),
        "actual_total": float(df[actual].sum()),
        "calc_total": float(df[calc].sum()),
        "diff_total": float(df[diff].sum()),
        "max_abs_diff": float(df[diff].abs().max() if len(df) else 0.0),
    }


def mismatch_table(df: pd.DataFrame, diff_col: str, cols: list[str], limit: int = 20) -> pd.DataFrame:
    out = df.loc[df[diff_col].abs() > TOLERANCE, cols].copy()
    if out.empty:
        return out
    return out.assign(_abs=df[diff_col].abs()).sort_values("_abs", ascending=False).drop(columns="_abs").head(limit)


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    supplement_hits = [
        path.name
        for path in ROOT.iterdir()
        if path.is_file()
        and (u(r"\u589e\u8865") in path.name or u(r"\u56de\u6b3e\u516c\u5f0f") in path.name or u(r"\u672a\u5230\u8d26\u671f\u4f59\u989d") in path.name)
    ]
    indicator_rows = load_indicator_rows()
    source, source_meta = load_source_cashflow()
    report, report_meta = load_report()
    report = attach_scope_flags(add_calculations(attach_source_cashflow(report, source)))

    checks = {
        "cashflow_source": summarize(report, "cashflow_source", "cashflow", "source_cashflow", "cashflow_diff"),
        "numerator_formula": summarize(report, "numerator_formula", "numerator", "calc_numerator", "numerator_diff"),
        "denominator_indicator_formula": summarize(
            report, "denominator_indicator_formula", "denominator", "calc_denominator_indicator", "denominator_indicator_diff"
        ),
        "denominator_best_visible_formula": summarize(
            report,
            "denominator_best_visible_formula",
            "denominator",
            "calc_denominator_best_visible",
            "denominator_best_visible_diff",
        ),
        "ratio_formula": summarize(report, "ratio_formula", "ratio", "calc_ratio", "ratio_diff"),
    }
    overall_status = "passed"
    for required in ["cashflow_source", "numerator_formula", "denominator_indicator_formula", "ratio_formula"]:
        if checks[required]["status"] != "passed":
            overall_status = "failed"
            break

    summary = {
        "status": overall_status,
        "report_month": REPORT_MONTH,
        "report_file": report_meta["report_file"],
        "report_layout_note": report_meta["layout_note"],
        "report_rows": int(len(report)),
        "project_codes": int(report["code_norm"].nunique()),
        "supplement_formula_workbook_found": supplement_hits,
        "indicator_rows": indicator_rows,
        "cashflow_source_meta": source_meta,
        "checks": checks,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))

    print("\nCASHFLOW_SOURCE_MISMATCH_TOP20")
    print(
        markdown_table(
            mismatch_table(
                report,
                "cashflow_diff",
                [
                    "region",
                    "line",
                    "project_code",
                    "project_name",
                    "cashflow",
                    "source_cashflow",
                    "cashflow_diff",
                    "project_level",
                    "project_status",
                ],
            )
        )
    )

    print("\nNUMERATOR_FORMULA_MISMATCH_TOP20")
    print(
        markdown_table(
            mismatch_table(
                report,
                "numerator_diff",
                [
                    "region",
                    "line",
                    "project_code",
                    "project_name",
                    "numerator",
                    "calc_numerator",
                    "numerator_diff",
                ],
            )
        )
    )

    print("\nDENOMINATOR_INDICATOR_FORMULA_MISMATCH_TOP20")
    print(
        markdown_table(
            mismatch_table(
                report,
                "denominator_indicator_diff",
                [
                    "region",
                    "line",
                    "project_code",
                    "project_name",
                    "denominator",
                    "calc_denominator_indicator",
                    "denominator_indicator_diff",
                    "water_receivable_alloc",
                    "project_level",
                    "project_status",
                ],
            )
        )
    )

    print("\nDENOMINATOR_BEST_VISIBLE_FORMULA_MISMATCH_TOP20")
    print(
        markdown_table(
            mismatch_table(
                report,
                "denominator_best_visible_diff",
                [
                    "region",
                    "line",
                    "project_code",
                    "project_name",
                    "denominator",
                    "calc_denominator_best_visible",
                    "denominator_best_visible_diff",
                    "water_receivable_alloc",
                    "project_level",
                    "project_status",
                ],
            )
        )
    )

    print("\nRATIO_FORMULA_MISMATCH_ALL")
    print(
        markdown_table(
            mismatch_table(
                report,
                "ratio_diff",
                [
                    "region",
                    "line",
                    "project_code",
                    "project_name",
                    "ratio",
                    "numerator",
                    "denominator",
                    "calc_ratio",
                    "ratio_diff",
                    "project_status_report",
                ],
                limit=50,
            )
        )
    )


if __name__ == "__main__":
    main()
