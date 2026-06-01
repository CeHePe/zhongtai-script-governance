from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path

import pandas as pd

from _project_root import find_project_root


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
PREVIOUS_YEAR_MONTH = "202412"
VALUE_TOL = 1e-6
AMOUNT_TOL = 1e-5
ZERO_DENOMINATOR = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


PROJECT = u(r"\u9879\u76ee")
RATIO = u(r"\u56de\u6b3e\u8425\u6536\u6bd4")
INDICATOR_LIST = u(r"\u6307\u6807\u6e05\u5355")
TOTAL_CASH = u(r"\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41")


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and all(token in path.name for token in tokens)
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
    return text


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def safe_ratio(numerator: float, denominator: float) -> float:
    if abs(denominator) <= ZERO_DENOMINATOR:
        return 0.0
    return numerator / denominator


def safe_growth(current: float, previous: float) -> float:
    if abs(previous) <= ZERO_DENOMINATOR:
        return 0.0
    return (current - previous) / abs(previous)


def load_1312(month: str) -> pd.DataFrame:
    path = find_workbook("1.3.1.2", month, PROJECT)
    df = pd.read_excel(path, header=None, dtype=object).iloc[2:].reset_index(drop=True).copy()
    if df.shape[1] != 27:
        raise RuntimeError(f"Unexpected 1.3.1.2 column count in {path.name}: {df.shape[1]}")
    df = df.rename(
        columns={
            0: "region",
            1: "project_code",
            2: "project_name",
            7: "revenue_ratio",
            9: "revenue_ratio_yoy",
            10: "total_cash",
            12: "total_cash_yoy",
        }
    )
    df["code_norm"] = df["project_code"].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    for column in ["revenue_ratio", "revenue_ratio_yoy", "total_cash", "total_cash_yoy"]:
        df[column] = as_number(df[column])
    df["report_file"] = path.name
    return df[
        [
            "region",
            "project_code",
            "code_norm",
            "project_name",
            "revenue_ratio",
            "revenue_ratio_yoy",
            "total_cash",
            "total_cash_yoy",
            "report_file",
        ]
    ]


def load_ratio_source() -> pd.DataFrame:
    path = find_workbook(RATIO, REPORT_MONTH, PROJECT)
    df = pd.read_excel(path, header=None, dtype=object).iloc[3:].reset_index(drop=True).copy()
    df["code_norm"] = df.iloc[:, 3].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["ratio_display"] = as_number(df.iloc[:, 32])
    df["ratio_numerator"] = as_number(df.iloc[:, 33])
    df["ratio_denominator"] = as_number(df.iloc[:, 34])
    out = df.groupby("code_norm", as_index=False).agg(
        ratio_display=("ratio_display", "sum"),
        ratio_numerator=("ratio_numerator", "sum"),
        ratio_denominator=("ratio_denominator", "sum"),
        ratio_source_rows=("code_norm", "size"),
    )
    out["calc_revenue_ratio"] = [
        safe_ratio(numerator, denominator)
        for numerator, denominator in zip(out["ratio_numerator"], out["ratio_denominator"])
    ]
    out["ratio_source_file"] = path.name
    return out


def load_152_total_cash() -> pd.DataFrame:
    path = find_workbook("1.5.2", REPORT_MONTH)
    df = pd.read_excel(path, header=None, dtype=object).iloc[5:].reset_index(drop=True).copy()
    df["code_norm"] = df.iloc[:, 1].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["source_total_cash"] = as_number(df.iloc[:, 10]) / 10000.0
    out = df.groupby("code_norm", as_index=False).agg(
        calc_total_cash=("source_total_cash", "sum"),
        total_cash_source_rows=("code_norm", "size"),
    )
    out["total_cash_source_file"] = path.name
    return out


def load_indicator_rows() -> list[dict[str, object]]:
    path = find_workbook("JKS_", INDICATOR_LIST)
    df = pd.read_excel(path, header=None, dtype=object)
    wanted_rows = [756, 1363, 1374, 1385]
    rows: list[dict[str, object]] = []
    for excel_row in wanted_rows:
        row = df.iloc[excel_row - 1]
        values = [None if pd.isna(value) else value for value in row.tolist()]
        compact = [value for value in values if value not in (None, "")]
        rows.append({"excel_row": excel_row, "values": compact})
    return rows


def summarize(
    base: pd.DataFrame,
    label: str,
    report_col: str,
    calc_col: str,
    tolerance: float,
    missing_col: str | None = None,
) -> dict[str, object]:
    diff = base[report_col] - base[calc_col]
    mismatch = diff.abs() > tolerance
    missing_source_rows = int(base[missing_col].isna().sum()) if missing_col else int(base[calc_col].isna().sum())
    return {
        "check": label,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(base)),
        "mismatch_rows": int(mismatch.sum()),
        "missing_source_rows": missing_source_rows,
        "report_total": float(base[report_col].sum()),
        "calc_total": float(base[calc_col].fillna(0.0).sum()),
        "diff_total": float(diff.fillna(0.0).sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
    }


def sample_mismatches(
    base: pd.DataFrame,
    report_col: str,
    calc_col: str,
    tolerance: float,
    extra_cols: list[str] | None = None,
    limit: int = 6,
) -> list[dict[str, object]]:
    extra_cols = extra_cols or []
    diff = base[report_col] - base[calc_col]
    bad = base.loc[diff.abs() > tolerance].copy()
    bad["diff"] = diff.loc[bad.index]
    cols = ["region", "project_code", "project_name", report_col, calc_col, "diff", *extra_cols]
    return bad[cols].head(limit).to_dict(orient="records")


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    current = load_1312(REPORT_MONTH)
    previous = load_1312(PREVIOUS_YEAR_MONTH).rename(
        columns={
            "revenue_ratio": "prev_revenue_ratio",
            "total_cash": "prev_total_cash_report",
            "report_file": "previous_report_file",
        }
    )
    ratio_source = load_ratio_source()
    total_cash_source = load_152_total_cash()

    base = (
        current.merge(
            previous[["code_norm", "prev_revenue_ratio", "prev_total_cash_report", "previous_report_file"]],
            on="code_norm",
            how="left",
        )
        .merge(ratio_source, on="code_norm", how="left")
        .merge(total_cash_source, on="code_norm", how="left")
    )

    base["calc_revenue_ratio"] = base["calc_revenue_ratio"].fillna(0.0)
    base["calc_total_cash"] = base["calc_total_cash"].fillna(0.0)
    base["prev_revenue_ratio"] = base["prev_revenue_ratio"].fillna(0.0)
    base["prev_total_cash_report"] = base["prev_total_cash_report"].fillna(0.0)
    base["calc_revenue_ratio_yoy"] = base["calc_revenue_ratio"] - base["prev_revenue_ratio"]
    base["calc_total_cash_yoy"] = [
        safe_growth(current_value, previous_value)
        for current_value, previous_value in zip(base["calc_total_cash"], base["prev_total_cash_report"])
    ]

    summaries = [
        summarize(base, u(r"\u56de\u6b3e\u8425\u6536\u6bd4\uff08%\uff09"), "revenue_ratio", "calc_revenue_ratio", VALUE_TOL, "ratio_source_rows"),
        summarize(base, u(r"\u56de\u6b3e\u8425\u6536\u6bd4\u540c\u6bd4\uff08%\uff09"), "revenue_ratio_yoy", "calc_revenue_ratio_yoy", VALUE_TOL, "previous_report_file"),
        summarize(base, u(r"\u603b\u56de\u6b3e\u989d\uff08\u4e07\u5143\uff09"), "total_cash", "calc_total_cash", AMOUNT_TOL, "total_cash_source_rows"),
        summarize(base, u(r"\u603b\u56de\u6b3e\u989d\u540c\u6bd4\uff08%\uff09"), "total_cash_yoy", "calc_total_cash_yoy", VALUE_TOL, "previous_report_file"),
    ]

    result = {
        "period": REPORT_MONTH,
        "dimension": "project",
        "rules": {
            "revenue_ratio_current": u(r"\u56de\u6b3e\u8425\u6536\u6bd4202512\u9879\u76ee.xlsx: \u56de\u6b3e\u8425\u6536\u6bd4\u5206\u5b50 / \u56de\u6b3e\u8425\u6536\u6bd4\u5206\u6bcd"),
            "revenue_ratio_yoy": u(r"202512\u590d\u7b97\u56de\u6b3e\u8425\u6536\u6bd4 - 202412\u76841.3.1.2\u9879\u76ee\u62a5\u8868\u56de\u6b3e\u8425\u6536\u6bd4"),
            "total_cash_current": u(r"1.5.2\u534a\u6536\u4ed8\u53e3\u5f84\u5e95\u8868\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41 / 10000"),
            "total_cash_yoy": u(r"(202512\u590d\u7b97\u603b\u56de\u6b3e\u989d - 202412\u76841.3.1.2\u9879\u76ee\u62a5\u8868\u603b\u56de\u6b3e\u989d) / abs(202412\u603b\u56de\u6b3e\u989d); \u4e0a\u5e74\u7edd\u5bf9\u503c<=1e-6\u63090"),
        },
        "source_files": sorted(
            {
                find_workbook("1.3.1.2", REPORT_MONTH, PROJECT).name,
                find_workbook("1.3.1.2", PREVIOUS_YEAR_MONTH, PROJECT).name,
                find_workbook(RATIO, REPORT_MONTH, PROJECT).name,
                find_workbook("1.5.2", REPORT_MONTH).name,
            }
        ),
        "missing_sources": [
            u(r"\u672a\u627e\u5230202412\u76841.5.2\u534a\u6536\u4ed8\u53e3\u5f84\u5e95\u8868\uff1b\u603b\u56de\u6b3e\u989d\u540c\u6bd4\u53ea\u80fd\u6309202412\u76841.3.1.2\u9879\u76ee\u62a5\u8868\u603b\u56de\u6b3e\u989d\u5217\u4f5c\u4e3a\u4e0a\u5e74\u57fa\u6570\u505a\u62a5\u8868\u53e3\u5f84\u6821\u9a8c\uff0c\u4e0d\u80fd\u5b8c\u6574\u8ffd\u6eaf\u4e0a\u5e74\u6e90\u8868\u3002")
        ],
        "indicator_rows": load_indicator_rows(),
        "source_coverage": {
            "report_rows": int(len(base)),
            "ratio_source_matched_rows": int(base["ratio_source_rows"].notna().sum()),
            "total_cash_source_matched_rows": int(base["total_cash_source_rows"].notna().sum()),
            "previous_report_matched_rows": int(base["previous_report_file"].notna().sum()),
            "previous_report_total_cash_nonzero_rows": int((base["prev_total_cash_report"].abs() > VALUE_TOL).sum()),
        },
        "summaries": summaries,
        "samples": {
            "revenue_ratio_yoy": sample_mismatches(
                base,
                "revenue_ratio_yoy",
                "calc_revenue_ratio_yoy",
                VALUE_TOL,
                ["calc_revenue_ratio", "prev_revenue_ratio"],
            ),
            "total_cash": sample_mismatches(
                base,
                "total_cash",
                "calc_total_cash",
                AMOUNT_TOL,
                ["total_cash_source_rows"],
            ),
        },
    }

    output_dir = ROOT / "local_outputs"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"validate_1312_project_cash_backflow_metrics_{REPORT_MONTH}.json"
    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nWrote {output_path}")


if __name__ == "__main__":
    main()
