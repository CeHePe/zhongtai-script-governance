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
TOL = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


PROJECT = u(r"\u9879\u76ee")
CURRENT_REPORT = "1.3.1.1"
REVENUE_RATIO = u(r"\u56de\u6b3e\u8425\u6536\u6bd4")
CURRENT_RATE = u(r"\u5f53\u671f\u56de\u6b3e\u7387")
PREVIOUS_YEAR_RATE = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387")
PAST_YEAR_RATE = u(r"\u5f80\u5e74\u56de\u6b3e\u7387")
CONTINUOUS_RATE = u(r"\u8fde\u7eed\u56de\u6b3e\u7387")


REPORT_COLUMNS = [
    "region",
    "project_code",
    "project_name",
    "geo",
    "province",
    "city",
    "district",
    "revenue_ratio",
    "revenue_ratio_mom",
    "revenue_ratio_yoy",
    "overall_rate",
    "overall_rate_mom",
    "overall_rate_yoy",
    "current_rate",
    "current_rate_mom",
    "current_rate_yoy",
    "arrears_rate",
    "arrears_rate_mom",
    "arrears_rate_yoy",
    "previous_year_rate",
    "previous_year_rate_mom",
    "previous_year_rate_yoy",
    "past_year_rate",
    "past_year_rate_mom",
    "past_year_rate_yoy",
    "n12_rate",
    "n12_rate_mom",
    "n12_rate_yoy",
    "n24_rate",
    "n24_rate_mom",
    "n24_rate_yoy",
    "n36_rate",
    "n36_rate_mom",
    "n36_rate_yoy",
    "overdue3_rate",
    "overdue3_rate_mom",
    "overdue3_rate_yoy",
]


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def clean_value(value: object) -> object:
    if pd.isna(value):
        return None
    return value


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


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens!r}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def safe_rate(numerator: float, denominator: float) -> float:
    if abs(denominator) <= TOL:
        return 0.0
    return numerator / denominator


def load_1311(month: str) -> pd.DataFrame:
    path = find_workbook(CURRENT_REPORT, month, PROJECT)
    df = pd.read_excel(path, header=None, dtype=object).iloc[2:].reset_index(drop=True).copy()
    if df.shape[1] != len(REPORT_COLUMNS):
        raise RuntimeError(f"Unexpected 1.3.1.1 column count in {path.name}: {df.shape[1]}")
    df.columns = REPORT_COLUMNS
    df["code_norm"] = df["project_code"].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    for column in REPORT_COLUMNS[7:]:
        df[column] = as_number(df[column])
    df["report_file"] = path.name
    return df


def load_indicator_rows() -> list[dict[str, object]]:
    indicator = find_workbook("JKS_", u(r"\u6307\u6807\u6e05\u5355"))
    df = pd.read_excel(indicator, sheet_name=0, dtype=object)
    name_col = df.columns[3]
    dim_col = df.columns[4]
    period_col = df.columns[5]

    wanted = {
        REVENUE_RATIO,
        PREVIOUS_YEAR_RATE,
        PAST_YEAR_RATE,
        u(r"\u5386\u6b20\u56de\u6b3e\u7387"),
        u(r"\u7efc\u5408\u56de\u6b3e\u7387"),
        "N+12" + CONTINUOUS_RATE,
        "N+24" + CONTINUOUS_RATE,
        "N+36" + CONTINUOUS_RATE,
    }
    rows: list[dict[str, object]] = []
    for idx, row in df.iterrows():
        metric = "" if pd.isna(row[name_col]) else str(row[name_col]).strip()
        relation = "" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip()
        component = "" if pd.isna(row.iloc[2]) else str(row.iloc[2]).strip()
        if str(row[dim_col]).strip() != PROJECT or str(row[period_col]).strip() != u(r"\u7d2f\u8ba1"):
            continue
        if metric not in wanted and relation not in wanted and component not in wanted and CURRENT_RATE not in relation:
            continue
        rows.append(
            {
                "excel_row": int(idx + 2),
                "serial": clean_value(row.iloc[0]),
                "relation": clean_value(row.iloc[1]),
                "component": clean_value(row.iloc[2]),
                "metric": clean_value(row[name_col]),
                "dimension": clean_value(row[dim_col]),
                "period": clean_value(row[period_col]),
                "method": clean_value(row.iloc[8]),
                "source_table": clean_value(row.iloc[10]),
                "logic": clean_value(row.iloc[12]),
            }
        )
    return rows


def load_ratio_source(tokens: tuple[str, ...], skip_rows: int, code_col: int, rate_col: int) -> pd.DataFrame:
    path = find_workbook(*tokens)
    df = pd.read_excel(path, header=None, dtype=object).iloc[skip_rows:].reset_index(drop=True).copy()
    df["code_norm"] = df.iloc[:, code_col].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["source_rate"] = as_number(df.iloc[:, rate_col])
    # Project底表通常一项目一行；若出现重复，保留逐行求和以暴露重复导致的差异。
    return df.groupby("code_norm", as_index=False).agg(source_rate=("source_rate", "sum"), source_rows=("code_norm", "size"))


def load_num_den_source(tokens: tuple[str, ...], skip_rows: int, code_col: int, numerator_col: int, denominator_col: int) -> pd.DataFrame:
    path = find_workbook(*tokens)
    df = pd.read_excel(path, header=None, dtype=object).iloc[skip_rows:].reset_index(drop=True).copy()
    df["code_norm"] = df.iloc[:, code_col].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["numerator"] = as_number(df.iloc[:, numerator_col])
    df["denominator"] = as_number(df.iloc[:, denominator_col])
    out = df.groupby("code_norm", as_index=False).agg(
        numerator=("numerator", "sum"),
        denominator=("denominator", "sum"),
        source_rows=("code_norm", "size"),
    )
    out["source_rate"] = [safe_rate(n, d) for n, d in zip(out["numerator"], out["denominator"])]
    return out


def load_sources() -> pd.DataFrame:
    revenue = load_ratio_source((REVENUE_RATIO, REPORT_MONTH, PROJECT), 3, 3, 32).rename(
        columns={"source_rate": "calc_revenue_ratio", "source_rows": "revenue_source_rows"}
    )
    current = load_num_den_source((CURRENT_RATE, REPORT_MONTH, PROJECT), 3, 3, 17, 28).rename(
        columns={
            "numerator": "current_numerator",
            "denominator": "current_denominator",
            "source_rate": "calc_current_rate",
            "source_rows": "current_source_rows",
        }
    )
    previous_year = load_num_den_source((PREVIOUS_YEAR_RATE, REPORT_MONTH, PROJECT), 2, 3, 12, 20).rename(
        columns={
            "numerator": "previous_year_numerator",
            "denominator": "previous_year_denominator",
            "source_rate": "calc_previous_year_rate",
            "source_rows": "previous_year_source_rows",
        }
    )
    past_year = load_num_den_source((PAST_YEAR_RATE, REPORT_MONTH, PROJECT), 2, 3, 10, 17).rename(
        columns={
            "numerator": "past_year_numerator",
            "denominator": "past_year_denominator",
            "source_rate": "calc_past_year_rate",
            "source_rows": "past_year_source_rows",
        }
    )

    base = (
        revenue.merge(current, on="code_norm", how="outer")
        .merge(previous_year, on="code_norm", how="outer")
        .merge(past_year, on="code_norm", how="outer")
    )
    for column in [
        "calc_revenue_ratio",
        "current_numerator",
        "current_denominator",
        "calc_current_rate",
        "previous_year_numerator",
        "previous_year_denominator",
        "calc_previous_year_rate",
        "past_year_numerator",
        "past_year_denominator",
        "calc_past_year_rate",
    ]:
        base[column] = base[column].fillna(0.0)
    base["calc_arrears_rate"] = [
        safe_rate(n, d)
        for n, d in zip(
            base["previous_year_numerator"] + base["past_year_numerator"],
            base["previous_year_denominator"] + base["past_year_denominator"],
        )
    ]
    base["calc_overall_rate"] = [
        safe_rate(n, d)
        for n, d in zip(
            base["current_numerator"] + base["previous_year_numerator"] + base["past_year_numerator"],
            base["current_denominator"] + base["previous_year_denominator"] + base["past_year_denominator"],
        )
    ]

    for n in [12, 24, 36]:
        source = load_ratio_source((f"N+{n}", CONTINUOUS_RATE, REPORT_MONTH, PROJECT), 3, 3, 18).rename(
            columns={"source_rate": f"calc_n{n}_rate", "source_rows": f"n{n}_source_rows"}
        )
        base = base.merge(source, on="code_norm", how="outer")
        base[f"calc_n{n}_rate"] = base[f"calc_n{n}_rate"].fillna(0.0)
    return base


def summarize(base: pd.DataFrame, label: str, report_col: str, calc_col: str) -> dict[str, object]:
    diff = base[report_col] - base[calc_col]
    mismatch = diff.abs() > TOL
    top = base.loc[mismatch, ["project_code", "project_name", report_col, calc_col]].copy()
    top["diff"] = diff[mismatch]
    top = top.reindex(top["diff"].abs().sort_values(ascending=False).index).head(8)
    return {
        "check": label,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(base)),
        "mismatch_rows": int(mismatch.sum()),
        "report_total": float(base[report_col].sum()),
        "calc_total": float(base[calc_col].sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "top_mismatches": top.to_dict("records"),
    }


def source_gap_counts(base: pd.DataFrame) -> dict[str, int]:
    checks = {
        "revenue_ratio": ("revenue_source_rows", "revenue_ratio"),
        "current_rate": ("current_source_rows", "current_rate"),
        "previous_year_rate": ("previous_year_source_rows", "previous_year_rate"),
        "past_year_rate": ("past_year_source_rows", "past_year_rate"),
        "n12_rate": ("n12_source_rows", "n12_rate"),
        "n24_rate": ("n24_source_rows", "n24_rate"),
        "n36_rate": ("n36_source_rows", "n36_rate"),
    }
    result: dict[str, int] = {}
    for label, (source_col, report_col) in checks.items():
        missing = base[source_col].isna()
        result[f"{label}_missing_source_rows"] = int(missing.sum())
        result[f"{label}_missing_source_nonzero_report_rows"] = int((missing & (base[report_col].abs() > TOL)).sum())
    return result


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    report = load_1311(REPORT_MONTH)
    previous_report = load_1311(PREVIOUS_YEAR_MONTH).drop_duplicates("code_norm", keep="first")
    sources = load_sources()
    base = report.merge(sources, on="code_norm", how="left")
    for column in [column for column in base.columns if column.startswith("calc_")]:
        base[column] = base[column].fillna(0.0)

    previous_values = previous_report.set_index("code_norm")[
        [
            "revenue_ratio",
            "overall_rate",
            "current_rate",
            "arrears_rate",
            "previous_year_rate",
            "past_year_rate",
            "n12_rate",
            "n24_rate",
            "n36_rate",
        ]
    ].rename(columns=lambda column: f"prev_{column}")
    base = base.merge(previous_values, on="code_norm", how="left")
    for metric in [
        "revenue_ratio",
        "overall_rate",
        "current_rate",
        "arrears_rate",
        "previous_year_rate",
        "past_year_rate",
        "n12_rate",
        "n24_rate",
        "n36_rate",
    ]:
        base[f"calc_{metric}_yoy"] = (base[metric] - base[f"prev_{metric}"]).fillna(0.0)

    value_checks = [
        ("回款营收比", "revenue_ratio", "calc_revenue_ratio"),
        ("综合回款率", "overall_rate", "calc_overall_rate"),
        ("当期回款率", "current_rate", "calc_current_rate"),
        ("历欠回款率", "arrears_rate", "calc_arrears_rate"),
        ("上一年回款率", "previous_year_rate", "calc_previous_year_rate"),
        ("往年回款率", "past_year_rate", "calc_past_year_rate"),
        ("N+12连续回款率", "n12_rate", "calc_n12_rate"),
        ("N+24连续回款率", "n24_rate", "calc_n24_rate"),
        ("N+36连续回款率", "n36_rate", "calc_n36_rate"),
    ]
    yoy_checks = [
        (f"{label}_同比", f"{report_col}_yoy", f"calc_{report_col}_yoy")
        for label, report_col, _ in value_checks
    ]

    result = {
        "indicator_rows": load_indicator_rows(),
        "source_files": {
            "current_report": find_workbook(CURRENT_REPORT, REPORT_MONTH, PROJECT).name,
            "previous_yoy_report": find_workbook(CURRENT_REPORT, PREVIOUS_YEAR_MONTH, PROJECT).name,
            "revenue_ratio": find_workbook(REVENUE_RATIO, REPORT_MONTH, PROJECT).name,
            "current_rate": find_workbook(CURRENT_RATE, REPORT_MONTH, PROJECT).name,
            "previous_year_rate": find_workbook(PREVIOUS_YEAR_RATE, REPORT_MONTH, PROJECT).name,
            "past_year_rate": find_workbook(PAST_YEAR_RATE, REPORT_MONTH, PROJECT).name,
            "n12": find_workbook("N+12", CONTINUOUS_RATE, REPORT_MONTH, PROJECT).name,
            "n24": find_workbook("N+24", CONTINUOUS_RATE, REPORT_MONTH, PROJECT).name,
            "n36": find_workbook("N+36", CONTINUOUS_RATE, REPORT_MONTH, PROJECT).name,
        },
        "scope": {
            "tested": [label for label, _, _ in value_checks] + [label for label, _, _ in yoy_checks],
            "not_tested": [u(r"\u73af\u6bd4\u6307\u6807"), u(r"\u653f\u4f01\u5355\u4e00\u4e1a\u6743\u8d85\u8d26\u671f\u4e09\u4e2a\u6708\u5185\u56de\u6b3e\u7387")],
        },
        "counts": {
            "current_report_rows": int(len(report)),
            "previous_yoy_report_rows": int(len(previous_report)),
            "current_duplicate_code_rows": int(report.duplicated("code_norm").sum()),
            "previous_duplicate_code_rows": int(previous_report.duplicated("code_norm").sum()),
            **source_gap_counts(base),
        },
        "value_summary": [summarize(base, *check) for check in value_checks],
        "yoy_summary": [summarize(base, *check) for check in yoy_checks],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
