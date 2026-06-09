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
MAIN_BUSINESS = u(r"\u4e3b\u8425\u4e1a\u52a1")
PREVIOUS_YEAR_RATE = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387")
PAST_YEAR_RATE = u(r"\u5f80\u5e74\u56de\u6b3e\u7387")
INDICATOR_LIST = u(r"\u6307\u6807\u6e05\u5355")


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


def safe_rate(numerator: float, denominator: float) -> float:
    if abs(denominator) <= TOL:
        return 0.0
    return numerator / denominator


def load_main_business(month: str) -> pd.DataFrame:
    path = find_workbook("1._", MAIN_BUSINESS, month, PROJECT)
    df = pd.read_excel(path, header=None, dtype=object).iloc[2:].reset_index(drop=True).copy()
    if df.shape[1] != 54:
        raise RuntimeError(f"Unexpected main-business column count in {path.name}: {df.shape[1]}")
    df = df.rename(
        columns={
            1: "region",
            2: "line",
            3: "project_code",
            4: "project_name",
            30: "arrears_rate",
            31: "prev_arrears_rate",
            32: "arrears_rate_yoy",
        }
    )
    df["code_norm"] = df["project_code"].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    for column in ["arrears_rate", "prev_arrears_rate", "arrears_rate_yoy"]:
        df[column] = as_number(df[column])
    df["report_file"] = path.name
    return df[
        [
            "region",
            "line",
            "project_code",
            "code_norm",
            "project_name",
            "arrears_rate",
            "prev_arrears_rate",
            "arrears_rate_yoy",
            "report_file",
        ]
    ]


def load_num_den_source(tokens: tuple[str, ...], numerator_col: int, denominator_col: int) -> pd.DataFrame:
    path = find_workbook(*tokens)
    df = pd.read_excel(path, header=None, dtype=object).iloc[2:].reset_index(drop=True).copy()
    df["code_norm"] = df.iloc[:, 3].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["numerator"] = as_number(df.iloc[:, numerator_col])
    df["denominator"] = as_number(df.iloc[:, denominator_col])
    out = df.groupby("code_norm", as_index=False).agg(
        numerator=("numerator", "sum"),
        denominator=("denominator", "sum"),
        source_rows=("code_norm", "size"),
    )
    out["source_file"] = path.name
    return out


def load_current_arrears_source() -> pd.DataFrame:
    previous_year = load_num_den_source((PREVIOUS_YEAR_RATE, REPORT_MONTH, PROJECT), 12, 20).rename(
        columns={
            "numerator": "previous_year_numerator",
            "denominator": "previous_year_denominator",
            "source_rows": "previous_year_source_rows",
            "source_file": "previous_year_source_file",
        }
    )
    past_year = load_num_den_source((PAST_YEAR_RATE, REPORT_MONTH, PROJECT), 10, 17).rename(
        columns={
            "numerator": "past_year_numerator",
            "denominator": "past_year_denominator",
            "source_rows": "past_year_source_rows",
            "source_file": "past_year_source_file",
        }
    )
    base = previous_year.merge(past_year, on="code_norm", how="outer")
    for column in [
        "previous_year_numerator",
        "previous_year_denominator",
        "past_year_numerator",
        "past_year_denominator",
    ]:
        base[column] = base[column].fillna(0.0)
    base["calc_arrears_rate"] = [
        safe_rate(numerator, denominator)
        for numerator, denominator in zip(
            base["previous_year_numerator"] + base["past_year_numerator"],
            base["previous_year_denominator"] + base["past_year_denominator"],
        )
    ]
    return base


def load_indicator_rows() -> list[dict[str, object]]:
    path = find_workbook("JKS_", INDICATOR_LIST)
    df = pd.read_excel(path, header=None, dtype=object)
    wanted_rows = [1608, 1619, 1641, 1696, 1784]
    rows: list[dict[str, object]] = []
    for excel_row in wanted_rows:
        row = df.iloc[excel_row - 1]
        compact = [None if pd.isna(value) else value for value in row.tolist()]
        compact = [value for value in compact if value not in (None, "")]
        rows.append({"excel_row": excel_row, "values": compact})
    return rows


def summarize(base: pd.DataFrame, label: str, report_col: str, calc_col: str) -> dict[str, object]:
    diff = base[report_col] - base[calc_col]
    mismatch = diff.abs() > TOL
    return {
        "check": label,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(base)),
        "mismatch_rows": int(mismatch.sum()),
        "report_total": float(base[report_col].sum()),
        "calc_total": float(base[calc_col].sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
    }


def sample_mismatches(base: pd.DataFrame, report_col: str, calc_col: str, limit: int = 8) -> list[dict[str, object]]:
    diff = base[report_col] - base[calc_col]
    bad = base.loc[diff.abs() > TOL].copy()
    bad["diff"] = diff.loc[bad.index]
    columns = ["region", "line", "project_code", "project_name", report_col, calc_col, "diff"]
    return bad[columns].head(limit).to_dict(orient="records")


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    current_report = load_main_business(REPORT_MONTH)
    previous_report = load_main_business(PREVIOUS_YEAR_MONTH).rename(
        columns={
            "arrears_rate": "previous_report_arrears_rate",
            "report_file": "previous_report_file",
        }
    )
    current_source = load_current_arrears_source()

    base = current_report.merge(
        current_source,
        on="code_norm",
        how="left",
    ).merge(
        previous_report[["code_norm", "previous_report_arrears_rate", "previous_report_file"]],
        on="code_norm",
        how="left",
    )

    base["calc_arrears_rate"] = base["calc_arrears_rate"].fillna(0.0)
    base["calc_prev_arrears_rate"] = base["previous_report_arrears_rate"].fillna(0.0)
    base["calc_arrears_rate_yoy"] = base["calc_arrears_rate"] - base["calc_prev_arrears_rate"]

    result = {
        "period": REPORT_MONTH,
        "report": "1._主营业务指标达成",
        "dimension": "project",
        "rules": {
            "current": "202512上一年回款率项目底表分子 + 202512往年回款率项目底表分子，再除以两张底表分母合计",
            "previous_same_period": "202412【1._主营业务指标达成】项目报表的历欠回款率累计完成值",
            "yoy": "当前复算累计完成值 - 去年同期累计完成值",
        },
        "source_files": sorted(
            {
                find_workbook("1._", MAIN_BUSINESS, REPORT_MONTH, PROJECT).name,
                find_workbook("1._", MAIN_BUSINESS, PREVIOUS_YEAR_MONTH, PROJECT).name,
                find_workbook(PREVIOUS_YEAR_RATE, REPORT_MONTH, PROJECT).name,
                find_workbook(PAST_YEAR_RATE, REPORT_MONTH, PROJECT).name,
            }
        ),
        "missing_sources": [
            "未找到202412上一年回款率/往年回款率项目底表；去年同期累计完成值按202412【1._主营业务指标达成】项目报表列校验。"
        ],
        "indicator_rows": load_indicator_rows(),
        "source_coverage": {
            "report_rows": int(len(base)),
            "previous_year_source_matched_rows": int(base["previous_year_source_rows"].notna().sum()),
            "past_year_source_matched_rows": int(base["past_year_source_rows"].notna().sum()),
            "any_current_source_matched_rows": int(
                (base["previous_year_source_rows"].notna() | base["past_year_source_rows"].notna()).sum()
            ),
            "previous_report_matched_rows": int(base["previous_report_file"].notna().sum()),
        },
        "summaries": [
            summarize(base, "历欠回款率累计完成值（%）", "arrears_rate", "calc_arrears_rate"),
            summarize(base, "去年同期累计完成值（%）", "prev_arrears_rate", "calc_prev_arrears_rate"),
            summarize(base, "同比增长率（%）", "arrears_rate_yoy", "calc_arrears_rate_yoy"),
        ],
        "samples": {
            "current": sample_mismatches(base, "arrears_rate", "calc_arrears_rate"),
            "previous_same_period": sample_mismatches(base, "prev_arrears_rate", "calc_prev_arrears_rate"),
            "yoy": sample_mismatches(base, "arrears_rate_yoy", "calc_arrears_rate_yoy"),
        },
    }

    output_dir = ROOT / "local_outputs"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"validate_main_business_project_arrears_rate_{REPORT_MONTH}.json"
    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nWrote {output_path}")


if __name__ == "__main__":
    main()
