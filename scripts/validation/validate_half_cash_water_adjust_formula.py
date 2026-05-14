from __future__ import annotations

import json
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
TOLERANCE = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


HALF_ATTRIBUTABLE = u(r"\u534a\u6536\u4ed8\u5f52\u6bcd\u51c0\u5229\u6da6")
PROJECT = u(r"\u9879\u76ee")
REGION = u(r"\u533a\u57df")
LINE = u(r"\u6761\u7ebf")
PROFESSIONAL_COMPANY = u(r"\u4e13\u4e1a\u516c\u53f8")
SPACE = u(r"\u7a7a\u95f4")

FORMULA_COLUMNS = {
    "water_backflow": u(r"\u4ee3\u6536\u6c34\u7535\u56de\u6b3e\u989d_\u534a\u6536\u4ed8"),
    "water_backflow_alloc": u(r"\u4ee3\u6536\u6c34\u7535\u56de\u6b3e\u989d_\u534a\u6536\u4ed8-\u636e\u5b9e\u5206\u644a"),
    "water_remaining_receivable": u(r"\u5f53\u671f\u4ee3\u6536\u6c34\u7535\u8d39\u5269\u4f59\u5e94\u6536_\u534a\u6536\u4ed8"),
    "water_current_receivable_alloc": u(r"\u5f53\u671f\u5e94\u6536\u6c34\u7535\u8d39_\u534a\u6536\u4ed8_\u636e\u5b9e\u5206\u644a"),
    "energy_income_tax": u(r"\u80fd\u8017\u6536\u5165_\u534a\u6536\u4ed8*0.06"),
    "single_current_unreceived": u(r"\u5f53\u5e74\u5355\u4e00\u4e1a\u6743\u672a\u5230\u8d26\u671f\u7684\u6c34\u7535\u8d39\u5e94\u6536"),
    "single_prev_unreceived": u(r"\u4e0a\u5e74\u5355\u4e00\u4e1a\u6743\u672a\u5230\u8d26\u671f\u7684\u6c34\u7535\u8d39\u5e94\u6536"),
    "water_adjust": u(r"\u6c34\u7535\u8d39\u8c03\u6574\u989d"),
}


def find_report(dimension_token: str) -> Path:
    # 用 Python 枚举中文文件名，避免终端编码影响路径匹配。
    prefix = f"{HALF_ATTRIBUTABLE}{REPORT_MONTH}"
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and prefix in path.name and dimension_token in path.name
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one report for {dimension_token}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def to_number(series: pd.Series) -> pd.Series:
    # 报表空值按 0 参与公式；非数字文本也按 0 处理。
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def valid_text_mask(series: pd.Series) -> pd.Series:
    text = series.astype(str).str.strip()
    return series.notna() & text.ne("") & ~text.isin(["0", "0.0", "nan", "None"])


def prepare_rows(df: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    # 各维度报表尾部都有空行或脏值，统一在公式校验前剔除。
    if key_columns:
        mask = valid_text_mask(df[key_columns[0]])
        return df[mask].copy()
    metric_mask = df[[FORMULA_COLUMNS["water_adjust"]]].apply(lambda col: to_number(col).abs() > TOLERANCE).any(axis=1)
    return df[metric_mask].copy()


def check_report(label: str, dimension_token: str, key_columns: list[str]) -> dict:
    # 这里只测指标清单里的公式一致性，不回推底层台账。
    df = pd.read_excel(find_report(dimension_token), dtype=object)
    df = prepare_rows(df, key_columns)
    missing = [column for column in FORMULA_COLUMNS.values() if column not in df.columns]
    if missing:
        return {"label": label, "status": "blocked", "missing_columns": missing}

    calc = (
        to_number(df[FORMULA_COLUMNS["water_backflow"]])
        + to_number(df[FORMULA_COLUMNS["water_backflow_alloc"]])
        - to_number(df[FORMULA_COLUMNS["water_remaining_receivable"]])
        - to_number(df[FORMULA_COLUMNS["water_current_receivable_alloc"]])
        - to_number(df[FORMULA_COLUMNS["energy_income_tax"]])
        + to_number(df[FORMULA_COLUMNS["single_current_unreceived"]])
        - to_number(df[FORMULA_COLUMNS["single_prev_unreceived"]])
    )
    report = to_number(df[FORMULA_COLUMNS["water_adjust"]])
    diff = calc - report

    detail_columns = key_columns if key_columns else []
    mismatch = df.loc[diff.abs() > TOLERANCE, detail_columns].copy()
    mismatch["water_adjust_report"] = report.loc[mismatch.index]
    mismatch["water_adjust_calc"] = calc.loc[mismatch.index]
    mismatch["diff"] = diff.loc[mismatch.index]
    if not key_columns:
        mismatch.insert(0, "row_index", mismatch.index)

    return {
        "label": label,
        "status": "passed" if mismatch.empty else "failed",
        "rows": int(len(df)),
        "mismatch_rows": int(len(mismatch)),
        "report_total": float(report.sum()),
        "calc_total": float(calc.sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "mismatch_details": mismatch.head(30).to_dict(orient="records"),
    }


def main() -> None:
    results = [
        check_report("project_water_adjust_formula", PROJECT, [u(r"\u533a\u57df"), u(r"\u6761\u7ebf"), u(r"\u9879\u76ee"), u(r"\u9879\u76ee.1")]),
        check_report("region_water_adjust_formula", REGION, [u(r"\u533a\u57df")]),
        check_report("line_water_adjust_formula", LINE, [u(r"\u6761\u7ebf")]),
        check_report("professional_company_water_adjust_formula", PROFESSIONAL_COMPANY, [u(r"\u4e13\u4e1a\u516c\u53f8")]),
        check_report("space_water_adjust_formula", SPACE, []),
    ]
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
