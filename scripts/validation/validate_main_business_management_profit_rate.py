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


MAIN_BUSINESS = u(r"\u4e3b\u8425\u4e1a\u52a1\u6307\u6807\u8fbe\u6210")
PROJECT = u(r"\u9879\u76ee")
REGION_LINE = u(r"\u533a\u57df\u6761\u7ebf")


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


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def safe_rate(profit: float, income: float) -> float:
    if abs(income) < TOL:
        return 0.0
    return profit / income


def literal_rate(profit: float, income: float) -> float:
    if income == 0:
        return 0.0
    return profit / income


def load_dimension(dimension: str) -> pd.DataFrame:
    if dimension == "project":
        token = PROJECT
        expected_cols = 54
        identity = {1: "region", 2: "line", 3: "project_code", 4: "project_name"}
        income_col, profit_col, rate_col = 45, 48, 51
    elif dimension == "region_line":
        token = REGION_LINE
        expected_cols = 62
        identity = {1: "region", 2: "line", 3: "owner"}
        income_col, profit_col, rate_col = 51, 54, 57
    else:
        raise ValueError(dimension)

    path = find_workbook("1._", MAIN_BUSINESS, REPORT_MONTH, token)
    raw = pd.read_excel(path, header=None, dtype=object).iloc[2:].reset_index(drop=True).copy()
    if raw.shape[1] != expected_cols:
        raise RuntimeError(f"Unexpected column count in {path.name}: {raw.shape[1]}")

    out = pd.DataFrame({name: raw.iloc[:, col] for col, name in identity.items()})
    out["cumulative_income"] = as_number(raw.iloc[:, income_col])
    out["management_profit"] = as_number(raw.iloc[:, profit_col])
    out["report_profit_rate"] = as_number(raw.iloc[:, rate_col])
    out["calc_profit_rate"] = [
        safe_rate(profit, income)
        for profit, income in zip(out["management_profit"], out["cumulative_income"])
    ]
    out["literal_formula_rate"] = [
        literal_rate(profit, income)
        for profit, income in zip(out["management_profit"], out["cumulative_income"])
    ]
    out["diff"] = out["report_profit_rate"] - out["calc_profit_rate"]
    out["abs_diff"] = out["diff"].abs()
    out["literal_formula_diff"] = out["report_profit_rate"] - out["literal_formula_rate"]
    out["literal_formula_abs_diff"] = out["literal_formula_diff"].abs()
    out["source_file"] = path.name
    return out


def summarize(df: pd.DataFrame, dimension: str) -> dict[str, object]:
    mismatch = df["abs_diff"] > TOL
    zero_income = df["cumulative_income"].abs() < TOL
    return {
        "dimension": dimension,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(df)),
        "mismatch_rows": int(mismatch.sum()),
        "zero_income_rows": int(zero_income.sum()),
        "max_abs_diff": float(df["abs_diff"].max() if len(df) else 0.0),
        "diff_total": float(df["diff"].sum()),
    }


def summarize_literal_formula(df: pd.DataFrame, dimension: str) -> dict[str, object]:
    mismatch = df["literal_formula_abs_diff"] > TOL
    return {
        "dimension": dimension,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(df)),
        "mismatch_rows": int(mismatch.sum()),
        "max_abs_diff": float(df["literal_formula_abs_diff"].max() if len(df) else 0.0),
    }


def samples(df: pd.DataFrame, limit: int = 10) -> list[dict[str, object]]:
    bad = df[df["abs_diff"] > TOL].copy()
    return bad.head(limit).to_dict(orient="records")


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    project = load_dimension("project")
    region_line = load_dimension("region_line")

    result = {
        "period": REPORT_MONTH,
        "report": "1._主营业务指标达成",
        "metric": "管报净利润率（%）",
        "rule": "管报净利润累计完成值（万元） / 累计营业收入（万元）；报表存储为小数比例，展示时可乘100%",
        "zero_denominator_rule": f"累计营业收入绝对值小于 {TOL} 万元时，复算值取0",
        "source_files": [project["source_file"].iloc[0], region_line["source_file"].iloc[0]],
        "summaries": [
            summarize(project, "项目"),
            summarize(region_line, "区域条线"),
        ],
        "literal_formula_summaries": [
            summarize_literal_formula(project, "项目"),
            summarize_literal_formula(region_line, "区域条线"),
        ],
        "samples": {
            "项目": samples(project),
            "区域条线": samples(region_line),
        },
    }

    output_dir = ROOT / "local_outputs"
    output_dir.mkdir(exist_ok=True)
    json_path = output_dir / f"validate_main_business_management_profit_rate_{REPORT_MONTH}.json"
    project_path = output_dir / f"validate_main_business_management_profit_rate_{REPORT_MONTH}_project.csv"
    region_line_path = output_dir / f"validate_main_business_management_profit_rate_{REPORT_MONTH}_region_line.csv"
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    project.to_csv(project_path, index=False, encoding="utf-8-sig")
    region_line.to_csv(region_line_path, index=False, encoding="utf-8-sig")

    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nWrote {json_path}")
    print(f"Wrote {project_path}")
    print(f"Wrote {region_line_path}")


if __name__ == "__main__":
    main()
