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
REPORT_NAME = u(r"\u533a\u57df\u53ca\u9879\u76ee\u56de\u6b3e\u7c7b\u578b\u5206\u6790")
OCCUPANCY = u(r"\u5165\u4f4f\u7387")
TOTAL_AMOUNT = u(r"\u56de\u6b3e\u603b\u91d1\u989d")
CASHFLOW = u(r"\u7d2f\u8ba1\u56de\u6536\u73b0\u91d1\u6d41")
LARGE_AR = "8.1"
REVENUE_RATIO = u(r"\u56de\u6b3e\u8425\u6536\u6bd4")
INDICATOR_LIST = "JKS_" + u(r"\u6570\u636e\u4e2d\u53f0\u4e8c\u671f") + "_" + u(r"\u6307\u6807\u6e05\u5355") + ".xlsx"


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
            if path.is_file() and path.suffix.lower() in {".xlsx", ".xls"} and all(token in path.name for token in tokens)
        ]
    if len(matches) != 1:
        target = exact_name or tokens
        raise RuntimeError(f"Expected one workbook for {target!r}, got {len(matches)}: {[p.name for p in matches]}")
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


def clean_value(value: object) -> object:
    if pd.isna(value):
        return None
    return value


def load_indicator_rows() -> list[dict[str, object]]:
    path = find_workbook(exact_name=INDICATOR_LIST)
    df = pd.read_excel(path, sheet_name=0, dtype=object)
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
        (u(r"\u5927\u989d\u6b20\u8d39\u5206\u6790"), OCCUPANCY, OCCUPANCY + "_" + PROJECT),
        (REVENUE_RATIO, CASHFLOW, CASHFLOW + "_" + PROJECT),
    ]

    rows: list[dict[str, object]] = []
    for relation, component, metric in wanted:
        target = df[
            df[relation_col].astype(str).str.strip().eq(relation)
            & df[component_col].astype(str).str.strip().eq(component)
            & df[metric_col].astype(str).str.strip().eq(metric)
            & df[dimension_col].astype(str).str.strip().eq(PROJECT)
            & df[period_col].astype(str).str.strip().eq(u(r"\u7d2f\u8ba1"))
        ].copy()
        if len(target) != 1:
            raise RuntimeError(f"Expected one indicator row for {(relation, component, metric)}, got {len(target)}")
        idx = target.index[0]
        row = target.iloc[0]
        rows.append(
            {
                "excel_row": int(idx + 2),
                "serial": clean_value(row.get(serial_col)),
                "relation": clean_value(row.get(relation_col)),
                "component": clean_value(row.get(component_col)),
                "metric": clean_value(row.get(metric_col)),
                "dimension": clean_value(row.get(dimension_col)),
                "period": clean_value(row.get(period_col)),
                "source_method": clean_value(row.get(method_col)),
                "source_table": clean_value(row.get(source_col)),
                "logic": clean_value(row.get(logic_col)),
            }
        )
    return rows


def load_target(month: str) -> pd.DataFrame:
    path = find_workbook(REPORT_NAME, month, PROJECT)
    raw = pd.read_excel(path, header=None, dtype=object)
    data = raw.iloc[2:].reset_index(drop=True).copy()
    data = data[data[2].astype(str).str.strip().ne("")].copy()
    out = pd.DataFrame(
        {
            "region": data[0],
            "line": data[1],
            "project_code": data[2].astype(str).str.strip(),
            "project_name": data[3].astype(str).str.strip(),
            "occupancy": as_number(data[7]),
            "occupancy_yoy": as_number(data[9]),
            "total_amount": as_number(data[10]),
            "total_amount_yoy": as_number(data[12]),
        }
    )
    out["code_norm"] = out["project_code"].map(normalize_code)
    out["report_file"] = path.name
    return out[out["code_norm"].ne("")].copy()


def load_occupancy_source() -> pd.DataFrame:
    path = find_workbook(LARGE_AR, REPORT_MONTH, PROJECT)
    raw = pd.read_excel(path, header=None, dtype=object)
    data = raw.iloc[2:].reset_index(drop=True).copy()
    data = data[data[3].astype(str).str.strip().ne("")].copy()
    out = pd.DataFrame(
        {
            "project_code": data[3].astype(str).str.strip(),
            "source_occupancy": as_number(data[9]),
        }
    )
    out["code_norm"] = out["project_code"].map(normalize_code)
    return out.groupby("code_norm", as_index=False).agg(
        source_occupancy=("source_occupancy", "sum"),
        occupancy_source_rows=("code_norm", "size"),
    )


def load_amount_source() -> pd.DataFrame:
    path = find_workbook(REVENUE_RATIO, REPORT_MONTH, PROJECT)
    raw = pd.read_excel(path, header=None, dtype=object)
    data = raw.iloc[3:].reset_index(drop=True).copy()
    data = data[data[3].astype(str).str.strip().ne("")].copy()
    out = pd.DataFrame(
        {
            "project_code": data[3].astype(str).str.strip(),
            "source_total_amount": as_number(data[5]) / 10000,
        }
    )
    out["code_norm"] = out["project_code"].map(normalize_code)
    return out.groupby("code_norm", as_index=False).agg(
        source_total_amount=("source_total_amount", "sum"),
        amount_source_rows=("code_norm", "size"),
    )


def add_yoy(current: pd.DataFrame, previous: pd.DataFrame) -> pd.DataFrame:
    prev = previous[["code_norm", "occupancy", "total_amount"]].rename(
        columns={"occupancy": "prev_occupancy", "total_amount": "prev_total_amount"}
    )
    out = current.merge(prev, on="code_norm", how="left")
    out["prev_occupancy"] = out["prev_occupancy"].fillna(0.0)
    out["prev_total_amount"] = out["prev_total_amount"].fillna(0.0)
    out["calc_occupancy_yoy"] = out["occupancy"] - out["prev_occupancy"]
    out["calc_total_amount_yoy"] = 0.0
    mask = out["prev_total_amount"].abs().gt(1.0)
    out.loc[mask, "calc_total_amount_yoy"] = (
        (out.loc[mask, "total_amount"] - out.loc[mask, "prev_total_amount"]) / out.loc[mask, "prev_total_amount"] * 100
    )
    return out


def summarize(df: pd.DataFrame, label: str, actual: str, calc: str, diff: str, missing_col: str | None = None) -> dict[str, object]:
    missing = pd.Series(False, index=df.index) if missing_col is None else df[missing_col].isna()
    mismatch = df[diff].abs().gt(TOL)
    return {
        "check": label,
        "status": "passed" if not mismatch.any() and not missing.any() else "failed",
        "rows": int(len(df)),
        "missing_rows": int(missing.sum()),
        "mismatch_rows": int(mismatch.sum()),
        "actual_total": float(df[actual].sum()),
        "calc_total": float(df.loc[~missing, calc].sum()),
        "diff_total": float(df.loc[~missing, diff].sum()),
        "max_abs_diff": float(df.loc[~missing, diff].abs().max() if (~missing).any() else 0.0),
    }


def sample_rows(df: pd.DataFrame, diff_col: str, cols: list[str], missing_col: str | None = None, limit: int = 20) -> list[dict[str, object]]:
    missing = pd.Series(False, index=df.index) if missing_col is None else df[missing_col].isna()
    subset = df[df[diff_col].abs().gt(TOL)].copy()
    if subset.empty:
        return []
    subset["_abs"] = subset[diff_col].abs()
    subset = subset.sort_values(["_abs", "project_code"], ascending=[False, True]).head(limit)
    return subset[cols].to_dict(orient="records")


def main() -> None:
    configure_stdout()
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    indicator_rows = load_indicator_rows()
    current = load_target(REPORT_MONTH)
    previous = load_target(PREVIOUS_YEAR_MONTH)
    occupancy_source = load_occupancy_source()
    amount_source = load_amount_source()

    result = current.merge(occupancy_source, on="code_norm", how="left").merge(amount_source, on="code_norm", how="left")
    result["occupancy_source_missing"] = result["source_occupancy"].isna()
    result["amount_source_missing"] = result["source_total_amount"].isna()
    result["source_occupancy"] = result["source_occupancy"].fillna(0.0)
    result["source_total_amount"] = result["source_total_amount"].fillna(0.0)
    result = add_yoy(result, previous)
    result["occupancy_diff"] = result["occupancy"] - result["source_occupancy"]
    result["total_amount_diff"] = result["total_amount"] - result["source_total_amount"]
    result["occupancy_yoy_diff"] = result["occupancy_yoy"] - result["calc_occupancy_yoy"]
    result["total_amount_yoy_diff"] = result["total_amount_yoy"] - result["calc_total_amount_yoy"]

    summaries = {
        "occupancy_from_8_1": summarize(result, "occupancy_from_8_1", "occupancy", "source_occupancy", "occupancy_diff", "source_occupancy"),
        "total_amount_from_revenue_ratio": summarize(
            result, "total_amount_from_revenue_ratio", "total_amount", "source_total_amount", "total_amount_diff", "source_total_amount"
        ),
        "occupancy_yoy_from_202412_report": summarize(
            result, "occupancy_yoy_from_202412_report", "occupancy_yoy", "calc_occupancy_yoy", "occupancy_yoy_diff"
        ),
        "total_amount_yoy_from_202412_report": summarize(
            result, "total_amount_yoy_from_202412_report", "total_amount_yoy", "calc_total_amount_yoy", "total_amount_yoy_diff"
        ),
    }

    output = {
        "status": "passed" if all(item["status"] == "passed" for item in summaries.values()) else "failed",
        "scope": {
            "report": REPORT_NAME,
            "month": REPORT_MONTH,
            "dimension": PROJECT,
            "skipped": [u(r"\u73af\u6bd4\u6307\u6807")],
            "yoy_base": f"{PREVIOUS_YEAR_MONTH} {REPORT_NAME}{PROJECT}",
            "row_count": int(len(result)),
        },
        "indicator_rows": indicator_rows,
        "source_files": {
            "target_current": current["report_file"].iloc[0] if not current.empty else None,
            "target_previous_year": previous["report_file"].iloc[0] if not previous.empty else None,
            "occupancy_source": find_workbook(LARGE_AR, REPORT_MONTH, PROJECT).name,
            "amount_source": find_workbook(REVENUE_RATIO, REPORT_MONTH, PROJECT).name,
        },
        "rule_confirmations": {
            "monthly_vs_cumulative": u(r"\u4e24\u4e2a\u6307\u6807\u5747\u6309\u6307\u6807\u6e05\u5355\u7684\u7d2f\u8ba1\u53e3\u5f84\uff1b\u73af\u6bd4\u6307\u6807\u672c\u6b21\u4e0d\u6d4b\u3002"),
            "d_exit_exclusion": u(r"\u6307\u6807\u6e05\u5355\u672a\u5199\u660e\u6392\u9664 D \u7c7b\u5df2\u64a4\u573a\uff1b\u9879\u76ee\u7ea7\u76f4\u63a5\u6309\u5e95\u8868\u9879\u76ee\u7f16\u53f7\u5bf9\u9f50\u3002"),
            "non_assessment_exclusion": u(r"\u6307\u6807\u6e05\u5355\u672a\u5199\u660e\uff08\u5229\u6da6\uff09\u975e\u8003\u6838\u9879\u76ee\u6392\u9664\u3002"),
            "high_dimension_adjustment": u(r"\u672c\u6b21\u4ec5\u9879\u76ee\u7ea7\uff0c\u4e0d\u6d89\u53ca\u9ad8\u7ef4\u9644\u52a0\u9879\u6216\u9879\u76ee\u6c47\u603b\u3002"),
        },
        "summaries": summaries,
        "samples": {
            "occupancy_from_8_1": sample_rows(
                result,
                "occupancy_diff",
                ["project_code", "project_name", "occupancy", "source_occupancy", "occupancy_diff"],
                "source_occupancy",
            ),
            "total_amount_from_revenue_ratio": sample_rows(
                result,
                "total_amount_diff",
                ["project_code", "project_name", "total_amount", "source_total_amount", "total_amount_diff"],
                "source_total_amount",
            ),
            "occupancy_yoy_from_202412_report": sample_rows(
                result,
                "occupancy_yoy_diff",
                ["project_code", "project_name", "occupancy_yoy", "calc_occupancy_yoy", "occupancy_yoy_diff"],
            ),
            "total_amount_yoy_from_202412_report": sample_rows(
                result,
                "total_amount_yoy_diff",
                ["project_code", "project_name", "total_amount_yoy", "calc_total_amount_yoy", "total_amount_yoy_diff"],
            ),
        },
    }
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
