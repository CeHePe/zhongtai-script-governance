from __future__ import annotations

import json
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
VALUE_TOLERANCE = 1e-5  # 万元；0.1 元以内视为展示/精度差异。
YOY_ZERO_DENOMINATOR = 1e-6  # 万元；上年绝对值过小按 0 处理。


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def as_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def load_112_project(month: str) -> pd.DataFrame:
    """读取 1.1.2 项目报表前三个半收付指标及同比列。"""
    project_token = u(r"\u9879\u76ee")
    df = pd.read_excel(find_workbook("1.1.2", month, project_token), header=None, dtype=object).iloc[3:].copy()
    df = df.rename(
        columns={
            0: "region",
            1: "line",
            2: "code",
            3: "name",
            4: "half_net_profit",
            5: "half_net_profit_yoy",
            6: "half_attr_profit",
            7: "half_attr_profit_yoy",
            8: "half_revenue",
            9: "half_revenue_yoy",
        }
    )
    df["code_norm"] = df["code"].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    metric_columns = [
        "half_net_profit",
        "half_net_profit_yoy",
        "half_attr_profit",
        "half_attr_profit_yoy",
        "half_revenue",
        "half_revenue_yoy",
    ]
    for column in metric_columns:
        df[column] = as_number(df[column])
    return df[["region", "line", "code", "code_norm", "name", *metric_columns]]


def load_152_source() -> pd.DataFrame:
    """1.5.2 半收付口径底表：收入、净利润源值单位为元，转万元。"""
    df = pd.read_excel(find_workbook("1.5.2"), header=None, dtype=object).iloc[5:].copy()
    df = df.rename(columns={1: "code", 12: "half_revenue_yuan", 29: "half_net_profit_yuan"})
    df["code_norm"] = df["code"].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["src_half_revenue"] = as_number(df["half_revenue_yuan"]) / 10000.0
    df["src_half_net_profit"] = as_number(df["half_net_profit_yuan"]) / 10000.0
    return df.groupby("code_norm", as_index=False).agg(
        src_half_revenue=("src_half_revenue", "sum"),
        src_half_net_profit=("src_half_net_profit", "sum"),
        src_152_rows=("code_norm", "size"),
    )


def load_half_attr_source() -> pd.DataFrame:
    """半收付归母净利润项目报表：归母净利润源值单位为元，转万元。"""
    half_attr = u(r"\u534a\u6536\u4ed8\u5f52\u6bcd\u51c0\u5229\u6da6")
    project = u(r"\u9879\u76ee")
    df = pd.read_excel(find_workbook(half_attr, "202512", project), dtype=object)
    # 该报表列顺序稳定：第 3 列为立项编码，第 35 列为半收付归母净利润。
    df["code_norm"] = df.iloc[:, 2].map(normalize_code)
    df = df[df["code_norm"].ne("")].copy()
    df["src_half_attr_profit"] = as_number(df.iloc[:, 34]) / 10000.0
    return df.groupby("code_norm", as_index=False).agg(
        src_half_attr_profit=("src_half_attr_profit", "sum"),
        src_attr_rows=("code_norm", "size"),
    )


def load_indicator_rows() -> list[dict[str, object]]:
    """留痕本次测试依赖的指标清单原始行。"""
    indicator = u(r"JKS_\u6570\u636e\u4e2d\u53f0\u4e8c\u671f_\u6307\u6807\u6e05\u5355")
    df = pd.read_excel(find_workbook(indicator), sheet_name=0, dtype=object)
    targets = {
        "半收付归母净利润_项目",
        "半收付收入_项目",
        "半收付净利润_不含计提费_项目",
    }
    records: list[dict[str, object]] = []
    for idx, row in df.iterrows():
        metric = row.iloc[3]
        if pd.isna(metric) or str(metric).strip() not in targets:
            continue
        if str(row.iloc[5]).strip() != "累计":
            continue
        records.append(
            {
                "excel_row": int(idx) + 2,
                "seq": row.iloc[0],
                "metric": str(row.iloc[3]).strip(),
                "org_dim": row.iloc[4],
                "period": row.iloc[5],
                "method": row.iloc[8],
                "source": None if pd.isna(row.iloc[10]) else row.iloc[10],
                "logic": row.iloc[12],
            }
        )
    return records


def calc_yoy(current: float, previous: float) -> float:
    if abs(previous) <= YOY_ZERO_DENOMINATOR:
        return 0.0
    return (current - previous) / abs(previous)


def summarize(base: pd.DataFrame, label: str, report_col: str, calc_col: str) -> dict[str, object]:
    diff = base[report_col] - base[calc_col]
    mismatch = diff.abs() > VALUE_TOLERANCE
    return {
        "check": label,
        "status": "passed" if not mismatch.any() else "failed",
        "rows": int(len(base)),
        "mismatch_rows": int(mismatch.sum()),
        "report_total": float(base[report_col].sum()),
        "calc_total": float(base[calc_col].sum()),
        "diff_total": float(diff.sum()),
        "max_abs_diff": float(diff.abs().max()),
    }


def main() -> None:
    current = load_112_project("202512")
    previous = load_112_project("202412")
    # 当前和上年均有一个同编码、项目名尾部空格不同的 0 值重复行；同比匹配时去重避免行扩张。
    previous_unique = (
        previous.drop_duplicates("code_norm", keep="first")
        .set_index("code_norm")[["half_net_profit", "half_attr_profit", "half_revenue"]]
        .rename(columns=lambda column: f"prev_{column}")
    )

    base = (
        current.merge(load_152_source(), on="code_norm", how="left")
        .merge(load_half_attr_source(), on="code_norm", how="left")
        .merge(previous_unique, on="code_norm", how="left")
    )

    missing_source = {
        "missing_152_codes": int(base["src_152_rows"].isna().sum()),
        "missing_152_nonzero_report_rows": int(
            (
                base["src_152_rows"].isna()
                & ((base["half_net_profit"].abs() > VALUE_TOLERANCE) | (base["half_revenue"].abs() > VALUE_TOLERANCE))
            ).sum()
        ),
        "missing_attr_codes": int(base["src_attr_rows"].isna().sum()),
        "missing_attr_nonzero_report_rows": int(
            (base["src_attr_rows"].isna() & (base["half_attr_profit"].abs() > VALUE_TOLERANCE)).sum()
        ),
        "missing_previous_codes": int(base["prev_half_net_profit"].isna().sum()),
        "missing_previous_nonzero_yoy_rows": int(
            (
                base["prev_half_net_profit"].isna()
                & (
                    (base["half_net_profit_yoy"].abs() > VALUE_TOLERANCE)
                    | (base["half_attr_profit_yoy"].abs() > VALUE_TOLERANCE)
                    | (base["half_revenue_yoy"].abs() > VALUE_TOLERANCE)
                )
            ).sum()
        ),
    }

    for column in [
        "src_half_revenue",
        "src_half_net_profit",
        "src_half_attr_profit",
        "prev_half_net_profit",
        "prev_half_attr_profit",
        "prev_half_revenue",
    ]:
        base[column] = base[column].fillna(0.0)

    base["calc_half_net_profit"] = base["src_half_net_profit"]
    base["calc_half_attr_profit"] = base["src_half_attr_profit"]
    base["calc_half_revenue"] = base["src_half_revenue"]
    for metric in ["half_net_profit", "half_attr_profit", "half_revenue"]:
        base[f"calc_{metric}_yoy"] = [
            calc_yoy(current_value, previous_value)
            for current_value, previous_value in zip(base[metric], base[f"prev_{metric}"])
        ]

    checks = [
        ("半收付净利润累计完成值（万元）", "half_net_profit", "calc_half_net_profit"),
        ("半收付归母净利润累计完成值（万元）", "half_attr_profit", "calc_half_attr_profit"),
        ("半收付收入累计完成值（万元）", "half_revenue", "calc_half_revenue"),
        ("半收付净利润同比增长率", "half_net_profit_yoy", "calc_half_net_profit_yoy"),
        ("半收付归母净利润同比增长率", "half_attr_profit_yoy", "calc_half_attr_profit_yoy"),
        ("半收付收入同比增长率", "half_revenue_yoy", "calc_half_revenue_yoy"),
    ]

    result = {
        "indicator_rows": load_indicator_rows(),
        "source_files": {
            "current_112": find_workbook("1.1.2", "202512", u(r"\u9879\u76ee")).name,
            "previous_112": find_workbook("1.1.2", "202412", u(r"\u9879\u76ee")).name,
            "source_152": find_workbook("1.5.2").name,
            "source_half_attr": find_workbook(
                u(r"\u534a\u6536\u4ed8\u5f52\u6bcd\u51c0\u5229\u6da6"), "202512", u(r"\u9879\u76ee")
            ).name,
        },
        "counts": {
            "current_rows": int(len(current)),
            "previous_rows": int(len(previous)),
            "current_duplicate_code_rows": int(current.duplicated("code_norm").sum()),
            "previous_duplicate_code_rows": int(previous.duplicated("code_norm").sum()),
            **missing_source,
        },
        "summary": [summarize(base, label, report_col, calc_col) for label, report_col, calc_col in checks],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
