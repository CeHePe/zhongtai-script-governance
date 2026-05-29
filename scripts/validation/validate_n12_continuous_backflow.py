from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from _project_root import find_project_root
from typing import Iterable

import pandas as pd


ROOT = find_project_root(__file__)
TOL = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


PROJECT = u(r"\u9879\u76ee")
REGION_HOUSING = u(r"\u533a\u57df\u4f4f\u5b85")
REGION_LINE = u(r"\u533a\u57df\u6761\u7ebf")
HOUSING = u(r"\u4f4f\u5b85")
NON_RELATED = u(r"\u975e\u5173\u8054\u65b9")
NON_ASSESS = u(r"\u975e\u8003\u6838\u9879\u76ee")
PREV_YEAR_RATE = u(r"\u4e0a\u4e00\u5e74\u56de\u6b3e\u7387")
RECEIVABLE_RATIO = u(r"\u56de\u6b3e\u8425\u6536\u6bd4")
STATUS_EXIT = u(r"\u5df2\u64a4\u573a")
STATUS_NOT_IN = u(r"\u672a\u8fdb\u573a")
MULTI_OWNERSHIP = u(r"\u591a\u4e1a\u6743")
ASSESS_PROJECT = u(r"\u8003\u6838\u9879\u76ee")
HOUSING_SERVICE = u(r"\u4f4f\u5b85\u670d\u52a1")


REPORT_COLS = [
    "idx",
    "region",
    "line",
    "project_code",
    "project_name",
    "owner_rcv",
    "big_owner_rev",
    "big_owner_ar",
    "gold_repayment_prev",
    "numerator",
    "revenue",
    "non_assess_revenue",
    "related_revenue",
    "non_assess_related_revenue",
    "gold_balance",
    "cutoff_income",
    "n12_not_due",
    "denominator",
    "rate",
]

REGION_COLS = [
    "idx",
    "region",
    "owner_rcv",
    "big_owner_rev",
    "big_owner_ar",
    "numerator",
    "revenue",
    "non_assess_revenue",
    "related_revenue",
    "non_assess_related_revenue",
    "offset_contract",
    "gold_balance",
    "cutoff_income",
    "n12_not_due",
    "denominator",
    "rate",
]

REGION_LINE_COLS = [
    "idx",
    "region",
    "line",
    "owner_rcv",
    "big_owner_rev",
    "big_owner_ar",
    "gold_repayment_prev",
    "numerator",
    "revenue",
    "non_assess_revenue",
    "related_revenue",
    "non_assess_related_revenue",
    "gold_balance",
    "cutoff_income",
    "n12_not_due",
    "denominator",
    "rate",
]

HOUSING_COLS = [
    "idx",
    "owner_rcv",
    "big_owner_rev",
    "big_owner_ar",
    "numerator",
    "revenue",
    "non_assess_revenue",
    "related_revenue",
    "non_assess_related_revenue",
    "offset_contract",
    "gold_balance",
    "cutoff_income",
    "n12_not_due",
    "denominator",
    "rate",
]

NUMERIC = [
    "owner_rcv",
    "big_owner_rev",
    "big_owner_ar",
    "gold_repayment_prev",
    "numerator",
    "revenue",
    "non_assess_revenue",
    "related_revenue",
    "non_assess_related_revenue",
    "gold_balance",
    "cutoff_income",
    "n12_not_due",
    "denominator",
    "rate",
]

SOURCE_COMPONENTS = []


@dataclass
class ComponentResult:
    name: str
    report_total: float
    source_total: float
    diff_total: float
    mismatch_count: int
    top_diffs: pd.DataFrame


def exact_code(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text in {"", "NAN", "NONE"}:
        return ""
    return text


def norm_code(value: object) -> str:
    text = exact_code(value)
    if len(text) > 1 and text[0].isalpha() and any(ch.isdigit() for ch in text[1:]):
        return text[1:]
    return text


def find_report(mark: str, excludes: Iterable[str] = ()) -> Path:
    matches = []
    for path in ROOT.glob("N+12*.xlsx"):
        name = path.name
        if mark in name and not any(ex in name for ex in excludes):
            matches.append(path)
    if len(matches) != 1:
        raise RuntimeError(f"Expected one N+12 report for {mark!r}, got {len(matches)}")
    return matches[0]


def find_optional_report(mark: str, excludes: Iterable[str] = ()) -> Path | None:
    try:
        return find_report(mark, excludes=excludes)
    except RuntimeError:
        return None


def find_workbook(*marks: str) -> Path:
    matches = [p for p in ROOT.glob("*.xlsx") if all(mark in p.name for mark in marks)]
    if len(matches) != 1:
        names = [p.name for p in matches]
        raise RuntimeError(f"Expected one workbook for {marks!r}, got {len(matches)}: {names}")
    return matches[0]


def normalize_numeric(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    out = df.copy()
    for col in cols:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out


def load_project_report() -> pd.DataFrame:
    path = find_report(PROJECT)
    df = pd.read_excel(path, header=None).iloc[3:].reset_index(drop=True)
    df.columns = REPORT_COLS
    df = normalize_numeric(df, NUMERIC)
    df["code_exact"] = df["project_code"].map(exact_code)
    df["code_norm"] = df["project_code"].map(norm_code)
    return df


def load_region_housing_report() -> pd.DataFrame:
    path = find_report(REGION_HOUSING)
    df = pd.read_excel(path, header=None).iloc[3:].reset_index(drop=True)
    df.columns = REGION_COLS
    return normalize_numeric(df, NUMERIC)


def load_region_line_report() -> pd.DataFrame:
    path = find_report(REGION_LINE)
    df = pd.read_excel(path, header=None).iloc[3:].reset_index(drop=True)
    df.columns = REGION_LINE_COLS
    return normalize_numeric(df, NUMERIC)


def load_housing_report() -> pd.DataFrame:
    path = find_report(HOUSING, excludes=[REGION_HOUSING, PROJECT])
    df = pd.read_excel(path, header=None).iloc[3:].reset_index(drop=True)
    df.columns = HOUSING_COLS
    return normalize_numeric(df, NUMERIC)


def source_lookup(source: pd.DataFrame, value_col: str) -> tuple[dict[str, float], dict[str, float]]:
    exact = source.groupby("code_exact", as_index=False)[value_col].sum()
    norm = source.groupby("code_norm", as_index=False).agg(
        value=(value_col, "sum"), exact_count=("code_exact", "nunique")
    )
    exact_map = dict(zip(exact["code_exact"], exact[value_col]))
    norm_map = dict(zip(norm.loc[norm["exact_count"] == 1, "code_norm"], norm.loc[norm["exact_count"] == 1, "value"]))
    return exact_map, norm_map


def apply_source(report_codes: pd.DataFrame, source: pd.DataFrame, value_col: str) -> pd.Series:
    exact_map, norm_map = source_lookup(source, value_col)
    values = []
    for _, row in report_codes.iterrows():
        exact = row["code_exact"]
        norm = row["code_norm"]
        if exact in exact_map:
            values.append(float(exact_map[exact]))
        elif norm in norm_map:
            values.append(float(norm_map[norm]))
        else:
            values.append(0.0)
    return pd.Series(values, index=report_codes.index)


def load_non_assess_codes() -> set[str]:
    path = next(p for p in ROOT.glob("*.xlsx") if p.name == u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26.xlsx"))
    df = pd.read_excel(path)
    code_col = next(col for col in df.columns if u(r"\u7acb\u9879\u7f16\u7801") in str(col))
    return {exact_code(v) for v in df[code_col].dropna()}


def load_ownership_map() -> pd.DataFrame:
    """用现有项目级辅助表补齐业权属性；项目查询当前版本不是项目辅助台账，不能用于本指标。"""
    frames: list[pd.DataFrame] = []

    half_profit = find_workbook(u(r"\u534a\u6536\u4ed8\u5f52\u6bcd\u51c0\u5229\u6da6"), "202512", PROJECT)
    half_df = pd.read_excel(half_profit, header=None).iloc[1:].reset_index(drop=True)
    half_df = half_df.iloc[:, [2, 35]].copy()
    half_df.columns = ["project_code", "ownership"]
    frames.append(half_df)

    recv = find_workbook(u(r"\u5e94\u6536\u8d26\u9f84\u53ca\u672a\u5230\u8d26\u671f\u91d1\u989d\u5e74\u5ea6\u5206\u5e03"), "202512")
    recv_df = pd.read_excel(recv, header=None).iloc[2:].reset_index(drop=True)
    recv_df = recv_df.iloc[:, [4, 7]].copy()
    recv_df.columns = ["project_code", "ownership"]
    frames.append(recv_df)

    combined = pd.concat(frames, ignore_index=True)
    combined["project_code"] = combined["project_code"].map(exact_code)
    combined["ownership"] = combined["ownership"].fillna("").astype(str).str.strip()
    combined = combined[(combined["project_code"] != "") & (combined["ownership"] != "")]
    combined = combined.drop_duplicates(["project_code", "ownership"])
    combined = combined.groupby("project_code", as_index=False).first()
    combined["is_multi"] = combined["ownership"] == MULTI_OWNERSHIP
    return combined


def load_project_query_flags() -> pd.DataFrame:
    path = next(p for p in ROOT.glob("*.xlsx") if p.name == u(r"\u9879\u76ee\u67e5\u8be2.xlsx"))
    df = pd.read_excel(path).iloc[:, [0, 5, 11]].copy()
    df.columns = ["project_code", "project_level", "project_status_query"]
    df["code_exact"] = df["project_code"].map(exact_code)
    return df[["code_exact", "project_level", "project_status_query"]]


def load_project_query_status() -> pd.DataFrame:
    path = next(p for p in ROOT.glob("*.xlsx") if p.name == u(r"\u9879\u76ee\u67e5\u8be2.xlsx"))
    df = pd.read_excel(path).iloc[:, [0, 11]].copy()
    df.columns = ["project_code", "project_status_query"]
    df["project_code"] = df["project_code"].astype(str)
    return df


def load_operating_202412() -> pd.DataFrame:
    path = find_workbook("1.5.1", "202412")
    raw = pd.read_excel(path, header=None).iloc[5:].reset_index(drop=True)
    df = pd.DataFrame(
        {
            "project_code": raw[1],
            "revenue": pd.to_numeric(raw[11], errors="coerce").fillna(0.0),
            "related_revenue": pd.to_numeric(raw[21], errors="coerce").fillna(0.0)
            + pd.to_numeric(raw[22], errors="coerce").fillna(0.0),
        }
    )
    df["code_exact"] = df["project_code"].map(exact_code)
    df["code_norm"] = df["project_code"].map(norm_code)
    return df[df["code_exact"] != ""]


def load_receivable_aging_202512() -> pd.DataFrame:
    path = find_workbook(u(r"\u5e94\u6536\u8d26\u9f84"), "202512")
    raw = pd.read_excel(path, header=None).iloc[2:].reset_index(drop=True)
    raw.columns = [
        "idx",
        "data_month",
        "type_name",
        "region",
        "project_code",
        "project_name",
        "customer_attr",
        "ownership_attr",
        "big_owner_ar_total",
        "age_current",
        "age_prev_year",
        "age_prev2",
        "age_prev3",
        "age_prev4",
        "not_due_total",
        "not_due_current",
        "not_due_prev_year",
        "not_due_prev2",
        "not_due_prev3",
        "not_due_prev4",
        "ownership_attr_dup",
        "assessment_flag",
    ]
    df = raw[
        (raw["data_month"].astype(str) == "2025-12")
        & (raw["type_name"].astype(str) == PROJECT)
        & (raw["customer_attr"].astype(str) == NON_RELATED)
    ].copy()
    df["code_exact"] = df["project_code"].map(exact_code)
    df["code_norm"] = df["project_code"].map(norm_code)
    df["big_owner_ar"] = pd.to_numeric(df["age_prev_year"], errors="coerce").fillna(0.0)
    df["n12_not_due"] = pd.to_numeric(df["not_due_prev_year"], errors="coerce").fillna(0.0)
    return df


def load_cutoff_202412() -> pd.DataFrame:
    path = find_workbook(u(r"\u622a\u6b62\u6027\u6536\u652f\u8c03\u6574\u53f0\u8d26"), "202412")
    raw = pd.read_excel(path, header=None).iloc[2:].reset_index(drop=True)
    raw.columns = [
        "idx",
        "type_name",
        "data_month",
        "project_name",
        "project_code",
        "region",
        "new_income",
        "prior_income",
        "new_cost",
    ]
    df = raw[(raw["type_name"].astype(str) == PROJECT) & (raw["data_month"].astype(str) == "2024-12")].copy()
    df["code_exact"] = df["project_code"].map(exact_code)
    df["code_norm"] = df["project_code"].map(norm_code)
    df["cutoff_income"] = pd.to_numeric(df["prior_income"], errors="coerce").fillna(0.0)
    return df


def load_gold_2024() -> pd.DataFrame:
    path = find_workbook(u(r"\u91d1\u5e01\u4f59\u989d\u53f0\u8d26"), "2024")
    raw = pd.read_excel(path, header=None).iloc[1:].reset_index(drop=True)
    raw.columns = ["idx", "year", "project_name", "project_code", "region", "team", "annual_gold"]
    df = raw[raw["year"].astype(str) == "2024"].copy()
    df["code_exact"] = df["project_code"].map(exact_code)
    df["code_norm"] = df["project_code"].map(norm_code)
    # 上一年金币余额取 2024 年12月累计当期金币，即年度应赠送金额全年累计。
    df["gold_balance"] = pd.to_numeric(df["annual_gold"], errors="coerce").fillna(0.0)
    return df


def load_offset_202512() -> pd.DataFrame:
    path = find_workbook(u(r"\u62b5\u623f\u53f0\u8d26"))
    raw = pd.read_excel(path, header=None).iloc[2:].reset_index(drop=True)
    raw.columns = [
        "idx",
        "data_month",
        "offset_month",
        "project_name",
        "project_code",
        "region",
        "service_team",
        "offset_type",
        "sell_status",
        "contract_amount",
        "assessment_ratio",
        "cash_amount",
        "cash_month",
        "note",
    ]
    df = raw[
        (raw["data_month"].astype(str) == "2025-12")
        & (raw["offset_month"].astype(str).str.startswith("2024"))
        & (raw["cash_month"].astype(str).between("2024-01", "2025-12"))
        & (raw["sell_status"].astype(str) == u(r"\u5df2\u53bb\u5316"))
    ].copy()
    df["code_exact"] = df["project_code"].map(exact_code)
    df["code_norm"] = df["project_code"].map(norm_code)
    contract = pd.to_numeric(df["contract_amount"], errors="coerce").fillna(0.0)
    cash = pd.to_numeric(df["cash_amount"], errors="coerce").fillna(0.0)
    df["offset_contract"] = contract - cash
    return df


def load_prev_year_rate_project_gold() -> pd.DataFrame:
    path = find_workbook(PREV_YEAR_RATE, "202512", PROJECT)
    raw = pd.read_excel(path, header=None).iloc[2:].reset_index(drop=True)
    raw.columns = [
        "idx",
        "region",
        "line",
        "project_code",
        "project_name",
        "col5",
        "col6",
        "col7",
        "col8",
        "col9",
        "col10",
        "col11",
        "col12",
        "col13",
        "col14",
        "gold_repayment_prev",
        "prev_numerator",
        "prev_ar",
        "col18",
        "col19",
        "col20",
        "gold_balance",
        "col22",
        "col23",
        "prev_denominator",
        "prev_rate",
    ]
    raw["gold_balance"] = pd.to_numeric(raw["gold_balance"], errors="coerce").fillna(0.0)
    return raw[["region", "line", "project_code", "gold_balance"]]


def load_current_gold_project() -> pd.DataFrame:
    path = find_workbook(RECEIVABLE_RATIO, "202512", PROJECT)
    raw = pd.read_excel(path, header=None).iloc[3:].reset_index(drop=True)
    raw.columns = [
        "idx",
        "region",
        "line",
        "project_code",
        "project_name",
        "num1",
        "num2",
        "num3",
        "num4",
        "num5",
        "num6",
        "num7",
        "num8",
        "num9",
        "num10",
        "num11",
        "num12",
        "revenue",
        "non_assess_revenue",
        "related_revenue",
        "non_assess_related_revenue",
        "discount1",
        "discount2",
        "discount3",
        "discount4",
        "rec1",
        "rec2",
        "rec3",
        "not_due",
        "prev_not_due",
        "current_gold",
        "ratio",
        "ratio_num",
        "ratio_den",
        "project_status",
    ]
    raw["current_gold"] = pd.to_numeric(raw["current_gold"], errors="coerce").fillna(0.0)
    raw["code_exact"] = raw["project_code"].map(exact_code)
    return raw[["region", "line", "project_code", "code_exact", "project_status", "current_gold"]]


def build_source(project_report: pd.DataFrame) -> pd.DataFrame:
    codes = project_report[["code_exact", "code_norm"]].drop_duplicates().reset_index(drop=True)

    op = load_operating_202412()
    aging = load_receivable_aging_202512()
    cutoff = load_cutoff_202412()
    gold = load_gold_2024()
    offset = load_offset_202512()

    source = codes.copy()
    source["revenue"] = apply_source(codes, op[["code_exact", "code_norm", "revenue"]], "revenue")
    source["related_revenue"] = apply_source(codes, op[["code_exact", "code_norm", "related_revenue"]], "related_revenue")
    source["big_owner_ar"] = apply_source(codes, aging[["code_exact", "code_norm", "big_owner_ar"]], "big_owner_ar")
    source["n12_not_due"] = apply_source(codes, aging[["code_exact", "code_norm", "n12_not_due"]], "n12_not_due")
    source["cutoff_income"] = apply_source(codes, cutoff[["code_exact", "code_norm", "cutoff_income"]], "cutoff_income")
    source["gold_balance"] = apply_source(codes, gold[["code_exact", "code_norm", "gold_balance"]], "gold_balance")
    source["offset_contract"] = apply_source(codes, offset[["code_exact", "code_norm", "offset_contract"]], "offset_contract")

    # 用户确认项目级这两项没有值，项目维度测试跳过，分母按报表项目级口径取 0。
    source["non_assess_revenue"] = 0.0
    source["non_assess_related_revenue"] = 0.0
    source["denominator"] = (
        source["revenue"]
        - source["non_assess_revenue"]
        - source["related_revenue"]
        + source["non_assess_related_revenue"]
        - source["offset_contract"]
        - source["gold_balance"]
        - source["cutoff_income"]
        - source["n12_not_due"]
    )
    return source


def compare_component(report: pd.DataFrame, source: pd.DataFrame, component: str) -> ComponentResult:
    report_group = report.groupby("code_exact", as_index=False).agg(
        project_name=("project_name", "first"),
        region=("region", "first"),
        report_value=(component, "sum"),
    )
    source_group = source.groupby("code_exact", as_index=False).agg(source_value=(component, "sum"))
    merged = report_group.merge(source_group, on="code_exact", how="left").fillna({"source_value": 0.0})
    merged["diff"] = merged["report_value"] - merged["source_value"]
    top = merged[merged["diff"].abs() > TOL].copy()
    top = top.sort_values("diff", key=lambda s: s.abs(), ascending=False).head(10)
    return ComponentResult(
        name=component,
        report_total=float(merged["report_value"].sum()),
        source_total=float(merged["source_value"].sum()),
        diff_total=float(merged["diff"].sum()),
        mismatch_count=int((merged["diff"].abs() > TOL).sum()),
        top_diffs=top[["code_exact", "project_name", "region", "report_value", "source_value", "diff"]],
    )


def internal_formula_check(label: str, df: pd.DataFrame) -> None:
    num_diff = (
        df["owner_rcv"] + df["big_owner_rev"] - df["big_owner_ar"] - df["gold_repayment_prev"] - df["numerator"]
    ).abs().max()
    den_diff = (
        df["revenue"]
        - df["non_assess_revenue"]
        - df["related_revenue"]
        + df["non_assess_related_revenue"]
        - df["gold_balance"]
        - df["cutoff_income"]
        - df["n12_not_due"]
        - df["denominator"]
    ).abs().max()
    valid = df["denominator"] != 0
    rate = pd.Series(0.0, index=df.index)
    rate.loc[valid] = df.loc[valid, "numerator"] / df.loc[valid, "denominator"]
    rate_diff = (rate - df["rate"]).abs().max()
    print(label)
    print(f"  numerator_formula_max_diff: {num_diff:.10f}")
    print(f"  denominator_formula_max_diff: {den_diff:.10f}")
    print(f"  rate_formula_max_diff: {rate_diff:.10f}")


def compare_rollup(project: pd.DataFrame, region: pd.DataFrame, housing: pd.DataFrame) -> None:
    roll_cols = [c for c in NUMERIC if c != "rate"]
    proj_region = project.groupby("region", as_index=False)[roll_cols].sum()
    proj_region["rate"] = proj_region["numerator"] / proj_region["denominator"]
    merged = proj_region.merge(region, on="region", suffixes=("_project", "_report"))
    print("\n[project_to_region_housing]")
    for _, row in merged.iterrows():
        diffs = {}
        for col in NUMERIC:
            diff = float(row[f"{col}_project"] - row[f"{col}_report"])
            if abs(diff) > TOL:
                diffs[col] = diff
        if diffs:
            print(f"{row['region']}: mismatch")
            for col, diff in diffs.items():
                print(f"  {col}: {diff:.10f}")
        else:
            print(f"{row['region']}: match")

    totals = project[roll_cols].sum()
    totals["rate"] = totals["numerator"] / totals["denominator"]
    print("\n[project_to_housing]")
    mismatch = False
    for col in NUMERIC:
        diff = float(totals[col] - housing.loc[0, col])
        if abs(diff) > TOL:
            mismatch = True
            print(f"  {col}: {diff:.10f}")
    if not mismatch:
        print("  match")


def compare_region_line_all_metrics() -> None:
    report_path = find_optional_report(REGION_LINE)
    if report_path is None:
        print("\n[region_line_all_metrics]")
        print("区域条线报表缺失，跳过")
        return

    project = load_project_report().copy()
    status = load_project_query_status()
    project["project_code"] = project["project_code"].astype(str)
    project = project.merge(status, on="project_code", how="left")
    report = load_region_line_report().copy()

    active_mask = ~project["project_status_query"].astype(str).isin([STATUS_EXIT, STATUS_NOT_IN])
    roll_cols = [c for c in NUMERIC if c != "rate"]
    full = project.groupby(["region", "line"], as_index=False)[roll_cols].sum()
    full["rate"] = full["numerator"] / full["denominator"]
    active = project[active_mask].groupby(["region", "line"], as_index=False)[roll_cols].sum()
    active["rate"] = active["numerator"] / active["denominator"]

    merged = full.merge(active, on=["region", "line"], suffixes=("_full", "_active"))
    merged = merged.merge(report[["region", "line"] + NUMERIC], on=["region", "line"], how="left")

    summary_rows: list[dict[str, object]] = []
    for metric in NUMERIC:
        report_total = float(report[metric].sum())
        full_total = float(full[metric].sum())
        active_total = float(active[metric].sum())
        diff_full = full_total - report_total
        diff_active = active_total - report_total
        if abs(diff_full) <= TOL:
            rule = "项目直接汇总"
            status_text = "PASS"
        elif abs(diff_active) <= TOL:
            rule = "项目汇总后排除已撤场/未进场"
            status_text = "PASS"
        else:
            rule = "未对平"
            status_text = "FAIL"
        summary_rows.append(
            {
                "metric": metric,
                "report_total": report_total,
                "project_full_total": full_total,
                "project_active_total": active_total,
                "diff_full": diff_full,
                "diff_active": diff_active,
                "rule": rule,
                "status": status_text,
            }
        )

    print("\n[region_line_all_metrics]")
    print(pd.DataFrame(summary_rows).to_string(index=False))


def compare_region_line_focus_metrics() -> None:
    report_path = find_optional_report(REGION_LINE)
    if report_path is None:
        print("\n[region_line_focus_metrics]")
        print("region_line report missing")
        return

    project = load_project_report().copy()
    ownership = load_ownership_map()
    non_assess_codes = load_non_assess_codes()
    report = load_region_line_report().copy()

    project = project.merge(ownership[["project_code", "ownership", "is_multi"]], on="project_code", how="left")
    project["is_non_assess"] = project["project_code"].isin(non_assess_codes)
    project["is_assess"] = ~project["is_non_assess"]
    project["line"] = project["line"].astype(str).str.strip()

    metric_specs = [
        {
            "metric": "owner_rcv",
            "source_col": "owner_rcv",
            "logic": "项目级 N+12小业主实收金额，按 多业权 且 考核项目 汇总",
            "mask": project["is_multi"].fillna(False) & project["is_assess"],
        },
        {
            "metric": "big_owner_rev",
            "source_col": "big_owner_rev",
            "logic": "项目级 N+12大业主营业收入，按 多业权 且 考核项目 汇总",
            "mask": project["is_multi"].fillna(False) & project["is_assess"],
        },
        {
            "metric": "big_owner_ar",
            "source_col": "big_owner_ar",
            "logic": "项目级 N+12大业主欠费余额，按 多业权 且 考核项目 汇总",
            "mask": project["is_multi"].fillna(False) & project["is_assess"],
        },
        {
            "metric": "non_assess_revenue",
            "source_col": "revenue",
            "logic": "非考核多业权项目的 项目级营业收入 汇总",
            "mask": project["is_multi"].fillna(False) & project["is_non_assess"],
        },
        {
            "metric": "non_assess_related_revenue",
            "source_col": "related_revenue",
            "logic": "非考核多业权项目的 项目级关联方营业收入 汇总",
            "mask": project["is_multi"].fillna(False) & project["is_non_assess"],
        },
    ]

    summary_rows: list[dict[str, object]] = []
    detail_frames: list[pd.DataFrame] = []
    for spec in metric_specs:
        grouped = (
            project.loc[spec["mask"]]
            .groupby(["region", "line"], as_index=False)[spec["source_col"]]
            .sum()
            .rename(columns={spec["source_col"]: "project_sum"})
        )
        merged = report[["region", "line", spec["metric"]]].merge(grouped, on=["region", "line"], how="left")
        merged["project_sum"] = merged["project_sum"].fillna(0.0)
        merged["diff"] = merged["project_sum"] - merged[spec["metric"]]
        merged["metric"] = spec["metric"]
        merged["logic"] = spec["logic"]
        detail_frames.append(merged)

        summary_rows.append(
            {
                "metric": spec["metric"],
                "logic": spec["logic"],
                "report_total": float(report[spec["metric"]].sum()),
                "project_total": float(merged["project_sum"].sum()),
                "diff": float(merged["diff"].sum()),
                "status": "PASS" if abs(float(merged["diff"].sum())) <= TOL else "FAIL",
            }
        )

    missing = project[project["ownership"].isna() | (project["ownership"].astype(str).str.strip() == "")]
    missing_summary = {
        "missing_projects": int(missing["project_code"].nunique()),
        "assess_owner_rcv": float(missing.loc[~missing["is_non_assess"], "owner_rcv"].sum()),
        "assess_big_owner_rev": float(missing.loc[~missing["is_non_assess"], "big_owner_rev"].sum()),
        "assess_big_owner_ar": float(missing.loc[~missing["is_non_assess"], "big_owner_ar"].sum()),
        "non_assess_revenue": float(missing.loc[missing["is_non_assess"], "revenue"].sum()),
        "non_assess_related_revenue": float(missing.loc[missing["is_non_assess"], "related_revenue"].sum()),
    }

    print("\n[region_line_focus_metrics]")
    print(pd.DataFrame(summary_rows).to_string(index=False))
    print("\n[region_line_focus_metric_details]")
    print(pd.concat(detail_frames, ignore_index=True).to_string(index=False))
    print("\n[ownership_mapping_gap]")
    print(pd.DataFrame([missing_summary]).to_string(index=False))


def compare_region_line_gold_logic() -> None:
    report_path = find_optional_report(REGION_LINE)
    if report_path is None:
        print("\n[region_line_gold_logic]")
        print("区域条线报表缺失，跳过")
        return

    report = load_region_line_report()[["region", "line", "gold_balance"]].rename(columns={"gold_balance": "report_gold"})
    prev_gold = (
        load_prev_year_rate_project_gold()
        .groupby(["region", "line"], as_index=False)["gold_balance"]
        .sum()
        .rename(columns={"gold_balance": "prev_gold_balance"})
    )

    current_gold = load_current_gold_project()
    current_all = (
        current_gold.groupby(["region", "line"], as_index=False)["current_gold"]
        .sum()
        .rename(columns={"current_gold": "current_gold_all"})
    )

    query_flags = load_project_query_flags()
    current_diag = current_gold.merge(query_flags, on="code_exact", how="left")
    current_diag = current_diag[
        ~(
            current_diag["project_level"].astype(str).str.startswith("D")
            & (current_diag["project_status_query"].astype(str) == STATUS_EXIT)
        )
    ].copy()
    current_d_exit = (
        current_diag.groupby(["region", "line"], as_index=False)["current_gold"]
        .sum()
        .rename(columns={"current_gold": "current_gold_exclude_d_exit"})
    )

    merged = report.merge(prev_gold, on=["region", "line"], how="left")
    merged = merged.merge(current_all, on=["region", "line"], how="left")
    merged = merged.merge(current_d_exit, on=["region", "line"], how="left").fillna(0.0)
    merged["diff_prev_logic"] = merged["prev_gold_balance"] - merged["report_gold"]
    merged["diff_current_logic"] = merged["current_gold_all"] - merged["report_gold"]
    merged["diff_current_exclude_d_exit"] = merged["current_gold_exclude_d_exit"] - merged["report_gold"]

    print("\n[region_line_gold_logic]")
    print(
        merged[
            [
                "region",
                "line",
                "report_gold",
                "current_gold_all",
                "diff_current_logic",
                "prev_gold_balance",
                "diff_prev_logic",
                "current_gold_exclude_d_exit",
                "diff_current_exclude_d_exit",
            ]
        ].to_string(index=False)
    )
    print("totals:")
    print(f"  report_gold: {merged['report_gold'].sum():.10f}")
    print(f"  current_gold_all: {merged['current_gold_all'].sum():.10f}")
    print(f"  prev_gold_balance: {merged['prev_gold_balance'].sum():.10f}")
    print(f"  current_gold_exclude_d_exit: {merged['current_gold_exclude_d_exit'].sum():.10f}")
    diagnose_region_line_prev_gold_gap()


def diagnose_region_line_prev_gold_gap() -> None:
    report_path = find_optional_report(REGION_LINE)
    if report_path is None:
        return

    project = load_project_report()[["region", "line", "project_code", "project_name", "gold_balance"]].copy()
    project["code_exact"] = project["project_code"].map(exact_code)
    project = project[project["line"].astype(str) == u(r"\u4f4f\u5b85\u670d\u52a1")].copy()
    query = load_project_query_flags()
    project = project.merge(query, on="code_exact", how="left")

    report = load_region_line_report()[["region", "line", "gold_balance"]].rename(columns={"gold_balance": "report_gold"})
    region_totals = (
        project.groupby(["region", "line"], as_index=False)["gold_balance"]
        .sum()
        .rename(columns={"gold_balance": "project_gold"})
        .merge(report, on=["region", "line"], how="left")
    )
    region_totals["diff"] = region_totals["project_gold"] - region_totals["report_gold"]

    excluded_mask = project["project_status_query"].astype(str).isin([STATUS_EXIT, STATUS_NOT_IN])
    excluded_totals = (
        project[excluded_mask]
        .groupby(["region", "line"], as_index=False)["gold_balance"]
        .sum()
        .rename(columns={"gold_balance": "excluded_gold"})
    )
    summary = region_totals.merge(excluded_totals, on=["region", "line"], how="left").fillna({"excluded_gold": 0.0})
    summary["diff_minus_excluded"] = summary["diff"] - summary["excluded_gold"]

    print("\n[region_line_prev_gold_gap_reason]")
    print(summary.to_string(index=False))

    detail = project[excluded_mask & (project["gold_balance"].abs() > TOL)].copy()
    if detail.empty:
        return

    detail = detail.sort_values(["region", "gold_balance"], ascending=[True, False])
    print("\n[region_line_prev_gold_gap_projects]")
    print(
        detail[
            [
                "region",
                "project_code",
                "project_name",
                "gold_balance",
                "project_level",
                "project_status_query",
            ]
        ].to_string(index=False)
    )


def main() -> None:
    project = load_project_report()

    print("[internal_formula_check]")
    internal_formula_check("project", project)
    region = None
    housing = None
    if find_optional_report(REGION_HOUSING) is not None:
        region = load_region_housing_report()
        internal_formula_check("region_housing", region)
    if find_optional_report(HOUSING, excludes=[REGION_HOUSING, PROJECT]) is not None:
        housing = load_housing_report()
        internal_formula_check("housing", housing)
    if find_optional_report(REGION_LINE) is not None:
        region_line = load_region_line_report()
        internal_formula_check("region_line", region_line)

    print("\n[project_source_component_checks]")
    print("当前工作区缺少旧脚本依赖的 202412 经营收支底表，项目源台账全量校验本次缓测")

    if region is not None and housing is not None:
        compare_rollup(project, region, housing)
    compare_region_line_focus_metrics()

    print("\n[blocked_components]")
    print("N+12小业主实收金额: 新视窗来源按用户要求缓测")
    print("N+12大业主营业收入: 金蝶来源按用户要求缓测")
    print("上一年在当期的金币回款: 未见小业主金币回款金额台账，且N+12项目报表没有单独列示该扣减项")
    print("非考核项目营业收入: 用户确认项目级没有值，本次跳过")
    print("非考核项目关联方营业收入: 用户确认项目级没有值，本次跳过")

    print("\n[summary]")
    print("failed_source_components: blocked")


if __name__ == "__main__":
    main()
