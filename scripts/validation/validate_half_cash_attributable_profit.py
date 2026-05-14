from __future__ import annotations

from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "202512"
REPORT_MONTH_DASHED = "2025-12"
TOLERANCE = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


HALF_ATTRIBUTABLE = u(r"\u534a\u6536\u4ed8\u5f52\u6bcd\u51c0\u5229\u6da6")
PROJECT = u(r"\u9879\u76ee")
REGION = u(r"\u533a\u57df")
LINE = u(r"\u6761\u7ebf")
REGION_LINE = u(r"\u533a\u57df\u6761\u7ebf")
INDICATOR = u(r"\u6307\u6807\u6e05\u5355")
QUERY = u(r"\u9879\u76ee\u67e5\u8be2")
NON_ASSESS = u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26")
PROFIT_NON_ASSESS = u(r"\u5229\u6da6\u7c7b\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26")
OTHER_ADJUST = u(r"\u5176\u4ed6\u8003\u6838\u8c03\u6574")
NON_DEV_INCENTIVE = u(r"\u975e\u53d1\u5c55\u4eba\u5458")
PLAN = u(r"\u5e26\u8d44\u644a\u9500")
RATIO = u(r"\u6bd4\u4f8b\u914d\u7f6e")
EARLY_END = u(r"\u63d0\u524d\u7ed3\u675f")
TERMINATED = u(r"\u5df2\u7ec8\u6b62")
NORMAL = u(r"\u6b63\u5e38")
FINANCE_CLOUD = u(r"\u8d22\u52a1\u4e91")
CAPITAL_ENTRY = u(r"\u5e26\u8d44\u8fdb\u573a")
HALF_OPERATING = "1.5.2-"
METRIC_COLUMNS = [
    "half_net_profit",
    "non_assess_half_net_profit",
    "minority_loss",
    "plan_capital",
    "actual_capital",
    "plan_smart",
    "actual_smart",
    "plan_quality",
    "actual_quality",
    "jka_2023_adjust",
    "other_adjust",
    "water_backflow",
    "water_backflow_alloc",
    "water_remaining_receivable",
    "water_current_receivable_alloc",
    "energy_income_tax",
    "single_current_unreceived",
    "single_prev_unreceived",
    "water_adjust",
    "overdue_performance_bond",
    "overdue_bid_bond",
    "cutoff_expense",
    "dev_incentive",
    "jl_jj_restore",
    "interest",
    "satellite_profit",
    "single_current_unreceived_2",
    "single_prev_unreceived_2",
    "prev_region_perf",
    "curr_region_perf",
    "half_attributable_profit",
]


def find_workbook(*tokens: str) -> Path:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and all(token in path.name for token in tokens)
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one workbook for {tokens}, got {len(matches)}: {[p.name for p in matches]}")
    return matches[0]


def find_finance_cloud_file() -> Path:
    """优先匹配最新的财务云/实际发生数文件，兼容 xlsx 和 xls。"""
    tokens = [FINANCE_CLOUD, u(r"\u5b9e\u9645\u53d1\u751f\u6570")]
    matches = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and any(token in path.name for token in tokens)
    ]
    if not matches:
        raise RuntimeError("Expected one finance-cloud workbook, got 0")
    matches.sort(key=lambda path: (path.suffix.lower() != ".xls", path.name, path.stat().st_mtime))
    return matches[-1]


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text in {"", "NAN", "NONE"}:
        return ""
    if len(text) >= 2 and text[0].isalpha() and any(ch.isdigit() for ch in text[1:]):
        return text[1:]
    return text


def numeric(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for column in columns:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


CAPITAL_ACTUAL_OVERRIDE: dict[str, float] = {}


def print_section(label: str, records: list[dict]) -> None:
    print(f"[{label}]")
    if not records:
        print("[]")
        return
    print(pd.DataFrame(records).to_json(force_ascii=True, orient="records"))


def load_indicator_final_rows() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(INDICATOR), sheet_name=0, dtype=object)
    cols = list(df.columns)
    name_col = cols[3]
    dim_col = cols[4]
    cycle_col = cols[5]
    method_col = cols[8]
    source_col = cols[10]
    logic_col = cols[12]
    rows = df[
        df[name_col].astype(str).isin([f"{HALF_ATTRIBUTABLE}_{PROJECT}", f"{HALF_ATTRIBUTABLE}_{REGION}"])
        & df[dim_col].astype(str).isin([PROJECT, REGION])
    ][[name_col, dim_col, cycle_col, method_col, source_col, logic_col]].fillna("")
    rows.columns = ["metric_name", "dimension", "cycle", "method", "source_table", "logic"]
    return rows


def load_indicator_related_sources() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(INDICATOR), sheet_name=0, dtype=object)
    cols = list(df.columns)
    relation_col = cols[1]
    name_col = cols[3]
    dim_col = cols[4]
    cycle_col = cols[5]
    source_col = cols[10]
    logic_col = cols[12]
    related = df[
        df[relation_col].astype(str).eq(HALF_ATTRIBUTABLE)
        | df[name_col].astype(str).str.contains(HALF_ATTRIBUTABLE, na=False, regex=False)
    ][[name_col, dim_col, cycle_col, source_col, logic_col]].fillna("")
    related.columns = ["metric_name", "dimension", "cycle", "source_table", "logic"]
    return related


def load_project_report() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(HALF_ATTRIBUTABLE, PROJECT), sheet_name=0)
    df.columns = [
        "region",
        "line",
        "project_code",
        "project_name",
        "half_net_profit",
        "non_assess_half_net_profit",
        "minority_loss",
        "plan_capital",
        "actual_capital",
        "plan_smart",
        "actual_smart",
        "plan_quality",
        "actual_quality",
        "jka_2023_adjust",
        "other_adjust",
        "water_backflow",
        "water_backflow_alloc",
        "water_remaining_receivable",
        "water_current_receivable_alloc",
        "energy_income_tax",
        "single_current_unreceived",
        "single_prev_unreceived",
        "water_adjust",
        "overdue_performance_bond",
        "overdue_bid_bond",
        "cutoff_expense",
        "dev_incentive",
        "jl_jj_restore",
        "interest",
        "satellite_profit",
        "single_current_unreceived_2",
        "single_prev_unreceived_2",
        "prev_region_perf",
        "curr_region_perf",
        "half_attributable_profit",
        "ownership_attr",
    ]
    df = numeric(df, list(df.columns[4:35]))
    df["code_norm"] = df["project_code"].map(normalize_code)
    df["code_exact"] = df["project_code"].astype(str).str.strip().str.upper()
    return df


def load_region_report() -> pd.DataFrame:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx"
        and HALF_ATTRIBUTABLE in path.name
        and REGION in path.name
        and REGION_LINE not in path.name
        and LINE not in path.name
        and PROJECT not in path.name
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one region workbook, got {len(matches)}: {[p.name for p in matches]}")
    df = pd.read_excel(matches[0], sheet_name=0)
    df.columns = ["region", *METRIC_COLUMNS, "ownership_attr"]
    df = numeric(df, list(df.columns[1:32]))
    return df


def load_line_report() -> pd.DataFrame:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx"
        and HALF_ATTRIBUTABLE in path.name
        and LINE in path.name
        and REGION_LINE not in path.name
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one line workbook, got {len(matches)}: {[p.name for p in matches]}")
    df = pd.read_excel(matches[0], sheet_name=0)
    df.columns = ["line", *METRIC_COLUMNS, "ownership_attr"]
    return numeric(df, METRIC_COLUMNS)


def load_region_line_report() -> pd.DataFrame:
    matches = [
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and HALF_ATTRIBUTABLE in path.name and REGION_LINE in path.name
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one region-line workbook, got {len(matches)}: {[p.name for p in matches]}")
    df = pd.read_excel(matches[0], sheet_name=0)
    df.columns = ["region", "line", *METRIC_COLUMNS, "ownership_attr"]
    return numeric(df, METRIC_COLUMNS)


def load_query() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(QUERY))
    column_map = {
        u(r"\u7acb\u9879\u7f16\u7801"): "project_code",
        u(r"\u9879\u76ee\u540d\u79f0"): "project_name",
        u(r"\u6240\u5c5e\u533a\u57df"): "region",
        u(r"\u7ba1\u7406\u5f52\u5c5e\u7247\u533a"): "charge_area",
        u(r"\u4e1a\u52a1\u5c5e\u6027"): "business_attr",
        u(r"\u9879\u76ee\u7b49\u7ea7"): "project_level",
        u(r"\u6536\u8d39\u9762\u79ef\uff08\u4e07\u65b9\uff09"): "charge_area_size",
        u(r"\u6240\u5c5e\u6cd5\u4eba"): "legal_entity",
        u(r"\u5408\u8d44\u516c\u53f8\u6807\u7b7e"): "jv_tag",
        u(r"\u7a7f\u900f\u6bd4\u4f8b"): "penetration_ratio",
        u(r"\u4e1a\u6743\u5c5e\u6027"): "ownership_attr",
        u(r"\u9879\u76ee\u72b6\u6001"): "project_status",
        u(r"\u8fdb\u573a\u65f6\u95f4"): "entry_date",
        u(r"\u5df2\u64a4\u573a\u65f6\u95f4"): "exit_date",
    }
    df = df.rename(columns={column: column_map.get(str(column).strip(), column) for column in df.columns})
    expected = [
        "project_code",
        "project_name",
        "region",
        "charge_area",
        "business_attr",
        "project_level",
        "charge_area_size",
        "legal_entity",
        "penetration_ratio",
        "ownership_attr",
        "project_status",
        "entry_date",
        "exit_date",
    ]
    missing = [column for column in expected if column not in df.columns]
    if missing:
        raise RuntimeError(f"Missing expected columns in project query: {missing}; got {list(df.columns)}")
    df = df[expected + [column for column in df.columns if column not in expected]]
    df["code_norm"] = df["project_code"].map(normalize_code)
    df["penetration_ratio"] = pd.to_numeric(df["penetration_ratio"], errors="coerce").fillna(1.0)
    return df


def load_non_assess_codes() -> set[str]:
    path = next(
        path
        for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and path.name == f"{NON_ASSESS}.xlsx"
    )
    df = pd.read_excel(path)
    code_col = next(column for column in df.columns if u(r"\u7acb\u9879\u7f16\u7801") in str(column))
    return {normalize_code(value) for value in df[code_col].dropna()}


def load_profit_non_assess_codes() -> set[str]:
    df = pd.read_excel(find_workbook(PROFIT_NON_ASSESS), header=5)
    code_col = next(column for column in df.columns if u(r"\u7acb\u9879\u7f16\u7801") in str(column))
    month_col = next(column for column in df.columns if u(r"\u6570\u636e\u5e74\u6708") in str(column))
    df = df[df[month_col].astype(str).str.strip().eq(REPORT_MONTH_DASHED)]
    return {normalize_code(value) for value in df[code_col].dropna()}


def load_1_5_2_project_profit() -> pd.DataFrame:
    raw = pd.read_excel(find_workbook(HALF_OPERATING), header=None)
    data = raw.iloc[5:].reset_index(drop=True).copy()
    data.columns = [f"c{i}" for i in range(data.shape[1])]
    data = data[data["c1"].astype(str).str.contains(r"\d", na=False)].copy()
    data["code_exact"] = data["c1"].astype(str).str.strip().str.upper()
    data["code_norm"] = data["c1"].map(normalize_code)
    data["half_net_profit_source"] = pd.to_numeric(data["c29"], errors="coerce").fillna(0.0)
    return data.groupby(["code_exact", "code_norm"], as_index=False)["half_net_profit_source"].sum()


def load_non_dev_incentive_region() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(NON_DEV_INCENTIVE), header=0)
    type_col = next(column for column in df.columns if u(r"\u7c7b\u578b") in str(column))
    month_col = next(column for column in df.columns if u(r"\u5e74\u6708") in str(column))
    region_col = next(column for column in df.columns if u(r"\u533a\u57df") in str(column))
    amount_col = next(column for column in df.columns if u(r"\u53d1\u5c55\u6fc0\u52b1\u91d1\u989d") in str(column))
    df = df[
        df[type_col].astype(str).str.strip().eq(REGION)
        & df[month_col].astype(str).str.strip().eq(REPORT_MONTH_DASHED)
    ].copy()
    df["amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0) * 0.81
    return df.groupby(region_col, as_index=False)["amount"].sum().rename(columns={region_col: "region"})


def load_interest_region() -> pd.DataFrame:
    fund = pd.read_excel(find_workbook(u(r"\u8d44\u91d1\u5229\u606f")), header=1)
    rel = pd.read_excel(find_workbook(u(r"\u6cd5\u4eba\u4e0e\u7ec4\u7ec7\u5173\u7cfb")), header=0)
    jv = pd.read_excel(find_workbook(u(r"\u5408\u8d44\u516c\u53f8")), header=1)
    fund = fund.rename(columns={fund.columns[0]: "period", fund.columns[1]: "legal", fund.columns[2]: "amount_tax"})
    rel = rel.rename(columns={rel.columns[0]: "legal", rel.columns[1]: "region", rel.columns[2]: "primary", rel.columns[3]: "enabled"})
    jv = jv.rename(columns={jv.columns[2]: "legal", jv.columns[3]: "tag"})
    fund = fund[fund["period"].astype(str).str.strip().eq(REPORT_MONTH_DASHED)].copy()
    fund["amount_tax"] = pd.to_numeric(fund["amount_tax"], errors="coerce").fillna(0.0)
    merged = fund.merge(rel, on="legal", how="left").merge(jv[["legal", "tag"]], on="legal", how="left")
    filtered = merged[
        merged["primary"].astype(str).eq(u(r"\u662f"))
        & merged["enabled"].astype(str).eq(u(r"\u542f\u7528"))
        & ~merged["tag"].astype(str).isin(["C", "D"])
    ].copy()
    filtered["amount"] = filtered["amount_tax"] / 1.06 * 0.81
    return filtered.groupby("region", as_index=False)["amount"].sum()


def half_formula(df: pd.DataFrame, include_region_perf: bool) -> pd.Series:
    calc = (
        df["half_net_profit"]
        - df["non_assess_half_net_profit"]
        - df["minority_loss"]
        - df["plan_capital"]
        + df["actual_capital"]
        - df["plan_smart"]
        + df["actual_smart"]
        - df["plan_quality"]
        + df["actual_quality"]
        - df["jka_2023_adjust"]
        + df["other_adjust"]
        + df["water_adjust"]
        - df["overdue_performance_bond"]
        - df["overdue_bid_bond"]
        + df["cutoff_expense"]
        + df["dev_incentive"]
        + df["jl_jj_restore"]
        + df["interest"]
        + df["satellite_profit"]
        + df["single_current_unreceived_2"]
        - df["single_prev_unreceived_2"]
    )
    if include_region_perf:
        calc = calc + df["prev_region_perf"] - df["curr_region_perf"]
    return calc


def observed_project_formula(df: pd.DataFrame) -> pd.Series:
    return (
        df["half_net_profit"]
        - df["non_assess_half_net_profit"]
        - df["minority_loss"]
        - df["plan_capital"]
        + df["actual_capital"]
        - df["plan_smart"]
        + df["actual_smart"]
        - df["plan_quality"]
        - df["actual_quality"]
        - df["jka_2023_adjust"]
        + df["other_adjust"]
        + df["water_adjust"]
        - df["overdue_performance_bond"]
        - df["overdue_bid_bond"]
        + df["cutoff_expense"]
        + df["dev_incentive"]
        + df["jl_jj_restore"]
        + df["interest"]
        + df["satellite_profit"]
        + df["single_current_unreceived_2"]
        - df["single_prev_unreceived_2"]
    )


def observed_region_formula(df: pd.DataFrame) -> pd.Series:
    return (
        df["half_net_profit"]
        - df["non_assess_half_net_profit"]
        - df["minority_loss"]
        - df["plan_capital"]
        + df["actual_capital"]
        - df["plan_smart"]
        - df["plan_quality"]
        - df["actual_quality"]
        - df["jka_2023_adjust"]
        + df["other_adjust"]
        + df["water_adjust"]
        - df["overdue_performance_bond"]
        - df["overdue_bid_bond"]
        + df["cutoff_expense"]
        + df["dev_incentive"]
        + df["jl_jj_restore"]
        + df["interest"]
        + df["satellite_profit"]
        + df["single_current_unreceived_2"]
        - df["single_prev_unreceived_2"]
        + df["prev_region_perf"]
        - df["curr_region_perf"]
    )


def water_formula(df: pd.DataFrame) -> pd.Series:
    return (
        df["water_backflow"]
        + df["water_backflow_alloc"]
        - df["water_remaining_receivable"]
        - df["water_current_receivable_alloc"]
        - df["energy_income_tax"]
        + df["single_current_unreceived"]
        - df["single_prev_unreceived"]
    )


def summarize_diff(label: str, calc: pd.Series, report: pd.Series) -> dict:
    diff = calc - report
    return {
        "check": label,
        "status": "passed" if (diff.abs() <= TOLERANCE).all() else "failed",
        "rows": int(len(diff)),
        "mismatch_rows": int((diff.abs() > TOLERANCE).sum()),
        "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
        "calc_total": float(calc.sum()),
        "report_total": float(report.sum()),
        "diff_total": float(diff.sum()),
    }


def summarize_dim_rollup(
    label: str,
    report_df: pd.DataFrame,
    calc_df: pd.DataFrame,
    keys: list[str],
    value_col: str,
    tolerance: float = TOLERANCE,
) -> dict:
    merged = report_df[keys + [value_col]].merge(
        calc_df[keys + [value_col]].rename(columns={value_col: "calc"}),
        on=keys,
        how="left",
    )
    merged["calc"] = merged["calc"].fillna(0.0)
    merged["diff"] = merged["calc"] - merged[value_col]
    mismatches = merged.loc[merged["diff"].abs() > tolerance, keys + [value_col, "calc", "diff"]]
    return {
        "check": label,
        "status": "passed" if mismatches.empty else "failed",
        "rows": int(len(merged)),
        "mismatch_rows": int(len(mismatches)),
        "report_total": float(merged[value_col].sum()),
        "calc_total": float(merged["calc"].sum()),
        "diff_total": float(merged["diff"].sum()),
        "max_abs_diff": float(merged["diff"].abs().max() if len(merged) else 0.0),
        "mismatches": mismatches.to_dict(orient="records"),
    }


def build_project_half_net_check(project_df: pd.DataFrame) -> pd.DataFrame:
    source = load_1_5_2_project_profit()
    check = project_df[["region", "line", "project_code", "project_name", "code_exact", "code_norm", "half_net_profit"]].merge(
        source[["code_exact", "half_net_profit_source"]],
        on="code_exact",
        how="left",
    )
    fallback = source.groupby("code_norm", as_index=False).agg(
        source_exact_count=("code_exact", "nunique"),
        fallback_half_net_profit_source=("half_net_profit_source", "sum"),
    )
    check = check.merge(fallback, on="code_norm", how="left")
    missing_exact = check["half_net_profit_source"].isna()
    no_prefix_code = ~check["code_exact"].str.match(r"^[A-Z]")
    use_fallback = missing_exact & no_prefix_code & check["source_exact_count"].eq(1)
    check.loc[use_fallback, "half_net_profit_source"] = check.loc[use_fallback, "fallback_half_net_profit_source"]
    check["half_net_profit_source"] = check["half_net_profit_source"].fillna(0.0)
    check["diff"] = check["half_net_profit"] - check["half_net_profit_source"]
    return check


def load_1_5_2_project_profit_components() -> pd.DataFrame:
    """读取 1.5.2 底表中项目级半收付净利润及税前利润组成字段。"""
    raw = pd.read_excel(find_workbook(HALF_OPERATING), header=None)
    data = raw.iloc[5:].reset_index(drop=True).copy()
    data.columns = [f"c{i}" for i in range(data.shape[1])]
    data = data[data["c1"].astype(str).str.contains(r"\d", na=False)].copy()
    data["code_exact"] = data["c1"].astype(str).str.strip().str.upper()
    data["code_norm"] = data["c1"].map(normalize_code)
    data["half_income"] = pd.to_numeric(data["c12"], errors="coerce").fillna(0.0)
    data["non_rec_income"] = pd.to_numeric(data["c13"], errors="coerce").fillna(0.0)
    data["business_cost"] = pd.to_numeric(data["c14"], errors="coerce").fillna(0.0)
    data["sales_fee"] = pd.to_numeric(data["c15"], errors="coerce").fillna(0.0)
    data["manage_fee"] = pd.to_numeric(data["c16"], errors="coerce").fillna(0.0)
    data["finance_fee"] = pd.to_numeric(data["c17"], errors="coerce").fillna(0.0)
    data["other_expense"] = pd.to_numeric(data["c20"], errors="coerce").fillna(0.0)
    data["profit_total_pre_mgmt"] = pd.to_numeric(data["c24"], errors="coerce").fillna(0.0)
    data["income_tax_pre_mgmt"] = pd.to_numeric(data["c27"], errors="coerce").fillna(0.0)
    data["half_net_profit_source"] = pd.to_numeric(data["c29"], errors="coerce").fillna(0.0)
    data["profit_total_user_formula"] = (
        data["half_income"]
        + data["non_rec_income"]
        - data["business_cost"]
        - data["sales_fee"]
        - data["manage_fee"]
        - data["finance_fee"]
        - data["other_expense"]
    )
    return data[
        [
            "code_exact",
            "code_norm",
            "profit_total_pre_mgmt",
            "income_tax_pre_mgmt",
            "half_net_profit_source",
            "profit_total_user_formula",
        ]
    ]


def build_region_line_half_net_tax_formula_check(project_df: pd.DataFrame) -> pd.DataFrame:
    """按税前利润汇总后重算区域条线所得税，便于定位高维税差。"""
    source = load_1_5_2_project_profit_components()
    mapping = project_df[
        ["region", "line", "project_code", "project_name", "code_exact", "code_norm", "half_net_profit"]
    ].copy()
    check = source.merge(mapping[["region", "line", "code_exact", "code_norm"]], on=["code_exact", "code_norm"], how="left")
    fallback = mapping.groupby("code_norm", as_index=False).agg(
        region_count=("region", "nunique"),
        line_count=("line", "nunique"),
        fallback_region=("region", "first"),
        fallback_line=("line", "first"),
    )
    check = check.merge(fallback, on="code_norm", how="left")
    missing = check["region"].isna() & check["region_count"].eq(1) & check["line_count"].eq(1)
    check.loc[missing, "region"] = check.loc[missing, "fallback_region"]
    check.loc[missing, "line"] = check.loc[missing, "fallback_line"]
    check = check[check["region"].notna() & check["line"].notna()].copy()

    project_source = check.groupby(["code_exact", "code_norm", "region", "line"], as_index=False).agg(
        project_profit_total_pre_mgmt=("profit_total_pre_mgmt", "sum"),
        project_half_net_profit_source=("half_net_profit_source", "sum"),
    )
    project_source["project_tax_recalc"] = project_source["project_profit_total_pre_mgmt"].clip(lower=0.0) * 0.19
    project_source["project_half_net_profit_recalc"] = (
        project_source["project_profit_total_pre_mgmt"] - project_source["project_tax_recalc"]
    )

    region_line_row = check.groupby(["region", "line"], as_index=False).agg(
        row_profit_total_pre_mgmt=("profit_total_pre_mgmt", "sum"),
        row_profit_total_user_formula=("profit_total_user_formula", "sum"),
        row_income_tax_sum=("income_tax_pre_mgmt", "sum"),
        row_half_net_profit_sum=("half_net_profit_source", "sum"),
    )
    region_line_row["row_tax_recalc"] = region_line_row["row_profit_total_pre_mgmt"].clip(lower=0.0) * 0.19
    region_line_row["row_half_net_profit_recalc"] = (
        region_line_row["row_profit_total_pre_mgmt"] - region_line_row["row_tax_recalc"]
    )

    region_line_project = project_source.groupby(["region", "line"], as_index=False).agg(
        project_profit_total_pre_mgmt=("project_profit_total_pre_mgmt", "sum"),
        project_half_net_profit_sum=("project_half_net_profit_source", "sum"),
        project_half_net_profit_recalc=("project_half_net_profit_recalc", "sum"),
    )

    report = load_region_line_report()[["region", "line", "half_net_profit"]].copy()
    result = (
        report.merge(region_line_row, on=["region", "line"], how="left")
        .merge(region_line_project, on=["region", "line"], how="left")
        .fillna(0.0)
    )
    result["diff_vs_project_sum"] = result["half_net_profit"] - result["project_half_net_profit_sum"]
    result["diff_vs_row_recalc"] = result["row_half_net_profit_recalc"] - result["half_net_profit"]
    result["diff_vs_project_recalc"] = result["project_half_net_profit_recalc"] - result["half_net_profit"]
    return result


def summarize_metric_rollups(project_df: pd.DataFrame, region_line_df: pd.DataFrame) -> list[dict]:
    calc = project_df.groupby(["region", "line"], as_index=False)[METRIC_COLUMNS].sum()
    rows = []
    for metric in METRIC_COLUMNS:
        merged = region_line_df[["region", "line", metric]].merge(
            calc[["region", "line", metric]].rename(columns={metric: "calc"}),
            on=["region", "line"],
            how="left",
        )
        merged["calc"] = merged["calc"].fillna(0.0)
        diff = merged["calc"] - merged[metric]
        rows.append(
            {
                "metric": metric,
                "status": "passed" if diff.abs().le(TOLERANCE).all() else "failed",
                "mismatch_rows": int(diff.abs().gt(TOLERANCE).sum()),
                "diff_total": float(diff.sum()),
                "abs_diff_total": float(diff.abs().sum()),
                "max_abs_diff": float(diff.abs().max() if len(diff) else 0.0),
            }
        )
    rows.sort(key=lambda item: (-item["abs_diff_total"], -item["mismatch_rows"], item["metric"]))
    return rows


def mark_d_exit(project_df: pd.DataFrame, query_df: pd.DataFrame) -> pd.DataFrame:
    merged = project_df.merge(
        query_df[["code_norm", "project_level", "project_status", "exit_date"]].drop_duplicates("code_norm"),
        on="code_norm",
        how="left",
    )
    exit_ym = pd.to_datetime(merged["exit_date"], errors="coerce").dt.strftime("%Y%m").fillna("")
    merged["is_d_exit_by_month"] = (
        merged["project_level"].astype(str).str.startswith("D")
        & merged["project_status"].astype(str).eq(u(r"\u5df2\u64a4\u573a"))
        & exit_ym.ne("")
        & exit_ym.le(REPORT_MONTH)
    )
    merged["is_d_exit_status"] = (
        merged["project_level"].astype(str).str.startswith("D")
        & merged["project_status"].astype(str).eq(u(r"\u5df2\u64a4\u573a"))
    )
    return merged


def load_other_adjust_project() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(OTHER_ADJUST), header=1)
    df = df[df[df.columns[1]].astype(str).eq(PROJECT)].copy()
    df["code_norm"] = df[df.columns[4]].map(normalize_code)
    amount_col = next(column for column in df.columns if HALF_ATTRIBUTABLE in str(column))
    df["amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    return df.groupby("code_norm", as_index=False)["amount"].sum()


def load_other_adjust_region() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(OTHER_ADJUST), header=1)
    df = df[df[df.columns[1]].astype(str).eq(REGION)].copy()
    amount_col = next(column for column in df.columns if HALF_ATTRIBUTABLE in str(column))
    df["amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    return df.groupby(df.columns[5], as_index=False)["amount"].sum().rename(columns={df.columns[5]: "region"})


def load_plan_ledger(query_df: pd.DataFrame) -> pd.DataFrame:
    plan_path = next(
        path for path in ROOT.iterdir()
        if path.suffix.lower() == ".xlsx" and PLAN in path.name and RATIO not in path.name
    )
    df = pd.read_excel(plan_path)
    df["code_norm"] = df[df.columns[2]].map(normalize_code)
    df["plan_type"] = df[df.columns[4]].astype(str)
    df["amount"] = pd.to_numeric(df[df.columns[6]], errors="coerce").fillna(0.0)
    df["years"] = pd.to_numeric(df[df.columns[7]].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    df["start_date"] = pd.to_datetime(df[df.columns[8]].astype(str) + "-01", errors="coerce")
    df["end_date"] = pd.to_datetime(df[df.columns[9]].astype(str) + "-01", errors="coerce")
    df["status"] = df[df.columns[11]].astype(str)
    df["early_end_date"] = pd.to_datetime(df[df.columns[12]].astype(str) + "-01", errors="coerce")
    df = df.merge(query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"), on="code_norm", how="left")
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)
    return df


def load_plan_ratio_map() -> dict[int, list[float]]:
    ratio_df = pd.read_excel(find_workbook(PLAN, RATIO), header=1)
    ratio_df = ratio_df.rename(columns={ratio_df.columns[0]: "years"})
    ratio_df["years_num"] = pd.to_numeric(ratio_df["years"].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    ratio_cols = [column for column in ratio_df.columns if u(r"\u644a\u9500\u6bd4\u4f8b") in str(column)]
    result: dict[int, list[float]] = {}
    for _, row in ratio_df.dropna(subset=["years_num"]).iterrows():
        result[int(row["years_num"])] = [
            0.0 if pd.isna(pd.to_numeric(row[column], errors="coerce")) else float(pd.to_numeric(row[column], errors="coerce"))
            for column in ratio_cols
        ]
    return result


def load_capital_actual_raw() -> pd.DataFrame:
    """读取财务云中的带资进场发生额，保留结束月重算所需的含税基数。"""
    df = pd.read_excel(find_finance_cloud_file(), dtype=object)
    def pick_column(*keywords: str) -> str:
        for column in df.columns:
            text = str(column).strip()
            if all(keyword in text for keyword in keywords):
                return column
        raise KeyError(f"未找到列: {keywords}")

    if "sub_project_code" not in df.columns:
        df["sub_project_code"] = df[pick_column(u(r"\u7acb\u9879"), u(r"\u7f16\u7801"))]
    if "share_type_name" not in df.columns:
        df["share_type_name"] = df[pick_column(u(r"\u644a\u9500"), u(r"\u7c7b\u578b"), u(r"\u540d\u79f0"))]
    if "tax_amount" not in df.columns:
        if "amount_dist" in df.columns:
            df["tax_amount"] = df["amount_dist"]
        else:
            df["tax_amount"] = df[pick_column(u(r"\u5206\u644a"), u(r"\u91d1\u989d"))]
    if "approved_time" not in df.columns:
        if "approval_time" in df.columns:
            df["approved_time"] = df["approval_time"]
        else:
            df["approved_time"] = df[pick_column(u(r"\u5ba1\u6279"), u(r"\u901a\u8fc7"), u(r"\u65f6\u95f4"))]

    df["code_norm"] = df["sub_project_code"].map(normalize_code)
    df["share_type_name"] = df["share_type_name"].astype(str).str.strip()
    df["tax_amount"] = pd.to_numeric(df["tax_amount"], errors="coerce").fillna(0.0)
    df["approved_month"] = pd.to_datetime(df["approved_time"], errors="coerce").dt.to_period("M")
    return df.loc[df["share_type_name"].eq(CAPITAL_ENTRY), ["code_norm", "tax_amount", "approved_month"]].copy()


def month_to_period(value: pd.Timestamp | pd.Period | object) -> pd.Period | None:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Period):
        return value
    return pd.Timestamp(value).to_period("M")


def amort_year_index(month_period: pd.Period, start_period: pd.Period) -> int:
    return ((month_period.year - start_period.year) * 12 + month_period.month - start_period.month) // 12


def capital_month_base(amount: float, ratios: dict[int, list[float]], years: int, month_period: pd.Period, start_period: pd.Period) -> float:
    share_list = ratios.get(years, [])
    year_idx = amort_year_index(month_period, start_period)
    share = share_list[year_idx] if 0 <= year_idx < len(share_list) else 0.0
    return amount * share / 12


def capital_prev_actual_base(code_norm: str, prev_year_end: pd.Period, actual_df: pd.DataFrame) -> float:
    """结束月调整时，优先使用已验证特例；否则取财务云截至上一年12月的带资进场含税额累计。"""
    if code_norm in CAPITAL_ACTUAL_OVERRIDE:
        return CAPITAL_ACTUAL_OVERRIDE[code_norm]
    matched = actual_df.loc[
        actual_df["code_norm"].eq(code_norm) & actual_df["approved_month"].notna() & actual_df["approved_month"].le(prev_year_end),
        "tax_amount",
    ]
    return float(matched.sum())


def compute_plan_values(plan_df: pd.DataFrame) -> pd.DataFrame:
    month_periods = list(pd.period_range("2025-01", REPORT_MONTH_DASHED, freq="M"))
    ratios = load_plan_ratio_map()
    capital_actual_df = load_capital_actual_raw()
    mapping = {
        u(r"\u5e26\u8d44\u644a\u9500"): "plan_capital",
        u(r"\u667a\u80fd\u5316\u6574\u6539"): "plan_smart",
        u(r"\u8d28\u6548\u63d0\u5347"): "plan_quality",
    }
    records: list[dict] = []
    for _, row in plan_df.iterrows():
        if pd.isna(row["start_date"]) or pd.isna(row["years"]) or row["plan_type"] not in mapping:
            continue

        start_period = month_to_period(row["start_date"])
        end_period = month_to_period(row["end_date"])
        early_end_period = month_to_period(row["early_end_date"])
        effective_end = early_end_period if row["status"] in {EARLY_END, TERMINATED} and early_end_period is not None else end_period
        total = 0.0
        for month_period in month_periods:
            if start_period is None or month_period < start_period:
                continue
            if effective_end is not None and month_period > effective_end:
                continue
            if row["plan_type"] == u(r"\u5e26\u8d44\u644a\u9500"):
                trigger_month = effective_end
                if trigger_month is not None and month_period == trigger_month:
                    # 序号 9 原文：摊销结束月 / 提前结束月改用“累计发生数 - 累计计划数 - 当年1月至触发月上一月计划数”重算。
                    prev_year_end = pd.Period(f"{trigger_month.year - 1}-12", freq="M")
                    year_start = pd.Period(f"{trigger_month.year}-01", freq="M")
                    trigger_prev = trigger_month - 1
                    prev_plan_base = sum(
                        capital_month_base(row["amount"], ratios, int(row["years"]), period, start_period)
                        for period in pd.period_range(start_period, min(prev_year_end, end_period), freq="M")
                    ) if end_period is not None and prev_year_end >= start_period else 0.0
                    current_year_prev_plan_base = sum(
                        capital_month_base(row["amount"], ratios, int(row["years"]), period, start_period)
                        for period in pd.period_range(max(year_start, start_period), min(trigger_prev, end_period), freq="M")
                    ) if end_period is not None and trigger_prev >= max(year_start, start_period) else 0.0
                    prev_actual_base = capital_prev_actual_base(row["code_norm"], prev_year_end, capital_actual_df)
                    monthly = (
                        (prev_actual_base - prev_plan_base - current_year_prev_plan_base)
                        * 0.81
                        / 1.06
                        * row["penetration_ratio"]
                    )
                else:
                    monthly_base = capital_month_base(row["amount"], ratios, int(row["years"]), month_period, start_period)
                    monthly = monthly_base * 0.81 / 1.06 * row["penetration_ratio"]
            elif row["plan_type"] == u(r"\u667a\u80fd\u5316\u6574\u6539"):
                monthly = row["amount"] / row["years"] / 12 * row["penetration_ratio"] * 0.81
            else:
                monthly = row["amount"] / row["years"] / 12 * row["penetration_ratio"] * 0.81 / 1.06
            total += monthly
        records.append({"code_norm": row["code_norm"], "component": mapping[row["plan_type"]], "amount": total})
    if not records:
        return pd.DataFrame(columns=["code_norm", "plan_capital", "plan_smart", "plan_quality"])
    wide = pd.DataFrame(records).groupby(["code_norm", "component"], as_index=False)["amount"].sum()
    wide = wide.pivot(index="code_norm", columns="component", values="amount").reset_index().fillna(0.0)
    for column in ["plan_capital", "plan_smart", "plan_quality"]:
        if column not in wide.columns:
            wide[column] = 0.0
    return wide[["code_norm", "plan_capital", "plan_smart", "plan_quality"]]


def missing_source_summary(related_sources: pd.DataFrame) -> list[dict]:
    aliases = {
        u(r"1.5.1-\u7ecf\u8425\u6536\u652f\u8868\u67e5\u8be2\u5e95\u8868\uff08\u6743\u8d23\u53e3\u5f84\uff09"): u(
            r"1.5.1-\u7efc\u7ba1\u533a\u57df\u53e3\u5f84"
        ),
        u(r"1.5.2-\u7ecf\u8425\u6536\u652f\u8868\u67e5\u8be2\u5e95\u8868\uff08\u534a\u6536\u4ed8\u53e3\u5f84\uff09"): u(
            r"1.5.2-\u7efc\u7ba1\u533a\u57df\u53e3\u5f84"
        ),
        u(r"\u5173\u8054\u65b9\u5e94\u6536\u53ca\u5b9e\u6536\u5e74\u5ea6\u5206\u5e03"): u(
            r"\u5173\u8054\u65b9\u6c34\u7535\u8d39\u5e94\u6536\u53ca\u5b9e\u6536\u5e74\u5ea6\u5206\u5e03"
        ),
        u(r"\u6cd5\u4eba\u516c\u53f8\u4e0e\u7ec4\u7ec7\u5173\u7cfb\u53f0\u8d26"): u(r"\u6cd5\u4eba\u4e0e\u7ec4\u7ec7\u5173\u7cfb"),
        u(r"\u57ab\u652f\u6c34\u7535\u8d39\u7269\u4e1a\u5206\u644a\u636e\u5b9e\u5206\u644a\u53f0\u8d26"): u(
            r"\u57ab\u652f\u6c34\u7535\u8d39\u7269\u4e1a\u5206\u644a\u636e\u5b9e\u5206\u644a"
        ),
    }
    skipped_sources = {
        u(r"\u57ab\u652f\u6c34\u7535\u8d39\u660e\u7ec6\u8868"),
        u(r"\u903e\u671f\u4fdd\u8bc1\u91d1\u53f0\u8d26"),
    }
    source_names: set[str] = set()
    for value in related_sources["source_table"]:
        text = str(value).strip()
        if not text:
            continue
        for part in text.replace("\r", "\n").split("\n"):
            part = part.strip()
            if part:
                source_names.add(part)
    available = [path.name for path in ROOT.iterdir() if path.suffix.lower() == ".xlsx"]
    records = []
    for source in sorted(source_names):
        if source in skipped_sources:
            records.append({"source_table": source, "status": "skipped_by_user"})
            continue
        alias = aliases.get(source, source)
        tokens = [token for token in source.replace("（", " ").replace("）", " ").replace("(", " ").replace(")", " ").split() if token]
        matched = any(source in name or alias in name or any(token in name for token in tokens) for name in available)
        records.append({"source_table": source, "status": "present" if matched else "missing_or_not_obvious"})
    return records


def main() -> None:
    indicator_rows = load_indicator_final_rows()
    related_sources = load_indicator_related_sources()
    project_df = load_project_report()
    region_df = load_region_report()
    line_df = load_line_report()
    region_line_df = load_region_line_report()
    query_df = load_query()
    non_assess_codes = load_non_assess_codes()
    profit_non_assess_codes = load_profit_non_assess_codes()
    project_flags = mark_d_exit(project_df, query_df)

    print_section("indicator_final_rows", indicator_rows.to_dict(orient="records"))
    print_section("required_source_tables", missing_source_summary(related_sources))

    checks = [
        summarize_diff("project_indicator_formula", half_formula(project_df, include_region_perf=False), project_df["half_attributable_profit"]),
        summarize_diff("region_indicator_formula", half_formula(region_df, include_region_perf=True), region_df["half_attributable_profit"]),
        summarize_diff("project_observed_formula", observed_project_formula(project_df), project_df["half_attributable_profit"]),
        summarize_diff("region_observed_formula", observed_region_formula(region_df), region_df["half_attributable_profit"]),
        summarize_diff("project_water_adjust_formula", water_formula(project_df), project_df["water_adjust"]),
        summarize_diff("region_water_adjust_formula", water_formula(region_df), region_df["water_adjust"]),
    ]
    print_section("formula_checks", checks)

    half_profit = build_project_half_net_check(project_df)
    print_section(
        "project_half_net_profit_1_5_2_check",
        [
            summarize_diff(
                "project_half_net_profit_1_5_2",
                half_profit["half_net_profit_source"],
                half_profit["half_net_profit"],
            )
        ]
        + half_profit.loc[
            (half_profit["half_net_profit_source"] - half_profit["half_net_profit"]).abs() > TOLERANCE,
            ["region", "project_code", "project_name", "half_net_profit", "half_net_profit_source"],
        ].head(30).to_dict(orient="records"),
    )

    # 条线和区域条线这层默认先看是否能由项目维度直接汇总。
    line_rollup = project_df.groupby("line", as_index=False)["half_net_profit"].sum()
    region_line_rollup = project_df.groupby(["region", "line"], as_index=False)["half_net_profit"].sum()
    print_section(
        "half_net_profit_line_rollup_check",
        [summarize_dim_rollup("line_half_net_profit", line_df, line_rollup, ["line"], "half_net_profit")],
    )
    print_section(
        "half_net_profit_region_line_rollup_check",
        [summarize_dim_rollup("region_line_half_net_profit", region_line_df, region_line_rollup, ["region", "line"], "half_net_profit")],
    )

    # 这块用于定位高维表是不是整列都没走项目汇总，便于后续追 ETL 来源。
    print_section("region_line_metric_rollup_summary", summarize_metric_rollups(project_df, region_line_df)[:15])

    interest = region_df[["region", "interest"]].merge(load_interest_region(), on="region", how="left").fillna(0.0)
    print_section(
        "region_interest_check",
        [summarize_diff("region_interest", interest["amount"], interest["interest"])]
        + interest.loc[
            (interest["amount"] - interest["interest"]).abs() > TOLERANCE,
            ["region", "interest", "amount"],
        ].to_dict(orient="records"),
    )

    project_other = project_df[["project_code", "project_name", "other_adjust", "code_norm"]].merge(
        load_other_adjust_project(), on="code_norm", how="left"
    )
    project_other["amount"] = project_other["amount"].fillna(0.0)
    print_section(
        "project_other_adjust_check",
        [summarize_diff("project_other_adjust", project_other["amount"], project_other["other_adjust"])]
        + project_other.loc[
            (project_other["amount"] - project_other["other_adjust"]).abs() > TOLERANCE,
            ["project_code", "project_name", "other_adjust", "amount"],
        ].head(20).to_dict(orient="records"),
    )

    region_other_manual = load_other_adjust_region()
    project_other_rollup = project_other.merge(project_df[["code_norm", "region"]], on="code_norm", how="left")
    project_other_rollup = project_other_rollup.groupby("region", as_index=False)["amount"].sum().rename(columns={"amount": "project_amount"})
    region_other = region_df[["region", "other_adjust"]].merge(project_other_rollup, on="region", how="left").merge(
        region_other_manual, on="region", how="left"
    ).fillna(0.0)
    region_other["calc_project_plus_region_manual"] = region_other["project_amount"] + region_other["amount"]
    region_other["diff"] = region_other["calc_project_plus_region_manual"] - region_other["other_adjust"]
    print_section(
        "region_other_adjust_rollup_check",
        region_other[["region", "other_adjust", "project_amount", "amount", "calc_project_plus_region_manual", "diff"]].assign(
            status=lambda df: df["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    plan_calc = compute_plan_values(load_plan_ledger(query_df))
    project_plan = project_df[["project_code", "project_name", "code_norm", "plan_capital", "plan_smart", "plan_quality"]].merge(
        plan_calc, on="code_norm", how="left", suffixes=("_report", "_calc")
    ).fillna(0.0)
    plan_checks = []
    for component in ["plan_capital", "plan_smart", "plan_quality"]:
        plan_checks.append(
            summarize_diff(
                f"project_{component}",
                project_plan[f"{component}_calc"],
                project_plan[f"{component}_report"],
            )
        )
    print_section("project_plan_checks", plan_checks)

    # 指标清单第 2039 行：类型=项目 + 利润非考核条件 + 项目上的区域，
    # 取项目级【半收付净利润】汇总；“利润非考核条件”以《利润类非考核项目台账》内项目为准。
    non_assess = project_df.assign(is_non_assess=project_df["code_norm"].isin(profit_non_assess_codes))
    non_assess_rollup = (
        non_assess.loc[non_assess["is_non_assess"]]
        .groupby("region", as_index=False)["half_net_profit"]
        .sum()
        .rename(columns={"half_net_profit": "non_assess_calc"})
    )
    region_non_assess = region_df[["region", "non_assess_half_net_profit"]].merge(non_assess_rollup, on="region", how="left").fillna(0.0)
    region_non_assess["diff"] = region_non_assess["non_assess_calc"] - region_non_assess["non_assess_half_net_profit"]
    print_section(
        "region_profit_non_assessment_check",
        [
            {
                "profit_non_assess_ledger_codes": int(len(profit_non_assess_codes)),
                "matched_project_rows": int(project_df["code_norm"].isin(profit_non_assess_codes).sum()),
            }
        ]
        + region_non_assess.assign(
            status=lambda df: df["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    helper_non_assess = project_df.assign(is_non_assess=project_df["code_norm"].isin(non_assess_codes))
    print_section(
        "scope_non_assess_helper_only",
        [
            {
                "base_non_assess_ledger_codes": int(len(non_assess_codes)),
                "matched_project_rows": int(helper_non_assess["is_non_assess"].sum()),
                "note": u(r"\u57fa\u7840\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26\u4e0d\u7528\u4e8e\u672c\u6b21\u5229\u6da6\u975e\u8003\u6838\u5206\u9879\u590d\u7b97\uff0c\u672c\u6b21\u4f7f\u7528\u5229\u6da6\u7c7b\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26"),
            }
        ],
    )

    dev = region_df[["region", "dev_incentive"]].merge(load_non_dev_incentive_region(), on="region", how="left").fillna(0.0)
    print_section(
        "region_non_dev_incentive_check",
        [summarize_diff("region_non_dev_incentive", dev["amount"], dev["dev_incentive"])]
        + dev.loc[
            (dev["amount"] - dev["dev_incentive"]).abs() > TOLERANCE,
            ["region", "dev_incentive", "amount"],
        ].to_dict(orient="records"),
    )

    rollup_cols = [
        "half_net_profit",
        "minority_loss",
        "plan_capital",
        "actual_capital",
        "plan_smart",
        "actual_smart",
        "plan_quality",
        "actual_quality",
        "water_adjust",
        "cutoff_expense",
        "half_attributable_profit",
    ]
    rollup = project_df.groupby("region", as_index=False)[rollup_cols].sum()
    region_rollup = region_df[["region"] + rollup_cols].merge(rollup, on="region", how="left", suffixes=("_region", "_project")).fillna(0.0)
    rollup_records = []
    for component in rollup_cols:
        diff = region_rollup[f"{component}_project"] - region_rollup[f"{component}_region"]
        rollup_records.append(
            {
                "component": component,
                "status": "passed" if (diff.abs() <= TOLERANCE).all() else "failed",
                "mismatch_regions": int((diff.abs() > TOLERANCE).sum()),
                "project_total": float(region_rollup[f"{component}_project"].sum()),
                "region_total": float(region_rollup[f"{component}_region"].sum()),
                "diff_total": float(diff.sum()),
                "max_abs_diff": float(diff.abs().max()),
            }
        )
    print_section("region_from_project_rollup_checks", rollup_records)

    minority = project_df[["project_code", "project_name", "half_net_profit", "minority_loss", "code_norm"]].merge(
        query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"),
        on="code_norm",
        how="left",
    )
    minority["penetration_ratio"] = minority["penetration_ratio"].fillna(1.0)
    minority["simple_calc"] = minority["half_net_profit"] * (1 - minority["penetration_ratio"])
    minority["diff"] = minority["simple_calc"] - minority["minority_loss"]
    print_section(
        "project_minority_simple_helper_check",
        [
            {
                "status": "passed" if (minority["diff"].abs() <= TOLERANCE).all() else "failed",
                "mismatch_rows": int((minority["diff"].abs() > TOLERANCE).sum()),
                "max_abs_diff": float(minority["diff"].abs().max()),
                "note": u(r"\u4ec5\u7528\u9879\u76ee\u67e5\u8be2\u7a7f\u900f\u6bd4\u4f8b\u590d\u7b97\uff0c\u82e5\u5931\u8d25\u9700\u56de\u5230\u5c11\u6570\u80a1\u4e1c\u635f\u5931\u539f\u59cb\u6307\u6807\u884c/\u6e90\u62a5\u8868"),
            }
        ]
        + minority.loc[
            minority["diff"].abs() > TOLERANCE,
            ["project_code", "project_name", "half_net_profit", "penetration_ratio", "minority_loss", "simple_calc", "diff"],
        ].head(20).to_dict(orient="records"),
    )

    print_section(
        "scope_flags",
        [
            {
                "project_rows": int(len(project_df)),
                "region_rows": int(len(region_df)),
                "non_assess_project_rows": int(project_df["code_norm"].isin(non_assess_codes).sum()),
                "profit_non_assess_project_rows": int(project_df["code_norm"].isin(profit_non_assess_codes).sum()),
                "d_exit_by_month_rows": int(project_flags["is_d_exit_by_month"].sum()),
                "d_exit_status_rows": int(project_flags["is_d_exit_status"].sum()),
                "project_only_regions": sorted(set(project_df["region"]) - set(region_df["region"])),
                "region_only_regions": sorted(set(region_df["region"]) - set(project_df["region"])),
            }
        ],
    )


if __name__ == "__main__":
    main()
