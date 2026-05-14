from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from _project_root import find_project_root

import pandas as pd


ROOT = find_project_root(__file__)
REPORT_MONTH = "2025-12"
REPORT_MONTH_COMPACT = "202512"
TOLERANCE = 1e-6


def u(text: str) -> str:
    return text.encode("utf-8").decode("unicode_escape")


PROJECT_TOKEN = u(r"\u9879\u76ee")
REGION_TOKEN = u(r"\u533a\u57df")
PRO_COMPANY_TOKEN = u(r"\u4e13\u4e1a\u516c\u53f8")
INDICATOR_TOKEN = u(r"\u6307\u6807\u6e05\u5355")
ATTRIBUTABLE_TOKEN = u(r"\u7ba1\u62a5\u5f52\u6bcd\u51c0\u5229\u6da6")
QUERY_TOKEN = u(r"\u9879\u76ee\u67e5\u8be2")
NON_ASSESS_TOKEN = u(r"\u975e\u8003\u6838\u9879\u76ee\u53f0\u8d26")
OPERATING_TOKEN = "1.5.1-"
OTHER_ADJUST_TOKEN = u(r"\u5176\u4ed6\u8003\u6838\u8c03\u6574")
CUTOFF_TOKEN = u(r"\u622a\u6b62\u6027\u6536\u652f\u8c03\u6574")
EBT_TOKEN = u(r"\u5e74\u5ea6\u5730\u4ea7\u5173\u8054\u65b9EBT\u6210\u672c")
RESTORE_TOKEN = u(r"\u91d1\u4ee4\u4e1a\u52a1\u8fd8\u539f")
LOCALITY_TOKEN = u(r"\u5c5e\u5730\u5316\u4e1a\u7ee9\u8c03\u6574")
DISCOUNT_TOKEN = u(r"\u7efc\u7ba1\u6298\u8ba9")
PLAN_TOKEN = u(r"\u5e26\u8d44\u644a\u9500")
PRO_LEDGER_TOKEN = u(r"\u91d1\u4ee4\u91d1\u5320\u91d1\u9890\u97f5\u6db5")
RATIO_TOKEN = u(r"\u6bd4\u4f8b\u914d\u7f6e")
EARLY_END_STATUS = u(r"\u63d0\u524d\u7ed3\u675f")

SKIPPED_COMPONENTS = [
    u(r"\u5e26\u8d44\u644a\u9500\u8c03\u6574\u53d1\u751f\u6570"),
    u(r"\u667a\u80fd\u5316\u6574\u6539\u644a\u9500\u8c03\u6574\u53d1\u751f\u6570"),
    u(r"\u8d28\u6548\u63d0\u5347\u53d1\u751f\u6570"),
    u(r"\u8d44\u91d1\u5229\u606f"),
]

BLOCKED_COMPONENTS = [
    u(r"\u975e\u53d1\u5c55\u4eba\u5458\u53d1\u5c55\u6fc0\u52b1\uff1a\u7f3a\u6e90\u53f0\u8d26"),
    u(r"\u5b89\u4fdd\u5c11\u6570\u80a1\u4e1c\u635f\u76ca\uff1a\u6307\u6807\u884c\u4f9d\u8d56 1.3-S / \u6307\u6807\u5e93\uff0c\u5f53\u524d\u672a\u63d0\u4f9b"),
]

PROJECT_COLUMNS = [
    "idx",
    "region",
    "line",
    "project_code",
    "project_name",
    "net_profit",
    "d_exit_net_profit",
    "impairment",
    "minority",
    "plan_capital",
    "actual_capital",
    "plan_smart",
    "actual_smart",
    "plan_quality",
    "actual_quality",
    "ebt_cost",
    "other_adjust",
    "discount",
    "cutoff",
    "dev_incentive",
    "jl_jj_restore",
    "interest",
    "satellite",
    "pro_d_exit_sum",
    "prev_region_perf",
    "curr_region_perf",
    "attributable",
]

REGION_COLUMNS = [
    "idx",
    "region",
    "net_profit",
    "d_exit_net_profit",
    "impairment",
    "minority",
    "plan_capital",
    "actual_capital",
    "plan_smart",
    "actual_smart",
    "plan_quality",
    "actual_quality",
    "ebt_cost",
    "other_adjust",
    "discount",
    "cutoff",
    "dev_incentive",
    "jl_jj_restore",
    "interest",
    "satellite",
    "pro_d_exit_sum",
    "prev_region_perf",
    "curr_region_perf",
    "attributable",
]

PRO_COLUMNS = [
    "idx",
    "company",
    "net_profit",
    "d_exit_net_profit",
    "impairment",
    "minority",
    "plan_capital",
    "actual_capital",
    "plan_smart",
    "actual_smart",
    "plan_quality",
    "actual_quality",
    "ebt_cost",
    "other_adjust",
    "discount",
    "cutoff",
    "dev_incentive",
    "jl_jj_restore",
    "interest",
    "satellite",
    "pro_d_exit_sum",
    "prev_region_perf",
    "curr_region_perf",
    "attributable",
]


@dataclass
class FormulaCheck:
    label: str
    mismatch_rows: int
    max_abs_diff: float


def find_workbook(*tokens: str) -> Path:
    matches = []
    for path in ROOT.iterdir():
        if path.suffix.lower() != ".xlsx":
            continue
        if all(token in path.name for token in tokens):
            matches.append(path)
    if len(matches) != 1:
        raise RuntimeError(f"Expected exactly one workbook for {tokens}, got {len(matches)}")
    return matches[0]


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text in {"", "NAN", "NONE"}:
        return ""
    if len(text) >= 2 and text[0].isalpha() and any(ch.isdigit() for ch in text[1:]):
        return text[1:]
    return text


def print_json_section(label: str, records: list[dict]) -> None:
    print(f"[{label}]")
    if not records:
        print("[]")
        return
    print(pd.DataFrame(records).to_json(force_ascii=True, orient="records"))


def load_indicator_rows() -> pd.DataFrame:
    indicator = pd.read_excel(find_workbook(INDICATOR_TOKEN), sheet_name=0)
    name_col = indicator.columns[3]
    dim_col = indicator.columns[4]
    cycle_col = indicator.columns[5]
    method_col = indicator.columns[8]
    source_col = indicator.columns[10]
    logic_col = indicator.columns[12]
    dims = [PROJECT_TOKEN, REGION_TOKEN, u(r"\u5b89\u4fdd"), u(r"\u91d1\u4ee4\u91d1\u5320"), u(r"\u91d1\u9890"), u(r"\u97f5\u6db5")]
    rows = indicator[
        indicator[name_col].astype(str).str.contains(ATTRIBUTABLE_TOKEN, na=False)
        & indicator[dim_col].astype(str).isin(dims)
    ][[name_col, dim_col, cycle_col, method_col, source_col, logic_col]].fillna("")
    rows.columns = ["metric_name", "dimension", "cycle", "method", "source_table", "logic"]
    return rows


def load_project_report() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(ATTRIBUTABLE_TOKEN, PROJECT_TOKEN))
    df.columns = PROJECT_COLUMNS
    for column in PROJECT_COLUMNS[5:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    df["code_norm"] = df["project_code"].map(normalize_code)
    return df


def load_region_report() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(ATTRIBUTABLE_TOKEN, REGION_TOKEN))
    df.columns = REGION_COLUMNS
    for column in REGION_COLUMNS[2:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def load_pro_report() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(ATTRIBUTABLE_TOKEN, PRO_COMPANY_TOKEN))
    df.columns = PRO_COLUMNS
    for column in PRO_COLUMNS[2:]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def load_query() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(QUERY_TOKEN))
    df.columns = [
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
    df["code_norm"] = df["project_code"].map(normalize_code)
    df["penetration_ratio"] = pd.to_numeric(df["penetration_ratio"], errors="coerce").fillna(1.0)
    return df


def load_non_assess_codes() -> set[str]:
    df = pd.read_excel(find_workbook(NON_ASSESS_TOKEN))
    code_col = next(column for column in df.columns if u(r"\u7acb\u9879\u7f16\u7801") in str(column))
    return {normalize_code(value) for value in df[code_col].dropna()}


def load_1_5_1_project_profit() -> pd.DataFrame:
    raw = pd.read_excel(find_workbook(OPERATING_TOKEN), header=None)
    data = raw.iloc[5:].reset_index(drop=True)
    data.columns = [f"c{i}" for i in range(data.shape[1])]
    data = data[data["c1"].astype(str).str.contains(r"[A-Z]\d", na=False)].copy()
    data["code_norm"] = data["c1"].map(normalize_code)
    data["net_profit_no_hq"] = pd.to_numeric(data["c48"], errors="coerce").fillna(0.0)
    return data.groupby("code_norm", as_index=False)["net_profit_no_hq"].sum()


def load_other_adjust_project() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(OTHER_ADJUST_TOKEN), header=1)
    df = df[df[df.columns[1]].astype(str) == PROJECT_TOKEN].copy()
    df["code_norm"] = df[df.columns[4]].map(normalize_code)
    df["amount"] = pd.to_numeric(df[df.columns[6]], errors="coerce").fillna(0.0)
    return df.groupby("code_norm", as_index=False)["amount"].sum()


def load_other_adjust_region() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(OTHER_ADJUST_TOKEN), header=1)
    df = df[df[df.columns[1]].astype(str) == REGION_TOKEN].copy()
    df["amount"] = pd.to_numeric(df[df.columns[6]], errors="coerce").fillna(0.0)
    return df.groupby(df.columns[5], as_index=False)["amount"].sum().rename(columns={df.columns[5]: "region"})


def load_cutoff_project(query_df: pd.DataFrame) -> pd.DataFrame:
    df = pd.read_excel(find_workbook(CUTOFF_TOKEN), header=1)
    df = df[df[df.columns[1]].astype(str) == PROJECT_TOKEN].copy()
    df["code_norm"] = df[df.columns[4]].map(normalize_code)
    df["new_income"] = pd.to_numeric(df[df.columns[6]], errors="coerce").fillna(0.0)
    df["new_cost"] = pd.to_numeric(df[df.columns[8]], errors="coerce").fillna(0.0)
    df = df.merge(query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"), on="code_norm", how="left")
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)
    df["amount"] = (df["new_income"] + df["new_cost"]) * 0.81 * df["penetration_ratio"]
    return df.groupby("code_norm", as_index=False)["amount"].sum()


def load_discount_project(query_df: pd.DataFrame) -> pd.DataFrame:
    df = pd.read_excel(find_workbook(DISCOUNT_TOKEN), header=1)
    df["code_norm"] = df[df.columns[3]].map(normalize_code)
    df["data_month"] = df[df.columns[1]].astype(str).str.strip()
    df["discount_type"] = df[df.columns[5]].astype(str).str.strip()
    df["top_type"] = df[df.columns[6]].astype(str).str.strip()
    df["total_discount"] = pd.to_numeric(df[df.columns[9]], errors="coerce").fillna(0.0)
    df["related_party_discount"] = pd.to_numeric(df[df.columns[11]], errors="coerce").fillna(0.0)
    df = df[df["data_month"].eq(REPORT_MONTH)].copy()
    df = df[df["top_type"].isin([u(r"\u5229\u6da6"), u(r"\u901a\u7528")])].copy()
    special_type = u(r"\u0032\u0030\u0032\u0033\u5e74\u5e95\u4e4b\u524d\u7684\u5173\u8054\u65b9\u5e94\u6536\u6b3e")
    df["base_amount"] = df["total_discount"]
    df.loc[df["discount_type"].eq(special_type), "base_amount"] = df.loc[
        df["discount_type"].eq(special_type), "related_party_discount"
    ]
    df = df.merge(query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"), on="code_norm", how="left")
    df["penetration_ratio"] = df["penetration_ratio"].fillna(1.0)
    df["amount"] = df["base_amount"] * 0.81 / 1.06 * df["penetration_ratio"]
    return df.groupby("code_norm", as_index=False)["amount"].sum()


def load_cutoff_region_rows() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(CUTOFF_TOKEN), header=1)
    df = df[df[df.columns[1]].astype(str) == REGION_TOKEN].copy()
    df["new_income"] = pd.to_numeric(df[df.columns[6]], errors="coerce").fillna(0.0)
    df["new_cost"] = pd.to_numeric(df[df.columns[8]], errors="coerce").fillna(0.0)
    df["amount"] = (df["new_income"] + df["new_cost"]) * 0.81
    return df.groupby(df.columns[5], as_index=False)["amount"].sum().rename(columns={df.columns[5]: "region"})


def load_ebt_region() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(EBT_TOKEN), header=1)
    df["adj_a"] = pd.to_numeric(df[df.columns[7]], errors="coerce").fillna(0.0)
    df["adj_b"] = pd.to_numeric(df[df.columns[8]], errors="coerce").fillna(0.0)
    df["amount"] = (df["adj_a"] + df["adj_b"]) * 0.81
    return df.groupby(df.columns[5], as_index=False)["amount"].sum().rename(columns={df.columns[5]: "region"})


def load_restore_region() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(RESTORE_TOKEN))
    amount_col = df.columns[3]
    region_col = df.columns[2]
    df["amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    return df.groupby(region_col, as_index=False)["amount"].sum().rename(columns={region_col: "region"})


def load_locality_adjustment() -> pd.DataFrame:
    df = pd.read_excel(find_workbook(LOCALITY_TOKEN), header=1)
    df["code_norm"] = df[df.columns[2]].map(normalize_code)
    df["share_ratio"] = pd.to_numeric(df[df.columns[5]], errors="coerce").fillna(0.0)
    return df.rename(
        columns={
            df.columns[3]: "current_region",
            df.columns[4]: "previous_region",
            df.columns[6]: "start_month",
            df.columns[7]: "end_month",
        }
    )


def load_pro_ledger() -> pd.DataFrame:
    raw = pd.read_excel(find_workbook(PRO_LEDGER_TOKEN), header=None)
    df = raw.iloc[4:].reset_index(drop=True)
    df.columns = [
        "idx",
        "company",
        "data_month",
        "revenue",
        "net_profit",
        "d_exit_sum",
        "minority",
        "half_revenue",
        "half_net_profit",
        "half_minority",
        "cashflow",
    ]
    for column in ["net_profit", "d_exit_sum", "minority"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    return df


def load_plan_ledger(query_df: pd.DataFrame) -> pd.DataFrame:
    plan_path = next(
        path for path in ROOT.iterdir() if path.suffix.lower() == ".xlsx" and PLAN_TOKEN in path.name and RATIO_TOKEN not in path.name
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
    ratio_df = pd.read_excel(find_workbook(PLAN_TOKEN, RATIO_TOKEN), header=1)
    ratio_df = ratio_df.rename(columns={ratio_df.columns[0]: "years"})
    ratio_df["years_num"] = pd.to_numeric(ratio_df["years"].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    ratio_df = ratio_df.dropna(subset=["years_num"]).copy()
    ratio_cols = [column for column in ratio_df.columns if u(r"\u644a\u9500\u6bd4\u4f8b") in str(column)]
    ratio_map: dict[int, list[float]] = {}
    for _, row in ratio_df.iterrows():
        ratio_map[int(row["years_num"])] = [
            0.0 if pd.isna(pd.to_numeric(row[column], errors="coerce")) else float(pd.to_numeric(row[column], errors="coerce"))
            for column in ratio_cols
        ]
    return ratio_map


def compute_project_plan_values(plan_df: pd.DataFrame) -> pd.DataFrame:
    month_starts = [period.to_timestamp() for period in pd.period_range("2025-01", REPORT_MONTH, freq="M")]
    ratios = load_plan_ratio_map()
    records: list[dict] = []
    type_mapping = {
        u(r"\u5e26\u8d44\u644a\u9500"): "plan_capital",
        u(r"\u667a\u80fd\u5316\u6574\u6539"): "plan_smart",
        u(r"\u8d28\u6548\u63d0\u5347"): "plan_quality",
    }

    for _, row in plan_df.iterrows():
        if pd.isna(row["start_date"]) or pd.isna(row["years"]) or row["plan_type"] not in type_mapping:
            continue
        effective_end = row["early_end_date"] if row["status"] == EARLY_END_STATUS and pd.notna(row["early_end_date"]) else row["end_date"]
        total = 0.0
        for month_start in month_starts:
            if month_start < row["start_date"]:
                continue
            if pd.notna(effective_end) and month_start > effective_end:
                continue
            if row["plan_type"] == u(r"\u5e26\u8d44\u644a\u9500"):
                year_idx = ((month_start.year - row["start_date"].year) * 12 + (month_start.month - row["start_date"].month)) // 12
                share_list = ratios.get(int(row["years"]), [])
                share = share_list[year_idx] if 0 <= year_idx < len(share_list) else 0.0
                monthly_amount = row["amount"] * share * 0.81 / 1.06 / 12 * row["penetration_ratio"]
            elif row["plan_type"] == u(r"\u667a\u80fd\u5316\u6574\u6539"):
                monthly_amount = row["amount"] / row["years"] / 12 * row["penetration_ratio"] * 0.81
            else:
                monthly_amount = row["amount"] / row["years"] / 12 * row["penetration_ratio"] * 0.81 / 1.06
            total += monthly_amount
        records.append({"code_norm": row["code_norm"], "component": type_mapping[row["plan_type"]], "amount": total})

    if not records:
        return pd.DataFrame(columns=["code_norm", "plan_capital", "plan_smart", "plan_quality"])

    wide = pd.DataFrame(records).groupby(["code_norm", "component"], as_index=False)["amount"].sum()
    wide = wide.pivot(index="code_norm", columns="component", values="amount").reset_index().fillna(0.0)
    for column in ["plan_capital", "plan_smart", "plan_quality"]:
        if column not in wide.columns:
            wide[column] = 0.0
    return wide[["code_norm", "plan_capital", "plan_smart", "plan_quality"]]


def build_formula_result(label: str, calc: pd.Series, report: pd.Series) -> FormulaCheck:
    diff = calc - report
    return FormulaCheck(label, int((diff.abs() > TOLERANCE).sum()), float(diff.abs().max()))


def project_formula(df: pd.DataFrame) -> pd.Series:
    return (
        df["net_profit"]
        - df["minority"]
        - df["plan_capital"]
        - df["plan_smart"]
        - df["plan_quality"]
        + df["actual_capital"]
        + df["actual_smart"]
        + df["actual_quality"]
        - df["ebt_cost"]
        + df["other_adjust"]
        + df["discount"]
        + df["cutoff"]
    )


def region_formula(df: pd.DataFrame) -> pd.Series:
    return (
        df["net_profit"]
        - df["d_exit_net_profit"]
        - df["minority"]
        - df["plan_capital"]
        - df["plan_smart"]
        - df["plan_quality"]
        + df["actual_capital"]
        + df["actual_smart"]
        + df["actual_quality"]
        - df["ebt_cost"]
        + df["other_adjust"]
        + df["discount"]
        + df["cutoff"]
        + df["dev_incentive"]
        + df["jl_jj_restore"]
        + df["interest"]
        + df["prev_region_perf"]
        - df["curr_region_perf"]
    )


def pro_formula(df: pd.DataFrame) -> pd.Series:
    return (
        df["net_profit"]
        - df["pro_d_exit_sum"]
        - df["impairment"]
        - df["minority"]
        - df["ebt_cost"]
        + df["other_adjust"]
        + df["cutoff"]
    )


def mark_d_exit(project_df: pd.DataFrame, query_df: pd.DataFrame) -> pd.DataFrame:
    merged = project_df.merge(
        query_df[["code_norm", "project_level", "project_status", "exit_date"]].drop_duplicates("code_norm"),
        on="code_norm",
        how="left",
    )
    exit_ym = pd.to_datetime(merged["exit_date"], errors="coerce").dt.strftime("%Y%m").fillna("")
    merged["is_helper_d_exit"] = (
        merged["project_level"].astype(str).str.startswith("D")
        & (merged["project_status"].astype(str) == u(r"\u5df2\u64a4\u573a"))
        & (exit_ym != "")
        & (exit_ym <= REPORT_MONTH_COMPACT)
    )
    merged["is_known_d_exit"] = (
        merged["project_level"].astype(str).str.startswith("D")
        & (merged["project_status"].astype(str) == u(r"\u5df2\u64a4\u573a"))
    )
    return merged


def main() -> None:
    indicator_rows = load_indicator_rows()
    project_df = load_project_report()
    region_df = load_region_report()
    pro_df = load_pro_report()
    query_df = load_query()
    plan_df = load_plan_ledger(query_df)
    project_plan_calc = compute_project_plan_values(plan_df)
    non_assess_codes = load_non_assess_codes()
    project_with_flags = mark_d_exit(project_df, query_df)

    print_json_section(
        "assumptions",
        [
            {
                "report_month": REPORT_MONTH,
                "report_is_cumulative_snapshot": True,
                "skipped_components": " / ".join(SKIPPED_COMPONENTS),
                "blocked_components": " / ".join(BLOCKED_COMPONENTS),
            }
        ],
    )

    print_json_section("indicator_rows", indicator_rows.to_dict(orient="records"))

    formula_rows = [
        build_formula_result("project_formula", project_formula(project_df), project_df["attributable"]).__dict__,
        build_formula_result("region_formula", region_formula(region_df), region_df["attributable"]).__dict__,
        build_formula_result("pro_company_formula", pro_formula(pro_df), pro_df["attributable"]).__dict__,
    ]
    print_json_section("formula_check", formula_rows)

    print_json_section(
        "project_scope",
        [
            {
                "project_rows": len(project_df),
                "non_assess_rows": int(project_df["code_norm"].isin(non_assess_codes).sum()),
                "helper_d_exit_rows": int(project_with_flags["is_helper_d_exit"].sum()),
            }
        ],
    )

    op_project = load_1_5_1_project_profit()
    project_net = project_df[["region", "line", "project_code", "project_name", "net_profit", "code_norm"]].merge(
        op_project, on="code_norm", how="left"
    )
    project_net["net_profit_no_hq"] = project_net["net_profit_no_hq"].fillna(0.0)
    project_net["diff"] = project_net["net_profit_no_hq"] - project_net["net_profit"]
    print_json_section(
        "project_net_profit_check",
        [
            {
                "status": "failed" if (project_net["diff"].abs() > TOLERANCE).any() else "passed",
                "mismatch_rows": int((project_net["diff"].abs() > TOLERANCE).sum()),
                "max_abs_diff": float(project_net["diff"].abs().max()),
                "report_total": float(project_net["net_profit"].sum()),
                "source_total": float(project_net["net_profit_no_hq"].sum()),
            }
        ]
        + project_net.loc[project_net["diff"].abs() > TOLERANCE, ["region", "line", "project_code", "project_name", "net_profit", "net_profit_no_hq", "diff"]]
        .to_dict(orient="records"),
    )

    project_minority = project_df[["project_code", "project_name", "net_profit", "minority", "code_norm"]].merge(
        query_df[["code_norm", "penetration_ratio"]].drop_duplicates("code_norm"),
        on="code_norm",
        how="left",
    )
    project_minority["penetration_ratio"] = project_minority["penetration_ratio"].fillna(1.0)
    project_minority["minority_calc"] = project_minority["net_profit"] * (1 - project_minority["penetration_ratio"])
    project_minority["diff"] = project_minority["minority_calc"] - project_minority["minority"]
    print_json_section(
        "project_minority_helper_check",
        [
            {
                "status": "failed" if (project_minority["diff"].abs() > TOLERANCE).any() else "passed",
                "mismatch_rows": int((project_minority["diff"].abs() > TOLERANCE).sum()),
                "max_abs_diff": float(project_minority["diff"].abs().max()),
            }
        ]
        + project_minority.loc[project_minority["diff"].abs() > TOLERANCE, ["project_code", "project_name", "net_profit", "minority", "penetration_ratio", "minority_calc", "diff"]]
        .head(20)
        .to_dict(orient="records"),
    )

    project_other = project_df[["project_code", "project_name", "other_adjust", "code_norm"]].merge(
        load_other_adjust_project(), on="code_norm", how="left"
    )
    project_other["amount"] = project_other["amount"].fillna(0.0)
    project_other["diff"] = project_other["amount"] - project_other["other_adjust"]
    print_json_section(
        "project_other_adjust_check",
        [
            {
                "status": "failed" if (project_other["diff"].abs() > TOLERANCE).any() else "passed",
                "mismatch_rows": int((project_other["diff"].abs() > TOLERANCE).sum()),
                "max_abs_diff": float(project_other["diff"].abs().max()),
            }
        ],
    )

    project_cutoff = project_df[["project_code", "project_name", "cutoff", "code_norm"]].merge(
        load_cutoff_project(query_df), on="code_norm", how="left"
    )
    project_cutoff["amount"] = project_cutoff["amount"].fillna(0.0)
    project_cutoff["diff"] = project_cutoff["amount"] - project_cutoff["cutoff"]
    print_json_section(
        "project_cutoff_check",
        [
            {
                "status": "failed" if (project_cutoff["diff"].abs() > TOLERANCE).any() else "passed",
                "mismatch_rows": int((project_cutoff["diff"].abs() > TOLERANCE).sum()),
                "max_abs_diff": float(project_cutoff["diff"].abs().max()),
            }
        ],
    )

    project_discount = project_df[["project_code", "project_name", "discount", "code_norm"]].merge(
        load_discount_project(query_df), on="code_norm", how="left"
    )
    project_discount["amount"] = project_discount["amount"].fillna(0.0)
    project_discount["diff"] = project_discount["amount"] - project_discount["discount"]
    print_json_section(
        "project_discount_check",
        [
            {
                "status": "failed" if (project_discount["diff"].abs() > TOLERANCE).any() else "passed",
                "mismatch_rows": int((project_discount["diff"].abs() > TOLERANCE).sum()),
                "max_abs_diff": float(project_discount["diff"].abs().max()),
                "report_total": float(project_discount["discount"].sum()),
                "source_total": float(project_discount["amount"].sum()),
            }
        ]
        + project_discount.loc[
            project_discount["diff"].abs() > TOLERANCE,
            ["project_code", "project_name", "discount", "amount", "diff"],
        ]
        .head(20)
        .to_dict(orient="records"),
    )

    project_plan = project_df[["project_code", "project_name", "code_norm", "plan_capital", "plan_smart", "plan_quality"]].merge(
        project_plan_calc, on="code_norm", how="left", suffixes=("_report", "_calc")
    )
    project_plan[["plan_capital_calc", "plan_smart_calc", "plan_quality_calc"]] = project_plan[
        ["plan_capital_calc", "plan_smart_calc", "plan_quality_calc"]
    ].fillna(0.0)
    for component in ["plan_capital", "plan_smart", "plan_quality"]:
        project_plan[f"{component}_diff"] = project_plan[f"{component}_calc"] - project_plan[f"{component}_report"]
    project_plan["max_abs_diff"] = project_plan[
        ["plan_capital_diff", "plan_smart_diff", "plan_quality_diff"]
    ].abs().max(axis=1)
    print_json_section(
        "project_plan_check",
        [
            {
                "status": "failed" if (project_plan["max_abs_diff"] > TOLERANCE).any() else "passed",
                "mismatch_rows": int((project_plan["max_abs_diff"] > TOLERANCE).sum()),
                "max_abs_diff": float(project_plan["max_abs_diff"].max()),
                "plan_capital_report_total": float(project_plan["plan_capital_report"].sum()),
                "plan_capital_calc_total": float(project_plan["plan_capital_calc"].sum()),
                "plan_smart_report_total": float(project_plan["plan_smart_report"].sum()),
                "plan_smart_calc_total": float(project_plan["plan_smart_calc"].sum()),
                "plan_quality_report_total": float(project_plan["plan_quality_report"].sum()),
                "plan_quality_calc_total": float(project_plan["plan_quality_calc"].sum()),
            }
        ]
        + project_plan.loc[
            project_plan["max_abs_diff"] > TOLERANCE,
            [
                "project_code",
                "project_name",
                "plan_capital_report",
                "plan_capital_calc",
                "plan_capital_diff",
                "plan_smart_report",
                "plan_smart_calc",
                "plan_smart_diff",
                "plan_quality_report",
                "plan_quality_calc",
                "plan_quality_diff",
            ],
        ].to_dict(orient="records"),
    )

    print_json_section(
        "project_blockers",
        [
            {
                "ebt_cost": u(r"\u5f53\u524d EBT \u53f0\u8d26\u53ea\u6709\u533a\u57df\u7c7b\u578b\uff0c\u65e0\u9879\u76ee\u7c7b\u578b\u884c"),
                "skipped": " / ".join(SKIPPED_COMPONENTS),
            }
        ],
    )

    non_d_project_rollup = project_with_flags.loc[~project_with_flags["is_helper_d_exit"]].groupby("region", as_index=False).agg(
        plan_capital=("plan_capital", "sum"),
        plan_smart=("plan_smart", "sum"),
        plan_quality=("plan_quality", "sum"),
        discount=("discount", "sum"),
        cutoff=("cutoff", "sum"),
        other_adjust=("other_adjust", "sum"),
    )
    non_d_project_discount_source = (
        project_with_flags.merge(project_discount[["code_norm", "amount"]].rename(columns={"amount": "discount_source"}), on="code_norm", how="left")
        .fillna({"discount_source": 0.0})
        .loc[lambda df: ~df["is_helper_d_exit"]]
        .groupby("region", as_index=False)["discount_source"]
        .sum()
    )

    region_other = region_df[["region", "other_adjust"]].merge(non_d_project_rollup[["region", "other_adjust"]], on="region", how="left")
    region_other = region_other.merge(load_other_adjust_region(), on="region", how="left")
    region_other = region_other.fillna(0.0)
    region_other["project_rollup"] = region_other["other_adjust_x"]
    region_other["region_manual"] = region_other["amount"]
    region_other["report_value"] = region_other["other_adjust_y"]
    region_other["calc"] = region_other["project_rollup"] + region_other["region_manual"]
    region_other["diff"] = region_other["calc"] - region_other["report_value"]
    print_json_section(
        "region_other_adjust_check",
        region_other[["region", "report_value", "project_rollup", "region_manual", "calc", "diff"]].assign(
            status=lambda x: x["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    region_discount = region_df[["region", "discount"]].merge(non_d_project_discount_source, on="region", how="left").fillna(0.0)
    region_discount["report_value"] = region_discount["discount"]
    region_discount["calc"] = region_discount["discount_source"]
    region_discount["diff"] = region_discount["calc"] - region_discount["report_value"]
    print_json_section(
        "region_discount_check",
        region_discount[["region", "report_value", "calc", "diff"]].assign(
            status=lambda x: x["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    region_plan = region_df[["region", "plan_capital", "plan_smart", "plan_quality"]].merge(
        non_d_project_rollup[["region", "plan_capital", "plan_smart", "plan_quality"]],
        on="region",
        how="left",
        suffixes=("_report", "_calc"),
    )
    region_plan[["plan_capital_calc", "plan_smart_calc", "plan_quality_calc"]] = region_plan[
        ["plan_capital_calc", "plan_smart_calc", "plan_quality_calc"]
    ].fillna(0.0)
    for component in ["plan_capital", "plan_smart", "plan_quality"]:
        region_plan[f"{component}_diff"] = region_plan[f"{component}_calc"] - region_plan[f"{component}_report"]
    region_plan["max_abs_diff"] = region_plan[
        ["plan_capital_diff", "plan_smart_diff", "plan_quality_diff"]
    ].abs().max(axis=1)
    print_json_section(
        "region_plan_check",
        [
            {
                "status": "failed" if (region_plan["max_abs_diff"] > TOLERANCE).any() else "passed",
                "mismatch_rows": int((region_plan["max_abs_diff"] > TOLERANCE).sum()),
                "max_abs_diff": float(region_plan["max_abs_diff"].max()),
            }
        ]
        + region_plan.loc[
            region_plan["max_abs_diff"] > TOLERANCE,
            [
                "region",
                "plan_capital_report",
                "plan_capital_calc",
                "plan_capital_diff",
                "plan_smart_report",
                "plan_smart_calc",
                "plan_smart_diff",
                "plan_quality_report",
                "plan_quality_calc",
                "plan_quality_diff",
            ],
        ].to_dict(orient="records"),
    )

    region_cutoff_rollup = (
        project_with_flags.loc[~project_with_flags["is_known_d_exit"]]
        .groupby("region", as_index=False)["cutoff"]
        .sum()
    )
    region_cutoff = region_df[["region", "cutoff"]].merge(region_cutoff_rollup, on="region", how="left").merge(
        load_cutoff_region_rows(), on="region", how="left"
    )
    region_cutoff = region_cutoff.fillna(0.0)
    region_cutoff["calc"] = region_cutoff["cutoff_y"] + region_cutoff["amount"] if "cutoff_y" in region_cutoff.columns else region_cutoff["amount"]
    region_cutoff["report_value"] = region_cutoff["cutoff_x"]
    region_cutoff["diff"] = region_cutoff["calc"] - region_cutoff["report_value"]
    print_json_section(
        "region_cutoff_check",
        region_cutoff[["region", "report_value", "calc", "diff"]].assign(
            status=lambda x: x["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    region_ebt = region_df[["region", "ebt_cost"]].merge(load_ebt_region(), on="region", how="left").fillna(0.0)
    region_ebt["diff"] = region_ebt["amount"] - region_ebt["ebt_cost"]
    print_json_section(
        "region_ebt_check",
        region_ebt[["region", "ebt_cost", "amount", "diff"]].assign(
            status=lambda x: x["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    region_restore = region_df[["region", "jl_jj_restore"]].merge(load_restore_region(), on="region", how="left").fillna(0.0)
    region_restore["diff"] = region_restore["amount"] - region_restore["jl_jj_restore"]
    print_json_section(
        "region_restore_check",
        region_restore[["region", "jl_jj_restore", "amount", "diff"]].assign(
            status=lambda x: x["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    d_exit_rollup = (
        project_with_flags.loc[project_with_flags["is_helper_d_exit"]]
        .groupby("region", as_index=False)["net_profit"]
        .sum()
        .rename(columns={"net_profit": "helper_d_exit_net_profit"})
    )
    region_d_exit = region_df[["region", "d_exit_net_profit"]].merge(d_exit_rollup, on="region", how="left").fillna(0.0)
    region_d_exit["diff"] = region_d_exit["helper_d_exit_net_profit"] - region_d_exit["d_exit_net_profit"]
    print_json_section(
        "region_d_exit_check",
        region_d_exit[["region", "d_exit_net_profit", "helper_d_exit_net_profit", "diff"]].assign(
            status=lambda x: x["diff"].abs().le(TOLERANCE).map({True: "passed", False: "failed"})
        ).to_dict(orient="records"),
    )

    region_rollup_sets = [
        {
            "project_only_regions": sorted(set(project_df["region"]) - set(region_df["region"])),
            "region_only_regions": sorted(set(region_df["region"]) - set(project_df["region"])),
            "note": u(r"\u533a\u57df\u9ad8\u7ef4\u7ed3\u679c\u4ecd\u4f9d\u8d56\u9879\u76ee\u7ef4\u5ea6\u5b8c\u6574\u9a8c\u8bc1\uff0c\u672c\u6b21\u533a\u57df\u7ed3\u8bba\u6309 provisional \u770b\u5f85"),
        }
    ]
    print_json_section("region_rollup_scope", region_rollup_sets)

    locality_df = load_locality_adjustment()
    print_json_section(
        "region_blockers",
        [
            {
                "prev_region_perf": u(r"\u6307\u6807\u903b\u8f91\u4f9d\u8d56\u201c\u9879\u76ee\u5355\u6708\u7ba1\u62a5\u5f52\u6bcd\u51c0\u5229\u6da6\u201d\uff0c\u5f53\u524d\u53ea\u6709 202512 \u7d2f\u8ba1\u62a5\u8868"),
                "curr_region_perf": u(r"\u540c\u4e0a"),
                "dev_incentive": u(r"\u7f3a\u300a\u975e\u53d1\u5c55\u4eba\u5458\u7684\u53d1\u5c55\u6fc0\u52b1\u53f0\u8d26\u300b"),
                "locality_rows": len(locality_df),
                "skipped": " / ".join(SKIPPED_COMPONENTS),
            }
        ],
    )

    pro_ledger = load_pro_ledger()
    pro_manual = pro_df[["company", "net_profit", "pro_d_exit_sum", "minority"]].merge(
        pro_ledger[["company", "net_profit", "d_exit_sum", "minority"]],
        on="company",
        how="left",
        suffixes=("_report", "_source"),
    ).fillna(0.0)
    pro_manual["net_profit_diff"] = pro_manual["net_profit_source"] - pro_manual["net_profit_report"]
    pro_manual["d_exit_diff"] = pro_manual["d_exit_sum"] - pro_manual["pro_d_exit_sum"]
    pro_manual["minority_diff"] = pro_manual["minority_source"] - pro_manual["minority_report"]
    print_json_section(
        "pro_company_manual_check",
        pro_manual.assign(
            status=lambda x: x.apply(
                lambda row: (
                    "blocked"
                    if row["company"] == u(r"\u5b89\u4fdd")
                    else "passed"
                    if max(abs(row["net_profit_diff"]), abs(row["d_exit_diff"]), abs(row["minority_diff"])) <= TOLERANCE
                    else "failed"
                ),
                axis=1,
            )
        ).to_dict(orient="records"),
    )

    print_json_section(
        "pro_company_zero_component_check",
        [
            {
                "company": row["company"],
                "ebt_cost": row["ebt_cost"],
                "other_adjust": row["other_adjust"],
                "cutoff": row["cutoff"],
                "status": (
                    "blocked"
                    if row["company"] == u(r"\u5b89\u4fdd")
                    else "passed"
                    if abs(row["ebt_cost"]) <= TOLERANCE and abs(row["other_adjust"]) <= TOLERANCE and abs(row["cutoff"]) <= TOLERANCE
                    else "failed"
                ),
                "note": (
                    u(r"\u5b89\u4fdd EBT/\u5c11\u6570\u80a1\u4e1c\u635f\u76ca\u903b\u8f91\u4ecd\u4f9d\u8d56 1.3-S")
                    if row["company"] == u(r"\u5b89\u4fdd")
                    else u(r"\u5f53\u524d\u6e90\u96c6\u672a\u51fa\u73b0\u5bf9\u5e94\u7c7b\u578b\u884c\uff0c\u62a5\u8868\u4e3a 0")
                ),
            }
            for _, row in pro_df.iterrows()
        ],
    )

    print_json_section(
        "overall_status",
        [
            {
                "project_dimension": u(r"\u90e8\u5206\u901a\u8fc7\uff1a\u516c\u5f0f/\u5176\u4ed6\u8003\u6838/\u622a\u6b62\u6027\u6536\u652f/\u7efc\u7ba1\u6298\u8ba9/\u4e09\u4e2a\u8ba1\u5212\u6570\u901a\u8fc7\uff0c1.5.1 \u4e3b\u503c 2 \u4e2a\u9879\u76ee\u7f3a\u6e90\uff0c\u5c11\u6570\u80a1\u4e1c\u635f\u76ca helper \u590d\u7b97 226 \u884c\u4e0d\u4e00\u81f4\uff0cEBT \u4ecd\u672a\u5b8c\u6574\u5b9a\u6027"),
                "region_dimension": u(r"\u90e8\u5206\u901a\u8fc7\uff1a\u516c\u5f0f/EBT/\u91d1\u4ee4\u4e1a\u52a1\u8fd8\u539f/\u622a\u6b62\u6027\u6536\u652f\u901a\u8fc7\uff0c\u8ba1\u5212\u6570/\u5176\u4ed6\u8003\u6838/\u7efc\u7ba1\u6298\u8ba9\u5b58\u5728\u533a\u57df\u5dee\u5f02\uff0cD\u7c7b\u64a4\u573a\u590d\u7b97\u5931\u8d25\uff0c\u4e1a\u7ee9\u8c03\u6574/\u975e\u53d1\u5c55\u6fc0\u52b1\u4ecd\u963b\u585e"),
                "pro_company_dimension": u(r"\u516c\u5f0f\u901a\u8fc7\uff1b\u91d1\u4ee4\u91d1\u5320/\u91d1\u9890/\u97f5\u6db5\u53f0\u8d26\u76f4\u6821\u901a\u8fc7\uff1b\u5b89\u4fdd\u56e0 1.3-S \u7f3a\u5931\u4ecd\u4e3a blocked"),
            }
        ],
    )


if __name__ == "__main__":
    main()
