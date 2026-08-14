import pandas as pd
from pathlib import Path
import datetime as dt
from Table_Preparation import prepare_data
from Data_Processing import process_data
from Storage import storage
import numpy as np

# ===============================
# 🔧 USER CONFIG — UPDATE ONLY THIS
# ===============================
update_month = "Jul"    # e.g. "Jan".."Dec" ( T - 1 )
current_Quarter = "Q2"

# Update Year here
year = 2026

# Update MMPR File here
MMPR     = r"C:\Users\balatarsini_avinitya\Downloads\MMPR 2607 - 11.8.26.xls"

# Update export root file here
ROOT_OUT = Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Estimation Market\MY")

# Update last month's lux report here
lux = Path(r"C:\Users\balatarsini_avinitya\Downloads\New LUX report summary_MY_June'26(Others market remain as Others) (1).xlsx")

# Update BR Data Split here
br = Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Actual Market\MY After Monthly Split\Output\MY BR Monthly split (2024 Q1 - 2026 Q2).xlsx")
br = pd.read_excel(br, header=0)

# ===============================
# 🔧 USER CONFIG — DON'T CHANGE HERE
# ===============================
# Instantiate helpers
DB= r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\storage.db"

p = prepare_data()
d = process_data()
s = storage(DB)

# Build current and last-month tags like "Oct'25"
now = dt.datetime.now()
MONTH_NUM_BY_ABBR = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

def month_year_tag(month_abbr: str, selected_year: int) -> str:
    month_key = str(month_abbr).strip().title()[:3]
    if month_key not in MONTH_NUM_BY_ABBR:
        raise ValueError(f"Invalid update_month={month_abbr!r}. Use Jan, Feb, Mar, etc.")
    return dt.datetime(selected_year, MONTH_NUM_BY_ABBR[month_key], 1).strftime("%b'%y")

def previous_month_tag(month_abbr: str, selected_year: int) -> str:
    month_key = str(month_abbr).strip().title()[:3]
    if month_key not in MONTH_NUM_BY_ABBR:
        raise ValueError(f"Invalid update_month={month_abbr!r}. Use Jan, Feb, Mar, etc.")

    month_num = MONTH_NUM_BY_ABBR[month_key]
    if month_num == 1:
        return dt.datetime(selected_year - 1, 12, 1).strftime("%b'%y")
    return dt.datetime(selected_year, month_num - 1, 1).strftime("%b'%y")

def months_through(month_abbr: str) -> list[str]:
    month_key = str(month_abbr).strip().title()[:3]
    if month_key not in MONTH_NUM_BY_ABBR:
        raise ValueError(f"Invalid update_month={month_abbr!r}. Use Jan, Feb, Mar, etc.")
    month_order = list(MONTH_NUM_BY_ABBR.keys())
    return month_order[:MONTH_NUM_BY_ABBR[month_key]]

def months_before(month_abbr: str) -> list[str]:
    month_key = str(month_abbr).strip().title()[:3]
    if month_key not in MONTH_NUM_BY_ABBR:
        raise ValueError(f"Invalid update_month={month_abbr!r}. Use Jan, Feb, Mar, etc.")
    return months_through(month_key)[:-1]

month_tag = month_year_tag(update_month, year)

def load_mmpr_mtd_evo_by_category(mmpr_path) -> dict[str, float]:
    """
    Read the Total row from MMPR/Sum and return Apr/Mar/etc. MTD Evo%
    by category. The source layout has category headers above paired
    Mth/Ytd columns, so this explicitly selects the Mth column.
    """
    mmpr = pd.read_excel(mmpr_path, sheet_name="Sum", header=None)

    category_row_idx = None
    for idx, row in mmpr.iterrows():
        labels = {str(v).strip() for v in row.dropna()}
        if {"Total Malaysia", "Skincare", "Make-up", "Fragrances"}.issubset(labels):
            category_row_idx = idx
            break
    if category_row_idx is None:
        raise ValueError("Could not find the category header row in MMPR/Sum.")

    metric_row_idx = None
    for idx in range(category_row_idx + 1, min(category_row_idx + 5, len(mmpr))):
        metrics = [str(v).strip().lower() for v in mmpr.iloc[idx].dropna()]
        if "mth" in metrics:
            metric_row_idx = idx
            break
    if metric_row_idx is None:
        raise ValueError("Could not find the Mth/Ytd metric row in MMPR/Sum.")

    total_row_idx = None
    for idx, row in mmpr.iterrows():
        if any(str(v).strip().lower() == "total" for v in row.iloc[:3].dropna()):
            total_row_idx = idx
    if total_row_idx is None:
        raise ValueError("Could not find the Total row in MMPR/Sum.")

    category_renames = {
        "Make-up": "Makeup",
        "Others": "Hair/Others",
        "Total Malaysia": "Total",
    }

    current_category = None
    evo = {}
    for col_idx in range(mmpr.shape[1]):
        raw_category = mmpr.iat[category_row_idx, col_idx]
        if pd.notna(raw_category) and str(raw_category).strip():
            current_category = str(raw_category).strip()

        metric = str(mmpr.iat[metric_row_idx, col_idx]).strip().lower()
        if current_category and metric == "mth" and current_category != "Total Malaysia":
            category = category_renames.get(current_category, current_category)
            evo[category] = pd.to_numeric(mmpr.iat[total_row_idx, col_idx], errors="coerce")

    evo = {category: float(value) for category, value in evo.items() if pd.notna(value)}
    if not evo:
        raise ValueError("No MTD Evo% values were extracted from MMPR/Sum.")
    return evo

def load_est_percentage_seed(last_month_tag: str) -> pd.DataFrame:
    """
    Prefer the previous monthly Est_percentage snapshot. If this is the first
    run in the new flow, fall back to the latest legacy/base percentage table.
    """
    attempts = [
        {"dataset": "Est_percentage", "category": last_month_tag, "latest": True},
        {"dataset": f"Est_percentage_{year - 1}", "category": last_month_tag, "latest": True},
        {"dataset": "Est_percentage", "category": None, "latest": True},
        {"dataset": f"Est_percentage_{year - 1}", "category": None, "latest": True},
    ]

    errors = []
    for attempt in attempts:
        try:
            df = s.load_snapshot(s.conn, **attempt)
            category_label = attempt["category"] or "latest available"
            print(f"Loaded estimation seed from dataset={attempt['dataset']!r}, category={category_label!r}")
            return df
        except Exception as exc:
            errors.append(f"{attempt}: {exc}")

    raise ValueError("Could not find an estimation percentage seed snapshot.\n" + "\n".join(errors))


def update_est_percentages_from_file():
    """
    Uses globals above. Only `update_month` is user-editable.
    Fixed settings inside:
      - sheet_name="Sum"
      - load snapshot:  dataset="Est_percentage", category=previous month from update_month
      - save snapshot:  dataset="Est_percentage", category=update_month/year
      - default renames: Make-up->Makeup, Others->Hair Care, Total Malaysia->Total
    Saves an Excel copy to ROOT_OUT named 'Est_Percentage_2025 {month_tag}.xlsx'
    """
    # -------- fixed constants (do not change) --------
    SHEET_NAME      = "Sum"
    LOAD_DATASET    = "Est_percentage"
    SAVE_DATASET    = "Est_percentage"
    DEFAULT_RENAMES = {
        "Make-up": "Makeup",
        "Others": "Hair Care",
        "Total Malaysia": "Total",
    }

    # -------------------------------------------------

    # Load the previous monthly snapshot and save a new snapshot for update_month.
    # Example: update_month="Mar", year=2026 -> load Feb'26, save Mar'26.
    last_month_tag = previous_month_tag(update_month, year)

    # 1) Load last month's snapshot. If it does not exist yet, use the latest
    # legacy/base estimation percentage table as the first seed.
    print(s.list_snapshots(s.conn))
    est_percentage_for_2025 = load_est_percentage_seed(last_month_tag)

    # 2) Read MMPR MTD Evo% from the Total row.
    mapping = load_mmpr_mtd_evo_by_category(MMPR)

    # 4) Update snapshot dataframe for the chosen update_month
    updated_est = est_percentage_for_2025.drop(columns=["Unnamed: 0"], errors="ignore").copy()

    if "Year" not in updated_est.columns:
        updated_est["Year"] = year

    # 1) Only do this if the year is missing
    if year not in updated_est["Year"].unique():
        # choose a template year to copy from, e.g. the first/only one (2025)
        template_year = updated_est["Year"].min()   # or max(), or 2025 explicitly

        template = updated_est[updated_est["Year"] == template_year].copy()

        # 2) set the value column (3rd column) to NaN
        value_col = template.columns[2]   # 0: Category, 1: Month, 2: your numeric column
        template[value_col] = np.nan

        # 3) change year to the new target year
        template["Year"] = year

        # 4) append to the original df
        updated_est = pd.concat([updated_est, template], ignore_index=True)

    mask = mask = (
        updated_est["Month"].eq(update_month)
        & updated_est["Category"].isin(mapping.keys())
        & updated_est["Year"].eq(year)
    )
    updated_est.loc[mask, "est_per"] = updated_est.loc[mask, "Category"].map(mapping)

    # 5) Save snapshot back to storage (under current month’s tag)
    table_name = s.save_snapshot(s.conn, SAVE_DATASET, updated_est, category=month_tag)
    print("saved to:", table_name)

    # 6) Save Excel copy to ROOT_OUT
    ROOT_OUT.mkdir(parents=True, exist_ok=True)
    excel_out_path = ROOT_OUT / f"{SAVE_DATASET} {month_tag}.xlsx"
    updated_est.to_excel(excel_out_path, index=False)
    print("wrote excel:", excel_out_path)

    return updated_est, mapping, month_tag, table_name, excel_out_path

from pathlib import Path
import datetime as dt
import pandas as pd

# ---------- small helper: debug folder + per-step Excel snapshots ----------
try:
    DEBUG_OUT = ROOT_OUT / "new_estimation_debug"    # use your global if available
except NameError:
    DEBUG_OUT = Path.cwd() / "new_estimation_debug"  # fallback to current dir

DEBUG_OUT.mkdir(parents=True, exist_ok=True)

def _ts():
    return dt.datetime.now().strftime("%Y%m%d_%H%M%S")

def _save_step(df: pd.DataFrame, name: str):
    """Save a DataFrame to the debug folder with a timestamped filename."""
    path = DEBUG_OUT / f"{name}__{_ts()}.xlsx"
    df.to_excel(path, index=False)
    print(f"[saved] {path}")


def new_estimation():
    selected_months = months_through(update_month)
    actual_months = months_before(update_month)

    # ----------------------------------------------------------------------
    # STEP 1: Load last saved estimation snapshot & normalize category labels
    # ----------------------------------------------------------------------
    est_percentage_for_2025 = s.load_snapshot(s.conn, dataset="Est_percentage", category=month_tag)
    mapping = {
        "Fragrance": "Fragrances",
        "Hair Care": "Hair/Others",
        "Make-up": "Makeup",
    }
    est_percentage_for_2025["Category"] = est_percentage_for_2025["Category"].replace(mapping)
    _save_step(est_percentage_for_2025, f"01_est_percentage_for_{year}_loaded_and_mapped")

    # ----------------------------------------------------------------------
    # STEP 2: Read LUX report (first 18 columns only)
    # ----------------------------------------------------------------------
    lux_report = p.read_file(path=lux,  sheet_name="Market Conso Data", header=1)
    lux_report = lux_report.iloc[:, :18]
    _save_step(lux_report, "02_lux_report_raw_first_18_cols")

    # ----------------------------------------------------------------------
    # STEP 3: Filter Total group for 2025 and 2024
    # ----------------------------------------------------------------------
    last_year = year - 1
    lux_report_2025 = lux_report[(lux_report["Year"] == year) & (lux_report["Group"] == "Total")]
    lux_report_2024 = lux_report[(lux_report["Year"] == last_year) & (lux_report["Group"] == "Total")]
    _save_step(lux_report_2025, f"03_lux_report_{year}_total")
    _save_step(lux_report_2024, f"03_lux_report_{last_year}_total")

    # ----------------------------------------------------------------------
    # STEP 4: Normalize to long format with Month/Value
    # ----------------------------------------------------------------------
    lux_report_2025 = p.normalize_df(
        lux_report_2025,
        id_vars=lux_report_2025.columns[:6],
        melt=True, var_name="Month", value_name="Value", detect_months=True
    )
    lux_report_2024 = p.normalize_df(
        lux_report_2024,
        id_vars=lux_report_2024.columns[:6],
        melt=True, var_name="Month", value_name="Value", detect_months=True
    )
    _save_step(lux_report_2025, f"04_lux_report_{year}_long")
    _save_step(lux_report_2024, f"04_lux_report_{last_year}_long")

    # ----------------------------------------------------------------------
    # STEP 5: Read MMPR MTD Evo% and estimate update month by category
    #         using last-year same-month category value.
    # ----------------------------------------------------------------------
    update_month_key = str(update_month).strip().title()[:3]
    mmpr_mtd_evo = load_mmpr_mtd_evo_by_category(MMPR)
    evo_df = pd.DataFrame(
        [{"Category": category, "est_per": evo} for category, evo in mmpr_mtd_evo.items()]
    )
    _save_step(evo_df, "05_mmpr_mtd_evo_by_category")

    channel_keys = ["Channel", "Category", "Brands", "Group", "Online/Offline", "Month"]
    last_year_update_month = lux_report_2024[
        lux_report_2024["Month"].astype(str).eq(update_month_key)
    ].copy()

    last_year_channel = (
        last_year_update_month
        .groupby(channel_keys, dropna=False, as_index=False)["Value"]
        .sum()
    )
    last_year_category = (
        last_year_channel
        .groupby(["Category", "Month"], dropna=False, as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "last_year_category_value"})
    )

    generated_category = (
        last_year_category
        .merge(evo_df, on="Category", how="left")
        .assign(
            Year=year,
            generated_category_value=lambda x: x["last_year_category_value"] * (1 + x["est_per"])
        )
    )

    missing_evo = generated_category.loc[
        generated_category["est_per"].isna(), "Category"
    ].drop_duplicates().tolist()
    if missing_evo:
        raise ValueError(f"Missing MMPR MTD Evo% for categories: {missing_evo}")

    channel_wob = (
        last_year_channel
        .merge(last_year_category, on=["Category", "Month"], how="left")
        .assign(
            channel_wob=lambda x: np.where(
                x["last_year_category_value"].ne(0),
                x["Value"] / x["last_year_category_value"],
                0,
            )
        )
    )
    _save_step(channel_wob, "06_last_year_same_month_channel_wob")
    _save_step(generated_category, "07_generated_category_values_from_mmpr_mtd_evo")

    estimated_update_month = (
        channel_wob
        .merge(
            generated_category[["Category", "Month", "generated_category_value"]],
            on=["Category", "Month"],
            how="left",
        )
        .assign(
            Year=year,
            Value=lambda x: x["generated_category_value"] * x["channel_wob"],
        )
    )
    estimated_update_month = estimated_update_month[
        ["Year", "Channel", "Category", "Brands", "Group", "Online/Offline", "Month", "Value"]
    ]
    _save_step(estimated_update_month, "08_estimated_update_month_split_by_channel_wob")

    # ----------------------------------------------------------------------
    # STEP 6: Keep actuals before update month, then append the estimated month.
    # ----------------------------------------------------------------------
    actual_prior_months = lux_report_2025[
        lux_report_2025["Month"].astype(str).isin(actual_months)
    ].copy()
    actual_prior_months = (
        actual_prior_months
        .groupby(["Year", *channel_keys], dropna=False, as_index=False)["Value"]
        .sum()
    )

    channel_split = pd.concat(
        [actual_prior_months, estimated_update_month],
        ignore_index=True,
    )
    channel_split["Channel"] = channel_split["Channel"].replace({"Tiktok": "TikTok"})
    channel_split = channel_split[
        channel_split["Month"].astype(str).isin(selected_months)
    ].copy()
    _save_step(channel_split, "09_actual_prior_months_plus_estimated_update_month")

    print(channel_split)

    # ----------------------------------------------------------------------
    # STEP 15: Add Brands/Group, map Online/Offline by Channel
    # ----------------------------------------------------------------------
    channel_split["Brands"] = "Market"
    channel_split["Group"] = "Total"

    cases = {
        "Boutiques (Offline)": "Offline",
        "Department Stores (Offline)": "Offline",
        "Department Stores (Online)":  "Online",
        "Sephora (Offline)":           "Offline",
        "Sephora (Online)":            "Online",
        "Tourist (Offline)":           "Offline",
        "Lazada":                      "Online",
        "Shopee":                      "Online",
        "TikTok":                      "Online",
        "Other e-tailers":             "Online",
        "E-Boutiques (Online)":        "Online",
        "Socom (Online)":              "Online",
        "Perfumery (Offline)":         "Offline", 
        "E-tailers (Online)":          "Online",
    }

    channel_split = d.map_row_by_any_exact(
        channel_split,
        cases=cases,
        new_col="Online/Offline",
        in_cols=["Channel"],
        first_wins=True,
        case_sensitive=False,      # <-- changed
        inplace=False,
        drop_empty_cols=True
    )
    _save_step(channel_split, "15_channel_split_after_channel_map")

    # ----------------------------------------------------------------------
    # STEP 16: Keep columns, rename Year_x -> Year
    # ----------------------------------------------------------------------
    columns_to_keep = ["Year", "Channel", "Category", "Brands", "Group", "Online/Offline", "Month", "Value"]
    channel_split = channel_split[columns_to_keep]
    _save_step(channel_split, "16_channel_split_pruned_and_year_renamed")
    print(channel_split)

    # ----------------------------------------------------------------------
    # STEP 17: Wide month output (12 columns), fill missing with 0
    # ----------------------------------------------------------------------
    chan_order = [
        "Boutiques (Offline)",
        "Department Stores (Offline)",
        "Department Stores (Online)",
        "E-Boutiques (Online)",
        "E-tailers (Online)",
        "Lazada",
        "Sephora (Offline)",
        "Sephora (Online)",
        "Shopee",
        "Socom (Online)",
        "TikTok",
        "Tourist (Offline)",
    ]
    cat_order = ["Fragrances", "Hair/Others", "Makeup", "Skincare"]

    # wide table (you already built channel_split above)
    channel_split = d.to_month_wide(
        channel_split,
        keys=("Year", "Channel", "Category", "Brands", "Group", "Online/Offline"),
        month_col="Month",
        value_col="Value",
        month_order=selected_months,
        fill_missing=True,
        fill_value=0,
        sort_rows=True,
    )

    # work on a copy
    df = channel_split.copy()

    # ✅ Complete ONLY Channel×Category grid (ignore other keys), zero-fill only NEW rows’ numeric columns
    grid = pd.MultiIndex.from_product([chan_order, cat_order], names=["Channel", "Category"])
    base = df.set_index(["Channel", "Category"])
    before_idx = base.index
    out17 = base.reindex(grid)

    # identify numeric columns to fill only on new rows
    value_cols = [c for c in out17.columns if pd.api.types.is_numeric_dtype(out17[c])]
    new_rows_mask = ~out17.index.isin(before_idx)
    out17.loc[new_rows_mask, value_cols] = 0  # leave existing rows' values untouched

    # reset + enforce order + sort
    out17 = out17.reset_index()
    out17["Channel"]  = pd.Categorical(out17["Channel"],  categories=chan_order, ordered=True)
    out17["Category"] = pd.Categorical(out17["Category"], categories=cat_order,  ordered=True)
    out17 = out17.sort_values(["Channel", "Category"], kind="mergesort").reset_index(drop=True)

    out17["Year"] = year
    out17["Brands"] = "Market"
    out17["Group"] = "Total"

    out17 = d.map_row_by_any_exact(
        out17,
        cases=cases,
        new_col="Online/Offline",
        in_cols=["Channel"],
        first_wins=True,
        case_sensitive=False,      # <-- changed
        inplace=False,
        drop_empty_cols=True
    )

    # save debug snapshot
    _save_step(out17, "17_channel_split_month_wide")
    print(out17)

    # === FINAL OUTPUT (outside debug) ===
    ROOT_OUT.mkdir(parents=True, exist_ok=True)
    final_filename = f"MY_Estimation_{year}_{update_month}_{_ts()}.xlsx"  # use month_tag instead of current_month if you prefer
    final_path = ROOT_OUT / final_filename
    out17.to_excel(final_path, index=False)
    print(f"[final saved] {final_path}")

    # Optionally return the last DF if you want:
    # return channel_split

def main():
    """
    Orchestrates the run:
      1) Sets runtime globals used by your functions (now, month_tag, year)
      2) Runs update_est_percentages_from_file()
      3) Runs new_estimation()
    """
    import datetime as dt

    # ---- runtime globals your functions rely on (no changes inside your funcs) ---
    global now, month_tag, year
    now = dt.datetime.now()
    month_tag = month_year_tag(update_month, year)   # e.g. update_month="Mar", year=2026 -> "Mar'26"

    print(f"[main] month_tag={month_tag} | year={year}")

    # ---- run step 1: update est percentages & save snapshot/Excel --------------
    try:
        upd = update_est_percentages_from_file()
        print("[main] update_est_percentages_from_file() completed.")
    except Exception as e:
        print("[main] ERROR in update_est_percentages_from_file():", e)
        upd = None

    # ---- run step 2: build new estimation outputs (writes debug excels) --------
    try:
        res = new_estimation()
        print("[main] new_estimation() completed.")
    except Exception as e:
        print("[main] ERROR in new_estimation():", e)
        res = None

    # return upd, res


if __name__ == "__main__":
    main()
