from Table_Preparation import prepare_data
from Data_Processing import process_data
from pathlib import Path
import datetime as dt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils import get_column_letter

# ========= Updates =========
bad_brands = ["The Ordinary"]  # Update brands to exclude it

SPLIT_PERC = [0.6, 0.4]  # global split percentages - Update the Percentage when needed
# Current: 0.6 for Etailer, 0.4 for Eboutique
# Update Current Status after changing

output_root = Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Actual Market") # Update the root folder to save where you want

# Put MY and SG BR files inside this folder, e.g. Raw BR\MY\file.xlsx and Raw BR\SG\file.xlsx
raw_br_root = Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Raw BR")

# Latest O+O reports used to split quarterly BR actuals into months.
ooo_report_by_country = {
    "MY": Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Actual Market\MY After Monthly Split\New LUX report summary_MY_June'26 (2).xlsx"),
    "SG": Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Actual Market\SG After Monthly Split\New LUX report summary_SG_June'26 (2).xlsx"),
}

monthly_start_year = 2024

# Update the Year and Quarter \here
target_year = 2026
target_quarter = "Q2"   # e.g. "Q1","Q2","Q3","Q4"

# Update the Password here
password = "BR.LOREAL"

# ========= Mapping For BR =========
# DO NOT CHANGE!!!!!!
# Call for classes 
p = prepare_data()
d = process_data()

# UPDATE IT IF NEEDED
COMBINED_CHANNEL_MAP = {
    "Brand website": "E-Boutiques (Online)",
    "All platforms": "E-tailers (Online)/E-Boutiques (Online)",
    "e-Sephora (All platforms)": "Sephora (Online)",

    # Offline buckets
    "Specialty Stores": "Sephora (Offline)",
    "Department Stores": "Department Stores (Offline)",
    "Standalone Boutiques": "Boutiques (Offline)",
    "Perfumery": "Perfumery (Offline)",

    # Online buckets
    "Chat & Shop": "E-Boutiques (Online)",
    "Lazada": "E-tailers (Online)",
    "Hermo": "E-tailers (Online)",
    "Prestomall": "E-tailers (Online)",
    "Shopee": "E-tailers (Online)",
    "Zalora": "E-tailers (Online)",
    "AMAZON": "E-tailers (Online)",
    "Phone Order Home Delivery": "E-tailers (Online)",
    "One Shop": "E-tailers (Online)",
    "Astro Go Shop": "E-tailers (Online)",
    "Tiktok": "E-tailers (Online)",
    "Phone Order Home Delivery/Click & Collect": "E-tailers (Online)",
}

MONTHS_BY_QUARTER = {
    "Q1": ["Jan", "Feb", "Mar"],
    "Q2": ["Apr", "May", "Jun"],
    "Q3": ["Jul", "Aug", "Sep"],
    "Q4": ["Oct", "Nov", "Dec"],
}
MONTH_ORDER = [m for months in MONTHS_BY_QUARTER.values() for m in months]
QUARTER_ORDER = {q: i for i, q in enumerate(MONTHS_BY_QUARTER, start=1)}

MONTHLY_CHANNEL_MAP = {
    "Lazada": "E-tailers (Online)",
    "Shopee": "E-tailers (Online)",
    "TikTok": "E-tailers (Online)",
    "Tiktok": "E-tailers (Online)",
    "Other e-tailers": "E-tailers (Online)",
    "Tourist Store (Offline)": "Tourist (Offline)",
}

ONLINE_OFFLINE_MAP = {
    "Boutiques (Offline)": "Offline",
    "Department Stores (Offline)": "Offline",
    "Department Stores (Online)": "Online",
    "E-Boutiques (Online)": "Online",
    "E-tailers (Online)": "Online",
    "Perfumery (Offline)": "Offline",
    "Sephora (Offline)": "Offline",
    "Sephora (Online)": "Online",
    "Tourist (Offline)": "Offline",
}

CATEGORY_MAP = {
    "Fragrance": "Fragrances",
    "Hair": "Hair/Others",
    "Hair Care": "Hair/Others",
    "Make-up": "Makeup",
}

# ========= FUNCTIONS ========= DO NOT CHANGE!!!!!!!!
def make_pandas_pivot(
    df: pd.DataFrame,
    *,
    include_totals: bool = False       # <- no Total columns
) -> pd.DataFrame:
    # Build pivot: rows = Category, Channel; cols = Year x Quarter; values = sum(ActualQuarter)
    piv = pd.pivot_table(
        df,
        index=["Category", "Channel"],
        columns=["Year", "Quarter"],
        values="ActualQuarter",
        aggfunc="sum",
        fill_value=0,
        margins=include_totals,
        margins_name="Total",
    ).sort_index()

    # Flatten multi-index columns like (2025, 'Q1') -> "2025 Q1"
    if isinstance(piv.columns, pd.MultiIndex):
        piv.columns = [f"{c[0]} {c[1]}" if c[1] != "" else f"{c[0]}" for c in piv.columns]
    else:
        piv.columns = [str(c) for c in piv.columns]

    # If totals were ever present, drop the Total columns explicitly
    if not include_totals:
        piv = piv.loc[:, [c for c in piv.columns if str(c).strip().lower() != "total"]]

    # Bring the row index out as columns so labels repeat down the rows
    piv = piv.reset_index()  # gives you Category, Channel columns with repeated labels
    
    # if you used the helper that flattens columns: "2025 Q3"
    col_key = f"{target_year} {target_quarter}"

    # id columns you want to preserve (adjust as needed)
    id_cols = [c for c in ["Category", "Channel"] if c in piv.columns]

    # ---------------------------
    # keep ONLY that YQ
    # ---------------------------
    if col_key in piv.columns:
        pivot_yq_only = piv.loc[:, id_cols + [col_key]]
    else:
        print(f"Warning: column '{col_key}' not found; available: {list(piv.columns)}")
        pivot_yq_only = piv.loc[:, id_cols]  # fallback
    
    desired = ["Channel", "Category", col_key]  
    pivot_yq_only = pivot_yq_only.reindex(columns=[c for c in desired if c in pivot_yq_only.columns])

    return pivot_yq_only

def infer_country_from_path(file_path: Path) -> str:
    parent_country = file_path.parent.name.upper()
    if parent_country in {"MY", "SG"}:
        return parent_country

    file_name = file_path.name.casefold()
    if "malaysia" in file_name:
        return "MY"
    if "singapore" in file_name:
        return "SG"

    return "UNKNOWN"

def find_br_files(raw_root: Path) -> list[Path]:
    return sorted(
        file
        for file in raw_root.rglob("*.xlsx")
        if not file.name.startswith("~$")
    )

def quarter_sort_key(year, quarter) -> tuple[int, int]:
    return int(year), QUARTER_ORDER.get(str(quarter).strip().upper(), 0)

def add_quarter_from_month(df: pd.DataFrame, month_col: str = "Month") -> pd.DataFrame:
    out = df.copy()
    month_to_quarter = {
        month: quarter
        for quarter, months in MONTHS_BY_QUARTER.items()
        for month in months
    }
    out["Quarter"] = out[month_col].map(month_to_quarter)
    return out

def latest_available_quarter(df_actual: pd.DataFrame) -> tuple[int, str]:
    actual = df_actual.dropna(subset=["Year", "Quarter"]).copy()
    if actual.empty:
        raise ValueError("No Year/Quarter rows found in BR actual data.")

    actual["Year"] = pd.to_numeric(actual["Year"], errors="coerce").astype("Int64")
    actual["Quarter"] = actual["Quarter"].astype(str).str.strip().str.upper()
    actual = actual.dropna(subset=["Year"])
    actual = actual[actual["Quarter"].isin(QUARTER_ORDER)]

    if actual.empty:
        raise ValueError("No valid Q1-Q4 quarters found in BR actual data.")

    latest = actual.sort_values(
        by=["Year", "Quarter"],
        key=lambda s: s.map(QUARTER_ORDER) if s.name == "Quarter" else s,
    ).iloc[-1]
    return int(latest["Year"]), latest["Quarter"]

def load_ooo_monthly_profile(ooo_path: str | Path) -> pd.DataFrame:
    ooo_path = Path(ooo_path)
    if not ooo_path.exists():
        raise FileNotFoundError(f"O+O report not found: {ooo_path}")

    lux = pd.read_excel(ooo_path, sheet_name="Market Conso Data", header=1)
    lux = lux.iloc[:, :18].copy()
    lux["Year"] = pd.to_numeric(lux["Year"], errors="coerce").astype("Int64")
    lux = lux[
        lux["Year"].notna()
        & lux["Group"].astype(str).str.strip().eq("Total")
        & lux["Brands"].astype(str).str.strip().eq("Market")
    ].copy()

    lux["Category"] = lux["Category"].replace(CATEGORY_MAP)
    lux["Channel"] = lux["Channel"].replace(MONTHLY_CHANNEL_MAP)

    monthly = p.normalize_df(
        lux,
        id_vars=["Year", "Channel", "Category", "Brands", "Group", "Online/Offline"],
        melt=True,
        var_name="Month",
        value_name="EstimateValue",
        detect_months=True,
    )
    monthly["EstimateValue"] = pd.to_numeric(monthly["EstimateValue"], errors="coerce").fillna(0)
    monthly = add_quarter_from_month(monthly)

    monthly = (
        monthly
        .groupby(["Year", "Quarter", "Month", "Channel", "Category"], dropna=False, as_index=False)["EstimateValue"]
        .sum()
    )
    return monthly

def allocate_quarterly_actuals_to_months(
    df_actual: pd.DataFrame,
    df_monthly_profile: pd.DataFrame,
    *,
    start_year: int,
    country: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    latest_year, latest_quarter = latest_available_quarter(df_actual)
    latest_key = quarter_sort_key(latest_year, latest_quarter)

    actual = df_actual.copy()
    actual["Year"] = pd.to_numeric(actual["Year"], errors="coerce").astype("Int64")
    actual["Quarter"] = actual["Quarter"].astype(str).str.strip().str.upper()
    actual["Category"] = actual["Category"].replace(CATEGORY_MAP)
    actual["ActualQuarter"] = pd.to_numeric(actual["ActualQuarter"], errors="coerce").fillna(0)
    actual = actual[
        actual["Year"].notna()
        & actual["Quarter"].isin(QUARTER_ORDER)
        & (actual["Year"] >= start_year)
    ].copy()
    actual = actual[
        actual.apply(lambda r: quarter_sort_key(r["Year"], r["Quarter"]) <= latest_key, axis=1)
    ].copy()

    quarterly_actual = (
        actual
        .groupby(["Year", "Quarter", "Channel", "Category"], dropna=False, as_index=False)["ActualQuarter"]
        .sum()
    )

    profile = df_monthly_profile.copy()
    profile["Year"] = pd.to_numeric(profile["Year"], errors="coerce").astype("Int64")
    profile["Quarter"] = profile["Quarter"].astype(str).str.strip().str.upper()
    profile = profile[
        profile["Year"].notna()
        & profile["Quarter"].isin(QUARTER_ORDER)
        & (profile["Year"] >= start_year)
    ].copy()
    profile = profile[
        profile.apply(lambda r: quarter_sort_key(r["Year"], r["Quarter"]) <= latest_key, axis=1)
    ].copy()

    keys = ["Year", "Quarter", "Channel", "Category"]
    profile_keys = profile[keys].drop_duplicates()
    actual_keys = quarterly_actual[keys].drop_duplicates()
    scaffold = pd.concat([profile_keys, actual_keys], ignore_index=True).drop_duplicates()
    scaffold["Month"] = scaffold["Quarter"].map(MONTHS_BY_QUARTER)
    scaffold = scaffold.explode("Month", ignore_index=True)

    profile = (
        scaffold
        .merge(profile, on=keys + ["Month"], how="left")
    )
    profile["EstimateValue"] = profile["EstimateValue"].fillna(0)

    quarter_estimate = (
        profile
        .groupby(keys, dropna=False, as_index=False)["EstimateValue"]
        .sum()
        .rename(columns={"EstimateValue": "EstimateTotal"})
    )

    monthly = (
        profile
        .merge(quarter_estimate, on=keys, how="left")
        .merge(quarterly_actual, on=keys, how="left")
    )
    monthly["ActualQuarter"] = monthly["ActualQuarter"].fillna(0)
    monthly["Actual"] = 0.0

    has_estimate = monthly["EstimateTotal"].ne(0)
    monthly.loc[has_estimate, "Actual"] = (
        monthly.loc[has_estimate, "EstimateValue"]
        / monthly.loc[has_estimate, "EstimateTotal"]
        * monthly.loc[has_estimate, "ActualQuarter"]
    )

    no_estimate_has_actual = monthly["EstimateTotal"].eq(0) & monthly["ActualQuarter"].ne(0)
    monthly.loc[no_estimate_has_actual, "Actual"] = monthly.loc[no_estimate_has_actual, "ActualQuarter"] / 3

    if country.upper() == "SG":
        dept_channels = ["Department Stores (Offline)", "Department Stores (Online)"]
        dept_mask = monthly["Channel"].isin(dept_channels)
        if dept_mask.any():
            dept_actual = monthly.loc[
                monthly["Channel"].eq("Department Stores (Offline)"),
                ["Year", "Quarter", "Category", "ActualQuarter"],
            ].drop_duplicates().rename(columns={"ActualQuarter": "_DeptQuarterActual"})

            dept_quarter_total = (
                monthly.loc[dept_mask]
                .groupby(["Year", "Quarter", "Category"], dropna=False, as_index=False)["EstimateValue"]
                .sum()
                .rename(columns={"EstimateValue": "_DeptQuarterEstimate"})
            )

            dept_calc = (
                monthly.loc[dept_mask, ["Year", "Quarter", "Month", "Channel", "Category", "EstimateValue"]]
                .merge(dept_actual, on=["Year", "Quarter", "Category"], how="left")
                .merge(dept_quarter_total, on=["Year", "Quarter", "Category"], how="left")
            )

            dept_allocated = (
                dept_calc["EstimateValue"]
                / dept_calc["_DeptQuarterEstimate"].replace(0, float("nan"))
                * dept_calc["_DeptQuarterActual"]
            ).fillna(0)

            dept_allocated = dept_allocated.astype(float)

            print("=" * 60)
            print("dept_mask:", dept_mask.sum())
            print("dept_calc:", len(dept_calc))
            print("dept_allocated:", len(dept_allocated))
            print("dept_actual:", len(dept_actual))

            print(
                dept_actual.groupby(
                    ["Year","Quarter","Category"]
                ).size().sort_values(ascending=False).head(20)
            )

            monthly.loc[dept_mask, "Actual"] = dept_allocated.astype(float).values
            monthly.loc[dept_mask, "ActualQuarter"] = dept_calc["_DeptQuarterActual"].fillna(0).to_numpy()

    monthly["Actual '000"] = monthly["Actual"] / 1000
    monthly["Brands"] = "Market"
    monthly["Group"] = "Total"
    monthly["Online/Offline"] = monthly["Channel"].map(ONLINE_OFFLINE_MAP)

    monthly["Month"] = pd.Categorical(monthly["Month"], categories=MONTH_ORDER, ordered=True)
    monthly["Quarter"] = pd.Categorical(monthly["Quarter"], categories=list(QUARTER_ORDER), ordered=True)
    monthly = monthly.sort_values(["Year", "Quarter", "Channel", "Category", "Month"], kind="mergesort")
    monthly["Month"] = monthly["Month"].astype(str)
    monthly["Quarter"] = monthly["Quarter"].astype(str)

    validation_frame = monthly.copy()
    validation_keys = keys.copy()
    if country.upper() == "SG":
        validation_frame["ValidationChannel"] = validation_frame["Channel"]
        validation_frame.loc[
            validation_frame["Channel"].isin(["Department Stores (Offline)", "Department Stores (Online)"]),
            "ValidationChannel",
        ] = "Department Stores (Offline+Online)"
        validation_keys = ["Year", "Quarter", "ValidationChannel", "Category"]

    validation = (
        validation_frame
        .groupby(validation_keys, dropna=False, as_index=False)
        .agg(
            MonthlyActual=("Actual", "sum"),
            ActualQuarter=("ActualQuarter", "first"),
            EstimateTotal=("EstimateTotal", "first"),
        )
    )
    if "ValidationChannel" in validation.columns:
        validation = validation.rename(columns={"ValidationChannel": "Channel"})
    validation["Difference"] = validation["MonthlyActual"] - validation["ActualQuarter"]
    validation["Status"] = validation["Difference"].abs().le(0.001).map({True: "OK", False: "CHECK"})

    return monthly.reset_index(drop=True), validation

def write_monthly_actual_workbook(
    monthly: pd.DataFrame,
    validation: pd.DataFrame,
    *,
    country: str,
    source_ooo_path: str | Path,
) -> Path:
    out_dir = output_root / f"{country} After Monthly Split" / "Output"
    out_dir.mkdir(parents=True, exist_ok=True)
    latest = (
        monthly[["Year", "Quarter"]]
        .drop_duplicates()
        .sort_values(
            by=["Year", "Quarter"],
            key=lambda s: s.map(QUARTER_ORDER) if s.name == "Quarter" else s,
        )
        .iloc[-1]
    )
    earliest = (
        monthly[["Year", "Quarter"]]
        .drop_duplicates()
        .sort_values(
            by=["Year", "Quarter"],
            key=lambda s: s.map(QUARTER_ORDER) if s.name == "Quarter" else s,
        )
        .iloc[0]
    )
    period_range = f"{int(earliest['Year'])} {earliest['Quarter']} - {int(latest['Year'])} {latest['Quarter']}"
    out_path = out_dir / f"{country} BR Monthly split ({period_range}).xlsx"

    data_source = d.to_month_wide(
        monthly,
        keys=("Year", "Channel", "Category", "Brands", "Group", "Online/Offline"),
        month_col="Month",
        value_col="Actual",
        month_order=MONTH_ORDER,
        fill_missing=True,
        fill_value=0,
        sort_rows=True,
        pre_aggregate=True,
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        monthly.to_excel(writer, sheet_name="Monthly Actual", index=False)
        data_source.to_excel(writer, sheet_name="Data Source", index=False)
        validation.to_excel(writer, sheet_name="Validation", index=False)

        for (year, quarter), quarter_df in monthly.groupby(["Year", "Quarter"], sort=True):
            sheet = f"{int(year)} {quarter}"
            quarter_df.to_excel(writer, sheet_name=sheet, index=False)

        pd.DataFrame(
            {
                "Item": ["Source O+O report", "Start year", "Latest BR quarter"],
                "Value": [
                    str(source_ooo_path),
                    monthly_start_year,
                    f"{int(latest['Year'])} {latest['Quarter']}",
                ],
            }
        ).to_excel(writer, sheet_name="Run Info", index=False)

    print("Wrote monthly actuals:", out_path)
    return out_path

def build_actuals_from_mcs(*, p, d, actual_file: str | Path, split: bool | None, country: str):
    # =========================================
    # 0) CONFIG & INPUTS
    # =========================================
    file = fr"{actual_file}"              # Upload Multi Channel Snapshot Report
    month_tag = dt.datetime.now().strftime("%b'%y")  # e.g. Oct'25

    # Ensure output dir exists (we'll save at the end only)
    root_dir = output_root / country
    root_dir.mkdir(parents=True, exist_ok=True)

    # =========================================
    # 1) READ + BASIC CLEAN (MCS)
    # =========================================
    mcs = p.read_file(file, password, header=12, sheet_name="Data") #If file has no password, remove it
    mcs = mcs[12:]  # drop header rows after reading
    mcs = p.normalize_df(mcs, to_lower=False, drop_all_na_cols=True)

    # normalize both sides: strip + casefold, then filter bad brands
    brand_norm = (
        mcs["Brand"].astype("string").str.strip().str.casefold()
    )
    bad_set = {b.strip().casefold() for b in bad_brands}
    mcs = mcs.loc[~brand_norm.isin(bad_set)].copy()

    # =========================================
    # 2) CHANNEL MAPPING (INTO HitLabel)
    # =========================================
    df2 = d.map_row_by_any_exact(
        mcs,
        cases=COMBINED_CHANNEL_MAP,        # <- global dict
        new_col="HitLabel",
        in_cols=["Channel", "e-platform"], # or None to scan all cols
        first_wins=True,
        case_sensitive=True,
        inplace=False,
        drop_empty_cols=True,
    )
    # (Removed early save; we save only at the end now)

    # =========================================
    # 3) SPLIT "All platforms" BUCKET (40/60)
    # =========================================
    if split is True:  # only split when explicitly True
        df3 = d.split_records(
            df2,
            row_filter={"HitLabel": "E-tailers (Online)/E-Boutiques (Online)"},
            split_col="HitLabel",
            parts=["E-tailers (Online)", "E-Boutiques (Online)"],
            perc=SPLIT_PERC,               # <- global list
            value_cols=["Sales"],
            drop_original=True,
            ensure_sum=True,  # second part balances (may go negative if earlier > 100%)
        )
    else:
        df3 = df2

    # Aggregate to Year/Category/Period/HitLabel level
    df_sum = (
        df3.groupby(["Year", "Category", "Period", "HitLabel"], as_index=False)["Sales"].sum()
    )

    df_sum["Sales"] = df_sum["Sales"] / 1000
    # (Removed split summary save; we save only the final renamed table now)

    # =========================================
    # 4) CLEAN + NORMALIZE (ACTUALS TABLE)
    # =========================================
    df4 = p.clean_to_numeric(df_sum, col="Sales")

    # Rename to match monthly schema for allocation step
    df_actual_renamed = df4.rename(columns={"Period": "Quarter", "Sales": "ActualQuarter", "HitLabel": "Channel"})[
        ["Category", "Year", "Quarter", "Channel", "ActualQuarter"]
    ]

    exclude_channels = {"E-tailers (Online)/E-Boutiques (Online)"}   # <-- change this
    norm_excl = {c.strip().casefold() for c in exclude_channels}

    df_actual_renamed = df_actual_renamed[
        ~df_actual_renamed["Channel"].astype(str).str.strip().str.casefold().isin(norm_excl)
    ].copy()

    # Normalize category labels (kept your custom mapping)
    mapping = {
        "Fragrance": "Fragrances",
        "Hair": "Hair/Others",
        "Make-up": "Makeup",
    }
    df_actual_renamed["Category"] = df_actual_renamed["Category"].replace(mapping)

    # =========================================
    # 5) SAVE AT THE END (Data + Pivot via pandas)
    # =========================================
    suffix = " - Include All Platform" if split is True else ""
    out_path = root_dir / f"BR After Split {country} {month_tag}{suffix}.xlsx"

    pivot_df = make_pandas_pivot(df_actual_renamed)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Sheet 1: final renamed data
        df_actual_renamed.to_excel(writer, sheet_name="Data", index=False)
        # Sheet 2: pandas pivot result
        pivot_df.to_excel(writer, sheet_name="Pivot")

    print("Wrote:", out_path)

    return df_actual_renamed

# ======= MAIN RUNNER ======= #FOR "ALL PLATFORM" - IF INCLUDED, SPLIT = TRUE AND VICE VERSA
if __name__ == "__main__":
    br_files = find_br_files(raw_br_root)
    if not br_files:
        raise FileNotFoundError(f"No BR Excel files found in: {raw_br_root}")

    for br_file in br_files:
        file_country = infer_country_from_path(br_file)
        print(f"Processing {file_country}: {br_file}")

        # Without Split
        # df_actuals = build_actuals_from_mcs(p=p, d=d, actual_file=br_file, split=False, country=file_country)

        # With Split
        df_actuals = build_actuals_from_mcs(p=p, d=d, actual_file=br_file, split=True, country=file_country)

        ooo_report = ooo_report_by_country.get(file_country)
        if ooo_report and Path(ooo_report).exists():
            monthly_profile = load_ooo_monthly_profile(ooo_report)
            monthly_actuals, monthly_validation = allocate_quarterly_actuals_to_months(
                df_actuals,
                monthly_profile,
                start_year=monthly_start_year,
                country=file_country,
            )
            write_monthly_actual_workbook(
                monthly_actuals,
                monthly_validation,
                country=file_country,
                source_ooo_path=ooo_report,
            )
        else:
            print(f"No O+O report configured for {file_country}; skipped monthly split.")
