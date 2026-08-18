import pandas as pd
from pathlib import Path
import datetime as dt
from Table_Preparation import prepare_data
from Data_Processing import process_data
from Storage import storage
import numpy as np

# === INPUT FILES (user-editable) =====================================================
MKT = r"C:\Users\balatarsini_avinitya\Downloads\Mkt est 2026_Jul.xlsx"  # Market-by-Channel summary source

current_month = "Jul"          # anchor month to operate on
year = 2026                    # target year for estimation/backfill
last_year = year - 1            # same month last year to include in final output
current_Quarter = "Q2"         # quarter to consider/backfill from (Q1..Q4)

ROOT_OUT = Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Estimation Market\SG")  # debug/output root

lux = Path(r"C:\Users\balatarsini_avinitya\Downloads\New LUX report summary_SG_June'26_Market Reinstated.xlsx")     # LUX market report

# Update BR Data Split here (post-channel-mapped, optionally split)
br = Path(r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\Actual Market\SG After Monthly Split\Output\SG BR Monthly split (2024 Q1 - 2026 Q2).xlsx")
br = pd.read_excel(br, header=0)  # Load the BR table (already prepared elsewhere)

# ===============================
# 🔧 USER CONFIG — DON'T CHANGE HERE
# ===============================
# Instantiate helpers (storage, prep, processing)
DB       = r"C:\Users\balatarsini_avinitya\Downloads\Final folders\LLD Market Estimation  V10.6.2026\LLD Market numbers\storage.db"

p = prepare_data()
d = process_data()
s = storage(DB)

# ---------- small helper: debug folder + per-step Excel snapshots ----------
try:
    DEBUG_OUT = ROOT_OUT / "new_estimation_debug"    # dedicated folder for intermediate outputs
except NameError:
    DEBUG_OUT = Path.cwd() / "new_estimation_debug"  # fallback if ROOT_OUT not defined

DEBUG_OUT.mkdir(parents=True, exist_ok=True)         # ensure folder exists

def _ts():
    # timestamp string for filenames
    return dt.datetime.now().strftime("%Y%m%d_%H%M%S")

def _save_step(df: pd.DataFrame, name: str):
    """Save a DataFrame to the debug folder with a timestamped filename."""
    path = DEBUG_OUT / f"{name}__{_ts()}.xlsx"
    df.to_excel(path, index=False)
    print(f"[saved] {path}")


# === Read MKT (Market by Channel Summary) and keep a compact slice ================
mkt = p.read_file(MKT, sheet_name="MarketbyChannel_Summary", header=0)
_save_step(mkt, "01_mkt_raw")

mkt = mkt.iloc[:12, :]
_save_step(mkt, "02_mkt_top12rows")

# Convert column headers to datetime where possible, then select the exact target month.
cols_dt = pd.to_datetime(mkt.columns, errors="coerce")

month_num = pd.to_datetime(current_month, format="%b").month
target_dt = pd.Timestamp(year=year, month=month_num, day=1)

match = np.where(cols_dt == target_dt)[0]

if len(match) == 0:
    raise ValueError(
        f"Column for {target_dt} not found. "
        f"Available datetime columns: {list(cols_dt.dropna())}"
    )

target_idx = int(match[0])
id_idx = 0

# Slice only ID + correct month column.
mkt_12_with_id = mkt.iloc[1:, [id_idx, target_idx]].copy()
_save_step(mkt_12_with_id, "03_mkt_target_month_only_before_scale")

val_col = mkt_12_with_id.columns[1]
mkt_12_with_id[val_col] = pd.to_numeric(mkt_12_with_id[val_col], errors="coerce") * 1000
_save_step(mkt_12_with_id, "04_mkt_target_month_scaled")

print("MKT: ")
print(mkt_12_with_id)

print("Detected datetime columns:")
print(pd.to_datetime(mkt.columns, errors="coerce"))

last_year_target_dt = pd.Timestamp(year=last_year, month=month_num, day=1)
last_year_match = np.where(cols_dt == last_year_target_dt)[0]

if len(last_year_match) == 0:
    raise ValueError(f"Column for {last_year_target_dt} not found.")

last_year_target_idx = int(last_year_match[0])

mkt_last_year_with_id = mkt.iloc[1:, [id_idx, last_year_target_idx]].copy()
_save_step(mkt_last_year_with_id, "05_mkt_last_year_before_scale")

last_year_val_col = mkt_last_year_with_id.columns[1]
mkt_last_year_with_id[last_year_val_col] = pd.to_numeric(
    mkt_last_year_with_id[last_year_val_col], errors="coerce"
) * 1000
_save_step(mkt_last_year_with_id, "06_mkt_last_year_scaled")

# === Read LUX report and normalize to long format ==================================
lux = pd.read_excel(lux, sheet_name="Market Conso Data", header = 1)
_save_step(lux, "06_lux_raw_sheet")

lux = lux.iloc[:, :18]  # keep first 18 columns (assumes wide months start here)
_save_step(lux, "07_lux_first18cols")

lux = p.normalize_df(
    lux,
    id_vars=lux.columns[:6],  # identifier columns (first 6)
    melt=True, var_name="Month", value_name="Value", detect_months=True  # melt months to rows
)
_save_step(lux, "08_lux_normalized_long")

# Filter out LUX 2025 + Group=Total + Department Stores Online/Offline
lux_report_2025 = lux[
    (lux["Year"] == 2025)
    & (lux["Group"] == "Total")
    & (lux["Channel"].isin(["Department Stores (Online)", "Department Stores (Offline)"]))
]
_save_step(lux_report_2025, "09_lux_2025_deptstores_total")

print(lux_report_2025)

# === WOB from BR: share-of-Category inside each (Channel, Year, Quarter) ===========
w_actual = d.calc_wob_from_totals(
    br,
    by=["Channel", "Year", "Quarter"],  # bucket: Channel x Year x Quarter
    sub_dim="Category",                 # split across Category within bucket
    value_cols=["ActualQuarter"],       # compute shares on 'ActualQuarter'
    as_percent=False,                   # return in fraction (0..1)
    keep_subtotals=True,                # keep 'Subtotal' rows if present
    return_wob_only=True,               # keep only keys + wob columns
)
_save_step(w_actual, "10_w_actual_wob_from_BR")

# === WOB from LUX 2025: share-of-Category inside each (Channel, Year, Month) =======
departmental_store = d.calc_wob_from_totals(
    lux_report_2025,
    by=["Channel", "Year", "Month"],    # bucket: Channel x Year x Month
    sub_dim="Category",                 # split across Category
    value_cols=["Value"],               # compute shares on 'Value'
    as_percent=False,
    keep_subtotals=True,
    return_wob_only=True,
)
departmental_store_2025 = departmental_store.copy()
departmental_store["Year"] = year       # reuse the 2025 mix as a proxy for selected year
_save_step(departmental_store, "11_departmental_store_wob_to_2026")

print(departmental_store)

# ----------------------------------------------------------------------
# STEP 12: Month explode helper (quarter -> months), then backfill quarters
#          NOTE: requires 'year' variable in scope (e.g., year = 2026)
# ----------------------------------------------------------------------
def backfill_quarters_from_prev_year(
    out: pd.DataFrame,
    selected_year: int,
    key_cols=("Category", "Channel"),
    start_quarter: str | None = None,   # e.g. "Q2"; if None, fill Q1..Q4
) -> pd.DataFrame:
    """
    Ensure that for each combo in key_cols, the selected_year has all quarters
    from start_quarter..Q4 (inclusive). If a quarter is missing, copy rows from
    (selected_year - 1) of the same quarter & keys, set Year=selected_year,
    and mark 'filled_from_prev_year' = True.
    """
    out = out.copy()

    QLIST = ["Q1", "Q2", "Q3", "Q4"]

    # --- normalize Quarter column to 'Q1'..'Q4'
    q = (out["Quarter"].astype(str).str.strip().str.upper())
    # Allow plain digits "1".."4" --> map to Q1..Q4
    q = q.str.replace(r"^([1-4])$", r"Q\1", regex=True)
    out["Quarter"] = q

    # --- determine which quarters to consider (full year or starting from a given quarter)
    if start_quarter is None:
        quarters_to_fill = QLIST
    else:
        sq = str(start_quarter).strip().upper()
        if sq in {"1","2","3","4"}:
            sq = f"Q{sq}"
        if sq not in QLIST:
            raise ValueError(f"start_quarter must be one of {QLIST}, got {start_quarter!r}")
        quarters_to_fill = QLIST[QLIST.index(sq):]  # inclusive: from start_quarter to Q4

    # Use combos that exist in selected_year or its previous year (scope of keys to consider)
    base = out.loc[out["Year"].isin([selected_year, selected_year - 1]), list(key_cols)].drop_duplicates()

    additions = []
    for _, key in base.iterrows():
        # mask to current key combination across the dataset
        mask_key = pd.Series(True, index=out.index)
        for c in key_cols:
            mask_key &= out[c].eq(key[c])

        for q in quarters_to_fill:
            # if target-year row for this quarter exists, skip
            has_q_this_year = out.loc[mask_key & (out["Year"] == selected_year) & (out["Quarter"] == q)]
            if len(has_q_this_year) > 0:
                continue

            # else, try to clone from previous year, same quarter
            prev = out.loc[mask_key & (out["Year"] == selected_year - 1) & (out["Quarter"] == q)]
            if len(prev) == 0:
                continue

            cloned = prev.copy()
            cloned["Year"] = selected_year
            cloned["filled_from_prev_year"] = True
            additions.append(cloned)

    # Append any cloned rows to fill missing quarters
    if additions:
        out = pd.concat([out, *additions], ignore_index=True)

    # Stable sort by common keys to keep output deterministic
    sort_cols = [c for c in ["Category", "Year", "Quarter", "Channel", "Month"] if c in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, kind="mergesort").reset_index(drop=True)

    return out

# Quarter -> months mapping (for exploding Q into 3 months)
months_by_quarter = {
    "Q1": ["Jan", "Feb", "Mar"],
    "Q2": ["Apr", "May", "Jun"],
    "Q3": ["Jul", "Aug", "Sep"],
    "Q4": ["Oct", "Nov", "Dec"],
}

# === Prepare 'out' from BR WOB: normalize, expand Quarter to months, backfill ======
out = w_actual.copy()
_save_step(out, "12_out_copy_initial")

out = p.normalize_df(out)  # standardize text/case etc.
_save_step(out, "13_out_normalized")

# Optionally convert numeric column(s) to numeric if needed
# out = p.clean_to_numeric(out, col = "ActualQuarter_wob")

# Expand quarter to individual months
out["Month"] = out["Quarter"].astype(str).str.strip().map(months_by_quarter)
out = out.explode("Month", ignore_index=True)
_save_step(out, "14_out_exploded_months")

# Backfill missing quarters for the selected year (using previous year's same quarter)
out = backfill_quarters_from_prev_year(
    out,
    selected_year=year,
    key_cols=("Category", "Channel"), 
    start_quarter=current_Quarter,   # start from current_Quarter (e.g., Q3) up to Q4
)
_save_step(out, "15_out_backfilled")

# Remove offline department store rows; they will be replaced by departmental_store mix later
out = out[out["Channel"] != "Department Stores (Offline)"]
_save_step(out, "16_out_drop_ds_offline")

# Keep only the columns needed for merging with MKT and LUX-based mixes
cols_to_keep = ["Category", "Year", "Month", "Channel", "ActualQuarter_wob"]
out = out[cols_to_keep].rename(columns = {"ActualQuarter_wob": "Value_wob"})
_save_step(out, "17_out_trimmed_renamed")

# Append the LUX-based departmental store (Online/Offline) cat-mix (now set to Year 2026)
out = pd.concat([out, departmental_store_2025, departmental_store], ignore_index=True)
_save_step(out, "18_out_plus_departmental_store")

print(out)
_save_step(out, "wob_sg")  # keep original save tag too

# === Clean source channel names and map to unified 'Channel' =======================
src = "Market By Channel"
# Reverse mapping: source "Market By Channel" -> unified 'Channel'
mapping_list_rev = [
    ("Boutique",                   "Boutiques (Offline)"),
    ("Departmental Stores",        "Department Stores (Offline)"),
    ("E-Rtailer (Tangs, etc)",     "Department Stores (Online)"),
    ("Eboutique",                  "E-Boutiques (Online)"),
    ("ETailer (Lazada, etc)",      "E-tailers (Online)"),
    ("Perfumery (Offline)",        "Perfumery(Offline)"),
    ("Sephora (Excl Sephora.com)", "Sephora (Offline)"),
    ("Sephora.com",                "Sephora (Online)"),
]

mapping_dict_rev = dict(mapping_list_rev)  # dict if needed (not used directly below)

def prepare_mkt_channel_values(df: pd.DataFrame, save_prefix: str) -> pd.DataFrame:
    out_mkt = df.copy()
    out_mkt[src] = (
        out_mkt[src].astype("string")
          .str.replace("\u00A0", " ", regex=False)  # replace NBSP with regular space
          .str.replace("*", "", regex=False)        # drop asterisks from labels
          .str.replace(r"\s+", " ", regex=True)     # collapse multiple spaces
          .str.strip()
    )
    _save_step(out_mkt, f"{save_prefix}_cleaned_source_channel_text")

    out_mkt = d.map_row_by_any_exact(
        out_mkt,
        cases=mapping_list_rev,
        new_col="Channel",
        in_cols=[src],
        first_wins=True,
        case_sensitive=False,      # match regardless of case
        inplace=False,
        drop_empty_cols=True
    )
    _save_step(out_mkt, f"{save_prefix}_mapped_to_unified_channel")

    # Drop unmapped rows and rename the kept value column to "Value" (2nd column in current slice)
    out_mkt = out_mkt.dropna(subset=["Channel"]).rename(columns={out_mkt.columns[1]: "Value"})
    _save_step(out_mkt, f"{save_prefix}_value_ready")
    return out_mkt

mkt_12_with_id = prepare_mkt_channel_values(mkt_12_with_id, "19_mkt_current_year")
mkt_last_year_with_id = prepare_mkt_channel_values(mkt_last_year_with_id, "19_mkt_last_year")

print(mkt_12_with_id)

# === Merge: attach WOB (Value_wob) to MKT values and compute estimation ===========
merge_new = (mkt_12_with_id.merge(out, on=["Channel"], how = "left")
             .assign(estimation = lambda x:x["Value"]* x["Value_wob"]))  # estimation = monthly market * WOB share
_save_step(merge_new, "22_merge_mkt_with_wob")

merge_new = merge_new[(merge_new["Year"] == year) & (merge_new["Month"] == current_month)]  # filter to current month & year
_save_step(merge_new, "23_merge_filtered_to_current_month")

merge_last_year = (mkt_last_year_with_id.merge(out, on=["Channel"], how="left")
                   .assign(estimation=lambda x: x["Value"] * x["Value_wob"]))
_save_step(merge_last_year, "23a_merge_last_year_mkt_with_wob")

merge_last_year = merge_last_year[
    (merge_last_year["Year"] == last_year)
    & (merge_last_year["Month"] == current_month)
]
_save_step(merge_last_year, "23b_merge_last_year_filtered_to_same_month")

merge_new = pd.concat([merge_last_year, merge_new], ignore_index=True)
_save_step(merge_new, "23c_merge_with_last_year_and_current_year")

print(merge_new)

# === Sorting order definitions (custom category & channel display order) ===========
cat_order = ["Skincare", "Makeup", "Hair/Others", "Fragrances"]
chan_order = [
    "Boutiques (Offline)", 
    "Department Stores (Offline)", 
    "Department Stores (Online)",
    "E-Boutiques (Online)",
    "E-tailers (Online)",
    "Sephora (Offline)",
    "Sephora (Online)",
    "Tourist (Offline)",
]

# === Add one "Tourist (Offline)" row per Category (if missing) and sort ============
df2 = merge_new.copy()
_save_step(df2, "24_df2_initial_copy")

channel_to_add = "Tourist (Offline)"   # channel to add per Category if not present

num_cols = df2.select_dtypes(include="number").columns.tolist()
rows = []

for row_year in sorted(df2["Year"].dropna().unique()):
    for cat in df2["Category"].dropna().unique():
        # skip if the (Year, Category, "Tourist (Offline)") already exists
        if (
            (df2["Year"] == row_year)
            & (df2["Category"] == cat)
            & (df2["Channel"] == channel_to_add)
        ).any():
            continue

        source = df2.loc[(df2["Year"] == row_year) & (df2["Category"] == cat)]
        if source.empty:
            continue

        # template from the first row of the year/category; then blank out values
        tmpl = source.iloc[0].to_dict()
        for k in tmpl:                      # default blanks for all fields
            tmpl[k] = pd.NA
        tmpl["Year"] = row_year
        tmpl["Category"] = cat
        tmpl["Month"] = current_month
        tmpl["Channel"] = channel_to_add
        # numeric columns default to 0 for consistency, except Year
        for c in num_cols:
            if c != "Year":
                tmpl[c] = 0
        rows.append(tmpl)

# append the constructed rows (if any) back to df2
if rows:
    df2 = pd.concat([df2, pd.DataFrame(rows)], ignore_index=True)
_save_step(df2, "25_df2_after_tourist_insertion")

# apply custom sorted categories/channels
df2["Category"] = pd.Categorical(df2["Category"], categories=cat_order, ordered=True)
df2["Channel"]  = pd.Categorical(df2["Channel"],  categories=chan_order, ordered=True)
_save_step(df2, "26_df2_after_categorical_set")

# final sort and clean index
df2 = df2.sort_values(["Year", "Category", "Channel"]).reset_index(drop=True)
print(df2)
_save_step(df2, "27_df2_sorted_final")

# save final step for validation/debugging (kept from your original)
_save_step(df2, "new")

ROOT_OUT.mkdir(parents=True, exist_ok=True)  # ensure parent output dir exists

final_filename = f"SG_Estimation_{year}_{current_month}_{_ts()}.xlsx"  # rename if you prefer
final_path = ROOT_OUT / final_filename

df2.to_excel(final_path, index=False)
print(f"[final saved] {final_path}")
