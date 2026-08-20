# %%
import os
import pandas as pd
import numpy as np
import datetime as dt

dir = os.getcwd()
# Specify the folders
generated_data_folder = f'{dir}/Generated Data'

# Specify the filename you want to use from the 'Generated Data' folder
desired_generated_data_file = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data\SG CPD&LDB Female&Male Skincare Topline_Final 2025-12-29 Corn (1)_CLEAN_merged.csv"

# Construct the full path to the specified generated data file
generated_data_path = os.path.join(generated_data_folder, desired_generated_data_file)

# Load the generated data
df = pd.read_csv(generated_data_path, header = 0)

# %% [markdown]
# ### Normalize Function

# %%
def normalize_headers_sales_chg(df: pd.DataFrame) -> pd.DataFrame:
    cols = (
        df.columns.astype(str)
          .str.replace('\u00A0', ' ', regex=False)  # NBSP -> space
          .str.replace(r'\s+', ' ', regex=True)     # collapse spaces
          .str.strip()
    )

    new_cols = []
    for c in cols:
        parts = [p.strip() for p in c.split('|')]
        left  = parts[0] if parts else c
        right = parts[1].strip() if len(parts) > 1 else ""

        # If left is an 'Unnamed' placeholder, prefer the right part (if present)
        if left.lower().startswith("unnamed") and right:
            left = right
            right = ""  # we've consumed the right as the name

        # Rule:
        # - normally keep only left (remove everything after |)
        # - if right contains "% chg", set to "<left> Sales Value % Chg"
        if "% chg" in right.lower():
            name = f"{left} % Chg"
        else:
            name = left

        name = name.strip()
        new_cols.append(name if name else c)  # fallback to original if empty

    # Deduplicate
    seen = {}
    uniq = []
    for name in new_cols:
        if name not in seen:
            seen[name] = 0
            uniq.append(name)
        else:
            seen[name] += 1
            uniq.append(f"{name}.{seen[name]}")

    out = df.copy()
    out.columns = uniq
    return out

# %% [markdown]
# ### Matching Function

# %%
import re
import unicodedata as u

# def pick_first_matches_and_remove(
#     df: pd.DataFrame,
#     order_list: list[str],
#     *,
#     take_cols: int = 23,
#     out: pd.DataFrame | None = None,
# ) -> tuple[pd.DataFrame, pd.DataFrame]:
#     """Pick FIRST match for each label (case/space-insensitive), append to `out`,
#     and DROP the matched row immediately so later searches won't find it again."""
#     if df.shape[1] < take_cols:
#         raise ValueError(f"DataFrame has only {df.shape[1]} columns; need ≥ {take_cols}.")
    
#     def norm_label(x: str) -> str:
#         x = str(x)
#         x = u.normalize("NFKC", x)                 # unify unicode variants
#         x = x.replace("\ufeff", "")                # BOM
#         x = x.replace("\u00A0", " ")               # NBSP
#         x = x.replace("\u2007", " ").replace("\u202F", " ")  # other thin spaces
#         x = x.replace("‐", "-").replace("-", "-").replace("–", "-").replace("—", "-")
#         x = re.sub(r"\s+", " ", x)                 # collapse spaces
#         return x.strip().casefold()                # trim + case-insensitive

#     def _build_pos_map(_df: pd.DataFrame) -> dict[str, int]:
#         fc = _df.iloc[:, 0].astype("string").map(norm_label)
#         tmp = fc.reset_index(drop=True).to_frame("label").dropna(subset=["label"]).reset_index().rename(columns={"index": "pos"})
#         return (tmp.drop_duplicates(subset="label", keep="first").set_index("label")["pos"].to_dict())

#     first_pos_by_label = _build_pos_map(df)

#     if out is None or out.empty:
#         out = pd.DataFrame(columns=df.columns[:take_cols])

#     for label in order_list:
#         key = norm_label(label)
#         if key in first_pos_by_label:
#             pos = int(first_pos_by_label[key])
#             out.loc[len(out)] = df.iloc[pos, :take_cols].tolist()

#             # DROP NOW so future searches can’t hit the earlier section’s rows
#             df = df.drop(df.index[pos]).reset_index(drop=True)

#             # Rebuild map because positions changed
#             first_pos_by_label = _build_pos_map(df)
#         else:
#             # No match: write a placeholder row showing the label we expected
#             out.loc[len(out)] = [label, *([pd.NA] * (take_cols - 1))]

#     return df.reset_index(drop=True), out

def pick_first_matches_and_remove(
    df: pd.DataFrame,
    order_list: list[str],
    *,
    take_cols: int = 23,
    out: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Pick FIRST match for each label (case/space-insensitive), except:
       - 'OTHERS' → pick the LAST match.
       Append to `out`, and DROP the matched row immediately."""

    if df.shape[1] < take_cols:
        raise ValueError(f"DataFrame has only {df.shape[1]} columns; need ≥ {take_cols}.")

    # ------------ Normalizer --------------------------------------------------
    def norm_label(x: str) -> str:
        x = str(x)
        x = u.normalize("NFKC", x)
        x = x.replace("\ufeff", "")
        x = x.replace("\u00A0", " ")
        x = x.replace("\u2007", " ").replace("\u202F", " ")
        x = x.replace("‐", "-").replace("-", "-").replace("–", "-").replace("—", "-")
        x = re.sub(r"\s+", " ", x)
        return x.strip().casefold()

    # ------------ Build FIRST-position map ------------------------------------
    def _build_first_map(_df: pd.DataFrame) -> dict[str, int]:
        fc = _df.iloc[:, 0].astype("string").map(norm_label)
        tmp = (
            fc.reset_index(drop=True)
              .to_frame("label")
              .dropna(subset=["label"])
              .reset_index()
              .rename(columns={"index": "pos"})
        )
        return (
            tmp.drop_duplicates(subset="label", keep="first")
               .set_index("label")["pos"]
               .to_dict()
        )

    # ------------ Added this part: Build LAST-position map -------------------------------------
    def _build_last_map(_df: pd.DataFrame) -> dict[str, int]:
        fc = _df.iloc[:, 0].astype("string").map(norm_label)
        tmp = (
            fc.reset_index(drop=True)
              .to_frame("label")
              .dropna(subset=["label"])
              .reset_index()
              .rename(columns={"index": "pos"})
        )
        return (
            tmp.drop_duplicates(subset="label", keep="last")
               .set_index("label")["pos"]
               .to_dict()
        )

    # initial maps
    first_pos = _build_first_map(df)
    last_pos  = _build_last_map(df)

    if out is None or out.empty:
        out = pd.DataFrame(columns=df.columns[:take_cols])

    # ======================================================================
    # Main loop
    # ======================================================================
    for label in order_list:
        key = norm_label(label)

        # special rule → pick LAST match for OTHERS
        if key == "others":
            pos_map = last_pos
        else:
            pos_map = first_pos

        if key in pos_map:
            pos = int(pos_map[key])
            out.loc[len(out)] = df.iloc[pos, :take_cols].tolist()

            # remove the matched row
            df = df.drop(df.index[pos]).reset_index(drop=True)

            # rebuild maps
            first_pos = _build_first_map(df)
            last_pos  = _build_last_map(df)

        else:
            # no match → placeholder row
            out.loc[len(out)] = [label, *([pd.NA] * (take_cols - 1))]

    return df.reset_index(drop=True), out


# %% [markdown]
# ### Men CPD

# %% Men CPD
# By Function
function_order = [
    "OIL CONTROL",
    "AGING",
    "BRIGHTENING",
    "HYDRATION + BASIC",
    "Hydration",
    "Basic",
]

# By Category
category_order = [
    "CLEANSER",
    "MOISTURISER",
    "FACE MASK",
    "SCRUB",
]

# By Brands (incl. sub-lines in the given order)
brand_order = [
    "TOTAL CPD",
    "Garnier",
    "AcnoFight",
    "Power White",
    "TurboLight",
    "Turbo Bright",
    "LOREAL DERMO EXPERTISE",
    "Expert", 
    "Hydenergetic",
    "Men",
    "Nivea",
    "Anti Aging",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Gatsby",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Men's Biore",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Shokobutsu",
    "Eversoft",
    "NanoWhite",
    "Sukin",
    "OTHER BRANDS",
]

# %% [markdown]
# ### Produce SG Men CPD Topline

# %%
# ---------------------------------------
# Load source CSV and normalize headers
# ---------------------------------------
df = normalize_headers_sales_chg(df)

# ---------------------------------------
# Slice “Sales Value” (first 23 cols) and “Unit” block (key + cols 47..68)
# ---------------------------------------
men_sales_value = df.iloc[:, :23].copy()                 # 0..22
men_sales_unit  = df.iloc[:, np.r_[0, 47:69]].copy()     # key + 47..68

# ---------------------------------------
# Align join key on first column (rename right if needed), then LEFT-merge
# ---------------------------------------
key = men_sales_value.columns[0]
if men_sales_unit.columns[0] != key:
    men_sales_unit = men_sales_unit.rename(columns={men_sales_unit.columns[0]: key})

# add suffixes before concatenation so names are unique
left  = men_sales_value.rename(columns=lambda c: c if c == key else f"{c}_value")
right = men_sales_unit.drop(columns=[key]).rename(columns=lambda c: f"{c}_unit")

merged = pd.concat([left, right], axis=1)

# ---------------------------------------
# Build the “TOTAL MALE SKINCARE” row from specific ranges:
#   - Value indices: 24..45
#   - Unit  indices: 70..91
# Scale by 1000 EXCEPT columns containing "Sales Value % Chg"
# ---------------------------------------
value_idx = list(range(24, 46))   # 24..45
unit_idx  = list(range(70, 92))   # 70..91

merged_total = [
    (pd.to_numeric(df.iloc[0, c], errors="coerce") /
     (1 if "% chg" in str(df.columns[c]).lower() else 1000))
    for c in (value_idx + unit_idx)
]

# ---------------------------------------
# Construct column headers for output: key + value headers + unit headers
# ---------------------------------------
cols = [key] \
     + [f"{c}_value" for c in df.columns[value_idx]] \
     + [f"{c}_unit"  for c in df.columns[unit_idx]]

# ---------------------------------------
# Initialize output DataFrame and append TOTAL row
# ---------------------------------------
out = pd.DataFrame(columns=cols)
out.loc[len(out)] = ["TOTAL MALE SKINCARE", *merged_total]

# ---------------------------------------
# By Function
# ----------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Function", *([pd.NA]*(len(out.columns) - 1))]

# ---------------------------------------
# Pick first matches for function_order; append to `out`, remove from `merged`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged,
    function_order,
    take_cols=45,
    out=out,          # or pass an existing DataFrame with matching columns
)

# ---------------------------------------
# Spacer + By Category header
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Category", *([pd.NA]*(len(out.columns) - 1))]

# ---------------------------------------
# Pick first matches for category_order; append to `out`, remove from `merged`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged, 
    category_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# Spacer + By Brands header
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Brands", *([pd.NA]*(len(out.columns) - 1))]

# ---------------------------------------
# Pick first matches for brand_order; append to `out`, remove from `merged`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged, 
    brand_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# Final alias for downstream usage
# ---------------------------------------
male_cpd = out


# %% [markdown]
# ### Female Skincare (Exc. Mass Medic)

# %% Female Skincare (Exc. Mass Medic)
# ---------------------------------------
# Function order
# ---------------------------------------
function_order = [
    "ANTI AGING",
    "BASIC",
    "HYDRATING",
    "OIL CONTROL",
    "PURIFYING",
    "WHITEN",
]

# ---------------------------------------
# Category order
# ---------------------------------------
category_order = [
    "CLEANSER",
    "MOISTURISER",
    "EYE MOISTURIZER",
    "MOISTURIZER CREAM",
    "MOISTURIZER GEL",
    "MOISTURIZER OTHERS",
    "MOISTURIZER LOTION",
    "MOISTURIZER TREATMENT",
    "MOISTURIZER WATER",
    "MASK",
    "SCRUB",
    "TONER",
    "MOISTURIZER ESSENCE",
    "MAKE UP REMOVER",
]

# ---------------------------------------
# Brand order (as provided; includes sub-rows like ANTI AGING, BASIC, etc.)
# ---------------------------------------
brand_order = [
    "TOTAL CPD",
    "Garnier",
    "ANTI AGING",
    "BASIC",
    "HYDRATING",
    "OIL CONTROL",
    "Whiten",
    "LOREAL DERMO EXPERTISE",
    "ANTI AGING",
    "BASIC",
    "HYDRATING",
    "OIL CONTROL",
    "WHITEN",
    "Maybelline",
    "Basic",
    "HADA LABO",
    "Bio Essence",
    "Eversoft",
    "Biore",
    "Olay",
    "Skintific",
    "ST IVES",
    "Nutox",
    "Safi",
    "Senka",
    "OTHERS",
    "EXCLUSIVE + PRIVATE LABELS",
]


# %% [markdown]
# ### Produce Female Skincare (Exc. Mass Medic) Topline Report

# %%
# ---------------------------------------
# Slice female Sales Value and Unit blocks for FEMALE SKINCARE
# Value block: cols 119..141  (23 columns)
# Unit  block: key at col 119 + cols 166..187
# ---------------------------------------
female_sales_value = df.iloc[:, 119:142].copy()               # 0..22
female_sales_unit  = df.iloc[:, np.r_[119, 166:188]].copy()   # key + 47..68

# ---------------------------------------
# Ensure both frames share the SAME join key (first column name)
# ---------------------------------------
key = female_sales_value.columns[0]
if female_sales_unit.columns[0] != key:
    female_sales_unit = female_sales_unit.rename(columns={female_sales_unit.columns[0]: key})

# ---------------------------------------
# Left-merge Unit block onto Value block (keep all Value rows)
# Suffixes make duplicate month headers distinguishable (_value vs _unit)
# ---------------------------------------
key = female_sales_value.columns[0]
if female_sales_unit.columns[0] != key:
    female_sales_unit = men_sales_unit.rename(columns={female_sales_unit.columns[0]: key})

# add suffixes before concatenation so names are unique
left  = female_sales_value.rename(columns=lambda c: c if c == key else f"{c}_value")
right = female_sales_unit.drop(columns=[key]).rename(columns=lambda c: f"{c}_unit")

merged = pd.concat([left, right], axis=1)

# ---------------------------------------
# Build the “TOTAL FEMALE SKINCARE (EXC. MASS MEDIC)” row from specific ranges:
#   - Value indices: 143..164
#   - Unit  indices: 189..210
# Scale by 1000 EXCEPT columns whose header contains "% chg"
# ---------------------------------------
value_idx = list(range(143, 165))   # 24..45
unit_idx  = list(range(189, 211))   # 70..91

merged_total = [
    (pd.to_numeric(df.iloc[0, c], errors="coerce") /
     (1 if "% chg" in str(df.columns[c]).lower() else 1000))
    for c in (value_idx + unit_idx)
]

# ---------------------------------------
# Construct column headers for output:
#   key + value headers (with _value) + unit headers (with _unit)
# ---------------------------------------
cols = [key] \
     + [f"{c}_value" for c in df.columns[value_idx]] \
     + [f"{c}_unit"  for c in df.columns[unit_idx]]

# ---------------------------------------
# Initialize output DataFrame and append TOTAL FEMALE row
# ---------------------------------------
out = pd.DataFrame(columns=cols)
out.loc[len(out)] = ["TOTAL FEMALE SKINCARE (EXC. MASS MEDIC)", *merged_total]

print(len(out.columns))  # sanity: number of columns in the output

# ---------------------------------------
# By Function section: add a blank spacer row
# Then pick first matches (by function_order) from `merged`
# and append them to `out` while removing matched rows from `merged`
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)

merged, out = pick_first_matches_and_remove(
    merged,
    function_order,
    take_cols=45,
    out=out,          # keep writing into the same output sheet
)

# ---------------------------------------
# By Category section: add a blank spacer row
# Then append first matches (by category_order), removing them from `merged`
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)

merged, out = pick_first_matches_and_remove(
    merged, 
    category_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# By Brands header + spacer, then append first matches by brand_order
# (pick_first_matches_and_remove appends rows to `out` and removes from `merged`)
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Brands", *([pd.NA]*(len(out.columns) - 1))]

# ---------------------------------------
# Pick first matches for brand_order; append to `out`, remove from `merged`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged, 
    brand_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# Final alias for downstream usage
# ---------------------------------------
female_cpd = out


# %% [markdown]
# ### SG LDB

# %% SG LDB
# ---------------------------------------
# Function order (Female)
# ---------------------------------------
function_order = [
    "Whiten",
    "FEMALE HYDRATION + BASIC",
    "Basic",
    "Hydrating",
    "FEMALE ANTI-AGING",
    "FEMALE OIL-CONTROL",
]

# ---------------------------------------
# Category order (Female)
# ---------------------------------------
category_order = [
    "CLEANSER",
    "MOISTURISER",
    "MASK",
    "SCRUB",
    "Toner + Moisturizer Water + Moisturizer Lotion",
    "MOISTURIZER LOTION",
    "MOISTURIZER WATER",
    "TONER",
    "MAKE UP REMOVER",
]

# ---------------------------------------
# Brand order (LDB Female)
# ---------------------------------------
brand_order = [
    "TOTAL LDB",

    "La Roche Posay",
    "Anti Aging",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Make Up Remover",

    "Vichy",
    "Anti Aging",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Make Up Remover",

    "Cerave",
    "Anti Aging",
    "Whiten",
    "Basic",
    "Hydrating",
    "Oil Control",

    "Avene",
    "Anti Aging",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Make Up Remover",

    "Neutrogena",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Make Up Remover",

    "Eucerin",
    "Anti Aging",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Make Up Remover",

    "Cetaphil",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Whiten",
    "Make Up Remover",

    "Bioderma",
    "Anti Aging",
    "Basic",
    "Hydrating",
    "Oil Control",
    "Make Up Remover",

    "Bionike",
    "Curel",
    "Evans",
    "Placentor Vegetal",
    "Sebamed",
    "Uriage",
    "Ego",
    "Mustela",
    "Physiogel",
    "QV",
    "Topicrem",
    "OTHERS",
    "EXCLUSIVE BRANDS",
]

# %% [markdown]
# ### Produce SG LDB Topline

# %%
# ---------------------------------------
# Slice female Sales Value and Unit blocks for FEMALE SKINCARE
# Value block: cols 119..141  (23 columns)
# Unit  block: key at col 119 + cols 166..187
# ---------------------------------------
female_sales_value = df.iloc[:, 215:238].copy()               # 0..22
female_sales_unit  = df.iloc[:, np.r_[215, 262:284]].copy()   # key + 47..68

# ---------------------------------------
# Ensure both frames share the SAME join key (first column name)
# ---------------------------------------
key = female_sales_value.columns[0]
if female_sales_unit.columns[0] != key:
    female_sales_unit = men_sales_unit.rename(columns={female_sales_unit.columns[0]: key})

# add suffixes before concatenation so names are unique
left  = female_sales_value.rename(columns=lambda c: c if c == key else f"{c}_value")
right = female_sales_unit.drop(columns=[key]).rename(columns=lambda c: f"{c}_unit")

merged = pd.concat([left, right], axis=1)

# ---------------------------------------
# Build the “TOTAL FEMALE SKINCARE (EXC. MASS MEDIC)” row from specific ranges:
#   - Value indices: 143..164
#   - Unit  indices: 189..210
# Scale by 1000 EXCEPT columns whose header contains "% chg"
# ---------------------------------------
value_idx = list(range(239, 261))   # 24..45
unit_idx  = list(range(285, 307))   # 70..91

merged_total = [
    (pd.to_numeric(df.iloc[0, c], errors="coerce") /
     (1 if "% chg" in str(df.columns[c]).lower() else 1000))
    for c in (value_idx + unit_idx)
]

# ---------------------------------------
# Construct column headers for output:
#   key + value headers (with _value) + unit headers (with _unit)
# ---------------------------------------
cols = [key] \
     + [f"{c}_value" for c in df.columns[value_idx]] \
     + [f"{c}_unit"  for c in df.columns[unit_idx]]

# ---------------------------------------
# Initialize output DataFrame and append TOTAL FEMALE row
# ---------------------------------------
out = pd.DataFrame(columns=cols)
out.loc[len(out)] = ["TOTAL FEMALE SKINCARE (EXC. MASS MEDIC)", *merged_total]

print(len(out.columns))  # sanity: number of columns in the output

# ---------------------------------------
# By Function section: add a blank spacer row
# Then pick first matches (by function_order) from `merged`
# and append them to `out` while removing matched rows from `merged`
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)

merged, out = pick_first_matches_and_remove(
    merged,
    function_order,
    take_cols=45,
    out=out,          # keep writing into the same output sheet
)

# ---------------------------------------
# By Category section: add a blank spacer row
# Then append first matches (by category_order), removing them from `merged`
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)

merged, out = pick_first_matches_and_remove(
    merged, 
    category_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# By Brands header + spacer, then append first matches by brand_order
# (pick_first_matches_and_remove appends rows to `out` and removes from `merged`)
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Brands", *([pd.NA]*(len(out.columns) - 1))]

# ---------------------------------------
# Pick first matches for brand_order; append to `out`, remove from `merged`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged, 
    brand_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# Final alias for downstream usage
# ---------------------------------------
female_ldb = out


# %% [markdown]
# ### Save final report

# %%
from pathlib import Path
import datetime as dt
import pandas as pd  # only needed for ExcelWriter in the combined-file example

dir = os.getcwd()

base = Path(dir).resolve()
# If CWD is ".../Topline Report/Topline Report Script", go up one to ".../Topline Report"
topline_root = base.parent if base.name.lower() == "topline report script" else base

# Month tag
month_tag_iso = dt.datetime.now().strftime("%Y-%m")

# Target: ...\Topline Report\Generated Report\Topline Report\2025-11
dest = topline_root / "Generated Report" / "Topline Report" / month_tag_iso
dest.mkdir(parents=True, exist_ok=True)

# --- 3) Save DataFrames to Excel in that folder ---
_ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

with pd.ExcelWriter(dest / f"Topline_SG_{month_tag_iso}_{_ts}.xlsx", engine="openpyxl") as writer:
    female_ldb.to_excel(writer, sheet_name="Female LDB", index=False)
    female_cpd.to_excel(writer, sheet_name="Female CPD", index=False)
    male_cpd.to_excel(writer, sheet_name="Male CPD", index=False)
print(f"[saved] {dest / f'Topline_{month_tag_iso}_{_ts}.xlsx'}")



