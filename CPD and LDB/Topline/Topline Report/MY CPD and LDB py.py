# %%
import os
import pandas as pd
import numpy as np
from pathlib import Path
import datetime as dt

dir = os.getcwd()
# Specify the folders
generated_data_folder = f'{dir}/Generated Data'
topline_template_folder = f'{dir}/Template/Topline Template'
generated_report_folder = f'{dir}/Generated Report'

desired_generated_data_file = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data\MY CPD Male Skincare Topline (Use This) 2025-12-26 Corn (1)_CLEAN_merged.csv" # Male SC
desired_generated_data_file1 = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data\MY CPD Female Skincare (MASS Only) Topline (Use This) 2025-12-26 Corn (1)_CLEAN_merged.csv" #Mass
desired_generated_data_file2 = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data\MY CPD Female Skincare (MASS + MASS MEDIC) Topline Corn (1)_CLEAN_merged.csv" #Mass + Mass Medic
desired_generated_data_file3 = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data\MY CPD Cosmetics Topline (Use This) Corn 251226 (1)_CLEAN_merged.csv" #Cosmetics
desired_generated_data_file4 = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data\MY LDB Skincare Topline (Use this) Corn (1)_CLEAN_merged.csv" #LDB

# Construct the full paths to the specified generated data files
male = os.path.join(generated_data_folder, desired_generated_data_file) # Male
female_mass = os.path.join(generated_data_folder, desired_generated_data_file1) # Female Mass Only
female_mass_medic = os.path.join(generated_data_folder, desired_generated_data_file2) # Female with Mass Medic
cosmetic = os.path.join(generated_data_folder, desired_generated_data_file3) # Cosmetic
ldb = os.path.join(generated_data_folder, desired_generated_data_file4) #ldb

print(male)
print(female_mass)
print(female_mass_medic)
print(cosmetic)
print(ldb)

# Load the generated data from the three sources
male = pd.read_csv(male, header = 0)
female_mass = pd.read_csv(female_mass, header = 0)
female_mass_medic = pd.read_csv(female_mass_medic, header = 0)
cosmetic = pd.read_csv(cosmetic, header = 0)
ldb = pd.read_csv(ldb, header = 0)

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

import unicodedata as u
from itertools import zip_longest

def show_codes(s):
    """Print raw, repr, and each character's code point + Unicode category."""
    s = str(s)
    print("RAW :", s)
    print("REPR:", repr(s))
    print("CODES:",
          [f"{c} -> U+{ord(c):04X} ({u.category(c)})" for c in s])

def diff_codes(a, b):
    """Side-by-side per-char diff with code points."""
    a, b = str(a), str(b)
    print(f"LEN A={len(a)}  LEN B={len(b)}")
    for i, (ca, cb) in enumerate(zip_longest(a, b, fillvalue="∅")):
        if ca != cb:
            ca_code = f"U+{ord(ca):04X} ({u.category(ca)})" if ca != "∅" else "∅"
            cb_code = f"U+{ord(cb):04X} ({u.category(cb)})" if cb != "∅" else "∅"
            print(f"[{i:02}] {ca!r:<3} {ca_code:<16} | {cb!r:<3} {cb_code}")

# target = "Anti Acne"              # from your list
# # cell   = df.iloc[row_idx, 0]      # value from the sheet

# show_codes(target)
# # show_codes(cell)
# # diff_codes(target, cell)

def pick_first_matches_and_remove(
    df: pd.DataFrame,
    order_list: list[str],
    *,
    take_cols: int = 23,
    out: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Pick FIRST match for each label (case/space-insensitive), append to `out`,
    and DROP the matched row immediately so later searches won't find it again."""
    if df.shape[1] < take_cols:
        raise ValueError(f"DataFrame has only {df.shape[1]} columns; need ≥ {take_cols}.")
    
    def norm_label(x: str) -> str:
        x = str(x)
        x = u.normalize("NFKC", x)
        # remove BOM/NBSP/thin spaces you had
        x = x.replace("\ufeff","").replace("\u00A0"," ").replace("\u2007"," ").replace("\u202F"," ")
        # remove soft hyphen + ALL format/control chars (covers \u200B, \u2060, etc.)
        x = x.replace("\u00AD", "")  # soft hyphen
        x = "".join(ch for ch in x if u.category(ch) not in {"Cf","Cc"})
        # unify dashes
        x = x.replace("–","-").replace("—","-").replace("‐","-")
        # collapse whitespace
        x = re.sub(r"\s+", " ", x).strip()
        return x.casefold()

    def _build_pos_map(_df: pd.DataFrame) -> dict[str, int]:
        fc = _df.iloc[:, 0].astype("string").map(norm_label)
        tmp = fc.reset_index(drop=True).to_frame("label").dropna(subset=["label"]).reset_index().rename(columns={"index": "pos"})
        return (tmp.drop_duplicates(subset="label", keep="first").set_index("label")["pos"].to_dict())

    first_pos_by_label = _build_pos_map(df)

    if out is None or out.empty:
        out = pd.DataFrame(columns=df.columns[:take_cols])

    for label in order_list:
        key = norm_label(label)
        target = key
        if key in first_pos_by_label:
            pos = int(first_pos_by_label[key])
            out.loc[len(out)] = df.iloc[pos, :take_cols].tolist()

            # DROP NOW so future searches can’t hit the earlier section’s rows
            df = df.drop(df.index[pos]).reset_index(drop=True)

            # Rebuild map because positions changed
            first_pos_by_label = _build_pos_map(df)
        else:
            # No match: write a placeholder row showing the label we expected
            out.loc[len(out)] = [label, *([pd.NA] * (take_cols - 1))]

            # # --- DEBUG: print what we searched for
            # print(f"[MISS] expected: {label!r}  | norm -> {key!r}")
            # show_codes(label)  # raw + repr + code points of your order_list label

            # # Build normalized view of the first column (what we're matching against)
            # col0_raw  = df.iloc[:, 0].astype("string")
            # col0_norm = col0_raw.map(norm_label)

            # # Show a few candidate rows from the sheet to compare (adjust the range if you want)
            # print("[CANDIDATES] First few first-column values with codes and diffs:")
            # for i in col0_norm.index[:5]:
            #     print(f"\nRow {i}:")
            #     print("RAW CELL:")
            #     show_codes(col0_raw.iloc[i])
            #     print("DIFF (normalized expected vs normalized cell):")
            #     diff_codes(key, col0_norm.iloc[i])

            # # (Optional) print the distinct normalized keys we actually have (first 20)
            # have_keys_sample = list(dict.fromkeys(col0_norm.tolist()))[:20]
            # print("\n[SAMPLE OF NORMALIZED KEYS IN DATA]:", have_keys_sample)

    return df.reset_index(drop=True), out

# %% Male Skincare
# ### Male Skincare

# %% Male Skincare
# ---------------------------------------
# Function order (Male CPD)
# ---------------------------------------
function_order = [
    "OIL CONTROL + ACNE",
    "WHITENING",
    "HYDRATION + BASIC",
    "Basic",
    "Hydration",
    "ANTI AGE",
]

# ---------------------------------------
# Category order (Male CPD)
# ---------------------------------------
category_order = [
    "CLEANSER",
    "MOISTURE",
    "FACE MASK",
]

# ---------------------------------------
# Brand order (Male CPD)
# (Includes brand headers and their sub-lines where provided)
# ---------------------------------------
brand_order = [
    "TOTAL CPD",

    "Garnier",
    "Turbo Light Oil Control",
    "Turbo Light",
    "Turbo Bright",
    "Acno Fight",
    "Power White",

    "L'Oreal Paris",
    "Men Expert",

    "Nivea",
    "Whitening",
    "Anti Acne",
    "Basic",
    "Hydration",
    "Anti Age",

    "Biore",
    "Anti Acne",
    "Basic",
    "Whitening",

    "Gatsby",
    "Basic",
    "Anti Acne",
    "Whitening",

    "Kahf",
    "Anti Acne",
    "Whitening",

    "Nano White",
    "Dashing",
    "Safi",
    "Bad Lab",
    "Fair & Handsome",
    "OTHER BRANDS",
    "Exclusive brand + Private label",
]


# %% [markdown]
# ### Produce Male Skincare Topline (MY)

# %% Male Skincare
# ---------------------------------------
# Normalize column headers (e.g., strip suffixes, unify "% Chg" labels, etc.)
# ---------------------------------------
df = normalize_headers_sales_chg(male)

# ---------------------------------------
# Split the sheet into two blocks:
#  - men_sales_value: first 23 columns (labels + VALUE measures)
#  - men_sales_unit : key column + UNIT measures (cols 47..68)
# ---------------------------------------
men_sales_value = df.iloc[:, 0:23].copy()
men_sales_unit  = df.iloc[:, np.r_[0, 47:69]].copy() 

# ---------------------------------------
# Align the join key (first column header) across both frames
# ---------------------------------------
key = men_sales_value.columns[0]
if men_sales_unit.columns[0] != key:
    men_sales_unit = men_sales_unit.rename(columns={men_sales_unit.columns[0]: key})

# ---------------------------------------
# Give each block distinct suffixes BEFORE concatenation so names are unique
#  - left  : keep key as-is; all other cols get "_value"
#  - right : drop the duplicate key; remaining cols get "_unit"
# Then place them side-by-side with concat (no merge cartesian expansion)
# ---------------------------------------
left  = men_sales_value.rename(columns=lambda c: c if c == key else f"{c}_value")
right = men_sales_unit.drop(columns=[key]).rename(columns=lambda c: f"{c}_unit")

merged = pd.concat([left, right], axis=1)

print(merged)

# ---------------------------------------
# Build the "TOTAL MALE SKINCARE" row from specific column slices on the raw df:
#   value_idx = 24..45, unit_idx = 70..91
# Scale by 1000 for level metrics, but DO NOT scale columns that are "% chg"
# ---------------------------------------
value_idx = list(range(24, 46))   # 24..45 (inclusive)
unit_idx  = list(range(70, 92))   # 70..91 (inclusive)

merged_total = [
    (pd.to_numeric(df.iloc[0, c], errors="coerce") /
     (1 if "% chg" in str(df.columns[c]).lower() else 1000))
    for c in (value_idx + unit_idx)
]

# ---------------------------------------
# Construct output frame schema: key + labeled value/unit columns
# ---------------------------------------
cols = [key] \
     + [f"{c}_value" for c in df.columns[value_idx]] \
     + [f"{c}_unit"  for c in df.columns[unit_idx]]

out = pd.DataFrame(columns=cols)

# Seed the first row: overall total
out.loc[len(out)] = ["TOTAL MALE SKINCARE", *merged_total]

# ---------------------------------------
# Section: By Function
#   - Insert a blank spacer row, then a section header row
#   - Pull first matches for each label in function_order and remove from `merged`
#   - take_cols=45 to match your intended leading columns width
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Function", *([pd.NA]*(len(out.columns) - 1))]

merged, out = pick_first_matches_and_remove(
    merged,
    function_order,
    take_cols=45,
    out=out,          # append results into existing `out`
)

# ---------------------------------------
# Section: By Category
#   - Spacer row + header
#   - Append first matches (and drop them from `merged`)
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Category", *([pd.NA]*(len(out.columns) - 1))]

merged, out = pick_first_matches_and_remove(
    merged, 
    category_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# Section: By Brands
#   - Spacer row + header
#   - Append first matches (and drop them from `merged`)
# ---------------------------------------
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = ["By Brands", *([pd.NA]*(len(out.columns) - 1))]

merged, out = pick_first_matches_and_remove(
    merged, 
    brand_order, 
    take_cols=45,
    out = out,
)

# ---------------------------------------
# Final DataFrame for MEN CPD
# ---------------------------------------
male_cpd = out


# %% [Mass Only]
# ### Female Skincare (Mass Only)

# %% Female Skincare (Mass Only)
# function_order = [
#     "BRIGHTENING",	
#     "BASIC + HYDRATION",	
#     "BASIC", 	
#     "HYDRATION",
#     "ANTI AGE",	
#     "ANTI ACNE",	
#     "SENSITIVE"	

# ]

function_order = [
    "BRIGHTENING",		
    "BASIC", 	
    "HYDRATION",
    "ANTI AGE",	
    "ANTI ACNE",	
    "SENSITIVE"	

]

category_order = [
    "CLEANSER",	
    "MOISTURE",	
    "CREAM/GEL/OIL/LOTION",
    "ESSENCE/SERUM",
    "COSMETIC / TREATMENT WATER",
    "THERMAL / SPRING WATER",
    "AMPOULE",
    "TONER",
    "FACE MASK",
    "FACE EYE",
    "MAKE UP REMOVER"	

]

brand_order = [
    "TOTAL CPD",
    "Garnier",
    "BRIGHTENING",
    "ANTI ACNE",
    "HYDRATION",
    "MAKE UP REMOVER",
    "L'oreal Paris",
    "ANTI ACNE",
    "ANTI AGE",
    "BASIC",
    "HYDRATION",
    "WHITENING",
    "Maybelline",
    "MAKE UP REMOVER",
    "SAFI",
    "BIO ESSENCE",
    "HADA LABO",
    "AIKEN",
    "OLAY",
    "NANO WHITE",
    "SKINTIFIC",
    "TORRIDEN",
    "NUTOX",
    "HIMALAYA",
    "SIMPLE",
    "CLINELLE",
    "ST.IVES",
    "NIVEA",
    "WARDAH",
    "OTHERS",
    "EXCLUSIVE + PRIVATE LABELS",
]


# %% Female Skincare (Mass Only)
# ### Produce Female Skincare Topline (Mass Only)

# %% Female Skincare (Mass Only)
# ---------------------------------------
# Normalize column headers (e.g., strip suffixes, unify "% Chg" labels, etc.)
# ---------------------------------------
df = normalize_headers_sales_chg(female_mass)

# ---------------------------------------
# Slice female Sales Value and Unit blocks for FEMALE SKINCARE
# Value block: cols 119..141  (23 columns)
# Unit  block: key at col 119 + cols 166..187
# ---------------------------------------
female_sales_value = df.iloc[:, :23].copy()               # 0..22
female_sales_unit  = df.iloc[:, np.r_[0, 47:69]].copy()   # key + 47..68

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
value_idx = list(range(24, 46))   # 24..45
unit_idx  = list(range(70, 92))   # 70..91

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

# %% [Mass + Medic Mass]
# ### Female Skincare (Plus Mass Medic)

# %% [Mass + Medic Mass]
# function_order = [
#     "WHITENING",	
#     "BASIC+HYDRATION",	
#     "BASIC", 	
#     "HYDRATION",
#     "ANTI AGE",	
#     "ANTI ACNE",	
#     "SENSITIVE"	
# ]

function_order = [
    "ANTI ACNE",
    "ANTI AGE",
    "BASIC", 	
    "HYDRATION",		
    "SENSITIVE",
    "WHITENING"
]

category_order = [
    "CLEANSER",	
    "MOISTURE",	
    "CREAM/GEL/OIL/LOTION",
    "ESSENCE/SERUM",
    "COSMETIC / TREATMENT WATER",
    "THERMAL / SPRING WATER",
    "AMPOULE",
    "TONER",
    "FACE MASK",
    "FACE EYE",
    "MAKE UP REMOVER"	

]

# brand_order = [
#     "TOTAL CPD",
#     "Garnier",
#     "WHITENING",
#     "ANTI ACNE",
#     "HYDRATION",
#     "MAKE UP REMOVER",
#     "L'oreal Paris",
#     "ANTI ACNE",
#     "ANTI AGE",
#     "HYDRATION",
#     "MAKE UP REMOVER",
#     "WHITENING",
#     "Maybelline",
#     "SAFI",
#     "BIO ESSENCE",
#     "HADA LABO",
#     "AIKEN",
#     "OLAY",
#     "NANO WHITE",
#     "NUTOX",
#     "HIMALAYA",
#     "SIMPLE",
#     "CLINELLE",
#     "ST.IVES",
#     "NIVEA",
#     "PHYSIOGEL",
#     "QV",
#     "SEBAMED",
#     "TORRIDEN",
#     "SKINTIFIC",
#     "NEUTROGENA",
#     "CETAPHIL",
#     "WARDAH",
#     "CERAVE",
#     "Others",
#     "EXCLUSIVE BRAND+PRIVATE LABEL",
# ]

brand_order = [
    "TOTAL CPD",
    "Garnier",
    "ANTI ACNE",
    "BLANK 1",
    "HYDRATION",
    "BLANK 2",
    "BLANK 3",
    "WHITENING",
    "MAKE UP REMOVER",
    "Maybelline",
    "L'oreal Paris",
    "ANTI ACNE",
    "ANTI AGE",
    "HYDRATION",
    "MAKE UP REMOVER",
    "WHITENING",
    "SAFI",
    "BIO ESSENCE",
    "HADA LABO",
    "AIKEN",
    "OLAY",
    "NANO WHITE",
    "NUTOX",
    "HIMALAYA",
    "SIMPLE",
    "CLINELLE",
    "ST.IVES",
    "NIVEA",
    "PHYSIOGEL",
    "QV",
    "SEBAMED",
    "TORRIDEN",
    "SKINTIFIC",
    "NEUTROGENA",
    "CETAPHIL",
    "WARDAH",
    "CERAVE",
    "Others",
    "EXCLUSIVE BRAND+PRIVATE LABEL",
]


# %% [Mass + Medic Mass]
# ### Produce Female Skincare Topline (Plus Mass Medic)

# %% [Mass + Medic Mass]
# ---------------------------------------
# Normalize column headers (e.g., strip suffixes, unify "% Chg" labels, etc.)
# ---------------------------------------
df = normalize_headers_sales_chg(female_mass_medic)

# ---------------------------------------
# Slice female Sales Value and Unit blocks for FEMALE SKINCARE
# Value block: cols 119..141  (23 columns)
# Unit  block: key at col 119 + cols 166..187
# ---------------------------------------
female_sales_value = df.iloc[:, :23].copy()               # 0..22
female_sales_unit  = df.iloc[:, np.r_[0, 47:69]].copy()   # key + 47..68

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
value_idx = list(range(24, 46))   # 24..45
unit_idx  = list(range(70, 92))   # 70..91

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

print(out)

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
female_medic = out

# %% [Cosmetic]
# ### Cosmetic

# %% Cosmetic Topline
# BRANDS (top list)
brands_order = [
    "TOTAL CPD",
    "Maybelline",
    "MBL WITHOUT COLOR SHOW",
    "MBL COLOR SHOW",
    "L'Oreal Paris",
    '3CE',
    "Silkygirl",
    "In 2 It",
    "Revlon",
    "Pixy",
    "Lipice",
    "Nivea",
    "Za",
    "Rimmel",
    "Barenbliss",
    "Dazzle Me",
    "Vaseline",
    "Wardah",
    "Skintific",
    "Makeover",
    "Other Brands",
    "Exclusive + Private Brands",
]

# TOTAL COSMETIC section (Face/Eye/Lip...)
total_cosmetic_order = [
    "COSMETIC",
    "Face",
    "blank1",
    "Eye",
    "Lip",
    "Nail",
    "Sets",
    "Assorted",
]

# TOTAL CPD brand + subcategory sequence (flat, preserves given order)
cpd_brand_order = [
    "TOTAL CPD",
    "Maybelline", "Face", "Eye", "Lip",
    "L'Oreal Paris", "Face", "Eye", "Lip",
    "3CE", "Face", "Eye", "Lip",
    "Silkygirl", "Face", "Eye", "Lip", "Nail", "Assorted",
    "In 2 It", "Face", "Eye", "Lip",
    "REVLON", "Face", "Eye", "Lip", "Nail",
    "Pixy", "Face",
    "Lipice", "Lip",
    "Nivea", "Lip",
    "Za", "Face", "Eye", "Lip_blank",
    "Rimmel", "Face", "Eye", "Lip", "Sets",
    "Barenbliss", "Face", "Eye", "Lip",
    "Dazzle Me", "Face", "Eye", "Lip",
    "Vaseline", "Lip",
    "Wardah", "Face", "Eye", "Lip", "Sets",
    "Skintific", "Face",
    "Makeover", "Face", "Eye", "Lip",
    "Other Brands",
    "Exclusive + Private Brands",
]


# %% [markdown]
# ### Cosmetic Topline

# %% Produce Cosmetic Topline
# ---------------------------------------
# Normalize column headers (e.g., strip suffixes, unify "% Chg" labels, etc.)
# ---------------------------------------
df = normalize_headers_sales_chg(cosmetic)

# ---------------------------------------
# BRANDS block (female make-up) — build side-by-side (Value + Unit)
# ---------------------------------------
female_sales_value = df.iloc[:, :23].copy()               # take first 23 cols (value section)
female_sales_unit  = df.iloc[:, np.r_[0, 47:69]].copy()   # take key col + unit block (47..68)

key = female_sales_value.columns[0]                       # first column name used as join key
if female_sales_unit.columns[0] != key:
    female_sales_unit = men_sales_unit.rename(columns={female_sales_unit.columns[0]: key})  # align key name (keep code as is)

# add suffixes before concatenation so names are unique
left  = female_sales_value.rename(columns=lambda c: c if c == key else f"{c}_value")  # suffix all non-key as _value
right = female_sales_unit.drop(columns=[key]).rename(columns=lambda c: f"{c}_unit")   # drop duplicate key and suffix _unit

merged = pd.concat([left, right], axis=1)                 # place Value + Unit blocks side-by-side

# column positions (from the original wide source) to extract top-row totals
value_idx = list(range(24, 46))   # 24..45  -> totals for value
unit_idx  = list(range(70, 92))   # 70..91  -> totals for unit

# build the top "TOTAL MAKE-UP" row (divide by 1000 except % chg columns)
merged_total = [
    (pd.to_numeric(df.iloc[0, c], errors="coerce") /
     (1 if "% chg" in str(df.columns[c]).lower() else 1000))
    for c in (value_idx + unit_idx)
]

# init output table with expected columns (assumes `cols` was prepared above)
out = pd.DataFrame(columns=cols)
out.loc[len(out)] = ["TOTAL MAKE-UP", *merged_total]      # append top total row
out.loc[len(out)] = [pd.NA] * len(out.columns)            # spacer row
out.loc[len(out)] = ["BRANDS", *([pd.NA]*(len(out.columns) - 1))]  # section header

# ---------------------------------------
# Pick first matches for brands_order; append to `out`, remove from `merged`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged,
    brands_order,
    take_cols=45,
    out=out,          
)

# two spacer rows between sections
out.loc[len(out)] = [pd.NA] * len(out.columns)
out.loc[len(out)] = [pd.NA] * len(out.columns)

# ---------------------------------------
# TOTAL COSMETIC section — rebuild Value + Unit view for that block
# ---------------------------------------
female_sales_value = df.iloc[:, 92:115].copy()               # value block for total cosmetic
female_sales_unit  = df.iloc[:, np.r_[92, 139:161]].copy()   # key + unit block

key = female_sales_value.columns[0]                          # key for this block
if female_sales_unit.columns[0] != key:
    female_sales_unit = men_sales_unit.rename(columns={female_sales_unit.columns[0]: key})  # align key (kept as is)

# add suffixes before concatenation so names are unique
left  = female_sales_value.rename(columns=lambda c: c if c == key else f"{c}_value")  # suffix non-key as _value
right = female_sales_unit.drop(columns=[key]).rename(columns=lambda c: f"{c}_unit")   # drop key duplicate, suffix _unit

merged = pd.concat([left, right], axis=1)                     # side-by-side again

# indices for totals (value + unit) for this section
value_idx = list(range(116, 138))   # indices for value totals
unit_idx  = list(range(162, 184))   # indices for unit totals

# recompute the top totals row (divide by 1000 except % chg columns)
merged_total = [
    (pd.to_numeric(df.iloc[0, c], errors="coerce") /
     (1 if "% chg" in str(df.columns[c]).lower() else 1000))
    for c in (value_idx + unit_idx)
]

# ---------------------------------------
# Pick first matches for total_cosmetic_order; append to `out`
# ---------------------------------------
merged, out = pick_first_matches_and_remove(
    merged,
    total_cosmetic_order,
    take_cols=45,
    out=out,          
)

# spacer row before next section
out.loc[len(out)] = [pd.NA] * len(out.columns)

# ---------------------------------------
# Pick first matches for CPD brand + subcategory sequence; append to `out`
# ---------------------------------------

merged, out = pick_first_matches_and_remove(
    merged,
    cpd_brand_order,
    take_cols=45,
    out=out,          
)

# final assembled table for cosmetic brand view
brand_cosmetic = out


# %% [markdown]
# ### Female LDB Skincare

# %% Female LDB Skincare
function_order = [
    "FEMALE BRIGHTENING",
    "FEMALE HYDRATION + BASIC",
    "Basic",
    "Hydration",
    "FEMALE ANTI-AGING",
    "FEMALE ANTI-ACNE + OIL CONTROL",
    "Anti Acne",
    "FEMALE SENSITIVE",
]

category_order = [
    "CLEANSER",
    "Moisture",
    "Cream/Gel/Oil/Lotion",
    "Essence/Serum",
    "Cosmetic/Treatment Water",
    "Ampoule",
    "TONER",
    "FACE MASK",
    "FACE EYE",
    "MAKE UP REMOVER",
]

brand_order = [
    "TOTAL LDB",
    "La Roche Posay",
    "Anti Acne",
    "Anti Age",
    "Basic",
    "Hydration",
    "Sensitive",
    "Whitening",
    "Make Up Remover",
    "Vichy",
    "anti-acne_blank", 
    "antiaging_blank",
    "Hydration_blank",
    "Make Up Remover_blank",
    "Cerave",
    "Basic",
    "Eucerin",
    "Anti Acne",
    "Anti Age",
    "Basic",
    "Hydration",
    "Sensitive",
    "Whitening",
    "Make Up Remover",
    "Cetaphil",
    "Anti Acne",
    "Basic",
    "Hydration",
    "Oil Control",
    "Whitening",
    "Neutrogena",
    "Anti Acne",
    "Anti Age",
    "Basic",
    "Hydration",
    "Oil Control",
    "Whitening",
    "Make Up Remover",
    "Avene",
    "anti acne",
    "anti age",
    "Basic",
    "Hydration",
    "Oil Control",
    "Sensitive",
    "Whitening",
    "Make Up Remover",
    "Physiogel",
    "Sebamed",
    "QV",
    "Hiruscar",
    "Uriage",
    "Bioderma",
    "OTHERS",
    "TOTAL PRIVATE LABEL & EXCLUSIVE BRANDS",
]


# %% [markdown]
# ### Produce Female Skincare LDB Topline

# %% Female LDB Skincare
# ---------------------------------------
# Normalize column headers (e.g., strip suffixes, unify "% Chg" labels, etc.)
# ---------------------------------------
df = normalize_headers_sales_chg(ldb)

# ---------------------------------------
# Slice female Sales Value and Unit blocks for FEMALE SKINCARE
# Value block: cols 119..141  (23 columns)
# Unit  block: key at col 119 + cols 166..187
# ---------------------------------------
female_sales_value = df.iloc[:, :23].copy()               # 0..22
female_sales_unit  = df.iloc[:, np.r_[0, 24:46]].copy()   # key + 47..68

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
value_idx = list(range(47, 69))   # 24..45
unit_idx  = list(range(70, 92))   # 70..91

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
out.loc[len(out)] = ["TOTAL FEMALE", *merged_total]

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

print(merged)

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

# %%
def norm_label(x: str) -> str:
    x = str(x)
    x = u.normalize("NFKC", x)
    # remove BOM/NBSP/thin spaces you had
    x = x.replace("\ufeff","").replace("\u00A0"," ").replace("\u2007"," ").replace("\u202F"," ")
    # remove soft hyphen + ALL format/control chars (covers \u200B, \u2060, etc.)
    x = x.replace("\u00AD", "")  # soft hyphen
    x = "".join(ch for ch in x if u.category(ch) not in {"Cf","Cc"})
    # unify dashes
    x = x.replace("–","-").replace("—","-").replace("‐","-")
    # collapse whitespace
    x = re.sub(r"\s+", " ", x).strip()
    return x.casefold()

# 1) Is there ANY 'anti acne' after normalization?
col0_norm = df.iloc[:, 0].astype("string").map(norm_label)
print("exact present? ", "anti acne" in set(col0_norm.dropna()))

# 2) Show rows that *contain* 'anti acne' ignoring spaces/hyphens/etc.
import re
hits = df.iloc[:,0].astype("string")
mask = hits.str.contains(r"a\s*n\s*t\s*i[^a-z0-9]+a\s*c\s*n\s*e", flags=re.I, regex=True, na=False)
print(hits[mask].head(20))

# 3) Looser: remove non-alphanumerics before comparing (one-off check)
def letters_only(s): 
    return re.sub(r"[^0-9a-z]+", "", norm_label(s))
present_compact = {letters_only(x) for x in df.iloc[:,0].astype("string")}
print("antiacne (compact) present? ", "antiacne" in present_compact)


# %% [markdown]
# ### Save Final Report

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

with pd.ExcelWriter(dest / f"Topline_MY_CPD_{month_tag_iso}_{_ts}.xlsx", engine="openpyxl") as writer:
    male_cpd.to_excel(writer, sheet_name="Male CPD", index=False)
    female_cpd.to_excel(writer, sheet_name="Female CPD", index=False)
    female_medic.to_excel(writer, sheet_name="Female CPD (Plus Mass Medic)", index=False)
    brand_cosmetic.to_excel(writer, sheet_name="Female Cosmetic", index=False)
    female_ldb.to_excel(writer, sheet_name="Female LDB", index=False)
print(f"[saved] {dest / f'Topline_{month_tag_iso}_{_ts}.xlsx'}")


# %%



