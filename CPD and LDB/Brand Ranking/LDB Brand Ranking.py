# BODYCARE CURRENT YEAR RULE
#
# False:
#   L'Oreal brands -> use latest available month
#   Non-L'Oreal brands + Market -> use last completed quarter only
#
# True:
#   All brands use latest available month
#
INCLUDE_CURRENT_NON_LOREAL_BODYCARE = False
import glob
import os
import re
import warnings
from datetime import datetime
from pathlib import Path
import pandas as pd
warnings.filterwarnings("ignore")
# Paths
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = Path(r"C:\Users\balatarsini_avinitya\Downloads\CPD & LDB O+O - June\loreal-report-automation (2)")
BRAND_RANKING_DIR = PROJECT_ROOT / "Brand Ranking"
NIELSEN_DIR = PROJECT_ROOT / "Data Source" / "Nielsen"
OMT_DIR = PROJECT_ROOT / "Data Source" / "OMT - O+O"
MAPPING_PATH = BRAND_RANKING_DIR / "Mapping.xlsx"
LDB_EXCEL_REPORT = SCRIPT_DIR / "LDB Excel report" / "MYSG ONE LDB CMI YTD Jun'26 (1).xlsx"
# LDB Data Output (Bodycare + Suncare source file).
# Adjust this path if the file lives somewhere else on your machine.
BODYCARE_FILE = SCRIPT_DIR / "LDB Data Output Jun'26.xlsx"
# Update these when running for a new reporting period.
NIELSEN_FILE = NIELSEN_DIR / "Nielsen Brand Ranking_Jun26 v270726.xlsx"
CURR_YEAR = 2026
PREV_YEAR = 2025
LAST_YEAR = 2024
WORKING_MONTH = 7
MASS_MEDIC = [
    "ACNE AID", "ACNES", "AVEENO", "BALNEUM", "BENZAC", "BIO-OIL", "CARMEX", "CERAVE", "CETAPHIL", "CUREL",
    "DERMATIX", "DERMAVEEN", "DIFFERIN", "DR.G", "DR.YU", "EGO", "EUBOS", "LACTACYD", "LINOLA", "MUSTELA",
    "NEUTROGENA", "PANOXYL", "PHYSIOGEL", "SEBAMED", "TOPICREM", "VANICREAM", "WIS", "XHEKPON", "FIRST AID BEAUTY",
    "AQUAPHOR", "NOBACTER", "LUBRIDERM", "NEOSPORIN", "DARROW", "DEXERYL", "ALERGIBON", "ALPHYGIENE", "BABIGOZ",
    "CANDERMYL", "GALDERMA", "GALDERMA OTHER", "HELIOBLOC", "HYDRODERM OMEGA", "IOCON", "IONIL", "MACROLANE",
    "MICROBAN", "MICROSUN", "NESTLE", "NUTRASPA", "OBSERVANCE", "PHYGIENE", "R-GEN", "SENTIAL", "ACHE", "ACNAID",
    "ACNE FREE", "ACOFAR", "ADDAX", "AKILDIA", "ALBOLENE", "AMLACTIN", "ANSEBIC", "AQUA SOAP", "AQUA-SOAP",
    "AVITIL", "AZULENNE", "BACCIDE", "BEAUTY PLUS", "BEDOOK", "BEPANTHEN/BEPANTHOL", "BETAGRANULOS", "BIAFINE",
    "BIOBLAS", "BIOCLIN", "BIOLIQ", "BIOXCIN", "BLUE LIZARD", "BODYSOL", "BONAVEN", "BOROLINE", "CERAMOL",
    "CERTAIN DRI", "CETOPIC", "CHICCO", "CICAMEL", "COOPER", "COTARYL", "CRISTALIA", "DECUBAL", "DERMAC",
    "DERMACTIVE", "DERMADRATE", "DERMAGE", "DERMAKERI", "DERMENA", "DERMON", "DERSUPRIL", "DEUMAVAN",
    "DOCTISSIMO PARAPHARMACIE", "DR.LI", "DR.LIDERMO", "DRAYEX", "DX2", "E45", "ELDOPAQUE", "EMOLIENTA", "EMOLIN",
    "EMOLIUM", "EPIMAX", "EVASOL", "FARMOQUIMICA", "FILTROSOL", "FLUOCIN", "FREI OEL (BOUHON)", "GALENCO", "GIFRER",
    "GILBERT", "GOLD BOND", "HAMILTON", "HIDRAFIL", "HIPOSOL", "HYALIX", "IDROVEL", "IHADA", "INFASIL",
    "INTERAPOTHEK", "IRALTONE", "ITANIDERM", "KAMILODERM", "KETOXIN", "KINERASE", "KORA", "LACTIBON",
    "LACTO CALAMINE", "LETI", "LIFAR", "LIPODERM", "LOTRIMIN", "MARQUE VERTE", "MICRORET", "MITOSYL", "MODERM",
    "MULTIDERMOL", "MUSSVITAL", "NEUTRA LICE", "NEUTRAPHARM", "NORDIN", "NUMIS", "NUMIS MED", "NURAPHARM",
    "NUTREM", "NUTRISIL", "OILATUM", "OILLAN", "OSMIN", "OTC IBERICA", "PANVEL DERMATIV", "PARABOTICA",
    "PHARMACTIV", "PHARMASEPT", "PHISOHEX", "PROCICAR", "REGENERUM", "RESTIV", "RESTIVOIL", "REVALESKIN", "ROCHE",
    "ROGE CAVAILLES", "ROYALCARE", "RUGARD (SCHEFFLER)", "SALILEX", "SALLVE", "SARNA", "SAUGELLA", "SEBORADIN",
    "SHADE", "SMOOTH-E", "SOLAR FOAM", "S-OLE", "SPECTRABAN", "STANHOME FAMILY EXPERT", "STIEFEL", "STIEPROX",
    "STIPROX", "STIPROXAL", "TARMED", "TRACTOPON", "TRI DERMA MD", "UREADERM", "UVEIL-PS", "UVESOL", "VEA",
    "VENUSIA", "VITA CITRAL", "VITALIFE", "ZODIAC", "QV", "BOBAI", "COLLAGE", "DERMAREST", "EPIZONE E",
    "GLAMY LAB", "LU MILD", "NOLAVER", "OXECURE", "RIUP", "SEBCUR", "SELENGENA", "SEROPIPE", "SHAAN",
    "STAR VILLE", "STRONGVILLE", "SYNOBAR", "UREMOL", "URISEC", "ZINPLEX",
]
NON_MASS_MEDIC_EXTRA = [
    "EUCERIN", "LA ROCHE POSAY", "VICHY", "AVENE", "DR MORITA", "HIRUSCAR", "URIAGE", "SKINCEUTICALS", "DECLEOR",
    "SANOFLORE", "AQUAPHOR", "BIODERMA", "ISDIN", "LIERAC", "FILORGA", "PROACTIV", "ROC", "WINONA", "DR CILABO",
    "EMOLIUM", "PHARMACERIS", "CAUDALIE", "NUXE", "RODAN", "FIELDS", "LIBREDERM", "MANTECORP", "DR. CI : LABO",
]
LOREAL_BRANDS = {
    "LA ROCHE POSAY",
    "VICHY",
    "SKINCEUTICALS",
    "CERAVE",
}
def load_non_mass_medic_brands() -> set[str]:
    mapping = pd.read_excel(MAPPING_PATH, sheet_name="Medic")
    from_mapping = mapping.iloc[:, 0].dropna().astype(str).str.upper().str.strip().tolist()
    return set(from_mapping + NON_MASS_MEDIC_EXTRA)
NON_MASS_MEDIC = load_non_mass_medic_brands()
MASS_MEDIC_SET = {brand.upper().strip() for brand in MASS_MEDIC}
def normalize_brand(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).upper().strip()
def _norm_cat(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).upper().strip()
def detect_category_from_row_general(row) -> str:
    # Look across multiple possible category fields and decide: FACECARE, SUN CARE, BODY CARE
    parts = []
    for k in ["Category", "Category L1", "Category L2", "Category L3", "Subcategory", "Sub Category", "Sub-Category"]:
        parts.append(_norm_cat(row.get(k)))
    combined = " ".join(parts)
    if "BODY" in combined:
        return "BODY CARE"
    if "PROTECT" in combined or "PROTECTION" in combined or "SUN" in combined:
        return "SUN CARE"
    # default to FACECARE when face/cleaning tokens appear or nothing else
    if "FACE" in combined or "CLEAN" in combined or "CLEANSING" in combined:
        return "FACECARE"
    return "FACECARE"
def detect_category_from_row_omt(row) -> str | None:
    # Strict rules for OMT raw files per user request:
    # - FACECARE: Category L1 == 'SKIN CARE' and Category L2 indicates face care / cleansing
    # - SUN CARE: Category L3 == 'FACE PROTECTION' only
    # - BODY CARE: Category L2 or L3 contains 'BODY'
    l1 = _norm_cat(row.get("Category L1"))
    if l1 != "SKIN CARE":
        return None
    l2 = _norm_cat(row.get("Category L2"))
    l3 = _norm_cat(row.get("Category L3"))
    if "BODY" in l2 or "BODY" in l3:
        return "BODY CARE"
    # SUN CARE must come from Category L3 == FACE PROTECTION only
    if l3 and ("FACE PROTECT" in l3 or "FACE PROTECTION" in l3 or l3 == "FACE PROTECTION"):
        return "SUN CARE"
    # FACECARE only from Category L2 indicating face care / cleansing
    # explicitly accept 'FACE CARE & CLEANSING'
    if l2 and ("FACE CARE & CLEANSING" in l2 or "FACE" in l2 or "CLEANS" in l2 or "CLEAN" in l2):
        return "FACECARE"
    return None
def normalize_time_period(df: pd.DataFrame) -> pd.DataFrame:
    df["Time Period"] = df["Time Period"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["Time Period"] = df["Time Period"].str.replace("YTD ", "YTD", regex=False)
    is_fy = df["Time Period"].str.isdigit() & (df["Time Period"].str.len() == 4)
    df.loc[is_fy, "Time Period"] = "FY" + df.loc[is_fy, "Time Period"]
    return df
def validate_time_periods(frames: dict[str, pd.DataFrame]) -> None:
    for sheet_name, df in frames.items():
        bad = df["Time Period"].astype(str).str.fullmatch(r"20\d{2}\.0")
        if bad.any():
            examples = sorted(df.loc[bad, "Time Period"].astype(str).unique())
            raise ValueError(f"Bad Time Period values in {sheet_name}: {examples}")
def period_to_time_period(period: str, country: str) -> str | None:
    period = str(period)
    year_match = re.search(r"(\d{4})", period)
    if "Cal Yr" in period and year_match:
        return "FY" + year_match.group(1)
    if country == "SG":
        if "YTD YA" in period:
            return f"YTD{PREV_YEAR}"
        if "YTD" in period:
            return f"YTD{CURR_YEAR}"
    ytd_match = re.search(r"YTD.*?(\d{2})\s*$", period)
    if ytd_match:
        return "YTD20" + ytd_match.group(1)
    return None
def map_nielsen_subdivision(row: pd.Series, country: str) -> str:
    if country == "MY":
        brand = normalize_brand(row.get("BRAND"))
        manufacturer = normalize_brand(row.get("MANUFACTURER"))
        if brand in MASS_MEDIC_SET:
            return "Mass Medic"
        if brand in NON_MASS_MEDIC or "TOTAL PRIVATE LABEL" in manufacturer:
            return "Exc. Mass Medic"
        return "Others"
    brand = normalize_brand(row.get("LOCAL BRAND"))
    manufacturer = normalize_brand(row.get("LOCAL MANUF"))
    if brand in MASS_MEDIC_SET:
        return "Mass Medic"
    if brand in NON_MASS_MEDIC or manufacturer == "TOTAL OTHERS":
        return "Exc. Mass Medic"
    return "Exc. Mass Medic"
def map_nielsen_brand(row: pd.Series, country: str) -> str:
    if country == "MY":
        brand = row.get("BRAND")
        manufacturer = row.get("MANUFACTURER")
    else:
        brand = row.get("LOCAL BRAND")
        manufacturer = row.get("LOCAL MANUF")
    return normalize_brand(brand if pd.notna(brand) else manufacturer)
def add_common_columns(df: pd.DataFrame, country: str, channel: str, flag: str) -> pd.DataFrame:
    df = normalize_time_period(df.copy())
    df["Country (Currency)"] = "MY (MYR'000)" if country == "MY" else "SG (SGD'000)"
    df["MY/SG"] = country
    df["Division"] = "LDB"
    df["Channel"] = channel
    df["Flag"] = flag
    df["Is_Loreal"] = df["Brand"].apply(lambda x: "Yes" if normalize_brand(x) in LOREAL_BRANDS else "No")
    return df[
        [
            "Country (Currency)",
            "MY/SG",
            "Category",
            "Brand",
            "Division",
            "Subdivision",
            "Time Period",
            "SO",
            "Channel",
            "Is_Loreal",
            "Flag",
        ]
    ]
def subdivision_from_report(value) -> str:
    value = str(value).strip()
    if value == "Mass Medical":
        return "Mass Medic"
    return "Exc. Mass Medic"
def time_period_from_year_month(df: pd.DataFrame) -> pd.DataFrame:
    fy = df.copy()
    fy["Time Period"] = "FY" + fy["Year"].astype(int).astype(str)
    ytd = df[df["Period"].astype(int) < WORKING_MONTH].copy()
    ytd["Time Period"] = "YTD" + ytd["Year"].astype(int).astype(str)
    return pd.concat([fy, ytd], ignore_index=True)
def get_last_completed_quarter_month(working_month: int) -> int:
    if working_month <= 3:
        return 0
    elif working_month <= 6:
        return 3
    elif working_month <= 9:
        return 6
    elif working_month <= 12:
        return 9
    return 12
def process_bodycare(country: str) -> pd.DataFrame:
    df = pd.read_excel(BODYCARE_FILE, sheet_name=country)
    # Columns confirmed:
    # A: MY/SG | B: Channel (Offline) | C: Year | D: Offline Est (Platform)
    # E: SKINCARE | F: Body Care/Suncare | G: Mass Medical/Non-Mass Medical
    # H: Brands | I: month | J: Sales Value
    rename_map = {
        df.columns[0]: "MY/SG",        # Col A
        df.columns[1]: "Channel",      # Col B
        df.columns[2]: "Year",         # Col C
        df.columns[5]: "CategoryRaw",  # Col F
        df.columns[6]: "MassType",     # Col G
        df.columns[7]: "Brand",        # Col H
        df.columns[8]: "Period",       # Col I
        df.columns[9]: "SO",           # Col J
    }
    df = df.rename(columns=rename_map)
    df = df[df["Year"].isin([LAST_YEAR, PREV_YEAR, CURR_YEAR])].copy()
    # Category from Col F: "Body Care" / "Suncare"
    category_map = {
        "BODY CARE": "BODY CARE",
        "BODYCARE": "BODY CARE",
        "SUN CARE": "SUN CARE",
        "SUNCARE": "SUN CARE",
    }
    df["CategoryRaw"] = df["CategoryRaw"].astype(str).str.strip().str.upper()
    df["Category"] = df["CategoryRaw"].map(category_map)
    unmapped = df.loc[df["Category"].isna(), "CategoryRaw"].unique()
    if len(unmapped):
        raise ValueError(
            f"Unrecognized Category values in {BODYCARE_FILE.name} ({country}): {list(unmapped)}"
        )
    # Brand cleanup
    df["Brand"] = (df["Brand"].astype(str).str.strip().str.upper())
    df.loc[
        df["Brand"].str.upper() == "MEDIC MARKET",
        "Brand"
    ] = "Market"
    # Subdivision
    df["Subdivision"] = (
        df["MassType"]
        .astype(str)
        .str.strip()
        .replace({
            "Mass Medical": "Mass Medic",
            "Non-Mass Medical": "Exc. Mass Medic",
        })
    )
    # Flag
    df["Flag"] = df["Brand"].apply(
        lambda x: "F" if str(x).strip() == "Market" else "WG"
    )
    # -------------------------------------------------
    # FY DATA
    # -------------------------------------------------
    fy = (
        df.groupby(
            [
                "Category",
                "Brand",
                "Subdivision",
                "Channel",
                "Flag",
                "Year",
            ],
            as_index=False,
            dropna=False,
        )["SO"]
        .sum()
    )
    fy["Time Period"] = "FY" + fy["Year"].astype(int).astype(str)
    # -------------------------------------------------
    # YTD DATA
    # -------------------------------------------------
    ytd_frames = []
    # Historical years
    historical = df[(df["Year"].isin([LAST_YEAR, PREV_YEAR])) & (df["Period"].astype(int) < WORKING_MONTH)].copy()
    if not historical.empty:
        hist_ytd = (
            historical.groupby(
                [
                    "Category",
                    "Brand",
                    "Subdivision",
                    "Channel",
                    "Flag",
                    "Year",
                ],
                as_index=False,
                dropna=False,
            )["SO"]
            .sum()
        )
        hist_ytd["Time Period"] = (
            "YTD" + hist_ytd["Year"].astype(int).astype(str)
        )
        ytd_frames.append(hist_ytd)
    # Current year
    current = df[df["Year"] == CURR_YEAR].copy()
    if not current.empty:
        if INCLUDE_CURRENT_NON_LOREAL_BODYCARE:
            current_ytd = current[
                current["Period"].astype(int) < WORKING_MONTH
            ].copy()
        else:
            quarter_cutoff = get_last_completed_quarter_month(
                WORKING_MONTH
            )
            loreal_mask = current["Brand"].apply(lambda x:normalize_brand(x) in LOREAL_BRANDS or str(x).strip() == "Market")
            loreal_rows = current[
                loreal_mask
                & (current["Period"].astype(int) < WORKING_MONTH)
            ]
            non_loreal_rows = current[
                (~loreal_mask)
                & (current["Period"].astype(int) <= quarter_cutoff)
            ]
            current_ytd = pd.concat(
                [loreal_rows, non_loreal_rows],
                ignore_index=True,
            )
        current_ytd = (
            current_ytd.groupby(
                [
                    "Category",
                    "Brand",
                    "Subdivision",
                    "Channel",
                    "Flag",
                    "Year",
                ],
                as_index=False,
                dropna=False,
            )["SO"]
            .sum()
        )
        current_ytd["Time Period"] = (
            "YTD" + current_ytd["Year"].astype(int).astype(str)
        )
        ytd_frames.append(current_ytd)
    ytd = pd.concat(ytd_frames, ignore_index=True)
    final = pd.concat([fy, ytd], ignore_index=True)
    final["Country (Currency)"] = (
        "MY (MYR'000)"
        if country == "MY"
        else "SG (SGD'000)"
    )
    final["MY/SG"] = country
    final["Division"] = "LDB"
    final["Is_Loreal"] = final["Brand"].apply(
        lambda x: (
            "Yes"
            if normalize_brand(x) in LOREAL_BRANDS
            else "No"
        )
    )
    return final[
        [
            "Country (Currency)",
            "MY/SG",
            "Category",
            "Brand",
            "Division",
            "Subdivision",
            "Time Period",
            "SO",
            "Channel",
            "Is_Loreal",
            "Flag",
        ]
    ]
def read_oo_data_report() -> pd.DataFrame:
    return pd.read_excel(LDB_EXCEL_REPORT, sheet_name="O+O Data", header=2)
def process_sg_tiktok_market() -> pd.DataFrame:
    df = read_oo_data_report()
    df = df[
        (df["Country (Currency)"].astype(str) == "SG (SGD'000)")
        & (df["Platform"].astype(str) == "TikTok")
        & (df["Source"].astype(str) == "TikTok Market Estimation")
        & (df["Brand"].astype(str).str.upper() == "MEDIC MARKET")
    ].copy()
    df = df[df["Year"].isin([LAST_YEAR, PREV_YEAR, CURR_YEAR])]
    category_map = {
        "FACECARE": "FACECARE",
        "FACE CARE & CLEANSING": "FACECARE",
        "SUNCARE": "SUN CARE",
        "SUN CARE": "SUN CARE",
        "BODYCARE": "BODY CARE",
        "BODY CARE": "BODY CARE",
    }
    df["Category"] = df["Category"].astype(str).str.upper().str.strip().map(category_map)
    df = df[df["Category"].notna()]
    df = time_period_from_year_month(df)
    df["Brand"] = "Market"
    df["Subdivision"] = df["Mass/Non-Mass"].apply(subdivision_from_report)
    df = df.groupby(["Category", "Brand", "Subdivision", "Time Period"], as_index=False, dropna=False)["Sellout_x"].sum()
    df = df.rename(columns={"Sellout_x": "SO"})
    return add_common_columns(df, "SG", "Online", "F")
def process_sg_departmental_store() -> tuple[pd.DataFrame, pd.DataFrame]:
    df = read_oo_data_report()
    df = df[
        (df["Country (Currency)"].astype(str) == "SG (SGD'000)")
        & (df["On/Offline"].astype(str) == "Offline")
        & (df["Platform"].astype(str) == "Departmental Store")
        & (df["Brand"].astype(str).str.upper().isin(["MEDIC MARKET", "SKINCEUTICALS"]))
    ].copy()
    df = df[df["Year"].isin([LAST_YEAR, PREV_YEAR, CURR_YEAR])]
    # Keep Facecare AND Sun Care rows (Skinceuticals sells both categories
    # in this channel; Medic Market total below is still restricted to
    # Facecare only, matching prior behaviour).
    category_map = {
        "FACECARE": "FACECARE",
        "FACE CARE & CLEANSING": "FACECARE",
        "SUNCARE": "SUN CARE",
        "SUN CARE": "SUN CARE",
    }
    df["Category"] = df["Category"].astype(str).str.upper().str.strip().map(category_map)
    df = df[df["Category"].notna()]
    df = time_period_from_year_month(df)
    df["Subdivision"] = df["Mass/Non-Mass"].apply(subdivision_from_report)
    market = df[
        (df["Brand"].astype(str).str.upper() == "MEDIC MARKET")
        & (df["Category"] == "FACECARE")
    ].copy()
    market["Brand"] = "Market"
    market = market.groupby(["Category", "Brand", "Subdivision", "Time Period"], as_index=False, dropna=False)["Sellout_x"].sum()
    market = market.rename(columns={"Sellout_x": "SO"})
    market = add_common_columns(market, "SG", "Offline", "F")
    skc = df[df["Brand"].astype(str).str.upper() == "SKINCEUTICALS"].copy()
    skc["Brand"] = "SKINCEUTICALS"
    skc = skc.groupby(["Category", "Brand", "Subdivision", "Time Period"], as_index=False, dropna=False)["Sellout_x"].sum()
    skc = skc.rename(columns={"Sellout_x": "SO"})
    skc = add_common_columns(skc, "SG", "Offline", "WG")
    return market, skc
def process_nielsen(country: str) -> pd.DataFrame:
    sheet_name = f"Nielsen {country} ACD"
    df = pd.read_excel(NIELSEN_FILE, sheet_name=sheet_name)
    df["Time Period"] = df["Periods"].apply(lambda value: period_to_time_period(value, country))
    df["Subdivision"] = df.apply(lambda row: map_nielsen_subdivision(row, country), axis=1)
    df["Brand"] = df.apply(lambda row: map_nielsen_brand(row, country), axis=1)
    # detect category (keep SUN CARE and BODY CARE when present)
    df["Category"] = df.apply(lambda r: detect_category_from_row_general(r), axis=1)
    df = df[df["Subdivision"] != "Others"]
    df = df[df["Time Period"].isin([f"FY{LAST_YEAR}", f"FY{PREV_YEAR}", f"FY{CURR_YEAR}", f"YTD{LAST_YEAR}", f"YTD{PREV_YEAR}", f"YTD{CURR_YEAR}"])]
    df = df.groupby(["Category", "Brand", "Subdivision", "Time Period"], as_index=False, dropna=False)["Sales Value"].sum()
    df = df.rename(columns={"Sales Value": "SO"})
    return add_common_columns(df, country, "Offline", "F")
def read_omt_files(country: str) -> pd.DataFrame:
    folder = OMT_DIR / f"{country} LDB"
    files = sorted(glob.glob(str(folder / f"OMT {country} LDB *.xlsx")))
    if not files:
        raise FileNotFoundError(f"No OMT files found in {folder}")
    frames = []
    for file in files:
        match = re.search(r"(\d{4})-(\d{2})", Path(file).stem)
        if not match:
            continue
        year, month = int(match.group(1)), int(match.group(2))
        if year not in [LAST_YEAR, PREV_YEAR, CURR_YEAR]:
            continue
        if year == CURR_YEAR and month >= WORKING_MONTH:
            continue
        frame = pd.read_excel(file, sheet_name="Export")
        frames.append(frame)
    if not frames:
        raise ValueError(f"No OMT files selected for {country}")
    return pd.concat(frames, ignore_index=True)
def map_omt_subdivision(brand: str) -> str:
    return "Mass Medic" if normalize_brand(brand) in MASS_MEDIC_SET else "Exc. Mass Medic"
def process_omt(country: str, flag: str) -> pd.DataFrame:
    df = read_omt_files(country)
    df = df[df["Category L1"] == "SKIN CARE"].copy()
    df[["Year", "Month"]] = df["Year Month"].str.split("-", expand=True).astype(int)
    # determine category strictly from Category L2/L3 in OMT raw files
    df["Category"] = df.apply(lambda r: detect_category_from_row_omt(r), axis=1)
    # drop rows that don't match the strict OMT rules
    df = df[df["Category"].notna()].copy()
    df["Subdivision"] = df["Brand"].apply(map_omt_subdivision)
    df = df.groupby(
        ["Brand", "Year", "Month", "Subdivision", "Category"],
        as_index=False,
        dropna=False,
    )["Total Est. Sales Local"].sum()
    fy = df.copy()
    fy["Time Period"] = "FY" + fy["Year"].astype(str)
    ytd = df[df["Month"] < WORKING_MONTH].copy()
    ytd["Time Period"] = "YTD" + ytd["Year"].astype(str)
    final = pd.concat([fy, ytd], ignore_index=True)
    final = final.groupby(["Category", "Brand", "Subdivision", "Time Period"], as_index=False, dropna=False)["Total Est. Sales Local"].sum()
    final = final.rename(columns={"Total Est. Sales Local": "SO"})
    return add_common_columns(final, country, "Online", flag)
def total_market(df: pd.DataFrame, country: str, channel: str) -> pd.DataFrame:
    market = df[(df["Channel"] == channel) & (df["Flag"] == "F")].copy()
    market["Brand"] = "Market"
    grouped = market.groupby(
        ["Country (Currency)", "MY/SG", "Category", "Brand", "Division", "Subdivision", "Time Period", "Channel", "Flag"],
        as_index=False,
        dropna=False,
    )["SO"].sum()
    grouped["Is_Loreal"] = "No"
    return grouped[[
        "Country (Currency)",
        "MY/SG",
        "Category",
        "Brand",
        "Division",
        "Subdivision",
        "Time Period",
        "SO",
        "Channel",
        "Is_Loreal",
        "Flag",
    ]]
def process_country(country: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    offline = process_nielsen(country)
    bodycare = process_bodycare(country)
    offline = pd.concat([offline, bodycare], ignore_index=True)
    online_wg = process_omt(country, "WG")
    online_f = process_omt(country, "F")
    if country == "SG":
        sg_tiktok_market = process_sg_tiktok_market()
        sg_dept_market, sg_dept_skc = process_sg_departmental_store()
        offline = pd.concat([offline, sg_dept_market, sg_dept_skc], ignore_index=True)
        online_f = pd.concat([online_f, sg_tiktok_market], ignore_index=True)
    oo_wg = pd.concat([offline[offline["Brand"] != "Market"], online_wg], ignore_index=True)
    oo_f = pd.concat([offline, online_f], ignore_index=True)
    final = pd.concat([total_market(oo_f, country, "Offline"), total_market(oo_f, country, "Online"), oo_wg], ignore_index=True)
    validate_time_periods({"offline": offline, "online_wg": online_wg, "online_f": online_f, "final": final})
    return offline, online_wg, online_f, final
def write_country_output(country: str, offline: pd.DataFrame, online_wg: pd.DataFrame, online_f: pd.DataFrame, final: pd.DataFrame) -> None:
    output = SCRIPT_DIR / f"{country} LDB Brand Ranking.xlsx"
    sheets = {
        "Offline": offline,
        "Online-wg": online_wg,
        "Online-f": online_f,
        "O+O": final,
    }
    validate_time_periods(sheets)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
def main() -> None:
    my_offline, my_online_wg, my_online_f, my_final = process_country("MY")
    sg_offline, sg_online_wg, sg_online_f, sg_final = process_country("SG")
    combined = pd.concat([my_final, sg_final], ignore_index=True)
    write_country_output("MY", my_offline, my_online_wg, my_online_f, my_final)
    write_country_output("SG", sg_offline, sg_online_wg, sg_online_f, sg_final)
    combined_sheets = {
        "MY LDB": my_final,
        "SG LDB": sg_final,
        "MY & SG LDB": combined,
    }
    validate_time_periods(combined_sheets)
    with pd.ExcelWriter(SCRIPT_DIR / "Split Data MYSG LDB Combined.xlsx", engine="openpyxl") as writer:
        for sheet_name, df in combined_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
if __name__ == "__main__":
    main()
