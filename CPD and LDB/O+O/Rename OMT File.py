import os
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime

# Single combined script with validation checks:
# 1) Scan all .xlsx files and extract Universe, Country, Year Month
# 2) For full years (2024, 2025) ensure months Jan-Dec are present for each country/type
# 3) For current year, ask user which month to process up to (default T-1) and ensure months are present
# 4) If validations pass, move files into consolidated folders `OMT {Country} {TYPE}`

base_dir = Path(r"C:\Users\balatarsini_avinitya\Downloads\OMT folder onecmi june") #downloaded files' folder

def determine_type_from_universe(universe: str):
    u = (universe or '').lower()
    if any(k in u for k in ('meic', 'medic', 'dermo')):
        return 'LDB'
    if 'mass' in u:
        return 'CPD'
    return None

# Scan phase
files_meta = []
skipped = []
for fpath in base_dir.rglob('*.xlsx'):
    if fpath.name.lower() in ('organize_omt.py', 'move_omt.py'):
        continue
    try:
        df = pd.read_excel(fpath, sheet_name=0, engine='openpyxl')
    except Exception as e:
        skipped.append((str(fpath), f'read error: {e}'))
        continue

    universe = ''
    country = ''
    year_month = None
    if 'Universe' in df.columns:
        vals = df['Universe'].dropna().astype(str)
        if len(vals) > 0:
            universe = vals.iloc[0]
    if 'Country' in df.columns:
        vals = df['Country'].dropna().astype(str)
        if len(vals) > 0:
            country = vals.iloc[0]
    if 'Year Month' in df.columns:
        ym_vals = df['Year Month'].dropna()
        if len(ym_vals) > 0:
            try:
                year_month = pd.to_datetime(ym_vals.max()).strftime('%Y-%m')
            except Exception:
                year_month = str(ym_vals.max())

    if not country:
        parts = fpath.parent.name.split()
        if len(parts) >= 2:
            country = parts[1]

    file_type = determine_type_from_universe(universe)
    if not file_type:
        parent_lower = fpath.parent.name.lower()
        if 'ldb' in parent_lower:
            file_type = 'LDB'
        elif 'cpd' in parent_lower:
            file_type = 'CPD'

    if not file_type or not country or not year_month:
        skipped.append((str(fpath), f'missing meta: universe="{universe}", country="{country}", year_month="{year_month}"'))
        continue

    files_meta.append({
        'path': fpath,
        'country': country,
        'type': file_type,
        'year_month': year_month,
    })

if not files_meta:
    print('No valid files found to process.')
    if skipped:
        print('Skipped files:')
        for s in skipped:
            print('-', s[0], '=>', s[1])
    sys.exit(1)

# Build mapping country/type -> set of months
from collections import defaultdict
mapping = defaultdict(set)
for m in files_meta:
    ym = m['year_month']
    mapping[(m['country'], m['type'])].add(ym)

def months_for_year(year):
    return {f"{year}-{str(m).zfill(2)}" for m in range(1,13)}

# Check full years
full_years = [2024, 2025]
errors = []
for (country, ftype), months in mapping.items():
    for y in full_years:
        expected = months_for_year(y)
        missing = sorted(expected - {m for m in months if m.startswith(f"{y}-")})
        if missing:
            errors.append((country, ftype, y, missing))

if errors:
    print('ERROR: Missing months detected for full years:')
    for country, ftype, year, missing in errors:
        print(f'- {country} {ftype} {year} missing months: {", ".join(missing)}')
    sys.exit(1)

# Current year check
now = datetime.now()
current_year = now.year
default_month_num = now.month - 1 if now.month > 1 else 12
default_year = current_year if now.month > 1 else current_year - 1
default_last = f"{default_year}-{str(default_month_num).zfill(2)}"

print(f"Default last month to process for {current_year} is {default_last} (T-1).")
resp = input(f"Enter last month to process for {current_year} in YYYY-MM (press Enter to accept {default_last}): ").strip()
if not resp:
    last_month_to_process = default_last
else:
    last_month_to_process = resp

# Validate format YYYY-MM
try:
    lm_year = int(last_month_to_process.split('-')[0])
    lm_month = int(last_month_to_process.split('-')[1])
    if lm_year != current_year:
        print(f'ERROR: Last month year ({lm_year}) must be {current_year}.')
        sys.exit(1)
    if not (1 <= lm_month <= 12):
        raise ValueError()
except Exception:
    print('ERROR: Invalid month format. Use YYYY-MM, e.g. 2026-03')
    sys.exit(1)

# For each country/type ensure months from Jan to last_month_to_process exist
from datetime import date
required_months = []
for m in range(1, lm_month + 1):
    required_months.append(f"{current_year}-{str(m).zfill(2)}")

curr_errors = []
for (country, ftype), months in mapping.items():
    missing = [rm for rm in required_months if rm not in months]
    if missing:
        curr_errors.append((country, ftype, missing))

if curr_errors:
    print('ERROR: Missing months for current year up to requested month:')
    for country, ftype, missing in curr_errors:
        print(f'- {country} {ftype} missing: {", ".join(missing)}')
    sys.exit(1)

print('Validation passed — moving files into consolidated folders...')

# Move phase
moved = 0
logs = []
for m in files_meta:
    fpath = m['path']
    country = m['country']
    ftype = m['type']
    year_month = m['year_month']
    dest_folder = base_dir / f"OMT {country} {ftype}"
    dest_folder.mkdir(exist_ok=True)
    base_name = f"OMT {country} {ftype} {year_month}"
    target_path = dest_folder / (base_name + f"{Path(fpath).suffix}")
    counter = 1
    while target_path.exists():
        target_path = dest_folder / f"{base_name} ({counter}){Path(fpath).suffix}"
        counter += 1
    try:
        Path(fpath).replace(target_path)
        moved += 1
        logs.append((str(fpath), str(target_path)))
    except Exception as e:
        skipped.append((str(fpath), f'move error: {e}'))

# Cleanup empty month folders
removed_dirs = []
for child in base_dir.iterdir():
    if not child.is_dir():
        continue
    if not child.name.startswith('OMT '):
        continue
    parts = child.name.split()
    if len(parts) >= 4:
        try:
            if not any(child.iterdir()):
                child.rmdir()
                removed_dirs.append(str(child))
        except Exception:
            pass

print('Processed files:', len(files_meta))
print('Moved:', moved)
print('Skipped:', len(skipped))
if logs:
    for o, n in logs:
        print('Moved:', os.path.basename(o), '->', os.path.basename(n))
if skipped:
    print('\nSkipped details:')
    for s in skipped:
        print('-', s[0], '=>', s[1])
if removed_dirs:
    print('\nRemoved empty folders:')
    for d in removed_dirs:
        print('-', d)