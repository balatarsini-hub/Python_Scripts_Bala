# Luxe Market

This repository contains Python scripts used for reporting and data preparation tasks. 
All my folders is now available in teams under "Bala" folder for reference

## LUXE Market Scripts

| Script | Purpose | Required files/folders | 
|---|---|---|
| `BR_Mapping.py` | Splits quarterly BR data into monthly data for MY and SG |*1. "Actual Market" folder <br> 2. "Raw BR" folder <br> 3. MY and SG O+O report <br> 4. BR Data*|
|`Estimation MY.py` | Estimates market values for MY mm'2026 |*1. MMPR for current <br> 2. Last month Luxe report <br> 3. BR After Split for latest quarter file*|
|`Estimation SG.py`| Estimates market values for MY mm'2026 |*1. Mkt est 2026_mm file <br> 2. Last month Luxe report <br> 3. BR After Split for latest quarter file*|

## How to use

1. Download my folder from teams because some may data base file and other folders to read the files
2. Read the instruction below for easy understanding before run the script.

## How to download
1. Select the script you need.
2. Download the `.py` file.

## Notes

- Make sure the required input files are available before running the script.
- Do not change the column names or file structure unless required.
- If a script requires specific supporting files, they will be mentioned in the relevant section.
- Some scripts needs to change the file/folders path. Some needs to change the data in those folder as we cannot change the path
- Save the all the file used for that script for easy reference
- Create new folder for each month

## Instruction
Download **LLD Market Estimation V10.6.2026** folder

<br>**BR_Mapping.py**
1. Save the latest O+O report in : *LLD Market Estimation  V10.6.2026 > LLD Market numbers >Actual Market > MY/SG After Monthly Split*
2. In *LLD Market Estimation  V10.6.2026 > LLD Market numbers > Raw BR* folder rename last quarter BR data folder. Eg. MY_Q1'26. Make sure new BR data folder named as "MY" and "SG"
3. Things to change in script: <br>*a. output_root - "Actual Market" folder path <br> b. raw_br_root - "Raw BR" folder path <br> c."ooo_report_by_country" - Luxe O+O report path*
4. Change the Quartertarget_year (if needed) and target_quarter

<br> **Estimation MY**
1. Create 1 folder and put currect working month **MMPR** file, **O+O report**, and **BR After Split file MY** file for easy ref
2. Download the last month's Luxe report. Filter to market 2026. Change the skincare with lowest value to Hair/Others. 
3. Things to change in script: <br>*a. MMPR - current working month MMPR <br> b. root_out = Estimation Market > MY folder <br> c. lux - last working month's (Market Hair/Others remain as Hair/Others) <br> d. br = MY BR Monthly split (the BR_Mapping.py' scripts output for MY) <br> e. update_month - current working month <br> f. current_Quarter - working month's quarter*

<br> **Estimation SG**
1. Download Mkt est 2026 for currect month file. Usually Siew Mun give this. Or can get from Celine.
2. Download the last month's Luxe report. Filter to market 2025 and 2026. Change the skincare with lowest value to Hair/Others.
3. Things to change in script: <br>*a. MKT - current working month Mkt est 2026 file <br> b. root_out = Estimation Market > SG folder <br> c. lux - last working month's (Market Hair/Others remain as Hair/Others) <br> d. br = MY BR Monthly split (the BR_Mapping.py' scripts output for SG) <br> e. update_month - current working month <br> f. current_Quarter - working month's quarter*
4. Change the market number both 2025 and 2026
