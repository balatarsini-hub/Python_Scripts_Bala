# Topline Report

This repository contains Python scripts used for reporting and data preparation tasks. 
All my folders is now available in teams under "Bala" folder for reference

## Topline Report Scripts

| Script | Purpose |Required files/folders | 
|---|---|---|
| `Topline Generate Data.py` | To clean the Raw file| Topline Nielsen File|
| `MY CPD and LDB.py` | To genereate Topline Report for MY| Clean_Merged file|
| `SG CPD and LDB.py` | To genereate Topline Report for SG| Clean_Merged file|
| `MY and SG Commentary Scripts` | To generate Commentary Scripts for MY and SG| Full Data file|


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
Download **Topline Report - July** folder

<br>**Topline Generate Data**

1. Replace the Nielsen file path. 
2. Run the script for all the files 1 by 1

<br> **MY CPD and LDB**
1. Change the file path for each
2. Data took from Topline Report - July > Generated Data

<br> **SG CPD and LDB**
1. Change the file path for each
2. Data took from Topline Report - July > Generated Data

<br> **MY and SG Commentary Scripts**
1. Change the file path for each
2. Data took from Topline Report - July > Full Data

Check the output in Generated Report with Raw data file
<br>*Do check the in the raw file for ZA lips in MY Cosmetic file have values or not. If got values change in **MY CPD and LDB** cosmetic scrips (ctrl+F > search cosmetic) change ZA "Lips_blank" to "Lips".

## <br> Steps:
1.	Get data from Neilsen
    - Go to favourites (Use the 1s with Corn)
    -	Before download, click all tabs 1 by 1 and let it refresh > Export (Select All reports)
    -	Download 6 files (1 for SG, MY: CPD: Cosmetics, Mass+Medic, Mass only, Male skincare; LDB: skincare)

2.	Place new raw data file path in `Topline Generate Data.py` to generate all the files
   
3.	Place the merged file from Topline generate data output (Topline Report - July > Generated Data) in `MY CPD and LDB.py` and `MY CPD and LDB.py`
    -	Use the output (topline report x2, generated report, topline report) do below:
    -	Check generated file with the raw data
    -	unit with unit file, value with value file
    -	Total CPD = sum of loreal brands (check sum)
    -	Check whether sum of all brands in generated file = 100
    -	After checking everything correct, copy and value paste in teams (in CPD and LDB O+O report)
    -	CPD O+O path in teams: one drive, aa, general, current year, one division excel, updated month, ldb cmi ytd and cpd cmi ytd
    -	Value paste everything 
    -	Make sure all months changed (table, title)
  
4.	Run commentary script
    -	Update and validate summary in final report
    -	share gain need do calculation
    -	for row 2, go to raw data total value table last month, check sales value and perc change
    -	ytd sentence 1: same sheet, ytd sales val % chg ya
    -	ytd sentence 2: sales val table total cpd/ ldb row go to ytd sales val % chg ya, cpd/ market from ytd sentence 1, share of sales value, ytd - (ytd -1) to get pp
    -	ytd sentence 3: copy the whole thing in sales val to a new sheet, select all brands, take ytd share of sales val - (ytd   -1) share of sales val, take the largest val among all brands
    -	for the sum table, check from sales val table and unit table, last 3 months val
    -	for market sum: evo of the latest month, in desc order

