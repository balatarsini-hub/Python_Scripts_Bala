# LDB O+O

This repository contains Python scripts used for reporting and data preparation tasks. 
All my folders is now available in teams under "Bala" folder for reference

## LDB O+O Scripts

| Script | Purpose/Output File name |Required files/folders | 
|---|---|---|
| `MY LDB O+O` | MY CPD JUN 2026 O+O| Nielsen O+O and OMT|
| `SG LDB O+O` | SG CPD JUN 2026 O+O| Nielsen O+O and OMT|
| `Rename OMT File` | Change OMT file name| OMT Data|
|`Bodycare estimate` | Bodycare pulse report split <br>- LDB Data Output Jul'26|*1. MY and SG Pulse Report <br> 2. MYSG ONE CPD CMI YTD Jun'26*|
|`Suncare_estimation_LDB - Including Competitor brands` | Suncare pulse report split <br>- Suncare LDB Data Output Jul'26|*1. MY and SG Pulse Report <br> 2. MYSG ONE CPD CMI YTD Jun'26*|


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
Download **CPD and LDB O+O and Offline Bodycare ingestion** folder

<br>**MY LDB O+O & SG LDB O+O**

1. Refresh Nielsen O+O file <br>
2. Put `Rename OMT File` output in CPD & LDB O+O - June > loreal-report-automation (2) > Data Source > OMT - O+O

<br> **Rename OMT File**
1. Download OMT Data put in 1 folder <br>
2. Change the path

<br> **Bodycare estimate**
1. Update files to the latest path <br>
2. Change REPORT_HEADER_ROW if needed

<br> **Suncare estimate**
1. Update files to the latest path <br>
2. Change REPORT_HEADER_ROW, MY_MARKET_ROW, SG_MARKET_ROW, brand's row if needed


