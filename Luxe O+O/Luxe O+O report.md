# Luxe O+O report for MY and SG

This repository contains instruction for Luxe O+O report preparation. 
All my folders is now available in teams under "Bala" folder for reference

## LUXE Market Scripts
*GR - Gold Retail file
MBGD - Monthly Brand Growth Distribution Count (Total Malaysia ) mm 2026
GR from Carrese - SG_Gold Retail - Sell Out by POS_2026_mm_mm*
DS Department Stores 
Ecomm - Ecomm Manual Adjustment

| Files Needed | Purpose | Where can find | 
|---|---|---|
| GR without Zone | Others stored as Hair/Others. Use to prepare data for MBGD file |2026: Last month's folder <br> 2025: In Team's:<br> MY-AA > Documents > General > 2025 > LLD > LLD O+O > Inputs Files|
| GR with Zone | Others store as Others. Contains all offline and Sephora offline channel's data. Use for Luxe report preparation |Can get from previous month's folder|
|Ecomm| Use to prepare data for MBGD file submission |In teams: <br> MY-MYSG LUXE SISOSIT > Documents > SISOSIT Report > [ADMIN only] SISOSIT Input Data > SISOSIT Input > Data Source <br> download it|
|OMS online data | To get E-Boutiques, Lazada, Shopee, and TikTok data | Ash has the script|
| SISOSIT | L'Oreal's internal report. O+O report needs to align with this file | In teams: <br> MY-MYSG LUXE SISOSIT > Documents > SISOSIT Report > (Open the latest T0* folder)|
|MMPR| Use to estimate the MY market number for current working month and get monthly brand ranking and evo% values | Get from Peggy (DBR person) monthly by submitting the MBGD file|
|Mkt est 2026_month| To estimate SG market values | Usually Siew Mun provide this / Can get from Celine|
|SG DS Brand Ranking 2026 - mm | To get DS brank ranking bt stores | SG DS Inputs / Siew Mun provide this|
|SG Rank Cleanup mm 26| To cleanup the data from SG DS Brand Ranking 2026 - mm file | From last month folder|
|GR from Carrese | To get Ds.com numbers | Can get from last month folder/ Caresse chat|
|MBGD | To get MY market estimation numbers | From Peggy's email|
|MBGD_working | Working for MBGD file from Peggy | Data from GR this year and last year same month |
|


### Working for each files:
<br>**a. GR without Zone**
<br> 1. In *Brand&Cat&Store_MY* sheet  Add 50 rows above the highlighted cells to avoid overlapping while refresh <br> 2. Change cell G2 to current working month (type manually) <br> 3.	Refresh Brand&Cat&Store_MY and Shu Workshop Online_MY sheet <br> 4. Update the Ecomm numbers in the highlighted cells ( filter to MY/SG → Filter out Aesop ) <br> 5. Remove the empty rows and Drag the formula until last row 

<br>**b. GR with zone**
<br> 1. In *Brand&Cat&Store_MY* & *Brand&Cat&Store_SG* sheet  Add 50 rows above the highlighted cells to avoid overlapping while refresh <br> 2. In *Brand&Cat&Store_MY* sheet cell G2, change to current working month (type manually) <br> 3. Refresh sheet Brand&Cat&Store_MY & Brand&Cat&Store_SG <br> 4. Remove the empty rows and Drag the formula until last row 

<br>**c. SG Rank Cleanup mm 26**
<br> 1. Copy paste the data from *SG DS Brand Ranking 2026 - mm* in SG Rank Cleanup mm 26 file in "Dec" sheet. Only copy ranking from 1 t0 30

<br>**d. GR from Carrese**
<br> 1. Refresh the *Gold Retail (DS.com)* sheet <br> 2. Filter to current working month <br> 3. Add Category axis if not available in the pivot table → Add *"Metier hierarchy (axis)"* in Rows below *"Signature hierarchy"* (Search *Metier hierarchy* in Field List

<br>**e. MBGD_working**
<br> 1. In *"DBR Data"* sheet:<br> Upper table: Change the data source with **2026** GR without Zone file <br> Lower table: Change the data source with **2025** GR without Zone file <br> Cross the numbers in *"DBR Data"* sheet with GR files and *"Brand Growth Report"* sheet with *"DBR Data"* sheet.
<br> Changing data source : Select the pivot table → PivotTable Analysis → Change Data Source → Click “↨” → Select Column A to D in GR file from row 11 until last row → enter → ok → close

<br>**f. MBGD**
<br> 1. Copy and paste the data for only the blue highlighted cells from the **MBGD_working** file, *"Brand Growth Report"* sheet <br> 2. Change the month in cell C3 <br> 3. Send to Peggy
