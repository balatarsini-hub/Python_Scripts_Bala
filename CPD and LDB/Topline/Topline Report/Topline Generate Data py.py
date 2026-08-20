# %%
#Importing functions
import os
import pandas as pd
import matplotlib as plt
import numpy as np
from openpyxl import load_workbook

import sys, subprocess
# subprocess.check_call([sys.executable, "-m", "pip", "install", "-U", "python-calamine", "pandas", "xlsxwriter"])
subprocess.check_call([
    sys.executable,
    "-m",
    "pip",
    "install",
    "-U",
    "python-calamine",
    "pandas",
    "xlsxwriter",
    "openpyxl"
])

# %%
# Specify the full path to your Excel file
dir = os.getcwd()
# xl = pd.ExcelFile(f'{dir}/../Raw Data from NIQD/MY CPD Cosmetics Topline (Use This) (1).xlsx')  # change naming to each NIQD data file

src = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Raw\MY LDB Skincare Topline (Use this) Corn (1).xlsx"
src = src.replace("\\", "/")
dst = src.replace(".xlsx", "_CLEAN.xlsx")

# read original with calamine (tolerant reader)
xl = pd.ExcelFile(src, engine="calamine")

# write out a clean copy using xlsxwriter
with pd.ExcelWriter(dst, engine="xlsxwriter") as writer:
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        df.to_excel(writer, sheet_name=sheet, index=False)

# now load the clean copy with openpyxl
excel_file = pd.ExcelFile(dst, engine="openpyxl")

# finally, get sheet names
sheet_names = excel_file.sheet_names

all_dfs = {}

# Function to clean column names
def clean_column_name(column_name):
    if isinstance(column_name, str):
        # Extract only the month part from the column name
        return column_name.split(" - ")[0]
    else:
        return column_name

# Function to remove footer rows
def remove_footer_rows(df):
    return df.iloc[:-5]

# Iterate over each sheet and process it
for sheet_name in sheet_names:
    df1 = excel_file.parse(sheet_name, header=8)
    df = pd.DataFrame(df1)
    
    # Initialize the previous_column with the first column name
    previous_column = df.columns[0]
    
    # Iterate through the columns and update 'Unnamed' columns based on the previous named column
    for i, column in enumerate(df.columns):
        if isinstance(column, str) and column.startswith('Unnamed'):
            df.rename(columns={column: previous_column}, inplace=True)
        else:
            previous_column = column
    
    # Clean the column names
    df.columns = [clean_column_name(col) for col in df.columns]
    
    # Remove footer rows
    df = remove_footer_rows(df)
    
    all_dfs[sheet_name] = df


# %% [markdown]
# # Removing the 'Index' Sheet

# %%
# Remove the 'Index' sheet
if 'Index' in all_dfs:
    del all_dfs['Index']

# 'all_data' is a dictionary with DataFrames for all sheets in the Excel file
# You can access each DataFrame using its sheet name
for sheet_name, df in all_dfs.items():
    print(f"Sheet Name: {sheet_name}")

# %% [markdown]
# # Report with Column

# %%
# Specify the folder for the cleaned data
output_folder = f'{dir}/Full Data'

# Ensure the output folder exists, create it if not
os.makedirs(output_folder, exist_ok=True)

# Assuming 'all_dfs' is your dictionary of DataFrames
# Create the output Excel file
base_file_name = os.path.splitext(os.path.basename(excel_file))[0]
output_excel_file = os.path.join(output_folder, f"{base_file_name}_full.xlsx")

# Save all cleaned DataFrames to a single Excel file
with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
    for sheet_name, df in all_dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"The cleaned DataFrames have been saved to: {output_excel_file}")

# %% [markdown]
# # Transform Dataframe to selected columns

# %%
# Define the minimum number of columns required
minimum_columns_required = 5 # this is because the sheet for table (Num,AVG,WTD) have only that particular month

# Define the columns to select by index
selected_columns_indices = [0, 1, 3, 4, 5, 7, 8, 9, 11, 12, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35,36]

# Iterate over each sheet and select specified columns based on their indices
for sheet_name, df in all_dfs.items():
    if df.shape[1] >= minimum_columns_required:
        if df.shape[1] == 37:
            df_cleaned = df.iloc[:, selected_columns_indices]
            all_dfs[sheet_name] = df_cleaned
        else:
            print(f"Total Columns in '{sheet_name}' not matching")
    else:
        print(f"Skipping '{sheet_name}' due to fewer than {minimum_columns_required} columns")

df.tail()

# %% [markdown]
# # Cleaned Output File

# %%
# Specify the folder for the cleaned data
output_folder = f'{dir}/Cleaned Generated Data'

os.makedirs(output_folder, exist_ok=True)
if not os.path.exists(output_folder):
    raise FileNotFoundError(f"Output folder could not be created: {output_folder}")


# Assuming 'all_dfs' is your dictionary of DataFrames
# Create the output Excel file
base_file_name = os.path.splitext(os.path.basename(excel_file))[0]
output_excel_file = os.path.join(output_folder, f"{base_file_name}_cleaned.xlsx")

# Save all cleaned DataFrames to a single Excel file
with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
    for sheet_name, df in all_dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"The cleaned DataFrames have been saved to: {output_excel_file}")

# %% [markdown]
# # Combining first 2 rows to perform calculation

# %%
# Read all sheets into a dictionary of DataFrames
all_dfs = all_dfs

# Create an empty DataFrame to store the concatenated data
concatenated_df_before = pd.concat(all_dfs.values(), axis=1)

# Concatenate the first two rows manually with a separator
concatenated_df_before.columns = concatenated_df_before.columns.astype(str) + ' | ' + concatenated_df_before.iloc[0].astype(str)

# drop the original first two rows
concatenated_df_before = concatenated_df_before.drop([0]).reset_index(drop=True)

# %% [markdown]
# # Perform Calculation to column name ends with % Chg YA

# %%
test = concatenated_df_before.copy()
# processed_columns = set()  # Set to track processed columns

# for column in test.columns:
#     if column.endswith('% Chg YA') and column not in processed_columns:
#         test[column] = test[column] / 100
#         processed_columns.add(column)

processed_columns = set()

for column in test.columns:
    if str(column).endswith('% Chg YA') and column not in processed_columns:
        test[column] = test[column] / 100
        processed_columns.add(column)


# %% [markdown]
# # Merge Output File (to populate monthly report)

# %%
# Get the base name of the original Excel file 
excel_file_name = os.path.splitext(os.path.basename(excel_file))[0]

# Specify the folder 
output_folder = r"C:\Users\balatarsini_avinitya\Downloads\Topline Report - July\Generated Data"

# Ensure the output folder exists, create it if not
if not os.path.exists(output_folder):
    print(f"Output folder does not exist. Creating: {output_folder}")
    os.makedirs(output_folder, exist_ok=True)
    
# os.makedirs(output_folder, exist_ok=True)

file_name = os.path.join(output_folder, f"{excel_file_name}_merged.csv")
print(excel_file_name)
# Save the modified DataFrame to the generated file
test.to_csv(file_name, index=False)

print(f"The modified DataFrame has been saved to: {file_name}")

# %%



