import pandas as pd
import openpyxl
from xlsxwriter import Workbook

# load the Excel file
file_path = "data cleaning/Marketing Analysis Data.xlsx" 
output_excel = "data cleaning/Updated_Marketing_Analysis.xlsx"

# read all sheets
sheets = pd.read_excel(file_path, sheet_name=None)

# list of sheets to skip
skip_sheets = [
    "Single Send (Content)",
    "Single Send BY EMAIL",
    "Single Send Summary",
    "Info Session Summary",
    "Admit Data - 2023 to 2025",
    "Admit Journey Summary",
    "Prospect Data - 2023 to 2025",
    "Prospect Journey Summary",
]

# copy all the rows from "Journey Name" and "Subject Line" 
# columns from every sheet (except the ones in the skip list) 
# into a new sheet called "journey_subject_keys.csv"
journey_subject_data = []
valid_sheets = {}
for sheet_name, df in sheets.items():
    #print(df.columns)

    if sheet_name not in skip_sheets:
        journey_subject_data.append(df[["Journey Name", "Subject Line"]])
        valid_sheets[sheet_name] = df

# combine all extracted data
journey_subject_df = pd.concat(journey_subject_data).drop_duplicates().reset_index(drop=True)

# generate a numerical key (aka primary key) for each row in this new sheet
journey_subject_df.insert(0, "Key", range(1, len(journey_subject_df) + 1))

# save new data to CSV
#journey_subject_keys_path = "/Users/nikhitavysyaraju/Downloads/INFSC 1740/data/journey_subject_keys.csv"
#journey_subject_df.to_csv(journey_subject_keys_path, index=False)

# Add the journey_subject_keys as a sheet
sheets["Journey Subject Keys"] = journey_subject_df

# map "Journey Name" and "Subject Line" to their generated keys
key_mapping = journey_subject_df.set_index(["Journey Name", "Subject Line"])['Key'].to_dict()

# for each VALID sheet in the excel workbook, 
for sheet_name, df in valid_sheets.items():
    # add a new blank first column that is called "Key"
    df.insert(0, "Key", df.set_index(["Journey Name", "Subject Line"]).index.map(key_mapping))
    
    # drop "Journey Name" and "Subject Line"
    df.drop(columns=["Journey Name", "Subject Line"], inplace=True)
    
    # save the updated DataFrame back
    sheets[sheet_name] = df

# save the updated workbook
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Processing complete. Updated Excel file saved.")
