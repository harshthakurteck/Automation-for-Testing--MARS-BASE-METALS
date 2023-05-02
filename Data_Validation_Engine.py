import pandas as pd
pd.set_option('display.max_columns', None) # Show all columns
pd.set_option('display.max_rows', None) # Show all rows
pd.set_option('display.expand_frame_repr', False) # Avoid wrapping the output to multiple lines

import openpyxl
import numpy as np
import io
import sys


# Create a StringIO object
output_buffer = io.StringIO()

# Redirect stdout to the StringIO object
sys.stdout = output_buffer

#Read the Excel file 
excel_file_path = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\Testing Status.xlsx"
page_name = 'Testing Tracker'
excel_file = pd.read_excel(excel_file_path, sheet_name=page_name, header = 1)

filter_release_name = excel_file.loc[excel_file['Release Name'] == "D1"]
print("Values in Release D1")
#null_column_name = 'Release Planned Date'
filter_release_name['Release Planned Date '] = filter_release_name['Release Planned Date'].replace({pd.NaT: np.nan, np.nan: ''})
#filter_release_name[null_column_name] = np.where(pd.isnull(filter_release_name[null_column_name]), pd.NaT, filter_release_name[null_column_name].fillna(''))

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
print(filter_release_name)

## Load the CSV files into pandas dataframes
file1 = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file1.csv"
file2a = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file2.csv"
file2b = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file2b.csv"
file3a = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file3.csv"
file3b = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file3b.csv"
file4a = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file4a.csv"
file4b = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file4b.csv"
file4c = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file4c.csv"
file5a = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file5a.csv"
file5b = r"C:\Users\hthakur2\OneDrive - Teck Resources Limited\Documents\Power BI- UAT Automation\powerbi_file5b.csv"

dataframes = []

dataframe1 = pd.read_csv(file1) #D1-DM
dataframes.append(dataframe1)

dataframe2a = pd.read_csv(file2a) #D2 HVC WO
dataframe2b = pd.read_csv(file2b) #D2 CDA Plant 
dataframes.append(dataframe2a)
dataframes.append(dataframe2b)

dataframe3a = pd.read_csv(file3a) #D3 RDM FMS
dataframe3b = pd.read_csv(file3b) #D3 RDM Plant
dataframes.append(dataframe3a)
dataframes.append(dataframe3b)

dataframe4a = pd.read_csv(file4a)
dataframe4b = pd.read_csv(file4b)
dataframe4c = pd.read_csv(file4c)
dataframes.append(dataframe4a)
dataframes.append(dataframe4b)
dataframes.append(dataframe4c)


dataframe5a = pd.read_csv(file5a)
dataframe5b = pd.read_csv(file5b)
dataframes.append(dataframe5a)
dataframes.append(dataframe5b)

# Define functions to perform data validations


# Define functions to perform data validations
def check_for_nan_values(dataframe):
    nan_values = dataframe[dataframe.isna().any(axis=1)].fillna('')
    if len(nan_values) > 0:
        print(f"\n{len(nan_values)} NaN values found in the following rows:")
        print("\n") 
        print(nan_values)
        print("\n")
    else:
        print("No NaN values found.")


#This function checks for negative values in numeric columns of the dataframe
def check_for_negative_values(dataframe):
    # Select only numeric columns
    numeric_cols = dataframe.select_dtypes(include=['float', 'int']).columns
    
    # Check for negative values
    negative_values = dataframe[numeric_cols].lt(0).sum()
    
    if negative_values.sum() > 0:
        print(f"\n{negative_values.sum()} Negative values found in the following rows:")
        print("\n")
        print(dataframe[dataframe[numeric_cols].lt(0).any(axis=1)])
        print("\n")

    else:
        print("No negative values found")
        print("\n")


def check_for_values_above_100(dataframe):
    # Get percentage columns
    percentage_cols = [col for col in dataframe.columns if "%" in col]
    if not percentage_cols:
        print("No percentage columns found. No value is above 100%")
        return None
    # Remove percent sign and convert to numeric
    percentage_cols_numeric = dataframe[percentage_cols].replace('%', '', regex=True).apply(pd.to_numeric, errors='coerce')
    # Check for values above 100%
    values_above_100 = percentage_cols_numeric[percentage_cols_numeric > 100]
    if values_above_100.isnull().values.any():
        values_above_100 = values_above_100.dropna(how='all')
    if not values_above_100.empty:
        print(f"\n{len(values_above_100)} values above 100% found in the following rows:")
        print("\n") 
        print(dataframe[dataframe.index.isin(values_above_100.index)])
        print("\n")
    else:
        print("No values above 100% found.")
        print("\n\n")        


def check_consecutive_values(df):
    date_range_col = next((col for col in ['Planning Week Date Range (Mon - Sun)', 'Month-Year'] if col in df.columns), None)
    if not date_range_col:
        print('Neither Planning Week Date Range nor Month-Year column found')
        return
    duplicates_found = False
    for col in df.columns:
        if col not in ['Planning Week Date Range (Mon - Sun)', 'Mine/Mill', 'Site Code']:
            # Find consecutive values that repeat at least three times
            for site_code, site_df in df.groupby('Site Code'):
                site_df = site_df.sort_values(by=[date_range_col])
                consecutive_count = 1
                last_value = None
                start_index = None
                for idx, row in site_df.iterrows():
                    value = row[col]
                    if value == last_value:
                        consecutive_count += 1
                        if consecutive_count == 3:
                            start_index = idx - 1
                    else:
                        consecutive_count = 1
                        last_value = value
                        start_index = None
                    if start_index is not None:
                        end_index = idx
                        date_range = row[date_range_col]
                        duplicates_found = True
                        if duplicates_found:
                            print('\nDuplicates found:')
                        print(f"Site Code: {site_code}, Column: {col}, Value: {value}, {date_range_col}: {date_range}")
    if not duplicates_found:
        print('No duplicate values found')


for i, df in enumerate(dataframes):
    print("\n\n" + f"CHECKS FOR FILE {i+1}...\n")
    
    check_for_nan_values(df)
    check_for_negative_values(df)
    check_for_values_above_100(df)
    check_consecutive_values(df)
   
    
sys.stdout = sys.__stdout__
output = output_buffer.getvalue()
with open("output.txt", "w") as f:
    # Write the output to the file
    f.write(output)




