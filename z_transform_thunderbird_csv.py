#Thunderbird has 3 different type of contacts, which can be exported as .csv files.
#Each .csv may have different combinations of data, eg in one the email addresses are empty and another they may be filles in; the same for phone.
#
#After concatenating all 3 .csv files and removing the 2 extraneous header rows (leaving row 1 as the only eader row),
#this .py reads the concatenated .csv file and tries to combine them into a single "matched" .csv which can be imported into Outlook People.
#
#Automatic Column Detection:
# The script reads the first row of the CSV file to detect column names dynamically. This approach assumes that the first row contains the header with column names.
#Adjustments: 
# Ensure that key_column is specified correctly according to your CSV file's structure. You may need to adjust how columns are detected or processed based on specific nuances in your data.
#Handling Empty Cells:
# The lambda functions prioritize non-empty cells (pd.notna(v)) when aggregating each column's data.
#
# This approach allows your Python script to dynamically handle CSV files with varying column names and perform aggregation based on a specified key column, automating much of the data processing task.
# Adjustments can be made based on your specific data requirements and structure.\
#

import pandas as pd

# Specify fully qualified paths using raw string literals
input_file = r'C:\TEMP\All_Sheets.csv'
output_file = r'C:\TEMP\All_Sheets_merged.csv'

# Read the CSV file with headers
df = pd.read_csv(input_file, header=0)  # Assuming headers are in the first row (index 0)

# Manually specify the key column (adjust as per your data structure)
key_column = 'Display Name'

# Get the number of rows in the input CSV file (excluding header)
input_rows = df.shape[0]  # Total rows including header: df.shape[0], excluding header: df.shape[0] - 1

# Get the column names from the DataFrame
columns = df.columns.tolist()

# Define aggregation functions to prioritize non-empty cells
agg_funcs = {}
for col in columns:
    if col != key_column:
        agg_funcs[col] = lambda x: next((v for v in x if pd.notna(v)), None)

# Group by the key column and aggregate using the defined functions
grouped = df.groupby(key_column).agg(agg_funcs).reset_index()

# Write the resulting DataFrame to a new CSV file
grouped.to_csv(output_file, index=False)

# Get the number of rows in the output CSV file
output_df = pd.read_csv(output_file)  # Read the output file back to count rows
output_rows = output_df.shape[0]

print(f"Number of rows in input CSV (excluding header): {input_rows - 1}")
print(f"Number of rows in output CSV: {output_rows}")
print(f"Processed data saved to {output_file}")
