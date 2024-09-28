import os
import pandas as pd

# Directory containing the Excel files
directory = '/Users/asishkumarg/Documents/BD_Reports_25-09-24/Blackduck-Extensions-25-09-24'

# List to store DataFrames from "Sheet 2"
dataframes = []

# Loop through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        filepath = os.path.join(directory, filename)
        
        # Load the specific sheet (Sheet 2) from each Excel file
        try:
            df = pd.read_excel(filepath, sheet_name='2. Open Source Components')
            df['Source_File'] = filename  # Optional: Add a column to track the source file
            dataframes.append(df)
        except ValueError:
            print(f"'Sheet 2' not found in {filename}. Skipping this file.")

# Concatenate all DataFrames from "Sheet 2"
combined_df = pd.concat(dataframes, ignore_index=True)

# Save the combined DataFrame to a new Excel file
output_path = 'combined_sheet_ext_Blackduck_Reports_extensions_25-09-24.xlsx'
combined_df.to_excel(output_path, index=False)

print(f"Data from 'Sheet 2' in all Excel files have been combined and saved as {output_path}")