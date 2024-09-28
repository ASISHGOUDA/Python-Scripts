import os
import pandas as pd

# Directory containing the Excel files
directory = '/Users/asishkumarg/Downloads/blackduck_fortify_mdh_26-09-24'

# Lists to store DataFrames for Blackduck and Fortify files
blackduck_dataframes = []
fortify_dataframes = []
ignored_files = []

# Loop through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Count occurrences of "fortify" and "SCA_RESULTS"
        fortify_count = filename.lower().count('fortify')
        blackduck_count = filename.lower().count('sca_results')

        # Classify the file
        if fortify_count > blackduck_count:
            # Load Fortify file and add Source_File column
            df = pd.read_excel(os.path.join(directory, filename), sheet_name='2. Source Code Issues')
            df['Source_File'] = filename  # Add the source file name
            fortify_dataframes.append(df)
        elif blackduck_count > 0:
            # Load Blackduck file and add Source_File column
            df = pd.read_excel(os.path.join(directory, filename), sheet_name='2. Open Source Components')
            df['Source_File'] = filename  # Add the source file name
            blackduck_dataframes.append(df)
        else:
            ignored_files.append((filename, "Does not match Blackduck or Fortify criteria"))

# Create a new Excel file to save combined reports
output_path = 'combined_reports_mdh-26-09-24.xlsx'

# Save DataFrames to separate sheets in the output Excel file
with pd.ExcelWriter(output_path) as writer:
    if blackduck_dataframes:
        pd.concat(blackduck_dataframes).to_excel(writer, sheet_name='Blackduck Reports', index=False)
    if fortify_dataframes:
        pd.concat(fortify_dataframes).to_excel(writer, sheet_name='Fortify Reports', index=False)

# Print the output filename
print(f"\nCombined reports have been saved to: {output_path}")

# Print counts and files processed successfully
print(f"\nTotal Blackduck files processed: {len(blackduck_dataframes)}")
if blackduck_dataframes:
    print("Blackduck Files:")
    for df in blackduck_dataframes:
        print(f"- {df['Source_File'].iloc[0]}")

print(f"\nTotal Fortify files processed: {len(fortify_dataframes)}")
if fortify_dataframes:
    print("Fortify Files:")
    for df in fortify_dataframes:
        print(f"- {df['Source_File'].iloc[0]}")

# Print ignored files and reasons
print("\nIgnored files:")
for file, reason in ignored_files:
    print(f"{file}: {reason}")
