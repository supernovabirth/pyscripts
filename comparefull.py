import pandas as pd

def read_excel_and_compare(sheet_a_path, sheet_b_path, output_file):
    # Read Excel sheets into Pandas DataFrames
    df_a = pd.read_excel(sheet_a_path)
    df_b = pd.read_excel(sheet_b_path)
    
    # Track differences
    differences = []
    
    # Track missing columns
    missing_columns = set()
    
    # Iterate over rows in dataframe A
    for index, row_a in df_a.iterrows():
        company_id = row_a['COMPANY_ID']
        # Find corresponding row in dataframe B with the same COMPANY_ID
        row_b = df_b[df_b['COMPANY_ID'] == company_id]
        
        # If no corresponding row found, skip
        if len(row_b) == 0:
            differences.append(f"Company with COMPANY_ID {company_id} is missing in sheet B")
            continue
        
        # Compare values in other columns
        for column in df_a.columns:
            if column == 'COMPANY_ID':
                continue  # Skip COMPANY_ID column
            if column not in df_b.columns:
                if column not in missing_columns:
                    missing_columns.add(column)
                    differences.append(f"Column '{column}' is missing in sheet B")
                continue
            value_a = row_a[column]
            value_b = row_b[column].iloc[0]  # Get the first matching row value
            if pd.isna(value_a) and pd.isna(value_b):
                continue  # Skip comparison if both values are missing
            if value_a != value_b:
                differences.append(f"Difference in {column} for COMPANY_ID {company_id}: "
                                   f"Sheet A: {value_a}, Sheet B: {value_b}")
    
    # Iterate over columns in dataframe B to check for extra columns in B
    for column in df_b.columns:
        if column == 'COMPANY_ID':
            continue  # Skip COMPANY_ID column
        if column not in df_a.columns:
            if column not in missing_columns:
                missing_columns.add(column)
                differences.append(f"Column '{column}' is extra in sheet B")
    
    # Write differences to output file
    with open(output_file, 'w') as f:
        for diff in differences:
            f.write(diff + '\n')

# Example usage
sheet_a_path = 'path/to/sheet_a.xlsx'
sheet_b_path = 'path/to/sheet_b.xlsx'
output_file = 'differences.txt'
read_excel_and_compare(sheet_a_path, sheet_b_path, output_file)
