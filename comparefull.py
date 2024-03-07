import pandas as pd

def read_column_names_to_skip(skip_file):
    try:
        with open(skip_file, 'r') as f:
            skip_columns = f.read().splitlines()
        return skip_columns
    except FileNotFoundError:
        print("Skip file not found.")
        return []

def read_excel_and_compare(sheet_a_path, sheet_b_path, output_file, skip_file):
    # Read Excel sheets into Pandas DataFrames
    df_a = pd.read_excel(sheet_a_path)
    df_b = pd.read_excel(sheet_b_path)
    
    # Read column names to skip from the skip file
    skip_columns = read_column_names_to_skip(skip_file)
    
    # Track differences
    differences = []
    
    # Track missing columns
    missing_columns = set()
    
    # Track columns with differences
    different_columns = set()
    
    # Track values present in one sheet but empty in the other
    empty_values = {'A': {}, 'B': {}}
    
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
            if column == 'COMPANY_ID' or column in skip_columns:
                continue  # Skip COMPANY_ID column and columns to be skipped
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
                different_columns.add(column)
            if pd.isna(value_a) and not pd.isna(value_b):
                empty_values['A'].setdefault(column, []).append(company_id)
            elif not pd.isna(value_a) and pd.isna(value_b):
                empty_values['B'].setdefault(column, []).append(company_id)
    
    # Iterate over columns in dataframe B to check for extra columns in B
    for column in df_b.columns:
        if column == 'COMPANY_ID' or column in skip_columns:
            continue  # Skip COMPANY_ID column and columns to be skipped
        if column not in df_a.columns:
            if column not in missing_columns:
                missing_columns.add(column)
                differences.append(f"Column '{column}' is extra in sheet B")
    
    # Write differences to output file
    with open(output_file, 'w') as f:
        for diff in differences:
            f.write(diff + '\n')
        f.write("\nColumns with differences:\n")
        for column in different_columns:
            f.write(column + '\n')
        f.write("\nValues present in one sheet but empty in the other:\n")
        for sheet, values in empty_values.items():
            for column, ids in values.items():
                f.write(f"In sheet {sheet}, column '{column}' has empty values for COMPANY_IDs: {', '.join(map(str, ids))}\n")

# Example usage
sheet_a_path = 'path/to/sheet_a.xlsx'
sheet_b_path = 'path/to/sheet_b.xlsx'
output_file = 'differences.txt'
skip_file = 'path/to/skip_columns.txt'
read_excel_and_compare(sheet_a_path, sheet_b_path, output_file, skip_file)
