import pandas as pd

def read_excel_and_compare(sheet_a_path, sheet_b_path):
    # Read Excel sheets into Pandas DataFrames
    df_a = pd.read_excel(sheet_a_path)
    df_b = pd.read_excel(sheet_b_path)
    
    # Track differences
    differences = []
    
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
            if row_a[column] != row_b[column].iloc[0]:  # Get the first matching row value
                differences.append(f"Difference in {column} for COMPANY_ID {company_id}: "
                                   f"Sheet A: {row_a[column]}, Sheet B: {row_b[column].iloc[0]}")
    
    # Iterate over rows in dataframe B to check for extra companies in B
    for index, row_b in df_b.iterrows():
        company_id = row_b['COMPANY_ID']
        # Find corresponding row in dataframe A with the same COMPANY_ID
        row_a = df_a[df_a['COMPANY_ID'] == company_id]
        
        # If no corresponding row found, it means company is extra in sheet B
        if len(row_a) == 0:
            differences.append(f"Company with COMPANY_ID {company_id} is extra in sheet B")
    
    # Print differences
    for diff in differences:
        print(diff)

# Example usage
sheet_a_path = 'path/to/sheet_a.xlsx'
sheet_b_path = 'path/to/sheet_b.xlsx'
read_excel_and_compare(sheet_a_path, sheet_b_path)
