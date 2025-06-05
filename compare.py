import pandas as pd
import numpy as np

def compare_excel_files(file1_path, file2_path, sheet_name1=None, sheet_name2=None, key_columns=None, output_path='excel_differences.xlsx'):
    # Read Excel files
    df1 = pd.read_excel(file1_path, sheet_name=sheet_name1)
    df2 = pd.read_excel(file2_path, sheet_name=sheet_name2)
    
    # Handle key columns
    temp_key = False
    if key_columns is None:
        df1['__temp_row_index__'] = range(1, len(df1) + 1)
        df2['__temp_row_index__'] = range(1, len(df2) + 1)
        key_columns = ['__temp_row_index__']
        temp_key = True
    else:
        missing1 = [col for col in key_columns if col not in df1.columns]
        missing2 = [col for col in key_columns if col not in df2.columns]
        if missing1:
            raise ValueError(f"Key columns {missing1} not found in first file.")
        if missing2:
            raise ValueError(f"Key columns {missing2} not found in second file.")
    
    # Merge dataframes
    merged = pd.merge(df1, df2, on=key_columns, how='outer', suffixes=('_file1', '_file2'), indicator=True)
    
    # Identify removed, added, and common rows
    removed = merged[merged['_merge'] == 'left_only'].copy()
    added = merged[merged['_merge'] == 'right_only'].copy()
    common = merged[merged['_merge'] == 'both'].copy()
    
    # Prepare results
    # Removed rows: original columns from first file
    if not removed.empty:
        removed_cols = key_columns + [col for col in df1.columns if col not in key_columns]
        removed = removed[[col if col in key_columns else f"{col}_file1" for col in removed_cols]]
        removed.columns = removed.columns.str.replace('_file1', '')
    else:
        removed = pd.DataFrame(columns=df1.columns)
    
    # Added rows: original columns from second file
    if not added.empty:
        added_cols = key_columns + [col for col in df2.columns if col not in key_columns]
        added = added[[col if col in key_columns else f"{col}_file2" for col in added_cols]]
        added.columns = added.columns.str.replace('_file2', '')
    else:
        added = pd.DataFrame(columns=df2.columns)
    
    # Changed rows: find differing cells
    changed_rows = []
    base_columns = [col for col in df1.columns if col not in key_columns and col in df2.columns]
    
    for _, row in common.iterrows():
        for col in base_columns:
            val1 = row[f"{col}_file1"]
            val2 = row[f"{col}_file2"]
            if val1 != val2 and not (pd.isna(val1) and pd.isna(val2)):
                key_vals = [row[k] for k in key_columns]
                changed_rows.append(key_vals + [col, val1, val2])
    
    changed_cols = key_columns + ['Column', 'Value_in_File1', 'Value_in_File2']
    changed = pd.DataFrame(changed_rows, columns=changed_cols) if changed_rows else pd.DataFrame(columns=changed_cols)
    
    # Remove temporary key if used
    if temp_key:
        removed = removed.drop(columns=key_columns, errors='ignore')
        added = added.drop(columns=key_columns, errors='ignore')
        changed = changed.drop(columns=key_columns, errors='ignore')
    
    # Write results to Excel
    with pd.ExcelWriter(output_path) as writer:
        removed.to_excel(writer, sheet_name='Removed', index=False)
        added.to_excel(writer, sheet_name='Added', index=False)
        changed.to_excel(writer, sheet_name='Changed', index=False)
    
    return output_path

# Example usage
if __name__ == "__main__":
    file1 = "file1.xlsx"
    file2 = "file2.xlsx"
    output = compare_excel_files(file1, file2, key_columns=['ID'], output_path='differences.xlsx')
    print(f"Differences saved to {output}")
