import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

def compare_excel_files(file1_path, file2_path, output_path, 
                        row_key_cols=None, col_header_row=0):
    """
    Compare two Excel files with flexible row/column matching
    
    Args:
        file1_path: Path to first Excel file
        file2_path: Path to second Excel file
        output_path: Output file path
        row_key_cols: List of columns to use as row identifiers
        col_header_row: Row index containing column headers (None for no headers)
    """
    # Read Excel files
    df1 = pd.read_excel(file1_path, header=col_header_row, dtype=str)
    df2 = pd.read_excel(file2_path, header=col_header_row, dtype=str)
    
    # Fill NaN values for consistent comparison
    df1 = df1.fillna("")
    df2 = df2.fillna("")
    
    # Default to position-based comparison if no keys provided
    if row_key_cols is None:
        return compare_by_position(df1, df2, output_path)
    
    # Align data using row keys and column headers
    return compare_by_keys(df1, df2, output_path, row_key_cols, col_header_row)

def compare_by_position(df1, df2, output_path):
    """Position-based comparison (original method)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"
    highlight = PatternFill(start_color='FFFF00', fill_type='solid')  # Yellow
    
    # Write headers
    ws['A1'] = "File 1"
    ws.cell(1, len(df1.columns) + 3, "File 2")
    
    # Write data
    max_row = max(len(df1), len(df2))
    max_col = max(len(df1.columns), len(df2.columns))
    
    for r in range(max_row):
        for c in range(max_col):
            # File 1 data
            if r < len(df1) and c < len(df1.columns):
                val1 = df1.iloc[r, c]
                cell1 = ws.cell(r+3, c+1, str(val1))
            else:
                cell1 = ws.cell(r+3, c+1, "MISSING")
                cell1.fill = highlight
            
            # File 2 data
            col2 = c + max_col + 3
            if r < len(df2) and c < len(df2.columns):
                val2 = df2.iloc[r, c]
                cell2 = ws.cell(r+3, col2, str(val2))
            else:
                cell2 = ws.cell(r+3, col2, "MISSING")
                cell2.fill = highlight
            
            # Compare
            if (r < min(len(df1), len(df2)) and (c < min(len(df1.columns), len(df2.columns))):
                if str(df1.iloc[r, c]) != str(df2.iloc[r, c]):
                    cell1.fill = highlight
                    cell2.fill = highlight
    
    # Add separator
    sep_col = max_col + 2
    for r in range(1, max_row + 4):
        ws.cell(r, sep_col, "|")
    
    wb.save(output_path)

def compare_by_keys(df1, df2, output_path, row_key_cols, col_header_row):
    """Key-based comparison for reordered rows/columns"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"
    highlight = PatternFill(start_color='FFFF00', fill_type='solid')  # Yellow
    
    # Prepare row keys
    df1 = df1.set_index(row_key_cols)
    df2 = df2.set_index(row_key_cols)
    
    # Align rows and columns
    all_rows = df1.index.union(df2.index)
    all_cols = df1.columns.union(df2.columns)
    
    df1 = df1.reindex(index=all_rows, columns=all_cols).fillna("")
    df2 = df2.reindex(index=all_rows, columns=all_cols).fillna("")
    
    # Reset index for display
    df1 = df1.reset_index()
    df2 = df2.reset_index()
    
    # Write headers
    ws['A1'] = "File 1"
    ws.cell(1, len(df1.columns) + 3, "File 2")
    
    # Write column headers
    for c, col in enumerate(df1.columns, 1):
        ws.cell(2, c, str(col))
    
    for c, col in enumerate(df2.columns, 1):
        ws.cell(2, c + len(df1.columns) + 2, str(col))
    
    # Write data and compare
    for r, (_, row1), (_, row2) in enumerate(zip(df1.iterrows(), df2.iterrows()), 3):
        for c, col in enumerate(df1.columns, 1):
            val1 = str(row1[col])
            cell1 = ws.cell(r, c, val1)
            
            val2 = str(row2[col])
            cell2 = ws.cell(r, c + len(df1.columns) + 2, val2)
            
            if val1 != val2:
                cell1.fill = highlight
                cell2.fill = highlight
    
    # Add separator
    sep_col = len(df1.columns) + 2
    for r in range(1, len(df1) + 4):
        ws.cell(r, sep_col, "|")
    
    # Highlight missing keys
    missing_keys = all_rows.difference(df1.index.intersection(df2.index))
    if not missing_keys.empty:
        key_cols = list(range(1, len(row_key_cols) + 1))
        
        for r in range(3, len(df1) + 3):
            key_vals = [str(ws.cell(r, c).value) for c in key_cols]
            if any(k in missing_keys for k in key_vals):
                for c in range(1, len(df1.columns) + len(df2.columns) + 3):
                    ws.cell(r, c).fill = PatternFill(
                        start_color='FF9999', fill_type='solid')  # Light red
    
    wb.save(output_path)

if __name__ == "__main__":
    # Example usage:
    # Position-based comparison:
    # compare_excel_files('file1.xlsx', 'file2.xlsx', 'output.xlsx')
    
    # Key-based comparison (columns A and B as row keys, header in row 0):
    compare_excel_files(
        'file1.xlsx',
        'file2.xlsx',
        'comparison.xlsx',
        row_key_cols=['ID', 'Category'],  # Column names or indices
        col_header_row=0                  # Header row index
    )
