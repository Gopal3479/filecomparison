import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

def compare_excel_files(file1_path, file2_path, output_path):
    # Read Excel files
    df1 = pd.read_excel(file1_path, header=None)
    df2 = pd.read_excel(file2_path, header=None)
    
    # Create a new workbook
    wb = Workbook()
    comparison_sheet = wb.active
    comparison_sheet.title = "Comparison"
    
    # Setup highlight style
    highlight_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Yellow
    
    # Determine max dimensions
    max_rows = max(len(df1), len(df2))
    max_cols = max(len(df1.columns), len(df2.columns))
    min_rows = min(len(df1), len(df2))
    min_cols = min(len(df1.columns), len(df2.columns))
    
    # Write headers
    comparison_sheet['A1'] = "File 1"
    comparison_sheet[get_column_letter(max_cols + 3) + '1'] = "File 2"
    
    # Write data and compare
    for r in range(max_rows):
        for c in range(max_cols):
            # File 1 data
            if r < len(df1) and c < len(df1.columns):
                val1 = df1.iloc[r, c]
                cell1 = comparison_sheet.cell(row=r+3, column=c+1, value=val1)
            else:
                cell1 = comparison_sheet.cell(row=r+3, column=c+1, value="MISSING")
                cell1.fill = highlight_fill
            
            # File 2 data (offset by max_cols + 2 columns)
            if r < len(df2) and c < len(df2.columns):
                val2 = df2.iloc[r, c]
                cell2 = comparison_sheet.cell(row=r+3, column=c+max_cols+3, value=val2)
            else:
                cell2 = comparison_sheet.cell(row=r+3, column=c+max_cols+3, value="MISSING")
                cell2.fill = highlight_fill
            
            # Compare values if within common dimensions
            if r < min_rows and c < min_cols:
                if (pd.isna(val1) and (pd.isna(val2)):
                    continue
                elif (pd.isna(val1) != (pd.isna(val2)) or (val1 != val2):
                    cell1.fill = highlight_fill
                    cell2.fill = highlight_fill
    
    # Add separation between files
    sep_col = max_cols + 2
    for r in range(1, max_rows + 3):
        comparison_sheet.cell(row=r, column=sep_col).value = "|"
    
    # Save the result
    wb.save(output_path)

# Example usage
if __name__ == "__main__":
    compare_excel_files('file1.xlsx', 'file2.xlsx', 'comparison_result.xlsx')
