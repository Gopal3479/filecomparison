import pandas as pd
import numpy as np # Added for robust numeric comparison
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

# Define highlighting styles
CELL_DIFF_FILL = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")     # Tomato
ROW_MATCH_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")     # Light Green
ROW_MISSING_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Light Gray
HEADER_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")        # Light Gray
TOTAL_ROW_FILL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")     # Light Blue

# Border style
THIN_BORDER = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

class ExcelComparator:
    """
    Compares two Excel sheets on a row-by-row basis and generates a detailed 
    side-by-side report highlighting the differences.

    Time Complexity: The overall performance is O(R * C), where R is the maximum
    number of rows and C is the total number of columns across both files. This
    is efficient for the intended task of cell-by-cell inspection.
    """
    def __init__(self, file1_path, file2_path, sheet1_name=None, sheet2_name=None, total_row_identifier: str = "Total"):
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.sheet1_name = sheet1_name
        self.sheet2_name = sheet2_name
        self.total_row_identifier = total_row_identifier
        
        if not self.sheet1_name:
            wb = load_workbook(file1_path, read_only=True)
            self.sheet1_name = wb.sheetnames[0]
            wb.close()
            
        if not self.sheet2_name:
            wb = load_workbook(file2_path, read_only=True)
            self.sheet2_name = wb.sheetnames[0]
            wb.close()
    
    def are_equal(self, a, b):
        """
        FIXED: Robustly checks if two values are equal, correctly handling NaNs, 
        datetimes, floating-point precision, and mixed data types.
        """
        if pd.isna(a) and pd.isna(b):
            return True
        if pd.isna(a) or pd.isna(b):
            return False
        
        # Handle datetime objects
        if isinstance(a, (datetime, pd.Timestamp)) or isinstance(b, (datetime, pd.Timestamp)):
            try:
                return pd.to_datetime(a) == pd.to_datetime(b)
            except (ValueError, TypeError):
                return False

        # Try robust numeric comparison
        try:
            return np.isclose(float(a), float(b), atol=1e-9, rtol=1e-9)
        except (ValueError, TypeError):
            # Fallback to case-insensitive string comparison
            return str(a).strip().lower() == str(b).strip().lower()

    def is_numeric(self, series):
        """Check if a series is numeric"""
        return pd.api.types.is_numeric_dtype(series)
    
    def create_side_by_side_sheet(self, df1, df2, output_wb):
        """Optimized side-by-side comparison with total row."""
        ws = output_wb.create_sheet("Side by Side Comparison")
        common_cols = list(set(df1.columns) & set(df2.columns))
        
        num_cols1 = [col for col in df1.columns if self.is_numeric(df1[col])]
        num_cols2 = [col for col in df2.columns if self.is_numeric(df2[col])]
        common_num_cols = list(set(num_cols1) & set(num_cols2))
        
        header_row = list(df1.columns) + [" | "] + list(df2.columns) + ["Match Status"]
        ws.append(header_row)
        
        # Create dataframes for summation, excluding total rows to prevent double-counting
        df1_for_sum, df2_for_sum = df1, df2
        if self.total_row_identifier:
            if not df1.empty:
                df1_for_sum = df1[~df1.iloc[:, 0].astype(str).str.contains(self.total_row_identifier, case=False, na=False)]
            if not df2.empty:
                df2_for_sum = df2[~df2.iloc[:, 0].astype(str).str.contains(self.total_row_identifier, case=False, na=False)]
        
        # Create total row using the filtered dataframes
        total_row, total_match = [], True
        for col in df1.columns:
            total_row.append(df1_for_sum[col].sum() if col in num_cols1 else "")
        total_row.append(" | ")
        for col in df2.columns:
            total_row.append(df2_for_sum[col].sum() if col in num_cols2 else "")
        
        for col in common_num_cols:
            if not self.are_equal(df1_for_sum[col].sum(), df2_for_sum[col].sum()):
                total_match = False
                break
        total_row.append("Matched" if total_match else "Not Matched")
        ws.append(total_row)
        
        # Apply formatting to header and total rows
        for cell in ws[1]: cell.fill, cell.font, cell.border = HEADER_FILL, Font(bold=True), THIN_BORDER
        for cell in ws[2]: cell.fill, cell.font, cell.border = TOTAL_ROW_FILL, Font(bold=True), THIN_BORDER
        
        # Precompute indices and comparison results
        separator_col, match_status_col = len(df1.columns) + 1, len(header_row)
        file1_end_col, file2_start_col = len(df1.columns), len(df1.columns) + 2
        
        match_status, diff_positions = [], []
        for i in range(max(len(df1), len(df2))):
            row_match, diffs_in_row = True, []
            if i < len(df1) and i < len(df2):
                for col in common_cols:
                    val1, val2 = df1.at[i, col], df2.at[i, col]
                    if not self.are_equal(val1, val2):
                        row_match = False
                        col_idx1 = df1.columns.get_loc(col) + 1
                        # BUG FIX: Correctly calculate the column index for the second dataframe
                        col_idx2 = df2.columns.get_loc(col) + file2_start_col
                        diffs_in_row.append((col_idx1, col_idx2))
            else:
                row_match = False
            
            match_status.append("Matched" if row_match else "Not Matched")
            diff_positions.append(diffs_in_row)
        
        # Write data rows and apply formatting in a single pass
        for i in range(max(len(df1), len(df2))):
            row_data = []
            if i < len(df1): row_data.extend(df1.iloc[i].values)
            else: row_data.extend([""] * len(df1.columns))
            row_data.append(" | ")
            if i < len(df2): row_data.extend(df2.iloc[i].values)
            else: row_data.extend([""] * len(df2.columns))
            row_data.append(match_status[i])
            ws.append(row_data)

            row_idx = i + 3
            fill_color = ROW_MATCH_FILL if match_status[i] == "Matched" else ROW_MISSING_FILL
            for col_idx in range(1, len(header_row) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = THIN_BORDER
                cell.fill = fill_color
            
            ws.cell(row=row_idx, column=separator_col).fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
            for col1, col2 in diff_positions[i]:
                ws.cell(row=row_idx, column=col1).fill = CELL_DIFF_FILL
                ws.cell(row=row_idx, column=col2).fill = CELL_DIFF_FILL

        # Highlight differences in total row
        if not total_match:
            for col in common_num_cols:
                if not self.are_equal(df1_for_sum[col].sum(), df2_for_sum[col].sum()):
                    col_idx1 = df1.columns.get_loc(col) + 1
                    col_idx2 = df2.columns.get_loc(col) + file2_start_col
                    ws.cell(row=2, column=col_idx1).fill = CELL_DIFF_FILL
                    ws.cell(row=2, column=col_idx2).fill = CELL_DIFF_FILL
        
        # Set column widths based on a sample of the data for performance
        for col_idx, column in enumerate(ws.columns, 1):
            max_length = max(len(str(cell.value or "")) for cell in column[:100])
            ws.column_dimensions[get_column_letter(col_idx)].width = (max_length + 2) * 1.2
        
        return len([s for s in match_status if s == "Matched"])
    
    def compare(self, output_file=None):
        try:
            df1 = pd.read_excel(self.file1_path, sheet_name=self.sheet1_name)
            df2 = pd.read_excel(self.file2_path, sheet_name=self.sheet2_name)
            
            output_wb = Workbook()
            output_wb.remove(output_wb.active)
            
            self.create_side_by_side_sheet(df1, df2, output_wb)
            
            if output_file:
                output_wb.save(output_file)
                return output_file
            return output_wb
            
        except Exception as e:
            raise Exception(f"Comparison error: {str(e)}")

# Example Usage
if __name__ == "__main__":
    # Create dummy files to demonstrate the numeric comparison fix
    data1 = {'ID': [1, 2, 3], 'Value': [100, 200.05, 300.0]}
    df1 = pd.DataFrame(data1)
    df1.to_excel("file_a.xlsx", index=False)
    
    # File B has an integer where A has a float, a tiny float difference, and a float where A has an int
    data2 = {'ID': [1, 2, 3], 'Value': [100.0, 200.050000001, 300]}
    df2 = pd.DataFrame(data2)
    df2.to_excel("file_b.xlsx", index=False)
    
    print("Created dummy files 'file_a.xlsx' and 'file_b.xlsx' to test numeric comparisons.")
    
    comparator = ExcelComparator(
        file1_path="file_a.xlsx",
        file2_path="file_b.xlsx",
    )
    
    try:
        # The new are_equal function will correctly see all 'Value' cells as matching.
        result_path = comparator.compare(output_file="comparison_results.xlsx")
        print(f"✅ Comparison complete. Results saved to: {result_path}")
    except Exception as e:
        print(f"❌ Error: {e}")
