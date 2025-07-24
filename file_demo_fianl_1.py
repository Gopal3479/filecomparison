import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

# Define highlighting styles
CELL_DIFF_FILL = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")    # Tomato
ROW_MATCH_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")    # Light Green
ROW_MISSING_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Light Gray
HEADER_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")       # Light Gray
TOTAL_ROW_FILL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")    # Light Blue

# Border style
THIN_BORDER = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

class ExcelComparator:
    def __init__(self, file1_path, file2_path, sheet1_name=None, sheet2_name=None):
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.sheet1_name = sheet1_name
        self.sheet2_name = sheet2_name
        
        if not self.sheet1_name:
            wb = load_workbook(file1_path, read_only=True)
            self.sheet1_name = wb.sheetnames[0]
            wb.close()
            
        if not self.sheet2_name:
            wb = load_workbook(file2_path, read_only=True)
            self.sheet2_name = wb.sheetnames[0]
            wb.close()
    
    def are_equal(self, a, b):
        """Check if two values are equal with type-specific comparisons"""
        if pd.isna(a) and pd.isna(b):
            return True
        if pd.isna(a) or pd.isna(b):
            return False
        if isinstance(a, (datetime, pd.Timestamp)) and isinstance(b, (datetime, pd.Timestamp)):
            return a == b
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            return abs(a - b) < 0.01  # Tolerance for floating point differences
        return str(a).strip() == str(b).strip()
    
    def is_numeric(self, series):
        """Check if a series is numeric"""
        return pd.api.types.is_numeric_dtype(series)
    
    def create_side_by_side_sheet(self, df1, df2, output_wb):
        """Optimized side-by-side comparison with total row"""
        ws = output_wb.create_sheet("Side by Side Comparison")
        common_cols = list(set(df1.columns) & set(df2.columns))
        
        # Identify numeric columns
        num_cols1 = [col for col in df1.columns if self.is_numeric(df1[col])]
        num_cols2 = [col for col in df2.columns if self.is_numeric(df2[col])]
        common_num_cols = list(set(num_cols1) & set(num_cols2))
        
        # Create header row
        header_row = list(df1.columns) + [" | "] + list(df2.columns) + ["Match Status"]
        ws.append(header_row)
        
        # Create total row
        total_row = []
        
        # File1 totals
        for col in df1.columns:
            if col in num_cols1:
                total_row.append(df1[col].sum())
            else:
                total_row.append("")
        
        # Separator
        total_row.append(" | ")
        
        # File2 totals
        for col in df2.columns:
            if col in num_cols2:
                total_row.append(df2[col].sum())
            else:
                total_row.append("")
        
        # Total row match status
        total_match = True
        for col in common_num_cols:
            val1 = df1[col].sum()
            val2 = df2[col].sum()
            if not self.are_equal(val1, val2):
                total_match = False
                break
        total_row.append("Matched" if total_match else "Not Matched")
        ws.append(total_row)
        
        # Apply header formatting
        for col_idx in range(1, len(header_row) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = HEADER_FILL
            cell.font = Font(bold=True)
            cell.border = THIN_BORDER
        
        # Apply total row formatting
        for col_idx in range(1, len(total_row) + 1):
            cell = ws.cell(row=2, column=col_idx)
            cell.fill = TOTAL_ROW_FILL
            cell.font = Font(bold=True)
            cell.border = THIN_BORDER
        
        # Precompute column indices
        separator_col = len(df1.columns) + 1
        match_status_col = len(header_row)
        file1_end_col = len(df1.columns)
        file2_start_col = file1_end_col + 2
        
        # Precompute comparison results
        match_status = []
        diff_positions = []
        
        for i in range(max(len(df1), len(df2))):
            row_match = True
            diffs_in_row = []
            
            # Only compare if both rows exist
            if i < min(len(df1), len(df2)):
                for col in common_cols:
                    val1 = df1.at[i, col]
                    val2 = df2.at[i, col]
                    if not self.are_equal(val1, val2):
                        row_match = False
                        col_idx1 = df1.columns.get_loc(col) + 1
                        col_idx2 = df2.columns.get_loc(col) + file2_start_col
                        diffs_in_row.append((col_idx1, col_idx2))
            
            # Handle missing rows
            if i >= len(df1) or i >= len(df2):
                row_match = False
            
            match_status.append("Matched" if row_match else "Not Matched")
            diff_positions.append(diffs_in_row)
        
        # Write data rows
        for i in range(max(len(df1), len(df2))):
            row_data = []
            
            # File1 data
            if i < len(df1):
                row_data.extend(df1.iloc[i].values)
            else:
                row_data.extend([""] * len(df1.columns))
            
            # Separator
            row_data.append(" | ")
            
            # File2 data
            if i < len(df2):
                row_data.extend(df2.iloc[i].values)
            else:
                row_data.extend([""] * len(df2.columns))
            
            # Match status
            row_data.append(match_status[i])
            
            ws.append(row_data)
            
            # Apply border to all cells
            for col_idx in range(1, len(header_row) + 1):
                ws.cell(row=i+3, column=col_idx).border = THIN_BORDER
        
        # Apply highlighting
        for i in range(len(match_status)):
            row_idx = i + 3  # Start after header and total row
            
            # Highlight separator column
            ws.cell(row=row_idx, column=separator_col).fill = PatternFill(
                start_color="000000", end_color="000000", fill_type="solid"
            )
            
            # Highlight entire row based on match status
            if match_status[i] == "Matched":
                fill = ROW_MATCH_FILL
            else:
                fill = ROW_MISSING_FILL
                
            # Apply to File1 section
            for col_idx in range(1, file1_end_col + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill
            
            # Apply to File2 section
            for col_idx in range(file2_start_col, file2_start_col + len(df2.columns)):
                ws.cell(row=row_idx, column=col_idx).fill = fill
            
            # Apply to status column
            ws.cell(row=row_idx, column=match_status_col).fill = fill
            
            # Highlight specific cell differences
            for col_idx1, col_idx2 in diff_positions[i]:
                ws.cell(row=row_idx, column=col_idx1).fill = CELL_DIFF_FILL
                ws.cell(row=row_idx, column=col_idx2).fill = CELL_DIFF_FILL
        
        # Highlight total row differences
        if not total_match:
            for col in common_num_cols:
                col_idx1 = df1.columns.get_loc(col) + 1
                col_idx2 = df2.columns.get_loc(col) + file2_start_col
                ws.cell(row=2, column=col_idx1).fill = CELL_DIFF_FILL
                ws.cell(row=2, column=col_idx2).fill = CELL_DIFF_FILL
        
        # Optimized column width calculation
        for col_idx in range(1, len(header_row) + 1):
            max_length = 0
            col_letter = get_column_letter(col_idx)
            
            # Check header
            max_length = max(max_length, len(str(header_row[col_idx-1])))
            
            # Check total row
            if col_idx <= len(total_row):
                max_length = max(max_length, len(str(total_row[col_idx-1])))
            
            # Check sample of data rows
            sample_rows = min(100, len(df1) + len(df2))
            for row_idx in range(3, min(3 + sample_rows, ws.max_row)):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_length = max(max_length, len(str(cell_value)))
            
            # Set width with buffer
            ws.column_dimensions[col_letter].width = (max_length + 2) * 1.2
        
        return len([s for s in match_status if s == "Matched"])
    
    def compare(self, output_file=None):
        try:
            # Read data
            df1 = pd.read_excel(self.file1_path, sheet_name=self.sheet1_name)
            df2 = pd.read_excel(self.file2_path, sheet_name=self.sheet2_name)
            
            # Create comparison workbook
            output_wb = Workbook()
            output_wb.remove(output_wb.active)
            
            # Create side-by-side comparison
            matched_count = self.create_side_by_side_sheet(df1, df2, output_wb)
            
            if output_file:
                output_wb.save(output_file)
                return output_file
            return output_wb
                
        except Exception as e:
            raise Exception(f"Comparison error: {str(e)}")

# Example usage
if __name__ == "__main__":
    comparator = ExcelComparator(
        file1_path="file1.xlsx",
        file2_path="file2.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet1"
    )
    
    try:
        result_path = comparator.compare(output_file="comparison_results.xlsx")
        print(f"Comparison saved to: {result_path}")
    except Exception as e:
        print(f"Error: {e}")
