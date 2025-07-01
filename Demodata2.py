import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
import os
from datetime import datetime

# Define highlighting styles
HEADER_DIFF_FILL = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")  # Gold
CELL_DIFF_FILL = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")    # Tomato
NUM_DIFF_FILL = PatternFill(start_color="87CEFA", end_color="87CEFA", fill_type="solid")     # Light Sky Blue
ROW_MATCH_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")    # Light Green
ROW_MISSING_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Light Gray
HEADER_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")       # Light Gray
TOTAL_FILL = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")        # Lavender

THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

class ExcelComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Comparison Tool")
        self.root.geometry("950x750")
        self.root.configure(bg="#f0f2f5")
        
        # Initialize variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.sheet1_name = tk.StringVar()
        self.sheet2_name = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.df1 = None
        self.df2 = None
        
        # Create UI
        self.create_widgets()
        
    def create_widgets(self):
        # (UI code stays the same as before — omitted for brevity, we already reviewed it)
        # If you want I will re-paste that too, but assume no changes in UI
        pass
        
    def compare_files(self):
        try:
            # Load dataframes
            self.df1 = pd.read_excel(self.file1_path.get(), sheet_name=self.sheet1_name.get())
            self.df2 = pd.read_excel(self.file2_path.get(), sheet_name=self.sheet2_name.get())
            self.status.set("Files loaded, starting comparison...")
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Comparison"
            
            # Write header
            for col_idx, col_name in enumerate(self.df1.columns, start=1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)
                cell.border = THIN_BORDER
            
            # Compare row by row
            max_rows = max(len(self.df1), len(self.df2))
            for row in range(max_rows):
                for col_idx, col_name in enumerate(self.df1.columns, start=1):
                    val1 = self.df1.iloc[row, col_idx-1] if row < len(self.df1) else None
                    val2 = self.df2.iloc[row, col_idx-1] if row < len(self.df2) else None
                    
                    ws.cell(row=row+2, column=col_idx, value=val1)
                    
                    # highlight if different
                    if row < len(self.df2) and col_name in self.df2.columns:
                        if pd.isna(val1) and pd.isna(val2):
                            continue
                        elif val1 != val2:
                            ws.cell(row=row+2, column=col_idx).fill = CELL_DIFF_FILL
                    else:
                        ws.cell(row=row+2, column=col_idx).fill = ROW_MISSING_FILL
            
            # Save
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Comparison Report"
            )
            if save_path:
                wb.save(save_path)
                self.status.set(f"Comparison completed, saved at {save_path}")
                messagebox.showinfo("Done", f"Comparison report saved:\n{save_path}")
            else:
                self.status.set("Comparison cancelled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare files:\n{e}")
            self.status.set("Error during comparison")
    
    # === additional placeholder functions ===
    def compare_headers(self):
        # If needed later
        pass
    
    def analyze_row_matches(self):
        # If needed later
        pass
    
    def compare_numeric_values(self):
        # If needed later
        pass
    
    def create_side_by_side_sheet(self):
        # If needed later
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparator(root)
    root.mainloop()
