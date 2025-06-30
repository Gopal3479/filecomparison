import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
import hashlib

# Define highlighting styles
HEADER_DIFF_FILL = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")  # Gold
CELL_DIFF_FILL = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")    # Tomato
NUM_DIFF_FILL = PatternFill(start_color="87CEFA", end_color="87CEFA", fill_type="solid")     # Light Sky Blue
ROW_MATCH_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")    # Light Green
ROW_MISSING_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Light Gray
TOTAL_FILL = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")        # Lavender
HEADER_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")       # Light Gray

# Border style
THIN_BORDER = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

class ExcelComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Row Comparison Tool")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f2f5")
        
        # Initialize variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.sheet1_name = tk.StringVar()
        self.sheet2_name = tk.StringVar()
        self.status = tk.StringVar(value="Ready to compare files")
        self.df1 = None
        self.df2 = None
        self.side_by_side_df = None
        
        # Create UI
        self.create_widgets()
        
    def create_widgets(self):
        # Header frame
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=80)
        header_frame.pack(fill="x", side="top")
        
        header_label = tk.Label(
            header_frame, 
            text="Excel Row Comparison Tool", 
            font=("Arial", 20, "bold"), 
            fg="white", 
            bg="#2c3e50"
        )
        header_label.pack(pady=20)
        
        # Main content frame
        main_frame = tk.Frame(self.root, bg="#f0f2f5")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # File selection section
        file_frame = tk.LabelFrame(
            main_frame, 
            text="File Selection", 
            font=("Arial", 12, "bold"), 
            bg="#f0f2f5", 
            padx=10, 
            pady=10
        )
        file_frame.pack(fill="x", pady=(0, 15))
        
        # File 1
        file1_frame = tk.Frame(file_frame, bg="#f0f2f5")
        file1_frame.pack(fill="x", pady=5)
        
        tk.Label(
            file1_frame, 
            text="File 1:", 
            font=("Arial", 10), 
            bg="#f0f2f5", 
            width=10, 
            anchor="w"
        ).pack(side="left")
        
        tk.Entry(
            file1_frame, 
            textvariable=self.file1_path, 
            width=50, 
            state="readonly",
            font=("Arial", 10)
        ).pack(side="left", padx=5, fill="x", expand=True)
        
        tk.Button(
            file1_frame, 
            text="Browse", 
            command=lambda: self.select_file(1), 
            bg="#3498db", 
            fg="white",
            font=("Arial", 10, "bold")
        ).pack(side="left")
        
        # File 2
        file2_frame = tk.Frame(file_frame, bg="#f0f2f5")
        file2_frame.pack(fill="x", pady=5)
        
        tk.Label(
            file2_frame, 
            text="File 2:", 
            font=("Arial", 10), 
            bg="#f0f2f5", 
            width=10, 
            anchor="w"
        ).pack(side="left")
        
        tk.Entry(
            file2_frame, 
            textvariable=self.file2_path, 
            width=50, 
            state="readonly",
            font=("Arial", 10)
        ).pack(side="left", padx=5, fill="x", expand=True)
        
        tk.Button(
            file2_frame, 
            text="Browse", 
            command=lambda: self.select_file(2), 
            bg="#3498db", 
            fg="white",
            font=("Arial", 10, "bold")
        ).pack(side="left")
        
        # Sheet selection section
        sheet_frame = tk.LabelFrame(
            main_frame, 
            text="Sheet Selection", 
            font=("Arial", 12, "bold"), 
            bg="#f0f2f5", 
            padx=10, 
            pady=10
        )
        sheet_frame.pack(fill="x", pady=(0, 15))
        
        # Sheet 1
        sheet1_frame = tk.Frame(sheet_frame, bg="#f0f2f5")
        sheet1_frame.pack(fill="x", pady=5)
        
        tk.Label(
            sheet1_frame, 
            text="File 1 Sheet:", 
            font=("Arial", 10), 
            bg="#f0f2f5", 
            width=15, 
            anchor="w"
        ).pack(side="left")
        
        tk.Entry(
            sheet1_frame, 
            textvariable=self.sheet1_name, 
            width=30, 
            font=("Arial", 10)
        ).pack(side="left", padx=5, fill="x", expand=True)
        
        # Sheet 2
        sheet2_frame = tk.Frame(sheet_frame, bg="#f0f2f5")
        sheet2_frame.pack(fill="x", pady=5)
        
        tk.Label(
            sheet2_frame, 
            text="File 2 Sheet:", 
            font=("Arial", 10), 
            bg="#f0f2f5", 
            width=15, 
            anchor="w"
        ).pack(side="left")
        
        tk.Entry(
            sheet2_frame, 
            textvariable=self.sheet2_name, 
            width=30, 
            font=("Arial", 10)
        ).pack(side="left", padx=5, fill="x", expand=True)
        
        # Options section
        options_frame = tk.LabelFrame(
            main_frame, 
            text="Comparison Options", 
            font=("Arial", 12, "bold"), 
            bg="#f0f2f5", 
            padx=10, 
            pady=10
        )
        options_frame.pack(fill="x", pady=(0, 15))
        
        self.highlight_missing = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame, 
            text="Highlight missing columns", 
            variable=self.highlight_missing, 
            bg="#f0f2f5", 
            font=("Arial", 10)
        ).pack(anchor="w", pady=3)
        
        self.highlight_cell_diffs = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame, 
            text="Highlight cell differences", 
            variable=self.highlight_cell_diffs, 
            bg="#f0f2f5", 
            font=("Arial", 10)
        ).pack(anchor="w", pady=3)
        
        self.highlight_row_matches = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame, 
            text="Highlight row matches/mismatches", 
            variable=self.highlight_row_matches, 
            bg="#f0f2f5", 
            font=("Arial", 10)
        ).pack(anchor="w", pady=3)
        
        self.create_num_table = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame, 
            text="Create numerical differences table", 
            variable=self.create_num_table, 
            bg="#f0f2f5", 
            font=("Arial", 10)
        ).pack(anchor="w", pady=3)
        
        self.create_totals = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame, 
            text="Show totals for numerical columns", 
            variable=self.create_totals, 
            bg="#f0f2f5", 
            font=("Arial", 10)
        ).pack(anchor="w", pady=3)
        
        # Action buttons
        button_frame = tk.Frame(main_frame, bg="#f0f2f5")
        button_frame.pack(fill="x", pady=20)
        
        compare_btn = tk.Button(
            button_frame, 
            text="Compare Excel Files", 
            command=self.compare_files, 
            bg="#27ae60", 
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
            width=20
        )
        compare_btn.pack(pady=10)
        
        # Status bar
        status_frame = tk.Frame(self.root, bg="#e0e0e0", height=30)
        status_frame.pack(fill="x", side="bottom")
        
        tk.Label(
            status_frame, 
            textvariable=self.status, 
            font=("Arial", 10), 
            bg="#e0e0e0", 
            anchor="w"
        ).pack(side="left", padx=10)
        
        # Legend
        legend_frame = tk.Frame(main_frame, bg="#f0f2f5")
        legend_frame.pack(fill="x", pady=10)
        
        tk.Label(
            legend_frame, 
            text="Highlight Legend:", 
            font=("Arial", 10, "bold"), 
            bg="#f0f2f5"
        ).pack(anchor="w")
        
        legend_inner = tk.Frame(legend_frame, bg="#f0f2f5")
        legend_inner.pack(fill="x", pady=5)
        
        # Header diff legend
        header_legend = tk.Frame(legend_inner, bg="#f0f2f5")
        header_legend.pack(side="left", padx=10)
        tk.Label(
            header_legend, 
            text="    ", 
            bg="#FFD700", 
            width=3, 
            height=1
        ).pack(side="left")
        tk.Label(
            header_legend, 
            text="Column Differences", 
            font=("Arial", 9), 
            bg="#f0f2f5"
        ).pack(side="left", padx=5)
        
        # Cell diff legend
        cell_legend = tk.Frame(legend_inner, bg="#f0f2f5")
        cell_legend.pack(side="left", padx=10)
        tk.Label(
            cell_legend, 
            text="    ", 
            bg="#FF6347", 
            width=3, 
            height=1
        ).pack(side="left")
        tk.Label(
            cell_legend, 
            text="Value Differences", 
            font=("Arial", 9), 
            bg="#f0f2f5"
        ).pack(side="left", padx=5)
        
        # Row match legend
        row_match_legend = tk.Frame(legend_inner, bg="#f0f2f5")
        row_match_legend.pack(side="left", padx=10)
        tk.Label(
            row_match_legend, 
            text="    ", 
            bg="#90EE90", 
            width=3, 
            height=1
        ).pack(side="left")
        tk.Label(
            row_match_legend, 
            text="Matched Rows", 
            font=("Arial", 9), 
            bg="#f0f2f5"
        ).pack(side="left", padx=5)
        
        # Row missing legend
        row_missing_legend = tk.Frame(legend_inner, bg="#f0f2f5")
        row_missing_legend.pack(side="left", padx=10)
        tk.Label(
            row_missing_legend, 
            text="    ", 
            bg="#D3D3D3", 
            width=3, 
            height=1
        ).pack(side="left")
        tk.Label(
            row_missing_legend, 
            text="Unmatched Rows", 
            font=("Arial", 9), 
            bg="#f0f2f5"
        ).pack(side="left", padx=5)
        
        # Total legend
        total_legend = tk.Frame(legend_inner, bg="#f0f2f5")
        total_legend.pack(side="left", padx=10)
        tk.Label(
            total_legend, 
            text="    ", 
            bg="#E6E6FA", 
            width=3, 
            height=1
        ).pack(side="left")
        tk.Label(
            total_legend, 
            text="Total Row", 
            font=("Arial", 9), 
            bg="#f0f2f5"
        ).pack(side="left", padx=5)
    
    def select_file(self, file_num):
        file_path = filedialog.askopenfilename(
            title=f"Select Excel File {file_num}",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            if file_num == 1:
                self.file1_path.set(file_path)
                try:
                    wb = load_workbook(file_path, read_only=True)
                    self.sheet1_name.set(wb.sheetnames[0])
                    wb.close()
                except:
                    self.sheet1_name.set("Sheet1")
            else:
                self.file2_path.set(file_path)
                try:
                    wb = load_workbook(file_path, read_only=True)
                    self.sheet2_name.set(wb.sheetnames[0])
                    wb.close()
                except:
                    self.sheet2_name.set("Sheet1")
    
    def are_equal(self, a, b):
        """Check if two values are equal, handling NaN cases"""
        if pd.isna(a) and pd.isna(b):
            return True
        if pd.isna(a) or pd.isna(b):
            return False
        return a == b
    
    def is_date(self, value):
        """Check if a value is a date"""
        return isinstance(value, (datetime, pd.Timestamp))
    
    def generate_row_hash(self, row, date_columns):
        """Generate a hash for a row, excluding date columns"""
        values = []
        for col, val in row.items():
            if col not in date_columns and not self.is_date(val) and not pd.isna(val):
                values.append(str(val))
        return hashlib.md5("|".join(values).encode()).hexdigest()
    
    def create_side_by_side_sheet(self, df1, df2, output_wb):
        """Create side-by-side comparison sheet with row matching"""
        # Create sheet
        ws = output_wb.create_sheet("Side by Side Comparison")
        
        # Get common columns
        common_cols = list(set(df1.columns) & set(df2.columns))
        
        # Identify date columns to exclude from matching
        date_cols1 = [col for col in df1.columns if pd.api.types.is_datetime64_any_dtype(df1[col])]
        date_cols2 = [col for col in df2.columns if pd.api.types.is_datetime64_any_dtype(df2[col])]
        date_columns = set(date_cols1 + date_cols2)
        
        # Create row hashes for matching (excluding date columns)
        df1['row_hash'] = df1.apply(lambda row: self.generate_row_hash(row, date_columns), axis=1)
        df2['row_hash'] = df2.apply(lambda row: self.generate_row_hash(row, date_columns), axis=1)
        
        # Create sets of hashes
        hash1_set = set(df1['row_hash'])
        hash2_set = set(df2['row_hash'])
        
        # Prepare data for side-by-side comparison
        side_by_side_data = []
        matched_indices = set()
        
        # Find matched rows (any row in df1 matches any row in df2)
        for idx1, row1 in df1.iterrows():
            if row1['row_hash'] in hash2_set:
                # Find the first matching row in df2
                match_idx = df2[df2['row_hash'] == row1['row_hash']].index[0]
                matched_indices.add(idx1)
                matched_indices.add(match_idx)
                
                row_data = []
                # Add File1 data
                for col in df1.columns:
                    if col != 'row_hash':
                        row_data.append(row1[col])
                # Add File2 data
                for col in df2.columns:
                    if col != 'row_hash':
                        row_data.append(df2.at[match_idx, col])
                row_data.append("Matched")
                side_by_side_data.append(row_data)
        
        # Add unmatched rows from df1
        for idx, row in df1.iterrows():
            if idx not in matched_indices:
                row_data = []
                # Add File1 data
                for col in df1.columns:
                    if col != 'row_hash':
                        row_data.append(row[col])
                # Add blank for File2
                for _ in df2.columns:
                    if _ != 'row_hash':
                        row_data.append(None)
                row_data.append("Not Matched (File1)")
                side_by_side_data.append(row_data)
        
        # Add unmatched rows from df2
        for idx, row in df2.iterrows():
            if idx not in matched_indices:
                row_data = []
                # Add blank for File1
                for _ in df1.columns:
                    if _ != 'row_hash':
                        row_data.append(None)
                # Add File2 data
                for col in df2.columns:
                    if col != 'row_hash':
                        row_data.append(row[col])
                row_data.append("Not Matched (File2)")
                side_by_side_data.append(row_data)
        
        # Create DataFrame for side-by-side view
        columns = [f"File1_{col}" for col in df1.columns if col != 'row_hash'] + \
                  [f"File2_{col}" for col in df2.columns if col != 'row_hash'] + \
                  ["Match Status"]
        self.side_by_side_df = pd.DataFrame(side_by_side_data, columns=columns)
        
        # Write headers
        header_row = ["File1"] * len(df1.columns) + ["File2"] * len(df2.columns) + [""]
        ws.append(header_row)
        
        col_names = [col for col in df1.columns if col != 'row_hash'] + \
                    [col for col in df2.columns if col != 'row_hash'] + \
                    ["Match Status"]
        ws.append(col_names)
        
        # Apply header styling
        for row in ws.iter_rows(min_row=1, max_row=2, max_col=len(col_names)):
            for cell in row:
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)
                cell.border = THIN_BORDER
        
        # Merge header cells
        if len(df1.columns) > 0:
            ws.merge_cells(
                start_row=1, 
                start_column=1, 
                end_row=1, 
                end_column=len(df1.columns)
            )
            file1_header = ws.cell(row=1, column=1)
            file1_header.value = "File1"
            file1_header.alignment = Alignment(horizontal='center')
        
        if len(df2.columns) > 0:
            start_col = len(df1.columns) + 1
            end_col = len(df1.columns) + len(df2.columns)
            ws.merge_cells(
                start_row=1, 
                start_column=start_col, 
                end_row=1, 
                end_column=end_col
            )
            file2_header = ws.cell(row=1, column=start_col)
            file2_header.value = "File2"
            file2_header.alignment = Alignment(horizontal='center')
        
        # Add totals row at the top
        if self.create_totals.get():
            totals_row = []
            
            # File1 totals
            for col in df1.columns:
                if col != 'row_hash' and pd.api.types.is_numeric_dtype(df1[col]):
                    totals_row.append(df1[col].sum())
                else:
                    totals_row.append("")
            
            # File2 totals
            for col in df2.columns:
                if col != 'row_hash' and pd.api.types.is_numeric_dtype(df2[col]):
                    totals_row.append(df2[col].sum())
                else:
                    totals_row.append("")
            
            totals_row.append("Totals")
            ws.append(totals_row)
            
            # Apply styling to totals row
            totals_row_num = ws.max_row
            for col_idx in range(1, len(totals_row) + 1):
                cell = ws.cell(row=totals_row_num, column=col_idx)
                cell.fill = TOTAL_FILL
                cell.font = Font(bold=True)
                cell.border = THIN_BORDER
        
        # Write data
        for _, row in self.side_by_side_df.iterrows():
            ws.append(row.tolist())
        
        # Apply styling and formatting
        start_row = 4 if self.create_totals.get() else 3
        for row_idx, row in enumerate(ws.iter_rows(min_row=start_row, max_row=ws.max_row), start_row):
            match_status = ws.cell(row=row_idx, column=len(col_names)).value
            
            # Apply row matching highlighting
            if self.highlight_row_matches.get():
                if match_status == "Matched":
                    for cell in row:
                        cell.fill = ROW_MATCH_FILL
                elif "Not Matched" in match_status:
                    for cell in row:
                        cell.fill = ROW_MISSING_FILL
            
            # Apply cell difference highlighting
            if self.highlight_cell_diffs.get() and match_status == "Matched":
                # Only compare common columns for matched rows
                for col_idx, col_name in enumerate(df1.columns, 1):
                    if col_name in common_cols:
                        file1_val = ws.cell(row=row_idx, column=col_idx).value
                        file2_val = ws.cell(row=row_idx, column=col_idx + len(df1.columns)).value
                        
                        if not self.are_equal(file1_val, file2_val):
                            ws.cell(row=row_idx, column=col_idx).fill = CELL_DIFF_FILL
                            ws.cell(row=row_idx, column=col_idx + len(df1.columns)).fill = CELL_DIFF_FILL
            
            # Apply borders
            for cell in row:
                cell.border = THIN_BORDER
        
        # Auto-size columns
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            if adjusted_width > 0:
                ws.column_dimensions[column].width = adjusted_width
        
        # Freeze panes
        ws.freeze_panes = ws.cell(row=3, column=1)
        
        matched_count = len([x for x in side_by_side_data if x[-1] == "Matched"])
        unmatched1_count = len([x for x in side_by_side_data if "File1" in x[-1]])
        unmatched2_count = len([x for x in side_by_side_data if "File2" in x[-1]])
        
        return matched_count, unmatched1_count, unmatched2_count
    
    def compare_files(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        sheet1 = self.sheet1_name.get()
        sheet2 = self.sheet2_name.get()
        
        if not file1 or not file2:
            messagebox.showerror("Error", "Please select both Excel files")
            return
        
        self.status.set("Reading files...")
        self.root.update()
        
        try:
            # Read data
            df1 = pd.read_excel(file1, sheet_name=sheet1)
            df2 = pd.read_excel(file2, sheet_name=sheet2)
            
            # Create comparison workbook
            output_wb = Workbook()
            output_wb.remove(output_wb.active)
            
            # 1. Compare headers
            self.compare_headers(df1, df2, output_wb)
            
            # 2. Create side-by-side comparison sheet
            matched_count, unmatched1_count, unmatched2_count = self.create_side_by_side_sheet(df1, df2, output_wb)
            
            # 3. Row matching analysis
            self.analyze_row_matches(df1, df2, output_wb, matched_count, unmatched1_count, unmatched2_count)
            
            # 4. Numerical differences
            if self.create_num_table.get():
                self.compare_numeric_values(df1, df2,
