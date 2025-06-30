import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import threading
import os

class ExcelComparatorApp:
    """
    A desktop application for comparing two Excel files with advanced features.
    Handles dynamic column ordering and differences between files.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Comparison Tool")
        self.root.geometry("600x450")
        self.root.configure(bg="#f0f0f0")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        """Creates and places all the UI widgets in the main window."""
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)

        title_label = tk.Label(main_frame, text="Advanced Excel Comparison Tool", font=("Helvetica", 18, "bold"), bg="#f0f0f0")
        title_label.pack(pady=(0, 20))

        file1_frame = tk.LabelFrame(main_frame, text="Select File 1", padx=10, pady=10, bg="#f0f0f0", font=("Helvetica", 11))
        file1_frame.pack(fill=tk.X, pady=10)
        
        tk.Entry(file1_frame, textvariable=self.file1_path, state="readonly", width=50).pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=4)
        tk.Button(file1_frame, text="Browse...", command=lambda: self.browse_file(self.file1_path)).pack(side=tk.LEFT, padx=(10, 0))

        file2_frame = tk.LabelFrame(main_frame, text="Select File 2", padx=10, pady=10, bg="#f0f0f0", font=("Helvetica", 11))
        file2_frame.pack(fill=tk.X, pady=10)

        tk.Entry(file2_frame, textvariable=self.file2_path, state="readonly", width=50).pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=4)
        tk.Button(file2_frame, text="Browse...", command=lambda: self.browse_file(self.file2_path)).pack(side=tk.LEFT, padx=(10, 0))

        self.compare_btn = tk.Button(main_frame, text="Compare Excel Files", command=self.start_comparison_thread, font=("Helvetica", 12, "bold"), bg="#4a90e2", fg="white", relief=tk.RAISED)
        self.compare_btn.pack(pady=20, ipady=8, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10, fill=tk.X)
        self.status_label = tk.Label(main_frame, text="Select two files to begin.", bg="#f0f0f0", anchor="w")
        self.status_label.pack(fill=tk.X)

    def browse_file(self, path_var):
        filepath = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=(("Excel Files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepath:
            path_var.set(filepath)
            self.status_label.config(text="File selected. Ready to compare.")

    def start_comparison_thread(self):
        if not self.file1_path.get() or not self.file2_path.get():
            messagebox.showerror("Error", "Please select both files before comparing.")
            return
        self.compare_btn.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        threading.Thread(target=self.run_comparison, daemon=True).start()

    def update_status(self, message, progress):
        self.root.after(0, self._update_status, message, progress)

    def _update_status(self, message, progress):
        self.status_label.config(text=message)
        self.progress_bar['value'] = progress
        self.root.update_idletasks()
        
    def run_comparison(self):
        try:
            self.update_status("Initializing...", 5)
            
            self.update_status("Reading File 1...", 10)
            df1, headers1 = self.read_excel_file(self.file1_path.get())
            self.update_status("Reading File 2...", 20)
            df2, headers2 = self.read_excel_file(self.file2_path.get())

            if df1.empty or df2.empty:
                raise ValueError("One or both Excel files appear to be empty or could not be read.")

            self.update_status("Comparing headers...", 30)
            header_comp_df = self.compare_headers(headers1, headers2)

            # --- Smart Column & Key Generation (based on common columns) ---
            self.update_status("Identifying common key columns...", 40)
            common_headers = list(set(headers1) & set(headers2))
            key_columns = self.find_key_columns(df1, df2, common_headers)
            if not key_columns:
                raise ValueError("No suitable common non-date string columns found for matching.")
            
            self.update_status("Creating matching keys...", 50)
            df1['__key'] = self.create_concatenated_key(df1, key_columns)
            df2['__key'] = self.create_concatenated_key(df2, key_columns)

            self.update_status("Calculating numerical totals...", 60)
            totals1 = df1.select_dtypes(include='number').sum()
            totals2 = df2.select_dtypes(include='number').sum()
            
            self.update_status("Performing side-by-side comparison...", 70)
            side_by_side_df, row_matching_df = self.create_comparison_data(df1, df2, headers1, headers2)

            self.update_status("Generating results file...", 85)
            self.generate_excel_report(header_comp_df, side_by_side_df, row_matching_df, totals1, totals2)
            
            self.update_status("Comparison complete! Report saved.", 100)
            messagebox.showinfo("Success", "Comparison finished successfully. Report 'Excel_Comparison_Report.xlsx' has been saved.")

        except Exception as e:
            self.update_status(f"Error: {e}", 0)
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.root.after(0, lambda: self.compare_btn.config(state=tk.NORMAL))

    def read_excel_file(self, path):
        df = pd.read_excel(path, header=None).dropna(how='all').reset_index(drop=True)
        if df.empty:
            return pd.DataFrame(), []

        header_row_index = 0
        for i, row in df.iterrows():
            if not row.isnull().all():
                header_row_index = i
                break
        
        headers = [str(h).strip() for h in df.iloc[header_row_index]]
        df.columns = headers
        df = df.iloc[header_row_index + 1:].reset_index(drop=True)
        
        if not df.empty and df.columns[0] is not None:
            first_col_str = df.iloc[:, 0].astype(str).str.strip().str.lower()
            total_rows = df[first_col_str.str.startswith('total', na=False)]
            if not total_rows.empty:
                total_row_index = total_rows.index[0]
                df = df.iloc[:total_row_index]

        # Convert numeric columns, coercing errors
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='ignore')

        return df.fillna(''), headers

    def compare_headers(self, headers1, headers2):
        h1_set = set(headers1)
        h2_set = set(headers2)
        common = sorted(list(h1_set.intersection(h2_set)))
        only_in_1 = sorted(list(h1_set - h2_set))
        only_in_2 = sorted(list(h2_set - h1_set))
        data = ([("Common", h) for h in common] +
                [("Only in File 1", h) for h in only_in_1] +
                [("Only in File 2", h) for h in only_in_2])
        return pd.DataFrame(data, columns=["Status", "Column Name"])

    def find_key_columns(self, df1, df2, common_headers):
        key_cols = []
        for col in common_headers:
            if df1[col].dtype == 'object' and df2[col].dtype == 'object':
                # A simple check to avoid date-like columns. A column is not a key if most values can be parsed as dates.
                try:
                    is_date_col = pd.to_datetime(df1[col], errors='coerce').notna().sum() / len(df1) > 0.7
                    if not is_date_col:
                        key_cols.append(col)
                except Exception:
                    key_cols.append(col)
        return key_cols

    def create_concatenated_key(self, df, key_columns):
        return df[key_columns].astype(str).apply(lambda x: '|'.join(x.str.strip().str.lower()), axis=1)

    def create_comparison_data(self, df1, df2, headers1, headers2):
        df1_renamed = df1.add_prefix('File1: ')
        df2_renamed = df2.add_prefix('File2: ')
        
        merged_df = pd.merge(
            df1_renamed, 
            df2_renamed, 
            left_on='File1: __key', 
            right_on='File2: __key', 
            how='outer'
        ).fillna('')
        
        def get_status(row):
            key1_present = isinstance(row.get('File1: __key'), str) and row.get('File1: __key') != ''
            key2_present = isinstance(row.get('File2: __key'), str) and row.get('File2: __key') != ''
            if key1_present and key2_present: return "Matched"
            if key1_present: return "Not Matched (File1)"
            return "Not Matched (File2)"

        merged_df['Match Status'] = merged_df.apply(get_status, axis=1)
        
        # --- Create Side-by-Side Data with all columns ---
        all_base_headers = sorted(list(set(headers1) | set(headers2)))
        sbs_cols = ['Match Status']
        for h in all_base_headers:
            sbs_cols.append(f'File1: {h}')
        for h in all_base_headers:
            sbs_cols.append(f'File2: {h}')
        
        # Reorder merged_df and add missing columns
        side_by_side_df = merged_df.reindex(columns=sbs_cols).fillna('')
        
        # --- Create Row Matching Analysis Data ---
        row_analysis_data = []
        for _, row in merged_df.iterrows():
            status = row['Match Status']
            key = row['File1: __key'] if status != "Not Matched (File2)" else row['File2: __key']
            details = "Row exists in both files." if status == "Matched" else f"Row only exists in {status.split('(')[1][:-1]}."
            file_loc = "Both" if status == "Matched" else status.split('(')[1][:-1]
            row_analysis_data.append([status, key, file_loc, details])
        
        row_matching_df = pd.DataFrame(row_analysis_data, columns=["Match Status", "Matching Key", "File", "Details"])
        
        return side_by_side_df, row_matching_df

    def generate_excel_report(self, header_df, sbs_df, row_match_df, totals1, totals2):
        output_path = "Excel_Comparison_Report.xlsx"
        
        lavender_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        gray_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        file1_header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        file2_header_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
        bold_font = Font(bold=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # --- Sheet 1: Side by Side Comparison ---
            totals_row = {"Match Status": "NUMERICAL TOTALS"}
            for col in sbs_df.columns:
                if col.startswith("File1: "):
                    base_col = col.replace("File1: ", "")
                    totals_row[col] = totals1.get(base_col, "")
                elif col.startswith("File2: "):
                    base_col = col.replace("File2: ", "")
                    totals_row[col] = totals2.get(base_col, "")
            
            totals_df = pd.DataFrame([totals_row])
            final_sbs_df = pd.concat([totals_df, sbs_df], ignore_index=True)
            final_sbs_df.to_excel(writer, sheet_name="Side by Side Comparison", index=False)
            ws_sbs = writer.sheets["Side by Side Comparison"]

            ws_sbs.cell(row=2, column=1).font = bold_font
            for cell in ws_sbs[2]: cell.fill = lavender_fill
            
            for i, col_name in enumerate(final_sbs_df.columns, 1):
                cell = ws_sbs.cell(row=1, column=i)
                if str(col_name).startswith("File1:"): cell.fill = file1_header_fill
                elif str(col_name).startswith("File2:"): cell.fill = file2_header_fill
                cell.font = bold_font
                ws_sbs.column_dimensions[chr(64 + i)].width = 25

            for r_idx, row in final_sbs_df.iloc[1:].iterrows():
                status = row['Match Status']
                fill = green_fill if status == "Matched" else gray_fill
                for cell in ws_sbs[r_idx + 2]: cell.fill = fill
            
            # --- Other Sheets ---
            header_df.to_excel(writer, sheet_name="Header Comparison", index=False)
            ws_header = writer.sheets["Header Comparison"]
            ws_header.column_dimensions['A'].width = 20
            ws_header.column_dimensions['B'].width = 40

            row_match_df.to_excel(writer, sheet_name="Row Matching Analysis", index=False)
            ws_row = writer.sheets["Row Matching Analysis"]
            ws_row.column_dimensions['A'].width = 20
            ws_row.column_dimensions['B'].width = 40
            ws_row.column_dimensions['C'].width = 15
            ws_row.column_dimensions['D'].width = 50

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparatorApp(root)
    root.mainloop()
