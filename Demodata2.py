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

        # --- Title ---
        title_label = tk.Label(main_frame, text="Advanced Excel Comparison Tool", font=("Helvetica", 18, "bold"), bg="#f0f0f0")
        title_label.pack(pady=(0, 20))

        # --- File Selection ---
        file1_frame = tk.LabelFrame(main_frame, text="Select File 1", padx=10, pady=10, bg="#f0f0f0", font=("Helvetica", 11))
        file1_frame.pack(fill=tk.X, pady=10)
        
        file1_entry = tk.Entry(file1_frame, textvariable=self.file1_path, state="readonly", width=50)
        file1_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=4)
        
        browse1_btn = tk.Button(file1_frame, text="Browse...", command=lambda: self.browse_file(self.file1_path))
        browse1_btn.pack(side=tk.LEFT, padx=(10, 0))

        file2_frame = tk.LabelFrame(main_frame, text="Select File 2", padx=10, pady=10, bg="#f0f0f0", font=("Helvetica", 11))
        file2_frame.pack(fill=tk.X, pady=10)

        file2_entry = tk.Entry(file2_frame, textvariable=self.file2_path, state="readonly", width=50)
        file2_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=4)
        
        browse2_btn = tk.Button(file2_frame, text="Browse...", command=lambda: self.browse_file(self.file2_path))
        browse2_btn.pack(side=tk.LEFT, padx=(10, 0))

        # --- Compare Button ---
        self.compare_btn = tk.Button(main_frame, text="Compare Excel Files", command=self.start_comparison_thread, font=("Helvetica", 12, "bold"), bg="#4a90e2", fg="white", relief=tk.RAISED)
        self.compare_btn.pack(pady=20, ipady=8, fill=tk.X)

        # --- Progress Bar and Status ---
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10, fill=tk.X)
        self.status_label = tk.Label(main_frame, text="Select two files to begin.", bg="#f0f0f0", anchor="w")
        self.status_label.pack(fill=tk.X)

    def browse_file(self, path_var):
        """Opens a file dialog to select an Excel file."""
        filepath = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=(("Excel Files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepath:
            path_var.set(filepath)
            self.status_label.config(text="File selected. Ready to compare.")

    def start_comparison_thread(self):
        """Starts the comparison process in a separate thread to keep the UI responsive."""
        if not self.file1_path.get() or not self.file2_path.get():
            messagebox.showerror("Error", "Please select both files before comparing.")
            return

        self.compare_btn.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        
        # Run the comparison in a new thread
        comparison_thread = threading.Thread(target=self.run_comparison)
        comparison_thread.daemon = True
        comparison_thread.start()

    def update_status(self, message, progress):
        """Updates the status label and progress bar from the main thread."""
        self.status_label.config(text=message)
        self.progress_bar['value'] = progress
        self.root.update_idletasks()
        
    def run_comparison(self):
        """The core logic for the Excel file comparison."""
        try:
            self.update_status("Initializing...", 5)
            
            # --- 1. Read Files ---
            self.update_status("Reading File 1...", 10)
            df1, headers1 = self.read_excel_file(self.file1_path.get())
            self.update_status("Reading File 2...", 20)
            df2, headers2 = self.read_excel_file(self.file2_path.get())

            if df1.empty or df2.empty:
                raise ValueError("One or both Excel files appear to be empty or could not be read.")

            # --- 2. Header Comparison ---
            self.update_status("Comparing headers...", 30)
            header_comp_df = self.compare_headers(headers1, headers2)

            # --- 3. Smart Column & Key Generation ---
            self.update_status("Identifying key columns...", 40)
            key_columns = self.find_key_columns(df1, headers1)
            if not key_columns:
                raise ValueError("No suitable non-date string columns found for matching.")
            
            self.update_status("Creating matching keys...", 50)
            df1['__key'] = self.create_concatenated_key(df1, key_columns)
            df2['__key'] = self.create_concatenated_key(df2, key_columns)

            # --- 4. Numerical Column Identification and Totals ---
            self.update_status("Calculating numerical totals...", 60)
            numeric_cols1 = self.find_numeric_columns(df1)
            numeric_cols2 = self.find_numeric_columns(df2)
            totals1 = df1[numeric_cols1].sum()
            totals2 = df2[numeric_cols2].sum()
            
            # --- 5. Side-by-Side and Row Matching ---
            self.update_status("Performing side-by-side comparison...", 70)
            side_by_side_df, row_matching_df = self.create_comparison_data(df1, df2, headers1, headers2)

            # --- 6. Generate Excel Report ---
            self.update_status("Generating results file...", 80)
            self.generate_excel_report(header_comp_df, side_by_side_df, row_matching_df, totals1, totals2, headers1, headers2)
            
            self.update_status("Comparison complete! Report saved.", 100)
            messagebox.showinfo("Success", "Comparison finished successfully. The report 'Excel_Comparison_Report.xlsx' has been saved.")

        except Exception as e:
            self.update_status("Error occurred.", 0)
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.compare_btn.config(state=tk.NORMAL)

    def read_excel_file(self, path):
        """Reads an excel file, finds the header, and filters out total rows."""
        df = pd.read_excel(path, header=None)
        
        # Find header row (first row with non-empty values)
        header_row_index = 0
        for i, row in df.iterrows():
            if not row.isnull().all():
                header_row_index = i
                break
        
        headers = list(df.iloc[header_row_index])
        df.columns = headers
        df = df.iloc[header_row_index + 1:].reset_index(drop=True)
        
        # Find 'Total' row and cut off data from there
        total_row_index = -1
        if not df.empty and df.columns[0] is not None:
             # Find 'Total' row and cut off data from there
            total_rows = df[df.iloc[:, 0].astype(str).str.strip().str.lower().str.startswith('total', na=False)]
            if not total_rows.empty:
                total_row_index = total_rows.index[0]
                df = df.iloc[:total_row_index]

        return df.fillna(''), headers

    def compare_headers(self, headers1, headers2):
        """Compares the headers of the two files."""
        h1_set = set(headers1)
        h2_set = set(headers2)
        common = sorted(list(h1_set.intersection(h2_set)))
        only_in_1 = sorted(list(h1_set - h2_set))
        only_in_2 = sorted(list(h2_set - h1_set))

        data = []
        for h in common: data.append(("Common", h))
        for h in only_in_1: data.append(("Only in File 1", h))
        for h in only_in_2: data.append(("Only in File 2", h))
        
        return pd.DataFrame(data, columns=["Status", "Column Name"])

    def find_key_columns(self, df, headers):
        """Finds non-date string columns to use for key generation."""
        key_cols = []
        # Attempt to convert object columns to datetime to identify date-like columns
        for col in headers:
            if df[col].dtype == 'object':
                try:
                    # If a significant portion can be converted to datetime, treat it as a date column.
                    if pd.to_datetime(df[col], errors='coerce').notna().sum() / len(df) < 0.8:
                        key_cols.append(col)
                except Exception:
                     key_cols.append(col)
        return key_cols

    def create_concatenated_key(self, df, key_columns):
        """Creates a single key by concatenating string columns."""
        return df[key_columns].astype(str).apply(lambda x: '|'.join(x.str.strip()), axis=1)

    def find_numeric_columns(self, df):
        """Identifies numeric columns in the DataFrame."""
        return df.select_dtypes(include='number').columns.tolist()

    def create_comparison_data(self, df1, df2, headers1, headers2):
        """Merges dataframes and creates side-by-side and row analysis data."""
        # Add file suffixes to original columns
        df1_renamed = df1.rename(columns={c: f"File1: {c}" for c in headers1})
        df2_renamed = df2.rename(columns={c: f"File2: {c}" for c in headers2})

        # Perform an outer merge on the key
        merged_df = pd.merge(df1_renamed, df2_renamed, left_on='File1: __key', right_on='File2: __key', how='outer')
        merged_df.fillna('', inplace=True)
        
        # --- Create Side-by-Side Data ---
        def get_status(row):
            key1 = row['File1: __key']
            key2 = row['File2: __key']
            if key1 and key2: return "Matched"
            if key1: return "Not Matched (File1)"
            return "Not Matched (File2)"

        merged_df['Match Status'] = merged_df.apply(get_status, axis=1)
        
        side_by_side_cols = ['Match Status'] + [f"File1: {h}" for h in headers1] + [f"File2: {h}" for h in headers2]
        side_by_side_df = merged_df[side_by_side_cols]

        # --- Create Row Matching Analysis Data ---
        row_analysis_data = []
        for _, row in merged_df.iterrows():
            status = row['Match Status']
            key = row['File1: __key'] if status != "Not Matched (File2)" else row['File2: __key']
            if status == "Matched":
                row_analysis_data.append([status, key, "Both", "Row exists in both files."])
            elif status == "Not Matched (File1)":
                 row_analysis_data.append([status, key, "File 1", "Row only exists in File 1."])
            else:
                 row_analysis_data.append([status, key, "File 2", "Row only exists in File 2."])
        
        row_matching_df = pd.DataFrame(row_analysis_data, columns=["Match Status", "Matching Key", "File", "Details"])
        
        return side_by_side_df, row_matching_df

    def generate_excel_report(self, header_df, sbs_df, row_match_df, totals1, totals2, headers1, headers2):
        """Creates and styles the final Excel report."""
        output_path = "Excel_Comparison_Report.xlsx"
        
        # --- Define Styles ---
        lavender_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        gray_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        file1_header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        file2_header_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
        bold_font = Font(bold=True)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # --- Sheet 1: Side by Side Comparison ---
            # Create totals row as a dataframe
            totals_data = {"Match Status": "NUMERICAL TOTALS"}
            for h in headers1: totals_data[f"File1: {h}"] = totals1.get(h, "")
            for h in headers2: totals_data[f"File2: {h}"] = totals2.get(h, "")
            totals_df = pd.DataFrame([totals_data], columns=sbs_df.columns)
            
            final_sbs_df = pd.concat([totals_df, sbs_df], ignore_index=True)
            final_sbs_df.to_excel(writer, sheet_name="Side by Side Comparison", index=False)
            ws_sbs = writer.sheets["Side by Side Comparison"]
            
            # Apply styling to SBS sheet
            # Style Totals row (row 2, as 1 is header)
            for cell in ws_sbs[2]:
                cell.fill = lavender_fill
                cell.font = bold_font

            # Style Headers (row 1)
            for i, col_name in enumerate(final_sbs_df.columns, 1):
                cell = ws_sbs.cell(row=1, column=i)
                if str(col_name).startswith("File1:"): cell.fill = file1_header_fill
                elif str(col_name).startswith("File2:"): cell.fill = file2_header_fill
                cell.font = bold_font

            # Style matched/unmatched rows
            for r_idx, row in final_sbs_df.iloc[1:].iterrows(): # Skip totals row in df
                status = row['Match Status']
                fill = green_fill if status == "Matched" else gray_fill
                for cell in ws_sbs[r_idx + 2]: # +2 to account for header and 0-indexing
                    cell.fill = fill
            
            # Set column widths
            for i, col in enumerate(final_sbs_df.columns):
                ws_sbs.column_dimensions[chr(65 + i)].width = 25

            # --- Sheet 2: Header Comparison ---
            header_df.to_excel(writer, sheet_name="Header Comparison", index=False)
            ws_header = writer.sheets["Header Comparison"]
            ws_header.column_dimensions['A'].width = 20
            ws_header.column_dimensions['B'].width = 40

            # --- Sheet 3: Row Matching Analysis ---
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
