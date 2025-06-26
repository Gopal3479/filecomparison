import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import shutil

class InputGUI:
    def __init__(self, master, config_filename="config.xlsx"):
        self.master = master
        self.config_filename = config_filename
        
        self.setup_ui()
        self.config_path = self.get_config_path()
        self.status_var.set(f"Ready | Config: {os.path.basename(self.config_path)}")
        
        # Set up input files directory
        self.input_files_dir = os.path.join(os.path.dirname(self.config_path), "input_files")
        os.makedirs(self.input_files_dir, exist_ok=True)

    def setup_ui(self):
        self.master.title("Commercial Trends Report Generator")
        self.master.geometry("500x400")
        
        # Modern theme setup
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.master.configure(bg='#f0f0f0')
        
        # Main container
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Commercial Trend Workbook", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            file_frame,
            text="Select Workbook",
            command=self.select_workbook,
            width=15
        ).pack(side=tk.LEFT)
        
        self.file_label = ttk.Label(file_frame, text="No file selected", wraplength=300)
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # Parameters section
        param_frame = ttk.LabelFrame(main_frame, text="Report Parameters", padding="10")
        param_frame.pack(fill=tk.X, pady=5)
        
        # Year input
        ttk.Label(param_frame, text="Year:").grid(row=0, column=0, sticky='w', pady=5)
        self.year_var = tk.StringVar()
        ttk.Entry(param_frame, textvariable=self.year_var, width=10).grid(row=0, column=1, sticky='w', pady=5)
        
        # Quarter input
        ttk.Label(param_frame, text="Quarter:").grid(row=1, column=0, sticky='w', pady=5)
        self.quarter_var = tk.StringVar(value="Q1")
        ttk.Combobox(
            param_frame,
            textvariable=self.quarter_var,
            values=["Q1", "Q2", "Q3", "Q4"],
            state="readonly",
            width=7
        ).grid(row=1, column=1, sticky='w', pady=5)
        
        # Division input
        ttk.Label(param_frame, text="Division:").grid(row=2, column=0, sticky='w', pady=5)
        self.division_var = tk.StringVar(value="North")
        ttk.Combobox(
            param_frame,
            textvariable=self.division_var,
            values=["North", "South", "East", "West", "Central"],
            state="readonly",
            width=10
        ).grid(row=2, column=1, sticky='w', pady=5)
        
        # Generate button
        ttk.Button(
            main_frame,
            text="Generate Report",
            command=self.generate_report,
            style='Accent.TButton'
        ).pack(pady=10)
        
        # Status bar
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(
            self.master,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Style configuration
        self.style.configure('Accent.TButton', font=('Arial', 10, 'bold'), foreground='white', background='#0078d7')

    def select_workbook(self):
        """Select and store the commercial trend workbook"""
        filetypes = (
            ('Excel files', '*.xlsx;*.xls'),
            ('All files', '*.*')
        )
        
        file_path = filedialog.askopenfilename(
            title='Select Commercial Trend Workbook',
            initialdir='~',
            filetypes=filetypes
        )
        
        if file_path:
            filename = os.path.basename(file_path)
            dest_path = os.path.join(self.input_files_dir, filename)
            
            if not os.path.exists(dest_path):
                try:
                    shutil.copy2(file_path, dest_path)
                    self.file_label.config(text=f"Using: {filename} (saved to input_files)")
                    self.status_var.set(f"File saved to input_files: {filename}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save file: {str(e)}")
                    self.file_label.config(text=f"Using: {filename} (not saved)")
                    self.status_var.set(f"Error saving file: {str(e)}")
            else:
                self.file_label.config(text=f"Using existing: {filename}")
                self.status_var.set(f"Using existing file: {filename}")
            
            self.selected_file = dest_path

    def generate_report(self):
        """Generate the commercial trends report"""
        # Validate inputs
        if not hasattr(self, 'selected_file'):
            messagebox.showwarning("Input Required", "Please select a commercial trend workbook first")
            return
        
        year = self.year_var.get().strip()
        if not year.isdigit() or len(year) != 4:
            messagebox.showerror("Input Error", "Please enter a valid 4-digit year")
            return
        
        quarter = self.quarter_var.get()
        division = self.division_var.get()
        
        self.status_var.set("Generating report...")
        self.master.update_idletasks()
        
        try:
            # Your report generation logic here
            # Example:
            output_filename = f"Commercial_Trends_{year}_{quarter}_{division}.xlsx"
            output_path = os.path.join(os.path.dirname(self.input_files_dir), "output", output_filename)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Simulate processing
            import time
            time.sleep(2)  # Replace with actual processing
            
            messagebox.showinfo(
                "Success", 
                f"Report generated successfully!\nSaved to: {output_path}"
            )
            self.status_var.set(f"Report generated: {output_filename}")
            
        except Exception as e:
            messagebox.showerror(
                "Processing Error", 
                f"An error occurred:\n{str(e)}"
            )
            self.status_var.set(f"Error: {str(e)}")

    def get_config_path(self):
        """Resolve config file path whether running as script or executable"""
        base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
        data_dir = os.path.join(base_path, "data")
        config_path = os.path.join(data_dir, self.config_filename)
        
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Config file not found at: {config_path}")
        
        return config_path

if __name__ == "__main__":
    root = tk.Tk()
    app = InputGUI(root)
    root.mainloop()
