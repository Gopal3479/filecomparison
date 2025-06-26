import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

# Resource path function for PyInstaller compatibility
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

class InputGUI:
    def __init__(self, master):
        self.master = master
        master.title("Report Generator")
        master.geometry("400x300")
        master.resizable(False, False)
        
        # Apply modern theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure colors
        master.configure(bg='#f0f0f0')
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10, 'bold'))
        self.style.configure('Header.TLabel', font=('Arial', 14, 'bold'))
        
        self.create_widgets()
        
    def create_widgets(self):
        # Header
        header_frame = ttk.Frame(self.master)
        header_frame.pack(pady=20, fill='x')
        
        ttk.Label(
            header_frame,
            text="Report Generation Parameters",
            style='Header.TLabel'
        ).pack()
        
        # Input fields
        form_frame = ttk.Frame(self.master)
        form_frame.pack(pady=10, padx=20, fill='x')
        
        # Year input
        ttk.Label(form_frame, text="Year:").grid(row=0, column=0, sticky='w', pady=5)
        self.year_var = tk.StringVar()
        year_entry = ttk.Entry(form_frame, textvariable=self.year_var, width=10)
        year_entry.grid(row=0, column=1, sticky='w', pady=5)
        year_entry.focus()
        
        # Quarter input
        ttk.Label(form_frame, text="Quarter:").grid(row=1, column=0, sticky='w', pady=5)
        self.quarter_var = tk.StringVar()
        quarter_combo = ttk.Combobox(
            form_frame, 
            textvariable=self.quarter_var,
            values=["Q1", "Q2", "Q3", "Q4"],
            state="readonly",
            width=7
        )
        quarter_combo.grid(row=1, column=1, sticky='w', pady=5)
        quarter_combo.current(0)
        
        # Division input
        ttk.Label(form_frame, text="Division:").grid(row=2, column=0, sticky='w', pady=5)
        self.division_var = tk.StringVar()
        division_combo = ttk.Combobox(
            form_frame, 
            textvariable=self.division_var,
            values=["North", "South", "East", "West", "Central"],
            state="readonly",
            width=10
        )
        division_combo.grid(row=2, column=1, sticky='w', pady=5)
        division_combo.current(0)
        
        # Submit button
        button_frame = ttk.Frame(self.master)
        button_frame.pack(pady=20)
        
        ttk.Button(
            button_frame,
            text="Generate Report",
            command=self.submit,
            width=15
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
        self.status_var.set("Ready")
    
    def submit(self):
        year = self.year_var.get().strip()
        quarter = self.quarter_var.get()
        division = self.division_var.get()
        
        # Validate inputs
        if not year:
            messagebox.showerror("Input Error", "Please enter a valid year")
            return
        if not year.isdigit() or len(year) != 4:
            messagebox.showerror("Input Error", "Year must be a 4-digit number")
            return
        
        self.status_var.set("Processing...")
        self.master.update_idletasks()
        
        try:
            # Get absolute paths
            config_path = resource_path(os.path.join('data', 'config.xlsx'))
            input_dir = resource_path(os.path.join('data', 'input_files'))
            output_dir = resource_path('output')
            
            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)
            
            # Import processing modules
            from data_cleaning import clean_data
            from data_processing import process_data
            from report_generation import generate_report
            
            # Execute processing pipeline
            cleaned_data = clean_data(input_dir, year, quarter, division)
            processed_data = process_data(cleaned_data, config_path)
            report_path = os.path.join(output_dir, f"report_{year}_{quarter}_{division}.pdf")
            generate_report(processed_data, report_path)
            
            messagebox.showinfo(
                "Success", 
                f"Report generated successfully!\nSaved to: {report_path}"
            )
            self.status_var.set("Report generated successfully")
            
        except Exception as e:
            messagebox.showerror(
                "Processing Error", 
                f"An error occurred:\n{str(e)}"
            )
            self.status_var.set("Error occurred")
            # For debugging during development
            if not hasattr(sys, '_MEIPASS'):
                raise e

if __name__ == "__main__":
    root = tk.Tk()
    app = InputGUI(root)
    root.mainloop()
