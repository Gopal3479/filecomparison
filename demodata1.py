import tkinter as tk
from tkinter import ttk
import os
import sys

class InputGUI:
    def __init__(self, master, config_filename="config.xlsx"):
        """
        Initialize GUI with config file from data folder
        
        Args:
            master: Tkinter root window
            config_filename: Name of config file in data folder
        """
        self.master = master
        self.config_filename = config_filename
        self.config_path = self.get_config_path()
        
        # Rest of your initialization code...
        self.setup_ui()
        self.status_var.set(f"Using config: {os.path.basename(self.config_path)}")

    def get_config_path(self):
        """Resolve config file path whether running as script or executable"""
        base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
        data_dir = os.path.join(base_path, "data")
        
        # Check if file exists in data folder
        config_path = os.path.join(data_dir, self.config_filename)
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Config file not found at: {config_path}")
        
        return config_path

    def setup_ui(self):
        """Initialize all UI components"""
        # Your existing UI setup code...
        pass

    def process_data(self):
        """Example method using the config path"""
        import pandas as pd
        try:
            df = pd.read_excel(self.config_path)
            # Process your data...
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read config: {str(e)}")
