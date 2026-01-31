"""
IIL Invoice Extractor v5.0 - Page Extractor
"""

import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import fitz
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF"])
    import fitz

class IILPageExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v5.0")
        self.root.geometry("650x550")
        
        self.green = '#01783f'
        self.yellow = '#c2d501'
        
        self.setup_ui()
    
    def setup_ui(self):
        header = tk.Frame(self.root, bg=self.green, height=70)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="IIL INVOICE EXTRACTOR",
            font=('Arial', 14, 'bold'),
            bg=self.green,
            fg='white'
        ).pack(pady=20)
        
        main = tk.Frame(self.root, padx=25, pady=15)
        main.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main, text="Source Folder:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=3)
        sf = tk.Frame(main)
        sf.pack(fill=tk.X, pady=5)
        self.source_entry = tk.Entry(sf, font=('Arial', 9))
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(sf, text="Browse", command=self.browse_source, bg=self.green, fg='white', padx=10).pack(side=tk.LEFT, padx=5)
        
        tk.Label(main, text="Destination Folder:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=3)
        df = tk.Frame(main)
        df.pack(fill=tk.X, pady=5)
        self.dest_entry = tk.Entry(df, font=('Arial', 9))
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(df, text="Browse", command=self.browse_dest, bg=self.green, fg='white', padx=10).pack(side=tk.LEFT, padx=5)
        
        tk.Label(main, text="Date From (YYYY-MM-DD, optional):", font=('Arial', 9)).pack(anchor='w', pady=3)
        self.date_from = tk.Entry(main, font=('Arial', 9))
        self.date_from.pack(fill=tk.X, pady=3)
        self.date_from.insert(0, "2024-01-01")
        
        tk.Label(main, text="Date To (YYYY-MM-DD, optional):", font=('Arial', 9)).pack(anchor='w', pady=3)
        self.date_to = tk.Entry(main, font=('Arial', 9))
        self.date_to.pack(fill=tk.X, pady=3)
        self.date_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        tk.Label(main, text="Customer Name or Account ID:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=3)
        self.search_entry = tk.Entry(main, font=('Arial', 10))
        self.search_entry.pack(fill=tk.X, pady=5)
        
        self.progress = ttk.Progressbar(main, length=500, mode='determinate')
        self.progress.pack(pady=10)
        
        self.status = tk.Label(main, text="Ready", font=('Arial', 9))
        self.status.pack(pady=5)
        
        tk.Button(
            main,
            text="EXTRACT PAGES",
            command=self.extract_pages,
            bg=self.yellow,
            font=('Arial', 11, 'bold'),
            padx=30,
            pady=10
        ).pack(pady=10)
    
    def browse_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)
    
    def browse_dest(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, folder)
    
    def extract_date(self, text):
        """Extract date from text - handles multiple formats, case-insensitive"""
        
        # Month mapping (case-insensitive)
        months = {
            'jan': 1, 'january': 1,
            'feb': 2, 'february': 2,
            'mar': 3, 'march': 3,
            'apr': 4, 'april': 4,
            'may': 5,
            'jun': 6, 'june': 6,
            'jul': 7, 'july': 7,
            'aug': 8, 'august': 8,
            'sep': 9, 'sept': 9, 'september': 9,
            'oct': 10, 'october': 10,
            'nov': 11, 'november': 11,
            'dec': 12, 'december': 12
        }
        
        patterns = [
            # 01-JAN-2026, 01-Jan-2026, 01-jan-2026
            r'(\d{1,2})[-\s](jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-\s](\d{4})',
            # JAN-01-2026, Jan-01-2026
            r'(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-\s](\d{1,2})[-\s](\d{4})',
            # 01 January 2026, 01 Jan 2026
            r'(\d{1,2})\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+(\d{4})',
            # DD/MM/YYYY
            r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})',
            # YYYY-MM-DD
            r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    groups = match.groups()
                    
                    # Check if first group is a month name
                    if groups[0].lower() in months:
                        # JAN-01-2026 format
                        month = months[groups[0].lower()]
                        day = int(groups[1])
                        year = int(groups[2])
                        return datetime(year, month, day)
                    
                    # Check if second group is a month name
                    elif len(groups) > 1 and groups[1].lower() in months:
                        # 01-JAN-2026 format
                        day = int(groups[0])
                        month = months[groups[1].lower()]
                        year = int(groups[2])
                        return datetime(year, month, day)
                    
                    # Numeric formats
                    elif len(groups[0]) == 4:
                        # YYYY-MM-DD
                        return datetime(int(groups[0]), int(groups[1]), int(groups[2]))
                    else:
                        # DD/MM/YYYY
                        return datetime(int(groups[2]), int(groups[1]), int(groups[0]))
                
                except Exception as e:
                    print(f"Date parse error: {e}")
                    continue
        
        return None
    
    def extract_pages(self):
        source = self.source_entry.get()
        dest = self.dest_entry.get()
        search = self.search_entry.get().lower().strip()
        
        if not source or not dest or not search:
            messagebox.showerror("Error", "Fill all required fields!")
            return
        
        if not os.path.exists(source):
            messagebox.showerror("Error", "Source folder doesn't exist!")
            return
        
        date_from = None
        date_to = None
        
        df_str = self.date_from.get().strip()
        dt_str = self.date_to.get().strip()
        
        if df_str:
            try:
                date_from = datetime.strptime(df_str, "%Y-%m-%d")
            except:
                pass
        
        if dt_str:
            try:
                date_to = datetime.strptime(dt_str, "%Y-%m-%d")
            except:
                pass
        
        os.makedirs(dest, exist_ok=True)
        
        pdf_files = [f for f in os.listdir(source) if f.lower().endswith('.pdf')]
        
        if not pdf*
