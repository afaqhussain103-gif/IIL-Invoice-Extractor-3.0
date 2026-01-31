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
        patterns = [
            r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})',
            r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                try:
                    groups = match.groups()
                    if len(groups[0]) == 4:
                        return datetime(int(groups[0]), int(groups[1]), int(groups[2]))
                    else:
                        return datetime(int(groups[2]), int(groups[1]), int(groups[0]))
                except:
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
        
        if not pdf_files:
            messagebox.showerror("Error", "No PDFs found!")
            return
        
        total = len(pdf_files)
        extracted = 0
        
        self.progress['maximum'] = total
        output_pdf = fitz.open()
        
        for idx, filename in enumerate(pdf_files):
            self.progress['value'] = idx + 1
            self.status.config(text=f"Scanning {idx+1}/{total}")
            self.root.update()
            
            try:
                doc = fitz.open(os.path.join(source, filename))
                
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = page.get_text()
                    
                    if search in text.lower():
                        if date_from or date_to:
                            page_date = self.extract_date(text)
                            if page_date:
                                if date_from and page_date < date_from:
                                    continue
                                if date_to and page_date > date_to:
                                    continue
                        
                        output_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        extracted += 1
                
                doc.close()
            except Exception as e:
                print(f"Error: {e}")
        
        if extracted > 0:
            output_name = f"{search.replace(' ', '_')}_extracted.pdf"
            output_pdf.save(os.path.join(dest, output_name))
            messagebox.showinfo("Success", f"Extracted {extracted} pages!\n\nSaved as: {output_name}")
        else:
            messagebox.showinfo("No Match", "No matching pages found!")
        
        output_pdf.close()
        self.progress['value'] = 0
        self.status.config(text="Ready")

if __name__ == "__main__":
    root = tk.Tk()
    app = IILPageExtractor(root)
    root.mainloop()
