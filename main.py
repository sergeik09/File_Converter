import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from PIL import Image
import pandas as pd
import json
import csv
from pathlib import Path

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Format Converter")
        self.root.geometry("670x300")
        self.root.configure(bg="#f0f0f0")
        
        # Supported conversions
        self.supported_conversions = {
            "Image": {
                "PNG": [".jpg", ".jpeg", ".bmp", ".gif"],
                "JPG": [".png", ".bmp", ".gif"],
                "BMP": [".png", ".jpg", ".jpeg", ".gif"],
                "GIF": [".png", ".jpg", ".jpeg", ".bmp"]
            },
            "Data": {
                "CSV": [".xlsx", ".json"],
                "JSON": [".csv", ".xlsx"],
                "XLSX": [".csv", ".json"]
            }
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="File Format Converter", font=("Helvetica", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=20)
        
        # File selection
        ttk.Label(main_frame, text="Select File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.file_path = tk.StringVar()
        file_entry = ttk.Entry(main_frame, textvariable=self.file_path, width=50)
        file_entry.grid(row=1, column=1, padx=5, pady=5)
        browse_btn = ttk.Button(main_frame, text="Browse", command=self.browse_file)
        browse_btn.grid(row=1, column=2, padx=5, pady=5)
        
        # Format selection
        ttk.Label(main_frame, text="Convert to:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.target_format = tk.StringVar()
        self.format_combo = ttk.Combobox(main_frame, textvariable=self.target_format, state="readonly")
        self.format_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Convert button
        convert_btn = ttk.Button(main_frame, text="Convert", command=self.convert_file)
        convert_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Status label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_path.set(file_path)
            self.update_target_formats(file_path)
            
    def update_target_formats(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        target_formats = []
        
        # Find supported conversions for the file type
        for category, formats in self.supported_conversions.items():
            for source_format, targets in formats.items():
                if f".{source_format.lower()}" == ext:
                    target_formats = [fmt.strip('.') for fmt in targets]
                    break
                    
        if target_formats:
            self.format_combo['values'] = target_formats
            if target_formats:
                self.format_combo.set(target_formats[0])
        else:
            self.format_combo['values'] = []
            self.format_combo.set('')
            messagebox.showwarning("Unsupported Format", "This file format is not supported for conversion.")
            
    def convert_file(self):
        source_path = self.file_path.get()
        if not source_path:
            messagebox.showerror("Error", "Please select a file to convert")
            return
            
        target_format = self.target_format.get()
        if not target_format:
            messagebox.showerror("Error", "Please select a target format")
            return
            
        try:
            self.progress['value'] = 0
            self.status_var.set("Converting...")
            
            source_ext = os.path.splitext(source_path)[1].lower()
            target_path = os.path.splitext(source_path)[0] + f".{target_format.lower()}"
            
            # Image conversion
            if source_ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif']:
                img = Image.open(source_path)
                # Convert RGBA to RGB if saving as JPEG
                if target_format.lower() in ['jpg', 'jpeg'] and img.mode == 'RGBA':
                    # Create a white background
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    # Paste the image onto the background using alpha channel as mask
                    background.paste(img, mask=img.split()[3])
                    img = background
                img.save(target_path)
                
            # Data conversion
            elif source_ext == '.csv':
                if target_format == 'JSON':
                    df = pd.read_csv(source_path)
                    df.to_json(target_path, orient='records')
                elif target_format == 'XLSX':
                    df = pd.read_csv(source_path)
                    df.to_excel(target_path, index=False)
                    
            elif source_ext == '.json':
                if target_format == 'CSV':
                    df = pd.read_json(source_path)
                    df.to_csv(target_path, index=False)
                elif target_format == 'XLSX':
                    df = pd.read_json(source_path)
                    df.to_excel(target_path, index=False)
                    
            elif source_ext == '.xlsx':
                if target_format == 'CSV':
                    df = pd.read_excel(source_path)
                    df.to_csv(target_path, index=False)
                elif target_format == 'JSON':
                    df = pd.read_excel(source_path)
                    df.to_json(target_path, orient='records')
                    
            self.progress['value'] = 100
            self.status_var.set("Conversion completed!")
            messagebox.showinfo("Success", f"File converted successfully!\nSaved as: {target_path}")
            
        except Exception as e:
            self.status_var.set("Error during conversion")
            messagebox.showerror("Error", f"An error occurred during conversion: {str(e)}")
            
if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()
