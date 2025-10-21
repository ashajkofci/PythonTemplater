# -*- coding: utf-8 -*-
"""
templater_gui.py

GUI application for CSV to DOCX template generation.
Multiplatform application using tkinter.
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from templater_core import (
    read_csv_any, get_placeholders_from_template, generate_documents
)


class TemplaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to DOCX Templater")
        self.root.geometry("800x700")
        
        # Data storage
        self.csv_path = None
        self.template_path = None
        self.csv_columns = []
        self.template_placeholders = []
        self.field_mapping = {}
        self.mapping_widgets = {}
        
        self.create_widgets()
    
    def create_widgets(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        current_row = 0
        
        # === File Selection Section ===
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)
        current_row += 1
        
        # CSV file
        ttk.Label(file_frame, text="CSV File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.csv_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.csv_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.select_csv).grid(row=0, column=2, padx=5)
        
        # Template file
        ttk.Label(file_frame, text="Template:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.template_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.template_label.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.select_template).grid(row=1, column=2, padx=5)
        
        # Output directory
        ttk.Label(file_frame, text="Output Dir:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_label = ttk.Label(file_frame, text="No directory selected", foreground="gray")
        self.output_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.select_output).grid(row=2, column=2, padx=5)
        
        # === Field Mapping Section ===
        self.mapping_frame = ttk.LabelFrame(main_frame, text="Field Mapping", padding="10")
        self.mapping_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.mapping_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(current_row, weight=1)
        current_row += 1
        
        # Scrollable frame for mappings
        self.mapping_canvas = tk.Canvas(self.mapping_frame, height=200)
        scrollbar = ttk.Scrollbar(self.mapping_frame, orient="vertical", command=self.mapping_canvas.yview)
        self.mapping_scrollframe = ttk.Frame(self.mapping_canvas)
        
        self.mapping_scrollframe.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_scrollframe, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.mapping_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.mapping_frame.rowconfigure(0, weight=1)
        self.mapping_frame.columnconfigure(0, weight=1)
        
        self.mapping_info = ttk.Label(self.mapping_scrollframe, 
                                     text="Select CSV and Template files to configure field mapping",
                                     foreground="gray")
        self.mapping_info.grid(row=0, column=0, columnspan=2, pady=20)
        
        # === Filename Configuration Section ===
        filename_frame = ttk.LabelFrame(main_frame, text="Filename Configuration", padding="10")
        filename_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)
        filename_frame.columnconfigure(1, weight=1)
        current_row += 1
        
        ttk.Label(filename_frame, text="Name Field:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.filename_field_var = tk.StringVar()
        self.filename_field_combo = ttk.Combobox(filename_frame, textvariable=self.filename_field_var, 
                                                 state="readonly")
        self.filename_field_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Label(filename_frame, text="Prefix:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.prefix_var = tk.StringVar(value="")
        ttk.Entry(filename_frame, textvariable=self.prefix_var).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Label(filename_frame, text="Suffix:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.suffix_var = tk.StringVar(value="")
        ttk.Entry(filename_frame, textvariable=self.suffix_var).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # === Options Section ===
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)
        current_row += 1
        
        self.zip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Create ZIP archive", variable=self.zip_var).grid(row=0, column=0, sticky=tk.W)
        
        # === Progress Section ===
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)
        current_row += 1
        
        self.progress_var = tk.StringVar(value="Ready")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        
        # === Action Buttons ===
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=current_row, column=0, pady=10)
        current_row += 1
        
        self.generate_btn = ttk.Button(button_frame, text="Generate Documents", 
                                       command=self.generate_documents, state="disabled")
        self.generate_btn.grid(row=0, column=0, padx=5)
        
        ttk.Button(button_frame, text="Exit", command=self.root.quit).grid(row=0, column=1, padx=5)
    
    def select_csv(self):
        filepath = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filepath:
            try:
                self.csv_path = filepath
                self.csv_label.config(text=os.path.basename(filepath), foreground="black")
                
                # Read CSV to get columns
                df = read_csv_any(filepath)
                self.csv_columns = list(df.columns)
                
                # Update filename field combobox
                self.filename_field_combo['values'] = [''] + self.csv_columns
                if self.csv_columns:
                    self.filename_field_var.set(self.csv_columns[0])
                
                self.update_mapping_ui()
                self.check_ready_to_generate()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read CSV file:\n{e}")
                self.csv_path = None
                self.csv_label.config(text="No file selected", foreground="gray")
    
    def select_template(self):
        filepath = filedialog.askopenfilename(
            title="Select Template File",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if filepath:
            try:
                self.template_path = filepath
                self.template_label.config(text=os.path.basename(filepath), foreground="black")
                
                # Get placeholders from template
                self.template_placeholders = get_placeholders_from_template(filepath)
                
                self.update_mapping_ui()
                self.check_ready_to_generate()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read template file:\n{e}")
                self.template_path = None
                self.template_label.config(text="No file selected", foreground="gray")
    
    def select_output(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_path = directory
            self.output_label.config(text=directory, foreground="black")
            self.check_ready_to_generate()
    
    def update_mapping_ui(self):
        # Clear existing mapping widgets
        for widget in self.mapping_scrollframe.winfo_children():
            widget.destroy()
        self.mapping_widgets.clear()
        
        if not self.template_placeholders or not self.csv_columns:
            if not self.template_placeholders and not self.csv_columns:
                msg = "Select CSV and Template files to configure field mapping"
            elif not self.template_placeholders:
                msg = "Select a Template file to see placeholders"
            else:
                msg = "Select a CSV file to see available columns"
            
            self.mapping_info = ttk.Label(self.mapping_scrollframe, text=msg, foreground="gray")
            self.mapping_info.grid(row=0, column=0, columnspan=2, pady=20)
            return
        
        # Create mapping controls
        ttk.Label(self.mapping_scrollframe, text="Template Placeholder", 
                 font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.mapping_scrollframe, text="CSV Column", 
                 font=('TkDefaultFont', 9, 'bold')).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        for idx, placeholder in enumerate(self.template_placeholders, start=1):
            ttk.Label(self.mapping_scrollframe, text=placeholder).grid(row=idx, column=0, padx=5, pady=2, sticky=tk.W)
            
            var = tk.StringVar()
            combo = ttk.Combobox(self.mapping_scrollframe, textvariable=var, 
                               values=[''] + self.csv_columns, state="readonly", width=30)
            combo.grid(row=idx, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
            
            # Try to auto-match placeholder to column
            auto_match = self.auto_match_field(placeholder)
            if auto_match:
                var.set(auto_match)
            
            self.mapping_widgets[placeholder] = var
        
        self.mapping_scrollframe.columnconfigure(1, weight=1)
    
    def auto_match_field(self, placeholder):
        """Try to automatically match a placeholder to a CSV column."""
        # Remove braces and convert to lowercase
        field_name = placeholder.strip('{}').lower()
        
        # Try exact match
        for col in self.csv_columns:
            if col.lower() == field_name:
                return col
        
        # Try partial match
        for col in self.csv_columns:
            if field_name in col.lower() or col.lower() in field_name:
                return col
        
        return None
    
    def check_ready_to_generate(self):
        if (self.csv_path and self.template_path and 
            hasattr(self, 'output_path') and self.output_path):
            self.generate_btn.config(state="normal")
        else:
            self.generate_btn.config(state="disabled")
    
    def generate_documents(self):
        # Build field mapping
        self.field_mapping = {}
        for placeholder, var in self.mapping_widgets.items():
            csv_col = var.get()
            if csv_col:
                self.field_mapping[placeholder] = csv_col
        
        if not self.field_mapping:
            messagebox.showwarning("Warning", "No field mappings configured. Please map at least one field.")
            return
        
        # Get filename configuration
        filename_field = self.filename_field_var.get() or None
        prefix = self.prefix_var.get()
        suffix = self.suffix_var.get()
        make_zip = self.zip_var.get()
        
        # Disable generate button during processing
        self.generate_btn.config(state="disabled")
        self.progress_bar['value'] = 0
        
        # Run generation in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.run_generation, 
                                 args=(filename_field, prefix, suffix, make_zip))
        thread.daemon = True
        thread.start()
    
    def run_generation(self, filename_field, prefix, suffix, make_zip):
        def progress_callback(current, total, message):
            progress = (current / total * 100) if total > 0 else 0
            self.root.after(0, lambda: self.progress_bar.config(value=progress))
            self.root.after(0, lambda: self.progress_var.set(message))
        
        try:
            generated_files, zip_path = generate_documents(
                csv_path=self.csv_path,
                template_path=self.template_path,
                outdir=self.output_path,
                field_mapping=self.field_mapping,
                filename_field=filename_field,
                filename_prefix=prefix,
                filename_suffix=suffix,
                make_zip=make_zip,
                progress_callback=progress_callback
            )
            
            message = f"Success! Generated {len(generated_files)} documents"
            if zip_path:
                message += f"\nZIP archive: {os.path.basename(zip_path)}"
            
            self.root.after(0, lambda: messagebox.showinfo("Success", message))
            self.root.after(0, lambda: self.progress_var.set(f"Complete - {len(generated_files)} files generated"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Generation failed:\n{e}"))
            self.root.after(0, lambda: self.progress_var.set("Error occurred"))
        finally:
            self.root.after(0, lambda: self.generate_btn.config(state="normal"))


def main():
    root = tk.Tk()
    app = TemplaterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
