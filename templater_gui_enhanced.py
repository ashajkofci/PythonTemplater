# -*- coding: utf-8 -*-
"""
templater_gui_enhanced.py

Enhanced GUI application for CSV to DOCX template generation.
Features:
- Multi-column mapping with priority (up to 5 columns per field)
- Column combination support
- Settings persistence per template+source combination
- Drag and drop support for files
"""
import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import hashlib
from pathlib import Path
from templater_core import (
    read_csv_any, get_placeholders_from_template, generate_documents
)


def get_config_dir():
    """Get OS-appropriate configuration directory"""
    if os.name == 'nt':  # Windows
        config_dir = Path(os.environ.get('APPDATA', Path.home())) / 'CSVTemplater'
    elif os.name == 'posix':
        if 'darwin' in os.sys.platform:  # macOS
            config_dir = Path.home() / 'Library' / 'Application Support' / 'CSVTemplater'
        else:  # Linux
            config_dir = Path.home() / '.config' / 'csvtemplater'
    else:
        config_dir = Path.home() / '.csvtemplater'
    
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def get_config_key(csv_path, template_path):
    """Generate a unique config key for a csv+template combination"""
    combined = f"{csv_path}|{template_path}"
    return hashlib.md5(combined.encode()).hexdigest()


class FieldMappingRow:
    """Represents a single field mapping row with multiple columns and priority"""
    def __init__(self, parent, placeholder, csv_columns, on_change_callback, on_delete_callback=None):
        self.parent = parent
        self.placeholder = placeholder
        self.csv_columns = csv_columns
        self.on_change_callback = on_change_callback
        self.on_delete_callback = on_delete_callback
        self.column_vars = []
        self.combine_var = tk.BooleanVar(value=False)
        
    def create_widgets(self, row_idx):
        """Create widgets for this mapping row"""
        frame = ttk.Frame(self.parent)
        frame.grid(row=row_idx, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=2)
        frame.columnconfigure(1, weight=1)
        
        # Placeholder label
        ttk.Label(frame, text=self.placeholder, width=15).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        # Container for column selectors
        columns_frame = ttk.Frame(frame)
        columns_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        columns_frame.columnconfigure(0, weight=1)
        
        self.columns_frame = columns_frame
        self.create_column_selector(0)
        
        # Combine checkbox
        ttk.Checkbutton(frame, text="Combine", variable=self.combine_var,
                       command=self._on_change).grid(row=0, column=2, padx=5)
        
        # Add column button
        add_btn = ttk.Button(frame, text="+", width=3, command=self.add_column_selector)
        add_btn.grid(row=0, column=3, padx=2)
        
        return frame
    
    def create_column_selector(self, idx):
        """Create a single column selector dropdown"""
        selector_frame = ttk.Frame(self.columns_frame)
        selector_frame.grid(row=idx, column=0, sticky=(tk.W, tk.E), pady=1)
        selector_frame.columnconfigure(1, weight=1)
        
        # Priority label
        ttk.Label(selector_frame, text=f"{idx+1}.", width=3).grid(row=0, column=0, sticky=tk.W)
        
        # Column combobox
        var = tk.StringVar()
        combo = ttk.Combobox(selector_frame, textvariable=var,
                           values=[''] + self.csv_columns, state="readonly", width=25)
        combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=2)
        combo.bind('<<ComboboxSelected>>', lambda e: self._on_change())
        
        self.column_vars.append(var)
        
        # Remove button (only if not the first one)
        if idx > 0:
            remove_btn = ttk.Button(selector_frame, text="-", width=3,
                                   command=lambda: self.remove_column_selector(idx))
            remove_btn.grid(row=0, column=2, padx=2)
    
    def add_column_selector(self):
        """Add another column selector (up to 5 total)"""
        if len(self.column_vars) < 5:
            self.create_column_selector(len(self.column_vars))
            self._on_change()
    
    def remove_column_selector(self, idx):
        """Remove a column selector"""
        if len(self.column_vars) > 1:
            # Remove the variable
            self.column_vars.pop(idx)
            # Recreate all selectors
            for widget in self.columns_frame.winfo_children():
                widget.destroy()
            for i in range(len(self.column_vars)):
                self.create_column_selector(i)
            self._on_change()
    
    def _on_change(self):
        """Callback when mapping changes"""
        if self.on_change_callback:
            self.on_change_callback()
    
    def get_mapping(self):
        """Get the mapping configuration for this field"""
        columns = [var.get() for var in self.column_vars if var.get()]
        if not columns:
            return None
        
        return {
            'columns': columns,
            'combine': self.combine_var.get()
        }
    
    def set_mapping(self, mapping_config):
        """Set the mapping from saved configuration"""
        if not mapping_config:
            return
        
        columns = mapping_config.get('columns', [])
        combine = mapping_config.get('combine', False)
        
        # Clear existing
        for widget in self.columns_frame.winfo_children():
            widget.destroy()
        self.column_vars.clear()
        
        # Add selectors for saved columns
        for i, col in enumerate(columns):
            self.create_column_selector(i)
            if i < len(self.column_vars):
                self.column_vars[i].set(col)
        
        # Set combine checkbox
        self.combine_var.set(combine)


class EnhancedTemplaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to DOCX Templater")
        self.root.geometry("900x750")
        
        # Data storage
        self.csv_path = None
        self.template_path = None
        self.csv_columns = []
        self.template_placeholders = []
        self.field_mapping_rows = {}
        self.config_dir = get_config_dir()
        
        self.create_widgets()
        self.setup_drag_drop()
    
    def setup_drag_drop(self):
        """Setup drag and drop for file inputs"""
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.handle_drop)
        except:
            # tkinterdnd2 not available, drag-drop won't work
            pass
    
    def handle_drop(self, event):
        """Handle file drop events"""
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip('{}')  # Remove curly braces if present
            ext = os.path.splitext(file_path)[1].lower()
            
            if ext == '.csv':
                self.load_csv(file_path)
            elif ext == '.docx':
                self.load_template(file_path)
    
    def create_widgets(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        current_row = 0
        
        # === File Selection Section ===
        file_frame = ttk.LabelFrame(main_frame, text="File Selection (or drag & drop)", padding="10")
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
        mapping_label_frame = ttk.LabelFrame(main_frame, text="Field Mapping (Priority: 1st→2nd→...)", padding="10")
        mapping_label_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        mapping_label_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(current_row, weight=1)
        current_row += 1
        
        # Buttons for mapping management
        btn_frame = ttk.Frame(mapping_label_frame)
        btn_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        ttk.Button(btn_frame, text="Reset Config", command=self.reset_config).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_frame, text="Tip: Use '+' to add fallback columns, 'Combine' to merge columns", 
                 foreground="gray").pack(side=tk.LEFT, padx=10)
        
        # Scrollable frame for mappings
        self.mapping_canvas = tk.Canvas(mapping_label_frame, height=250)
        scrollbar = ttk.Scrollbar(mapping_label_frame, orient="vertical", command=self.mapping_canvas.yview)
        self.mapping_scrollframe = ttk.Frame(self.mapping_canvas)
        
        self.mapping_scrollframe.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_scrollframe, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.mapping_canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        mapping_label_frame.rowconfigure(1, weight=1)
        mapping_label_frame.columnconfigure(0, weight=1)
        
        self.mapping_info = ttk.Label(self.mapping_scrollframe, 
                                     text="Select CSV and Template files to configure field mapping",
                                     foreground="gray")
        self.mapping_info.grid(row=0, column=0, columnspan=4, pady=20)
        
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
            self.load_csv(filepath)
    
    def load_csv(self, filepath):
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
            self.load_config()
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
            self.load_template(filepath)
    
    def load_template(self, filepath):
        try:
            self.template_path = filepath
            self.template_label.config(text=os.path.basename(filepath), foreground="black")
            
            # Get placeholders from template
            self.template_placeholders = get_placeholders_from_template(filepath)
            
            self.update_mapping_ui()
            self.load_config()
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
        self.field_mapping_rows.clear()
        
        if not self.template_placeholders or not self.csv_columns:
            if not self.template_placeholders and not self.csv_columns:
                msg = "Select CSV and Template files to configure field mapping"
            elif not self.template_placeholders:
                msg = "Select a Template file to see placeholders"
            else:
                msg = "Select a CSV file to see available columns"
            
            self.mapping_info = ttk.Label(self.mapping_scrollframe, text=msg, foreground="gray")
            self.mapping_info.grid(row=0, column=0, columnspan=4, pady=20)
            return
        
        # Create header
        header_frame = ttk.Frame(self.mapping_scrollframe)
        header_frame.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(header_frame, text="Placeholder", font=('TkDefaultFont', 9, 'bold'), width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="CSV Columns (in priority order)", font=('TkDefaultFont', 9, 'bold')).pack(side=tk.LEFT, padx=5)
        
        # Create mapping rows
        for idx, placeholder in enumerate(self.template_placeholders, start=1):
            row = FieldMappingRow(self.mapping_scrollframe, placeholder, self.csv_columns, 
                                 self.on_mapping_change)
            row.create_widgets(idx)
            self.field_mapping_rows[placeholder] = row
            
            # Try to auto-match placeholder to column
            auto_match = self.auto_match_field(placeholder)
            if auto_match and row.column_vars:
                row.column_vars[0].set(auto_match)
        
        self.mapping_scrollframe.columnconfigure(0, weight=1)
    
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
    
    def on_mapping_change(self):
        """Callback when any mapping changes"""
        self.save_config()
    
    def get_config_path(self):
        """Get the config file path for current csv+template combination"""
        if not self.csv_path or not self.template_path:
            return None
        
        config_key = get_config_key(self.csv_path, self.template_path)
        return self.config_dir / f"{config_key}.json"
    
    def save_config(self):
        """Save current configuration"""
        config_path = self.get_config_path()
        if not config_path:
            return
        
        config = {
            'csv_path': self.csv_path,
            'template_path': self.template_path,
            'mappings': {},
            'filename_field': self.filename_field_var.get(),
            'prefix': self.prefix_var.get(),
            'suffix': self.suffix_var.get(),
            'create_zip': self.zip_var.get()
        }
        
        # Save all field mappings
        for placeholder, row in self.field_mapping_rows.items():
            mapping = row.get_mapping()
            if mapping:
                config['mappings'][placeholder] = mapping
        
        try:
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"Failed to save config: {e}")
    
    def load_config(self):
        """Load saved configuration"""
        config_path = self.get_config_path()
        if not config_path or not config_path.exists():
            return
        
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            # Load field mappings
            mappings = config.get('mappings', {})
            for placeholder, mapping_config in mappings.items():
                if placeholder in self.field_mapping_rows:
                    self.field_mapping_rows[placeholder].set_mapping(mapping_config)
            
            # Load filename configuration
            self.filename_field_var.set(config.get('filename_field', ''))
            self.prefix_var.set(config.get('prefix', ''))
            self.suffix_var.set(config.get('suffix', ''))
            self.zip_var.set(config.get('create_zip', False))
            
            self.progress_var.set("Configuration loaded")
        except Exception as e:
            print(f"Failed to load config: {e}")
    
    def reset_config(self):
        """Reset configuration for current csv+template combination"""
        config_path = self.get_config_path()
        if config_path and config_path.exists():
            try:
                config_path.unlink()
                self.progress_var.set("Configuration reset")
                # Reload UI
                self.update_mapping_ui()
                self.filename_field_var.set('')
                self.prefix_var.set('')
                self.suffix_var.set('')
                self.zip_var.set(False)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to reset config:\n{e}")
        else:
            messagebox.showinfo("Info", "No saved configuration to reset")
    
    def check_ready_to_generate(self):
        if (self.csv_path and self.template_path and 
            hasattr(self, 'output_path') and self.output_path):
            self.generate_btn.config(state="normal")
        else:
            self.generate_btn.config(state="disabled")
    
    def generate_documents(self):
        # Build field mapping with priority and combination support
        field_mapping = {}
        for placeholder, row in self.field_mapping_rows.items():
            mapping = row.get_mapping()
            if mapping:
                columns = mapping['columns']
                combine = mapping['combine']
                
                if combine and len(columns) > 1:
                    # Combine columns with space
                    field_mapping[placeholder] = ' '.join(columns)
                else:
                    # Use priority - first available non-empty column
                    field_mapping[placeholder] = columns[0] if columns else None
                    # Store additional columns for fallback
                    if len(columns) > 1:
                        field_mapping[f"{placeholder}_fallback"] = columns[1:]
        
        if not field_mapping:
            messagebox.showwarning("Warning", "No field mappings configured. Please map at least one field.")
            return
        
        # Get filename configuration
        filename_field = self.filename_field_var.get() or None
        prefix = self.prefix_var.get()
        suffix = self.suffix_var.get()
        make_zip = self.zip_var.get()
        
        # Save configuration before generating
        self.save_config()
        
        # Disable generate button during processing
        self.generate_btn.config(state="disabled")
        self.progress_bar['value'] = 0
        
        # Run generation in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.run_generation, 
                                 args=(field_mapping, filename_field, prefix, suffix, make_zip))
        thread.daemon = True
        thread.start()
    
    def run_generation(self, field_mapping, filename_field, prefix, suffix, make_zip):
        def progress_callback(current, total, message):
            progress = (current / total * 100) if total > 0 else 0
            self.root.after(0, lambda: self.progress_bar.config(value=progress))
            self.root.after(0, lambda: self.progress_var.set(message))
        
        try:
            generated_files, zip_path = generate_documents(
                csv_path=self.csv_path,
                template_path=self.template_path,
                outdir=self.output_path,
                field_mapping=field_mapping,
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
    try:
        root = TkinterDnD.Tk()
    except:
        # Fallback to regular Tk if tkinterdnd2 is not available
        root = tk.Tk()
        print("Note: Drag and drop functionality requires tkinterdnd2 package")
    
    app = EnhancedTemplaterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
