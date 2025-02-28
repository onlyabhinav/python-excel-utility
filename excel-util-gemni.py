import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
from datetime import datetime

class ExcelUtilityApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Python Excel Utility")
        self.geometry("1200x800")

        self.excel_file_path = None
        self.sheet_name = None
        self.df = pd.DataFrame()
        self.all_columns = []
        self.selected_columns = []
        self.filtered_df = pd.DataFrame()
        self.sorted_df = pd.DataFrame()
        self.filter_criteria = {}
        self.sort_criteria = {}
        self.saved_configurations = {}
        self.config_file = "column_configurations.json"
        self.load_configurations() # Load configurations at startup

        self.create_widgets()

    def create_widgets(self):
        # --- File Selection ---
        file_frame = ttk.LabelFrame(self, text="1. File Selection")
        file_frame.pack(pady=10, padx=10, fill=tk.X)

        self.file_button = ttk.Button(file_frame, text="Select Excel File", command=self.select_excel_file)
        self.file_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=5, pady=5)

        # --- Sheet Management ---
        sheet_frame = ttk.LabelFrame(self, text="2. Sheet Management")
        sheet_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(sheet_frame, text="Select Sheet:").pack(side=tk.LEFT, padx=5, pady=5)
        self.sheet_dropdown = ttk.Combobox(sheet_frame, state="readonly")
        self.sheet_dropdown.pack(side=tk.LEFT, padx=5, pady=5)
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.select_sheet)

        # --- Column Selection ---
        column_frame = ttk.LabelFrame(self, text="3. Column Selection")
        column_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(column_frame, text="Available Columns:").pack(side=tk.LEFT, padx=5, pady=5)
        self.column_listbox = tk.Listbox(column_frame, selectmode=tk.MULTIPLE, height=5)
        self.column_listbox.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

        select_buttons_frame = ttk.Frame(column_frame)
        select_buttons_frame.pack(side=tk.LEFT, padx=5, pady=5, anchor="n")

        ttk.Button(select_buttons_frame, text="Select Columns", command=self.select_columns).pack(pady=5, fill=tk.X)
        ttk.Button(select_buttons_frame, text="Clear Selection", command=self.clear_column_selection).pack(pady=5, fill=tk.X)

        ttk.Label(column_frame, text="Selected Columns:").pack(side=tk.LEFT, padx=5, pady=5)
        self.selected_column_listbox = tk.Listbox(column_frame, height=5)
        self.selected_column_listbox.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

        # --- Save/Load Column Configurations ---
        config_frame = ttk.LabelFrame(self, text="4. Save/Load Column Configurations")
        config_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(config_frame, text="Configuration Name:").pack(side=tk.LEFT, padx=5, pady=5)
        self.config_name_entry = ttk.Entry(config_frame)
        self.config_name_entry.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        ttk.Button(config_frame, text="Save Configuration", command=self.save_column_config).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(config_frame, text="Load Configuration", command=self.load_column_config).pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(config_frame, text="Saved Configurations:").pack(side=tk.LEFT, padx=5, pady=5)
        self.config_dropdown = ttk.Combobox(config_frame, state="readonly")
        self.config_dropdown.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)
        self.config_dropdown['values'] = list(self.saved_configurations.keys()) # Moved here after dropdown creation
        self.config_dropdown.bind("<<ComboboxSelected>>", self.populate_config_name) # Moved here after dropdown creation


        # --- Data Display ---
        data_frame = ttk.LabelFrame(self, text="5. Data Display")
        data_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.treeview = ttk.Treeview(data_frame, show="headings")
        self.treeview_x_scrollbar = ttk.Scrollbar(data_frame, orient="horizontal", command=self.treeview.xview)
        self.treeview_y_scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=self.treeview.yview)
        self.treeview.configure(xscrollcommand=self.treeview_x_scrollbar.set, yscrollcommand=self.treeview_y_scrollbar.set)

        self.treeview_x_scrollbar.pack(side="bottom", fill="x")
        self.treeview_y_scrollbar.pack(side="right", fill="y")
        self.treeview.pack(side="left", fill="both", expand=True)

        # --- Filtering ---
        filter_frame = ttk.LabelFrame(self, text="6. Filtering")
        filter_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(filter_frame, text="Column:").pack(side=tk.LEFT, padx=5, pady=5)
        self.filter_column_dropdown = ttk.Combobox(filter_frame, state="readonly")
        self.filter_column_dropdown.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(filter_frame, text="Condition:").pack(side=tk.LEFT, padx=5, pady=5)
        self.filter_condition_dropdown = ttk.Combobox(filter_frame, state="readonly", values=["equals", "contains", "greater than", "less than", "starts with", "ends with"])
        self.filter_condition_dropdown.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(filter_frame, text="Value:").pack(side=tk.LEFT, padx=5, pady=5)
        self.filter_value_entry = ttk.Entry(filter_frame)
        self.filter_value_entry.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        ttk.Button(filter_frame, text="Apply Filter", command=self.apply_filter).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(filter_frame, text="Clear Filter", command=self.clear_filter).pack(side=tk.LEFT, padx=5, pady=5)

        # --- Sorting ---
        sort_frame = ttk.LabelFrame(self, text="7. Sorting")
        sort_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(sort_frame, text="Column:").pack(side=tk.LEFT, padx=5, pady=5)
        self.sort_column_dropdown = ttk.Combobox(sort_frame, state="readonly")
        self.sort_column_dropdown.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(sort_frame, text="Order:").pack(side=tk.LEFT, padx=5, pady=5)
        self.sort_order_dropdown = ttk.Combobox(sort_frame, state="readonly", values=["Ascending", "Descending"])
        self.sort_order_dropdown.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Button(sort_frame, text="Apply Sort", command=self.apply_sort).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(sort_frame, text="Clear Sort", command=self.clear_sort).pack(side=tk.LEFT, padx=5, pady=5)

        # --- Export Options ---
        export_frame = ttk.LabelFrame(self, text="8. Export Options")
        export_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(export_frame, text="Export Format:").pack(side=tk.LEFT, padx=5, pady=5)
        self.export_format_dropdown = ttk.Combobox(export_frame, state="readonly", values=["Excel", "CSV", "TXT"]) # Added TXT here
        self.export_format_dropdown.pack(side=tk.LEFT, padx=5, pady=5)
        self.export_format_dropdown.set("Excel") # Default to Excel

        ttk.Button(export_frame, text="Export Data", command=self.export_data).pack(side=tk.LEFT, padx=5, pady=5)

        # --- Status Bar ---
        self.status_bar = ttk.Label(self, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # --- Functionalities ---
    def select_excel_file(self):
        filetypes = (("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        filepath = filedialog.askopenfilename(title="Select an Excel File", filetypes=filetypes)
        if filepath:
            try:
                self.excel_file_path = filepath
                self.file_label.config(text=os.path.basename(self.excel_file_path))
                self.load_sheet_names()
                self.status_message("Excel file loaded successfully.")
            except Exception as e:
                self.status_message(f"Error loading file: {e}")
                messagebox.showerror("Error", f"Could not load Excel file.\n{e}")

    def load_sheet_names(self):
        try:
            xls = pd.ExcelFile(self.excel_file_path)
            sheet_names = xls.sheet_names
            self.sheet_dropdown['values'] = sheet_names
            if sheet_names:
                self.sheet_dropdown.set(sheet_names[0]) # Select first sheet by default
                self.select_sheet() # Automatically load the first sheet
            else:
                self.sheet_dropdown.set('')
            self.status_message("Sheet names loaded.")
        except Exception as e:
            self.status_message(f"Error loading sheet names: {e}")
            messagebox.showerror("Error", f"Could not load sheet names.\n{e}")
            self.sheet_dropdown['values'] = []
            self.sheet_dropdown.set('')

    def select_sheet(self, event=None):
        self.sheet_name = self.sheet_dropdown.get()
        if not self.sheet_name:
            return
        try:
            self.df = pd.read_excel(self.excel_file_path, sheet_name=self.sheet_name)
            self.all_columns = list(self.df.columns)
            self.column_listbox.delete(0, tk.END)
            self.filter_column_dropdown['values'] = self.all_columns
            self.sort_column_dropdown['values'] = self.all_columns
            for col in self.all_columns:
                self.column_listbox.insert(tk.END, col)
            self.selected_columns = []
            self.selected_column_listbox.delete(0, tk.END)
            self.clear_filter()
            self.clear_sort()
            self.status_message(f"Sheet '{self.sheet_name}' loaded.")
        except Exception as e:
            self.status_message(f"Error loading sheet data: {e}")
            messagebox.showerror("Error", f"Could not load sheet data for '{self.sheet_name}'.\n{e}")

    def select_columns(self):
        selected_indices = self.column_listbox.curselection()
        self.selected_columns = [self.column_listbox.get(i) for i in selected_indices]
        self.update_selected_column_listbox()
        self.update_data_display()
        self.status_message("Columns selected.")

    def clear_column_selection(self):
        self.selected_columns = []
        self.update_selected_column_listbox()
        self.update_data_display()
        self.status_message("Column selection cleared.")

    def update_selected_column_listbox(self):
        self.selected_column_listbox.delete(0, tk.END)
        for col in self.selected_columns:
            self.selected_column_listbox.insert(tk.END, col)

    def update_data_display(self):
        self.treeview.delete(*self.treeview.get_children())
        self.treeview["columns"] = self.selected_columns if self.selected_columns else self.all_columns
        self.treeview.column("#0", width=0, stretch=tk.NO) # Hide index column
        self.treeview.heading("#0", text="")

        for col in self.treeview["columns"]:
            self.treeview.heading(col, text=col)
            self.treeview.column(col, anchor=tk.W, width=100)

        df_to_display = self.sorted_df if not self.sorted_df.empty else (self.filtered_df if not self.filtered_df.empty else self.df)
        display_cols = self.selected_columns if self.selected_columns else self.all_columns

        if not df_to_display.empty:
            for index, row in df_to_display.iterrows():
                values = [row[col] for col in display_cols if col in df_to_display.columns] # Ensure column exists in filtered/sorted df
                self.treeview.insert("", tk.END, values=values)
        self.status_message("Data display updated.")

    def apply_filter(self):
        filter_column = self.filter_column_dropdown.get()
        filter_condition = self.filter_condition_dropdown.get()
        filter_value = self.filter_value_entry.get()

        if not filter_column or not filter_condition or filter_value == '':
            self.status_message("Please select a column, condition, and enter a filter value.")
            return

        try:
            original_df = self.sorted_df if not self.sorted_df.empty else self.df # Filter from sorted if sorted, else from original
            if original_df.empty:
                original_df = self.df

            if filter_condition == "equals":
                self.filtered_df = original_df[original_df[filter_column].astype(str).str.lower() == filter_value.lower()]
            elif filter_condition == "contains":
                self.filtered_df = original_df[original_df[filter_column].astype(str).str.lower().str.contains(filter_value.lower(), na=False)]
            elif filter_condition == "starts with":
                self.filtered_df = original_df[original_df[filter_column].astype(str).str.lower().str.startswith(filter_value.lower(), na=False)]
            elif filter_condition == "ends with":
                self.filtered_df = original_df[original_df[filter_column].astype(str).str.lower().str.endswith(filter_value.lower(), na=False)]
            elif filter_condition == "greater than":
                self.filtered_df = original_df[original_df[filter_column] > pd.to_numeric(filter_value, errors='coerce')].dropna(subset=[filter_column])
            elif filter_condition == "less than":
                self.filtered_df = original_df[original_df[filter_column] < pd.to_numeric(filter_value, errors='coerce')].dropna(subset=[filter_column])
            else:
                self.filtered_df = original_df # Fallback, should not happen
            self.sort_criteria = {} # Clear sort when new filter is applied for correct starting point
            self.sorted_df = pd.DataFrame()
            self.update_data_display()
            self.filter_criteria = {'column': filter_column, 'condition': filter_condition, 'value': filter_value}
            self.status_message("Filter applied.")

        except Exception as e:
            self.status_message(f"Error applying filter: {e}")
            messagebox.showerror("Error", f"Could not apply filter.\n{e}")
            self.filtered_df = pd.DataFrame()

    def clear_filter(self):
        self.filtered_df = pd.DataFrame()
        self.filter_criteria = {}
        self.update_data_display()
        self.status_message("Filter cleared.")

    def apply_sort(self):
        sort_column = self.sort_column_dropdown.get()
        sort_order = self.sort_order_dropdown.get()

        if not sort_column or not sort_order:
            self.status_message("Please select a column and sort order.")
            return

        try:
            df_to_sort = self.filtered_df if not self.filtered_df.empty else self.df
            if not df_to_sort.empty:
                ascending = sort_order == "Ascending"
                self.sorted_df = df_to_sort.sort_values(by=sort_column, ascending=ascending)
                self.sort_criteria = {'column': sort_column, 'order': sort_order}
                self.update_data_display()
                self.status_message(f"Data sorted by '{sort_column}' in {sort_order} order.")
            else:
                self.status_message("No data to sort.")
        except Exception as e:
            self.status_message(f"Error applying sort: {e}")
            messagebox.showerror("Error", f"Could not apply sort.\n{e}")
            self.sorted_df = pd.DataFrame()

    def clear_sort(self):
        self.sorted_df = pd.DataFrame()
        self.sort_criteria = {}
        self.update_data_display()
        self.status_message("Sort cleared.")

    def export_data(self):
        export_format = self.export_format_dropdown.get()
        if export_format not in ["Excel", "CSV", "TXT"]: # Added TXT here
            self.status_message("Please select a valid export format.")
            return

        export_df = self.sorted_df if not self.sorted_df.empty else (self.filtered_df if not self.filtered_df.empty else self.df)
        if export_df.empty:
            self.status_message("No data to export.")
            return

        filename_base = os.path.splitext(os.path.basename(self.excel_file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filter_str = ""
        sort_str = ""

        if self.filter_criteria:
            filter_col_abbrv = "".join([word[0].upper() for word in self.filter_criteria['column'].split()])
            filter_condition_abbrv = "".join([word[0].upper() for word in self.filter_criteria['condition'].split()])
            filter_str = f"_F-{filter_col_abbrv}_{filter_condition_abbrv}_{self.filter_criteria['value']}"

        if self.sort_criteria:
            sort_col_abbrv = "".join([word[0].upper() for word in self.sort_criteria['column'].split()])
            sort_order_abbrv = self.sort_criteria['order'][0].upper() # A or D
            sort_str = f"_S-{sort_col_abbrv}_{sort_order_abbrv}"

        default_filename = f"{filename_base}{filter_str}{sort_str}_{timestamp}"
        if export_format == "Excel":
            default_filename += ".xlsx"
            filetypes = (("Excel files", "*.xlsx"), ("All files", "*.*"))
            filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes, initialfile=default_filename, title="Export to Excel")
            if filepath:
                try:
                    export_df[self.selected_columns if self.selected_columns else self.all_columns].to_excel(filepath, index=False)
                    self.status_message(f"Data exported to '{filepath}' in Excel format.")
                except Exception as e:
                    self.status_message(f"Error exporting to Excel: {e}")
                    messagebox.showerror("Error", f"Could not export to Excel.\n{e}")
        elif export_format == "CSV":
            default_filename += ".csv"
            filetypes = (("CSV files", "*.csv"), ("All files", "*.*"))
            filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=filetypes, initialfile=default_filename, title="Export to CSV")
            if filepath:
                try:
                    export_df[self.selected_columns if self.selected_columns else self.all_columns].to_csv(filepath, index=False)
                    self.status_message(f"Data exported to '{filepath}' in CSV format.")
                except Exception as e:
                    self.status_message(f"Error exporting to CSV: {e}")
                    messagebox.showerror("Error", f"Could not export to CSV.\n{e}")
        elif export_format == "TXT": # Added TXT export here
            default_filename += ".txt"
            filetypes = (("Text files", "*.txt"), ("All files", "*.*"))
            filepath = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=filetypes, initialfile=default_filename, title="Export to TXT")
            if filepath:
                try:
                    export_df[self.selected_columns if self.selected_columns else self.all_columns].to_csv(filepath, index=False, sep='\t') # Tab separated
                    self.status_message(f"Data exported to '{filepath}' in TXT format.")
                except Exception as e:
                    self.status_message(f"Error exporting to TXT: {e}")
                    messagebox.showerror("Error", f"Could not export to TXT.\n{e}")

    def save_column_config(self):
        config_name = self.config_name_entry.get()
        if not config_name:
            self.status_message("Please enter a configuration name.")
            return

        if config_name in self.saved_configurations:
            if not messagebox.askyesno("Confirm Overwrite", f"Configuration '{config_name}' already exists. Overwrite?"):
                return

        self.saved_configurations[config_name] = self.selected_columns
        self.config_dropdown['values'] = list(self.saved_configurations.keys())
        self.save_configurations() # Save to file
        self.config_dropdown.set(config_name)
        self.status_message(f"Configuration '{config_name}' saved.")

    def load_column_config(self):
        config_name = self.config_dropdown.get()
        if not config_name:
            self.status_message("Please select a configuration to load.")
            return

        if config_name not in self.saved_configurations:
            self.status_message("Configuration not found.")
            return

        loaded_columns = self.saved_configurations[config_name]
        valid_columns = [col for col in loaded_columns if col in self.all_columns] # Check if columns still exist in current sheet

        if len(valid_columns) < len(loaded_columns):
            mismatched_cols = set(loaded_columns) - set(valid_columns)
            messagebox.showwarning("Configuration Mismatch", f"Some columns in the configuration '{config_name}' are not found in the current sheet:\n{', '.join(mismatched_cols)}.\nLoading only valid columns.")

        self.selected_columns = valid_columns
        self.update_selected_column_listbox()
        self.update_data_display()
        self.config_name_entry.delete(0, tk.END) # Clear entry after loading
        self.config_name_entry.insert(0, config_name)
        self.status_message(f"Configuration '{config_name}' loaded.")

    def populate_config_name(self, event=None):
        config_name = self.config_dropdown.get()
        if config_name:
            self.config_name_entry.delete(0, tk.END)
            self.config_name_entry.insert(0, config_name)

    def save_configurations(self):
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.saved_configurations, f)
        except Exception as e:
            self.status_message(f"Error saving configurations: {e}")

    def load_configurations(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    self.saved_configurations = json.load(f)
                if hasattr(self, 'config_dropdown'): # Check if config_dropdown is created before setting values
                    self.config_dropdown['values'] = list(self.saved_configurations.keys())
        except Exception as e:
            self.status_message(f"Error loading configurations: {e}")

    def status_message(self, message):
        self.status_bar.config(text=message)
        self.status_bar.update_idletasks() # Force update status bar immediately

if __name__ == "__main__":
    app = ExcelUtilityApp()
    app.mainloop()