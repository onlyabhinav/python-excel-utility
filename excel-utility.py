import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import json
from datetime import datetime

class ExcelUtilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Utility")
        self.root.geometry("1000x700")
        
        # Variables to store data
        self.excel_file_path = None
        self.current_df = None
        self.sheets = []
        self.selected_sheet = None
        self.columns = []
        self.selected_columns = []
        
        # Create frames
        self.create_frames()
        
        # Column configurations directory
        self.configs_dir = "column_configs"
        if not os.path.exists(self.configs_dir):
            os.makedirs(self.configs_dir)
        
        # Create widgets
        self.create_widgets()
        
        # Configure grid weights
        self.configure_grid()
    
    def create_frames(self):
        # Top frame for file selection and sheet selection
        self.top_frame = ttk.LabelFrame(self.root, text="Excel File Selection")
        self.top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        # Middle frame for column selection
        self.middle_frame = ttk.LabelFrame(self.root, text="Column Selection")
        self.middle_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        
        # Filter frame
        self.filter_frame = ttk.LabelFrame(self.root, text="Filter Options")
        self.filter_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        
        # Bottom frame for data display
        self.bottom_frame = ttk.LabelFrame(self.root, text="Data View")
        self.bottom_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
    
    def create_widgets(self):
        # File selection widgets
        self.file_button = ttk.Button(self.top_frame, text="Select Excel File", command=self.select_file)
        self.file_button.grid(row=0, column=0, padx=5, pady=5)
        
        self.file_label = ttk.Label(self.top_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, padx=5, pady=5)
        
        # Sheet selection widgets
        self.sheet_label = ttk.Label(self.top_frame, text="Select Sheet:")
        self.sheet_label.grid(row=1, column=0, padx=5, pady=5)
        
        self.sheet_combobox = ttk.Combobox(self.top_frame, state="disabled")
        self.sheet_combobox.grid(row=1, column=1, padx=5, pady=5)
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # Column selection widgets
        self.columns_listbox_label = ttk.Label(self.middle_frame, text="Available Columns:")
        self.columns_listbox_label.grid(row=0, column=0, padx=5, pady=5)
        
        self.columns_listbox = tk.Listbox(self.middle_frame, selectmode=tk.MULTIPLE, width=30, height=6)
        self.columns_listbox.grid(row=1, column=0, padx=5, pady=5)
        
        self.column_buttons_frame = ttk.Frame(self.middle_frame)
        self.column_buttons_frame.grid(row=1, column=1, padx=5, pady=5)
        
        self.select_column_button = ttk.Button(self.column_buttons_frame, text=">>", command=self.add_column)
        self.select_column_button.pack(pady=2)
        
        self.remove_column_button = ttk.Button(self.column_buttons_frame, text="<<", command=self.remove_column)
        self.remove_column_button.pack(pady=2)
        
        self.selected_columns_label = ttk.Label(self.middle_frame, text="Selected Columns:")
        self.selected_columns_label.grid(row=0, column=2, padx=5, pady=5)
        
        self.selected_columns_listbox = tk.Listbox(self.middle_frame, selectmode=tk.MULTIPLE, width=30, height=6)
        self.selected_columns_listbox.grid(row=1, column=2, padx=5, pady=5)
        
        # Column configuration buttons
        self.column_config_frame = ttk.Frame(self.middle_frame)
        self.column_config_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5)
        
        self.view_data_button = ttk.Button(self.column_config_frame, text="View Data", command=self.view_data)
        self.view_data_button.grid(row=0, column=0, padx=5, pady=5)
        
        self.save_config_button = ttk.Button(self.column_config_frame, text="Save Column Selection", command=self.save_column_config)
        self.save_config_button.grid(row=0, column=1, padx=5, pady=5)
        
        self.load_config_button = ttk.Button(self.column_config_frame, text="Load Column Selection", command=self.load_column_config)
        self.load_config_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Filter options
        self.filter_column_label = ttk.Label(self.filter_frame, text="Filter Column:")
        self.filter_column_label.grid(row=0, column=0, padx=5, pady=5)
        
        self.filter_column_combobox = ttk.Combobox(self.filter_frame, state="disabled")
        self.filter_column_combobox.grid(row=0, column=1, padx=5, pady=5)
        
        self.filter_condition_label = ttk.Label(self.filter_frame, text="Condition:")
        self.filter_condition_label.grid(row=0, column=2, padx=5, pady=5)
        
        self.filter_condition_combobox = ttk.Combobox(self.filter_frame, values=["equals", "contains", "greater than", "less than", "starts with", "ends with"], state="disabled")
        self.filter_condition_combobox.grid(row=0, column=3, padx=5, pady=5)
        
        self.filter_value_label = ttk.Label(self.filter_frame, text="Value:")
        self.filter_value_label.grid(row=0, column=4, padx=5, pady=5)
        
        self.filter_value_entry = ttk.Entry(self.filter_frame, state="disabled")
        self.filter_value_entry.grid(row=0, column=5, padx=5, pady=5)
        
        self.apply_filter_button = ttk.Button(self.filter_frame, text="Apply Filter", command=self.apply_filter, state="disabled")
        self.apply_filter_button.grid(row=0, column=6, padx=5, pady=5)
        
        self.clear_filter_button = ttk.Button(self.filter_frame, text="Clear Filter", command=self.clear_filter, state="disabled")
        self.clear_filter_button.grid(row=0, column=7, padx=5, pady=5)
        
        # Sort options
        self.sort_column_label = ttk.Label(self.filter_frame, text="Sort Column:")
        self.sort_column_label.grid(row=1, column=0, padx=5, pady=5)
        
        self.sort_column_combobox = ttk.Combobox(self.filter_frame, state="disabled")
        self.sort_column_combobox.grid(row=1, column=1, padx=5, pady=5)
        
        self.sort_order_label = ttk.Label(self.filter_frame, text="Order:")
        self.sort_order_label.grid(row=1, column=2, padx=5, pady=5)
        
        self.sort_order_combobox = ttk.Combobox(self.filter_frame, values=["Ascending", "Descending"], state="disabled")
        self.sort_order_combobox.grid(row=1, column=3, padx=5, pady=5)
        
        self.apply_sort_button = ttk.Button(self.filter_frame, text="Apply Sort", command=self.apply_sort, state="disabled")
        self.apply_sort_button.grid(row=1, column=4, padx=5, pady=5)
        
        self.clear_sort_button = ttk.Button(self.filter_frame, text="Clear Sort", command=self.clear_sort, state="disabled")
        self.clear_sort_button.grid(row=1, column=5, padx=5, pady=5)
        
        # Data display treeview
        self.tree_frame = ttk.Frame(self.bottom_frame)
        self.tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create scrollbars
        self.tree_y_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical")
        self.tree_y_scroll.pack(side="right", fill="y")
        
        self.tree_x_scroll = ttk.Scrollbar(self.tree_frame, orient="horizontal")
        self.tree_x_scroll.pack(side="bottom", fill="x")
        
        # Create treeview
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_y_scroll.set, xscrollcommand=self.tree_x_scroll.set)
        self.tree.pack(fill="both", expand=True)
        
        # Configure scrollbars
        self.tree_y_scroll.config(command=self.tree.yview)
        self.tree_x_scroll.config(command=self.tree.xview)
        
        # Export button
        self.export_button = ttk.Button(self.bottom_frame, text="Export Filtered Data", command=self.export_data, state="disabled")
        self.export_button.pack(pady=5)
    
    def configure_grid(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)
        
        self.top_frame.columnconfigure(1, weight=1)
        self.middle_frame.columnconfigure(0, weight=1)
        self.middle_frame.columnconfigure(2, weight=1)
        
        for i in range(8):
            self.filter_frame.columnconfigure(i, weight=1)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.excel_file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            
            try:
                # Get sheet names
                self.sheets = pd.ExcelFile(file_path).sheet_names
                self.sheet_combobox.config(values=self.sheets, state="readonly")
                
                # Reset other controls
                self.columns_listbox.delete(0, tk.END)
                self.selected_columns_listbox.delete(0, tk.END)
                self.clear_treeview()
                
                # Enable sheet selection
                if self.sheets:
                    self.sheet_combobox.current(0)
                    self.on_sheet_selected(None)
            
            except Exception as e:
                messagebox.showerror("Error", f"Error opening Excel file: {str(e)}")
    
    def on_sheet_selected(self, event):
        selected_sheet = self.sheet_combobox.get()
        if selected_sheet:
            self.selected_sheet = selected_sheet
            
            try:
                # Read the selected sheet
                self.current_df = pd.read_excel(self.excel_file_path, sheet_name=selected_sheet)
                
                # Update columns list
                self.columns = list(self.current_df.columns)
                
                # Update columns listbox
                self.columns_listbox.delete(0, tk.END)
                for col in self.columns:
                    self.columns_listbox.insert(tk.END, col)
                
                # Clear selected columns and treeview
                self.selected_columns = []
                self.selected_columns_listbox.delete(0, tk.END)
                self.clear_treeview()
            
            except Exception as e:
                messagebox.showerror("Error", f"Error reading sheet: {str(e)}")
    
    def add_column(self):
        selected_indices = self.columns_listbox.curselection()
        for i in selected_indices:
            column = self.columns_listbox.get(i)
            if column not in self.selected_columns:
                self.selected_columns.append(column)
                self.selected_columns_listbox.insert(tk.END, column)
    
    def remove_column(self):
        selected_indices = self.selected_columns_listbox.curselection()
        # Convert to list to avoid issues with changing indices
        selected_indices = list(selected_indices)
        # Start from the end to avoid index shifting problems
        for i in reversed(selected_indices):
            column = self.selected_columns_listbox.get(i)
            self.selected_columns.remove(column)
            self.selected_columns_listbox.delete(i)
    
    def view_data(self):
        if not self.current_df is None and self.selected_columns:
            try:
                # Clear current treeview
                self.clear_treeview()
                
                # Get dataframe with selected columns
                display_df = self.current_df[self.selected_columns].copy()
                
                # Configure treeview columns
                self.tree["columns"] = self.selected_columns
                self.tree.column("#0", width=0, stretch=tk.NO)  # Hide first column
                
                for col in self.selected_columns:
                    self.tree.column(col, anchor=tk.W, width=100)
                    self.tree.heading(col, text=col, anchor=tk.W)
                
                # Insert data
                for i, row in display_df.iterrows():
                    values = [row[col] for col in self.selected_columns]
                    self.tree.insert("", tk.END, text=str(i), values=values)
                
                # Enable filter and sort comboboxes
                self.filter_column_combobox.config(values=self.selected_columns, state="readonly")
                self.filter_condition_combobox.config(state="readonly")
                self.filter_value_entry.config(state="normal")
                self.apply_filter_button.config(state="normal")
                self.clear_filter_button.config(state="normal")
                
                self.sort_column_combobox.config(values=self.selected_columns, state="readonly")
                self.sort_order_combobox.config(state="readonly")
                self.apply_sort_button.config(state="normal")
                self.clear_sort_button.config(state="normal")
                
                self.export_button.config(state="normal")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error displaying data: {str(e)}")
    
    def apply_filter(self):
        if self.current_df is None or not self.selected_columns:
            return
        
        filter_column = self.filter_column_combobox.get()
        filter_condition = self.filter_condition_combobox.get()
        filter_value = self.filter_value_entry.get()
        
        if not filter_column or not filter_condition or not filter_value:
            messagebox.showinfo("Info", "Please complete all filter fields")
            return
        
        try:
            # Get dataframe with selected columns
            display_df = self.current_df[self.selected_columns].copy()
            
            # Apply filter based on condition
            if filter_condition == "equals":
                try:
                    # Try to convert to numeric if possible
                    numeric_value = float(filter_value)
                    filtered_df = display_df[display_df[filter_column] == numeric_value]
                except ValueError:
                    # If not numeric, use string comparison (case insensitive)
                    filtered_df = display_df[display_df[filter_column].astype(str).str.lower() == filter_value.lower()]
            
            elif filter_condition == "contains":
                filtered_df = display_df[display_df[filter_column].astype(str).str.contains(filter_value, case=False, na=False)]
            
            elif filter_condition == "greater than":
                try:
                    numeric_value = float(filter_value)
                    filtered_df = display_df[display_df[filter_column] > numeric_value]
                except ValueError:
                    messagebox.showerror("Error", "Value must be numeric for 'greater than' condition")
                    return
            
            elif filter_condition == "less than":
                try:
                    numeric_value = float(filter_value)
                    filtered_df = display_df[display_df[filter_column] < numeric_value]
                except ValueError:
                    messagebox.showerror("Error", "Value must be numeric for 'less than' condition")
                    return
            
            elif filter_condition == "starts with":
                filtered_df = display_df[display_df[filter_column].astype(str).str.lower().str.startswith(filter_value.lower(), na=False)]
            
            elif filter_condition == "ends with":
                filtered_df = display_df[display_df[filter_column].astype(str).str.lower().str.endswith(filter_value.lower(), na=False)]
            
            # Update treeview with filtered data
            self.clear_treeview()
            
            for i, row in filtered_df.iterrows():
                values = [row[col] for col in self.selected_columns]
                self.tree.insert("", tk.END, text=str(i), values=values)
            
            # Show count of filtered rows
            messagebox.showinfo("Filter Applied", f"Filter applied. {len(filtered_df)} rows match the criteria.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error applying filter: {str(e)}")
    
    def clear_filter(self):
        if self.current_df is None or not self.selected_columns:
            return
        
        # Reset filter fields
        self.filter_column_combobox.set("")
        self.filter_condition_combobox.set("")
        self.filter_value_entry.delete(0, tk.END)
        
        # Reload data
        self.view_data()
    
    def apply_sort(self):
        if self.current_df is None or not self.selected_columns:
            return
        
        sort_column = self.sort_column_combobox.get()
        sort_order = self.sort_order_combobox.get()
        
        if not sort_column or not sort_order:
            messagebox.showinfo("Info", "Please select a column and sort order")
            return
        
        try:
            # Get current data from treeview (which might be filtered)
            current_data = []
            for item in self.tree.get_children():
                item_values = self.tree.item(item)["values"]
                current_data.append(item_values)
            
            # Create DataFrame from current treeview data
            if current_data:
                current_df = pd.DataFrame(current_data, columns=self.selected_columns)
                
                # Apply sorting to the current data (which might be filtered)
                ascending = True if sort_order == "Ascending" else False
                sorted_df = current_df.sort_values(by=sort_column, ascending=ascending)
                
                # Update treeview with sorted data
                self.clear_treeview()
                
                for i, row in sorted_df.iterrows():
                    values = [row[col] for col in self.selected_columns]
                    self.tree.insert("", tk.END, text=str(i), values=values)
            else:
                messagebox.showinfo("Info", "No data to sort")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error applying sort: {str(e)}")
    
    def clear_sort(self):
        if self.current_df is None or not self.selected_columns:
            return
        
        # Reset sort fields
        self.sort_column_combobox.set("")
        self.sort_order_combobox.set("")
        
        # Reload data
        self.view_data()
    
    def clear_treeview(self):
        # Clear all items from treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def export_data(self):
        if not self.tree.get_children():
            messagebox.showinfo("Info", "No data to export")
            return
        
        try:
            # Generate default filename based on filters
            default_filename = self.generate_export_filename()
            
            # Ask for save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=default_filename
            )
            
            if not file_path:
                return
            
            # Get data from treeview
            data = []
            for item in self.tree.get_children():
                values = self.tree.item(item)["values"]
                data.append(values)
            
            # Create DataFrame
            export_df = pd.DataFrame(data, columns=self.selected_columns)
            
            # Export based on file extension
            if file_path.endswith('.csv'):
                export_df.to_csv(file_path, index=False)
            else:
                export_df.to_excel(file_path, index=False)
            
            messagebox.showinfo("Success", f"Data exported successfully to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting data: {str(e)}")
    
    def generate_export_filename(self):
        """Generate a short descriptive filename based on current filters"""
        base_name = os.path.splitext(os.path.basename(self.excel_file_path))[0] if self.excel_file_path else "export"
        sheet_name = self.selected_sheet if self.selected_sheet else ""
        
        # Get filter info
        filter_col = self.filter_column_combobox.get()
        filter_cond = self.filter_condition_combobox.get()
        filter_val = self.filter_value_entry.get()
        
        # Get sort info
        sort_col = self.sort_column_combobox.get()
        sort_order = self.sort_order_combobox.get()
        
        filename_parts = [base_name]
        
        # Add sheet info if available
        if sheet_name:
            # Shorten sheet name if too long
            if len(sheet_name) > 10:
                sheet_name = sheet_name[:8] + ".."
            filename_parts.append(sheet_name)
        
        # Add filter info if available
        if filter_col and filter_cond and filter_val:
            # Shorten column name if too long
            if len(filter_col) > 8:
                filter_col = filter_col[:6] + ".."
            
            # Shorten condition
            cond_map = {
                "equals": "eq",
                "contains": "cont",
                "greater than": "gt",
                "less than": "lt",
                "starts with": "sw",
                "ends with": "ew"
            }
            short_cond = cond_map.get(filter_cond, filter_cond[:2])
            
            # Shorten value if too long
            if len(filter_val) > 8:
                filter_val = filter_val[:6] + ".."
            
            filter_str = f"{filter_col}-{short_cond}-{filter_val}"
            filename_parts.append(filter_str)
        
        # Add sort info if available
        if sort_col and sort_order:
            # Shorten column name
            if len(sort_col) > 8:
                sort_col = sort_col[:6] + ".."
            
            # Shorten order
            short_order = "asc" if sort_order == "Ascending" else "desc"
            
            sort_str = f"{sort_col}-{short_order}"
            filename_parts.append(sort_str)
        
        # Join parts with underscores
        result = "_".join(filename_parts)
        
        # Sanitize filename (remove invalid characters)
        result = ''.join(c for c in result if c.isalnum() or c in ' -_.')
        
        # Add date stamp for uniqueness
        from datetime import datetime
        date_stamp = datetime.now().strftime("%m%d")
        result += f"_{date_stamp}"
        
        return result
    
    def save_column_config(self):
        if not self.selected_columns:
            messagebox.showinfo("Info", "No columns selected to save")
            return
        
        try:
            # Ask for configuration name
            config_name = simpledialog.askstring("Save Configuration", "Enter a name for this column configuration:")
            if not config_name:
                return
            
            # Sanitize filename
            config_name = ''.join(c for c in config_name if c.isalnum() or c in ' -_')
            
            # Create config data
            config_data = {
                'sheet_name': self.selected_sheet,
                'columns': self.selected_columns
            }
            
            # Save to file
            config_file = os.path.join(self.configs_dir, f"{config_name}.json")
            with open(config_file, 'w') as f:
                json.dump(config_data, f)
            
            messagebox.showinfo("Success", f"Column configuration '{config_name}' saved successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving configuration: {str(e)}")
    
    def load_column_config(self):
        try:
            # Get list of available configurations
            config_files = [f for f in os.listdir(self.configs_dir) if f.endswith('.json')]
            
            if not config_files:
                messagebox.showinfo("Info", "No saved configurations found")
                return
            
            # Create a simple dialog to select a configuration
            config_window = tk.Toplevel(self.root)
            config_window.title("Load Column Configuration")
            config_window.geometry("300x400")
            config_window.transient(self.root)
            config_window.grab_set()
            
            # Label
            ttk.Label(config_window, text="Select a configuration:").pack(pady=10)
            
            # Listbox for configurations
            config_listbox = tk.Listbox(config_window, width=40, height=15)
            config_listbox.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
            
            # Populate listbox
            config_names = [os.path.splitext(f)[0] for f in config_files]
            for name in config_names:
                config_listbox.insert(tk.END, name)
            
            # Load button
            def on_load():
                if not config_listbox.curselection():
                    messagebox.showinfo("Info", "Please select a configuration")
                    return
                
                selected_idx = config_listbox.curselection()[0]
                selected_name = config_names[selected_idx]
                config_file = os.path.join(self.configs_dir, f"{selected_name}.json")
                
                try:
                    with open(config_file, 'r') as f:
                        config_data = json.load(f)
                    
                    sheet_name = config_data.get('sheet_name')
                    columns = config_data.get('columns', [])
                    
                    # Check if current sheet matches saved configuration
                    if sheet_name != self.selected_sheet:
                        if not messagebox.askyesno("Warning", 
                                                 f"This configuration was saved for sheet '{sheet_name}', but " 
                                                 f"current sheet is '{self.selected_sheet}'. Continue anyway?"):
                            return
                    
                    # Check if all columns exist in current sheet
                    missing_columns = [col for col in columns if col not in self.columns]
                    if missing_columns:
                        if not messagebox.askyesno("Warning", 
                                                 f"Some columns in this configuration don't exist in the current sheet: "
                                                 f"{', '.join(missing_columns)}. Continue with available columns?"):
                            return
                        
                        # Remove missing columns
                        columns = [col for col in columns if col in self.columns]
                    
                    # Apply configuration
                    self.selected_columns = []
                    self.selected_columns_listbox.delete(0, tk.END)
                    
                    for col in columns:
                        if col in self.columns:  # Double-check column exists
                            self.selected_columns.append(col)
                            self.selected_columns_listbox.insert(tk.END, col)
                    
                    config_window.destroy()
                    
                    # If columns were loaded, view the data
                    if self.selected_columns:
                        self.view_data()
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Error loading configuration: {str(e)}")
            
            ttk.Button(config_window, text="Load Selected Configuration", command=on_load).pack(pady=10)
            ttk.Button(config_window, text="Cancel", command=config_window.destroy).pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading configurations: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUtilityApp(root)
    root.mainloop()