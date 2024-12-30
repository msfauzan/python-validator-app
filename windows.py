import tkinter as tk
import ttkbootstrap as ttkb
from tkinter import messagebox, filedialog
import db_utils

class ManageMappingWindow(ttkb.Toplevel):
    """Window untuk mengelola mapping."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Manage Mapping")
        self.sort_order = {"column": "Keyword", "direction": "asc"}
        self.minsize(600, 600)
        self.resizable(True, True)
        self.app = parent  # Store reference to parent App instance
        self.create_widgets()
        self.populate_treeview()

    def create_widgets(self):
        """Membuat widgets untuk ManageMappingWindow."""
        main_frame = ttkb.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        search_frame = ttkb.Frame(main_frame)
        search_frame.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E), pady=(0, 10))

        ttkb.Label(search_frame, text="Search:").pack(side=ttkb.LEFT, padx=(0, 5))
        self.search_entry = ttkb.Entry(search_frame)
        self.search_entry.pack(side=ttkb.LEFT, expand=True, fill=ttkb.X)
        self.search_entry.bind("<Return>", self.search_mapping)

        search_button = ttkb.Button(
            search_frame, text="Search", command=self.search_mapping
        )
        search_button.pack(side=ttkb.LEFT, padx=(5, 0))

        reset_button = ttkb.Button(
            search_frame, text="Reset", command=self.reset_search
        )
        reset_button.pack(side=ttkb.LEFT, padx=(5, 0))

        self.tree = ttkb.Treeview(
            main_frame, columns=("Keyword", "Category"), show="headings"
        )
        self.tree.grid(row=1, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        self.tree.heading(
            "Keyword",
            text="Keyword",
            command=lambda: self.sort_by_column("Keyword"),
        )
        self.tree.heading(
            "Category",
            text="Category",
            command=lambda: self.sort_by_column("Category"),
        )

        self.tree.column("Keyword", width=200)
        self.tree.column("Category", width=100, anchor="center")

        scrollbar = ttkb.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview
        )
        scrollbar.grid(row=1, column=1, sticky=(ttkb.N, ttkb.S))
        self.tree.configure(yscrollcommand=scrollbar.set)

        ttkb.Label(main_frame, text="Keyword:").grid(
            row=2, column=0, sticky=(ttkb.W), padx=5, pady=5
        )
        self.keyword_entry = ttkb.Entry(main_frame)
        self.keyword_entry.grid(row=3, column=0, sticky=(ttkb.W, ttkb.E), padx=5, pady=5)

        ttkb.Label(main_frame, text="Category:").grid(
            row=4, column=0, sticky=(ttkb.W), padx=5, pady=5
        )
        self.category_combobox = ttkb.Combobox(main_frame, values=db_utils.CATEGORIES)
        self.category_combobox.grid(row=5, column=0, sticky=(ttkb.W, ttkb.E), padx=5, pady=5)

        button_frame = ttkb.Frame(main_frame)
        button_frame.grid(row=6, column=0, pady=10)

        ttkb.Button(button_frame, text="Add", command=self.add_mapping).grid(
            row=0, column=0, padx=5
        )
        ttkb.Button(button_frame, text="Update", command=self.update_mapping).grid(
            row=0, column=1, padx=5
        )
        ttkb.Button(button_frame, text="Delete", command=self.delete_mapping).grid(
            row=0, column=2, padx=5
        )
        ttkb.Button(button_frame, text="Import SQL", command=self.import_sql).grid(
            row=0, column=3, padx=5
        )
        ttkb.Button(button_frame, text="Close", command=self.destroy).grid(
            row=0, column=4, padx=5
        )
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

    def reset_search(self):
        """Reset pencarian dan menampilkan semua data."""
        self.search_entry.delete(0, tk.END)
        self.populate_treeview()

    def populate_treeview(self):
        """Mengisi Treeview dengan data mapping."""
        if not self.tree.winfo_exists():
            return  # Jika Treeview sudah tidak ada, keluar dari fungsi

        self.tree.delete(*self.tree.get_children())
        mapping_data = db_utils.get_mapping_data("ref_mapping_penerima")

        sorted_data = sorted(
            mapping_data.items(),
            key=lambda item: item[0 if self.sort_order["column"] == "Keyword" else 1],
            reverse=(self.sort_order["direction"] == "desc"),
        )

        for keyword, category in sorted_data:
            self.tree.insert("", "end", values=(keyword, category))

    def add_mapping(self):
        """Menambahkan data mapping."""
        keyword = self.keyword_entry.get()
        category = self.category_combobox.get()

        if keyword and category:
            if db_utils.add_mapping_data(
                "ref_mapping_penerima", keyword, category
            ) and db_utils.add_mapping_data("ref_mapping_pembayar", keyword, category):
                self.show_success_message("Mapping added successfully!")
                self.populate_treeview()
                self.keyword_entry.delete(0, tk.END)
                self.category_combobox.set("")
                # Update validator reference data
                if hasattr(self.app, 'validator'):
                    self.app.validator.reload_reference_data()
            else:
                self.show_error_message(
                    "Failed to add mapping. Keyword might already exist."
                )
        else:
            self.show_error_message("Please enter both keyword and category.")

    def update_mapping(self):
        """Mengupdate data mapping."""
        selected_item = self.tree.selection()
        if selected_item:
            keyword = self.tree.item(selected_item, "values")[0]
            category = self.category_combobox.get()
            if category:
                db_utils.update_mapping_data("ref_mapping_penerima", keyword, category)
                db_utils.update_mapping_data(
                    "ref_mapping_pembayar", keyword, category
                )
                self.show_success_message("Mapping updated successfully!")
                self.populate_treeview()
                self.keyword_entry.delete(0, tk.END)
                self.category_combobox.set("")
                # Update validator reference data
                if hasattr(self.app, 'validator'):
                    self.app.validator.reload_reference_data()
            else:
                self.show_error_message("Please select a category.")
        else:
            self.show_error_message("Please select an item to update.")

    def delete_mapping(self):
        """Menghapus data mapping."""
        selected_item = self.tree.selection()
        if selected_item:
            keyword = self.tree.item(selected_item, "values")[0]
            if messagebox.askyesno(
                "Confirm Delete",
                f"Are you sure you want to delete the mapping for '{keyword}'?",
            ):
                db_utils.delete_mapping_data("ref_mapping_penerima", keyword)
                db_utils.delete_mapping_data("ref_mapping_pembayar", keyword)
                self.show_success_message("Mapping deleted successfully!")
                self.populate_treeview()
                # Update validator reference data
                if hasattr(self.app, 'validator'):
                    self.app.validator.reload_reference_data()
        else:
            self.show_error_message("Please select an item to delete.")

    def sort_by_column(self, column):
        """Mengurutkan data berdasarkan kolom."""
        if (
            self.sort_order["column"] == column
            and self.sort_order["direction"] == "asc"
        ):
            self.sort_order["direction"] = "desc"
        else:
            self.sort_order["column"] = column
            self.sort_order["direction"] = "asc"
        self.populate_treeview()

    def search_mapping(self, event=None):
        """Mencari mapping berdasarkan keyword."""
        search_term = self.search_entry.get().lower()
        self.tree.delete(*self.tree.get_children())
        mapping_data = db_utils.get_mapping_data("ref_mapping_penerima")

        for keyword, category in mapping_data.items():
            if search_term in keyword.lower() or search_term in category.lower():
                self.tree.insert("", "end", values=(keyword, category))

    def show_success_message(self, message):
        """Menampilkan success message."""
        messagebox.showinfo("Success", message)

    def show_error_message(self, message):
        """Menampilkan error message."""
        messagebox.showerror("Error", message)

    def import_sql(self):
        """Handle import SQL file."""
        try:
            file_path = filedialog.askopenfilename(
                title="Select SQL File",
                filetypes=[("SQL files", "*.sql"), ("All files", "*.*")]
            )
            
            if file_path:
                if messagebox.askyesno(
                    "Confirm Import",
                    "This will execute SQL commands on the database. Are you sure you want to continue?"
                ):
                    success, message = db_utils.execute_sql_file(file_path)
                    if success:
                        messagebox.showinfo("Success", message)
                        self.populate_treeview()  # Refresh data after import
                        # Update validator reference data if app exists
                        if hasattr(self.app, 'validator'):
                            self.app.validator.reload_reference_data()
                    else:
                        messagebox.showerror("Error", message)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import SQL: {str(e)}")
            db_utils.log_error(f"Error in import_sql: {str(e)}")

class ManageBankCodesWindow(ttkb.Toplevel):
    """Window untuk mengelola bank codes."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Manage Bank Codes")
        self.sort_order = {"column": "Code", "direction": "asc"}
        self.minsize(600, 600)
        self.resizable(True, True)
        self.create_widgets()
        self.populate_treeview()

    def create_widgets(self):
        """Membuat widgets untuk ManageBankCodesWindow."""
        main_frame = ttkb.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        # Add search frame
        search_frame = ttkb.Frame(main_frame)
        search_frame.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E), pady=(0, 10))

        ttkb.Label(search_frame, text="Search:").pack(side=ttkb.LEFT, padx=(0, 5))
        self.search_entry = ttkb.Entry(search_frame)
        self.search_entry.pack(side=ttkb.LEFT, expand=True, fill=ttkb.X)
        self.search_entry.bind("<Return>", self.search_bank_codes)

        search_button = ttkb.Button(
            search_frame, text="Search", command=self.search_bank_codes
        )
        search_button.pack(side=ttkb.LEFT, padx=(5, 0))

        reset_button = ttkb.Button(
            search_frame, text="Reset", command=self.reset_search
        )
        reset_button.pack(side=ttkb.LEFT, padx=(5, 0))

        # Treeview
        self.tree = ttkb.Treeview(
            main_frame, columns=("Code", "Name"), show="headings"
        )
        self.tree.grid(row=1, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        self.tree.heading(
            "Code", text="Code", command=lambda: self.sort_by_column("Code")
        )
        self.tree.heading(
            "Name", text="Name", command=lambda: self.sort_by_column("Name")
        )

        self.tree.column("Code", width=100, anchor="center")
        self.tree.column("Name", width=200)

        scrollbar = ttkb.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview
        )
        scrollbar.grid(row=1, column=1, sticky=(ttkb.N, ttkb.S))
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Entry widgets
        ttkb.Label(main_frame, text="Code:").grid(row=2, column=0, sticky=ttkb.W, padx=5, pady=5)
        self.code_entry = ttkb.Entry(main_frame)
        self.code_entry.grid(row=3, column=0, sticky=(ttkb.W, ttkb.E), padx=5, pady=5)

        ttkb.Label(main_frame, text="Name:").grid(row=4, column=0, sticky=ttkb.W, padx=5, pady=5)
        self.name_entry = ttkb.Entry(main_frame)
        self.name_entry.grid(row=5, column=0, sticky=(ttkb.W, ttkb.E), padx=5, pady=5)

        # Buttons
        button_frame = ttkb.Frame(main_frame)
        button_frame.grid(row=6, column=0, pady=10)

        ttkb.Button(button_frame, text="Add", command=self.add_bank_code).grid(row=0, column=0, padx=5)
        ttkb.Button(button_frame, text="Update", command=self.update_bank_code).grid(row=0, column=1, padx=5)
        ttkb.Button(button_frame, text="Delete", command=self.delete_bank_code).grid(row=0, column=2, padx=5)
        ttkb.Button(button_frame, text="Import SQL", command=self.import_sql).grid(
            row=0, column=3, padx=5
        )
        ttkb.Button(button_frame, text="Close", command=self.destroy).grid(row=0, column=4, padx=5)

        # Configure grid weights
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

    def reset_search(self):
        """Reset pencarian dan menampilkan semua data."""
        self.search_entry.delete(0, tk.END)
        self.populate_treeview()

    def populate_treeview(self):
        """Mengisi Treeview dengan data bank codes."""
        self.tree.delete(*self.tree.get_children())
        bank_codes = db_utils.get_bank_codes()
        
        sorted_data = sorted(
            bank_codes.items(),
            key=lambda item: item[0 if self.sort_order["column"] == "Code" else 1],
            reverse=(self.sort_order["direction"] == "desc"),
        )

        for code, name in sorted_data:
            self.tree.insert("", "end", values=(code, name))

    def search_bank_codes(self, event=None):
        """Mencari bank codes berdasarkan keyword."""
        search_term = self.search_entry.get().lower()
        self.tree.delete(*self.tree.get_children())
        bank_codes = db_utils.get_bank_codes()

        for code, name in bank_codes.items():
            if search_term in code.lower() or search_term in name.lower():
                self.tree.insert("", "end", values=(code, name))

    def sort_by_column(self, column):
        """Mengurutkan data berdasarkan kolom."""
        if self.sort_order["column"] == column and self.sort_order["direction"] == "asc":
            self.sort_order["direction"] = "desc"
        else:
            self.sort_order["column"] = column
            self.sort_order["direction"] = "asc"
        self.populate_treeview()

    def add_bank_code(self):
        """Menambahkan data bank code."""
        code = self.code_entry.get()
        name = self.name_entry.get()

        if code and name:
            if db_utils.add_bank_code(code, name):
                messagebox.showinfo("Success", "Bank code added successfully!")
                self.populate_treeview()
                self.code_entry.delete(0, tk.END)
                self.name_entry.delete(0, tk.END)
            else:
                messagebox.showerror(
                    "Error", "Failed to add bank code. Code might already exist."
                )
        else:
            messagebox.showerror("Error", "Please enter both code and name.")

    def update_bank_code(self):
        """Mengupdate data bank code."""
        selected_item = self.tree.selection()
        if selected_item:
            code = self.tree.item(selected_item, "values")[0]
            name = self.name_entry.get()
            if name:
                db_utils.update_bank_code(code, name)
                messagebox.showinfo("Success", "Bank code updated successfully!")
                self.populate_treeview()
                self.code_entry.delete(0, tk.END)
                self.name_entry.delete(0, tk.END)
            else:
                messagebox.showerror("Error", "Please enter a name.")
        else:
            messagebox.showerror("Error", "Please select an item to update.")

    def delete_bank_code(self):
        """Menghapus data bank code."""
        selected_item = self.tree.selection()
        if selected_item:
            code = self.tree.item(selected_item, "values")[0]
            if messagebox.askyesno(
                "Confirm Delete",
                f"Are you sure you want to delete the bank code '{code}'?",
            ):
                db_utils.delete_bank_code(code)
                messagebox.showinfo("Success", "Bank code deleted successfully!")
                self.populate_treeview()
        else:
            messagebox.showerror("Error", "Please select an item to delete.")

    def import_sql(self):
        """Handle import SQL file."""
        try:
            file_path = filedialog.askopenfilename(
                title="Select SQL File",
                filetypes=[("SQL files", "*.sql"), ("All files", "*.*")]
            )
            
            if file_path:
                if messagebox.askyesno(
                    "Confirm Import",
                    "This will execute SQL commands on the database. Are you sure you want to continue?"
                ):
                    success, message = db_utils.execute_sql_file(file_path)
                    if success:
                        messagebox.showinfo("Success", message)
                        self.populate_treeview()  # Refresh data after import
                        # Update validator reference data if app exists
                        if hasattr(self.app, 'validator'):
                            self.app.validator.reload_reference_data()
                    else:
                        messagebox.showerror("Error", message)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import SQL: {str(e)}")
            db_utils.log_error(f"Error in import_sql: {str(e)}")

class ManageStatusMappingWindow(ttkb.Toplevel):
    """Window untuk mengelola mapping status."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Manage Status Mapping")
        self.sort_order = {"column": "Keyword", "direction": "asc"}
        self.minsize(600, 600)
        self.resizable(True, True)
        self.create_widgets()
        self.populate_treeview()

    def create_widgets(self):
        """Membuat widgets untuk ManageStatusMappingWindow."""
        main_frame = ttkb.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        # Add search frame
        search_frame = ttkb.Frame(main_frame)
        search_frame.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E), pady=(0, 10))

        ttkb.Label(search_frame, text="Search:").pack(side=ttkb.LEFT, padx=(0, 5))
        self.search_entry = ttkb.Entry(search_frame)
        self.search_entry.pack(side=ttkb.LEFT, expand=True, fill=ttkb.X)
        self.search_entry.bind("<Return>", self.search_status)

        search_button = ttkb.Button(
            search_frame, text="Search", command=self.search_status
        )
        search_button.pack(side=ttkb.LEFT, padx=(5, 0))

        reset_button = ttkb.Button(
            search_frame, text="Reset", command=self.reset_search
        )
        reset_button.pack(side=ttkb.LEFT, padx=(5, 0))

        # Treeview
        self.tree = ttkb.Treeview(
            main_frame, columns=("Keyword", "Status"), show="headings"
        )
        self.tree.grid(row=1, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        self.tree.heading(
            "Keyword", text="Keyword", command=lambda: self.sort_by_column("Keyword")
        )
        self.tree.heading(
            "Status", text="Status", command=lambda: self.sort_by_column("Status")
        )

        self.tree.column("Keyword", width=200)
        self.tree.column("Status", width=100, anchor="center")

        scrollbar = ttkb.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview
        )
        scrollbar.grid(row=1, column=1, sticky=(ttkb.N, ttkb.S))
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Entry widgets
        ttkb.Label(main_frame, text="Keyword:").grid(row=2, column=0, sticky=ttkb.W, padx=5, pady=5)
        self.keyword_entry = ttkb.Entry(main_frame)
        self.keyword_entry.grid(row=3, column=0, sticky=(ttkb.W, ttkb.E), padx=5, pady=5)

        ttkb.Label(main_frame, text="Status:").grid(row=4, column=0, sticky=ttkb.W, padx=5, pady=5)
        self.status_entry = ttkb.Entry(main_frame)
        self.status_entry.grid(row=5, column=0, sticky=(ttkb.W, ttkb.E), padx=5, pady=5)

        # Buttons
        button_frame = ttkb.Frame(main_frame)
        button_frame.grid(row=6, column=0, pady=10)

        ttkb.Button(button_frame, text="Add", command=self.add_mapping).grid(row=0, column=0, padx=5)
        ttkb.Button(button_frame, text="Delete", command=self.delete_mapping).grid(row=0, column=1, padx=5)
        ttkb.Button(button_frame, text="Import SQL", command=self.import_sql).grid(
            row=0, column=2, padx=5
        )
        ttkb.Button(button_frame, text="Close", command=self.destroy).grid(row=0, column=3, padx=5)

        # Configure grid weights
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

    def reset_search(self):
        """Reset pencarian dan menampilkan semua data."""
        self.search_entry.delete(0, tk.END)
        self.populate_treeview()

    def populate_treeview(self):
        """Mengisi Treeview dengan data status mapping."""
        self.tree.delete(*self.tree.get_children())
        status_mapping = db_utils.get_status_mapping()
        
        data_list = []
        for keyword, statuses in status_mapping.items():
            for status in statuses:
                data_list.append((keyword, status))
        
        sorted_data = sorted(
            data_list,
            key=lambda item: item[0 if self.sort_order["column"] == "Keyword" else 1],
            reverse=(self.sort_order["direction"] == "desc"),
        )

        for keyword, status in sorted_data:
            self.tree.insert("", "end", values=(keyword, status))

    def search_status(self, event=None):
        """Mencari status mapping berdasarkan keyword."""
        search_term = self.search_entry.get().lower()
        self.tree.delete(*self.tree.get_children())
        status_mapping = db_utils.get_status_mapping()

        for keyword, statuses in status_mapping.items():
            for status in statuses:
                if search_term in keyword.lower() or search_term in status.lower():
                    self.tree.insert("", "end", values=(keyword, status))

    def sort_by_column(self, column):
        """Mengurutkan data berdasarkan kolom."""
        if self.sort_order["column"] == column and self.sort_order["direction"] == "asc":
            self.sort_order["direction"] = "desc"
        else:
            self.sort_order["column"] = column
            self.sort_order["direction"] = "asc"
        self.populate_treeview()

    def add_mapping(self):
        keyword = self.keyword_entry.get().strip().upper()
        status = self.status_entry.get().strip().upper()
        
        if keyword and status:
            if db_utils.add_status_mapping(keyword, status):
                messagebox.showinfo("Success", "Status mapping added successfully!")
                self.populate_treeview()
                self.keyword_entry.delete(0, "end")
                self.status_entry.delete(0, "end")
            else:
                messagebox.showerror("Error", "Failed to add status mapping")
        else:
            messagebox.showerror("Error", "Please enter both keyword and status")

    def delete_mapping(self):
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            keyword, status = item["values"]
            
            if messagebox.askyesno("Confirm", f"Delete mapping for {keyword} - {status}?"):
                if db_utils.delete_status_mapping(keyword, status):
                    messagebox.showinfo("Success", "Status mapping deleted successfully!")
                    self.populate_treeview()
                else:
                    messagebox.showerror("Error", "Failed to delete status mapping")
        else:
            messagebox.showerror("Error", "Please select a mapping to delete")

    def import_sql(self):
        """Handle import SQL file."""
        try:
            file_path = filedialog.askopenfilename(
                title="Select SQL File",
                filetypes=[("SQL files", "*.sql"), ("All files", "*.*")]
            )
            
            if file_path:
                if messagebox.askyesno(
                    "Confirm Import",
                    "This will execute SQL commands on the database. Are you sure you want to continue?"
                ):
                    success, message = db_utils.execute_sql_file(file_path)
                    if success:
                        messagebox.showinfo("Success", message)
                        self.populate_treeview()  # Refresh data after import
                        # Update validator reference data if app exists
                        if hasattr(self.app, 'validator'):
                            self.app.validator.reload_reference_data()
                    else:
                        messagebox.showerror("Error", message)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import SQL: {str(e)}")
            db_utils.log_error(f"Error in import_sql: {str(e)}")