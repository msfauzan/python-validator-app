import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import db_utils
import json
import ttkbootstrap as ttkb
from data_validator import DataValidator
from tool_tip import ToolTip
from windows import ManageMappingWindow, ManageBankCodesWindow, ManageStatusMappingWindow
import subprocess

# Load konfigurasi
config = db_utils.load_config()

# Constants dari config, dengan fallback value jika config tidak ada
DATABASE_NAME = config.get("database", {}).get("name", "reference_data.db") if config else "reference_data.db"
CATEGORIES = config.get("ui", {}).get("categories", ["B0", "C0", "C9", "D0", "E0", "F1", "F2", "Z9"]) if config else ["B0", "C0", "C9", "D0", "E0", "F1", "F2", "Z9"]
N1_STT_CODES = config.get("validation", {}).get("n1_stt_codes", ["1NNN", "1000", "1901", "1902", "1903", "1904", "1905", "1911", "1912", "1906", "1907", "2NNN", "2000", "2901", "2902", "2903", "2904", "2905", "2911", "2912", "2906", "2907"]) if config else ["1NNN", "1000", "1901", "1902", "1903", "1904", "1905", "1911", "1912", "1906", "1907", "2NNN", "2000", "2901", "2902", "2903", "2904", "2905", "2911", "2912", "2906", "2907"]
FUZZY_MATCH_THRESHOLD = config.get("validation", {}).get("fuzzy_match_threshold", 0.9) if config else 0.9
ICON_PATH = config.get("ui", {}).get("icon_path", "icon.ico") if config else "icon.ico"

# Constants untuk nama kolom
COL_NAMA_PENERIMA = "nama_penerima"
COL_KATEGORI_PENERIMA = "kategori_penerima"
COL_NAMA_PEMBAYAR = "nama_pembayar"
COL_KATEGORI_PEMBAYAR = "kategori_pembayar"
COL_KODE_BANK = "cKdBank"
COL_STATUS_PENERIMA = "status_penerima"
COL_STATUS_PEMBAYAR = "status_pembayar"
COL_STT = "stt"
COL_TAHUN = "tahun"
COL_BULAN = "bulan"

class App:
    def __init__(self, root):
        if not isinstance(root, ttkb.Window):
            raise ValueError("root harus instance dari ttkb.Window")
        self.root = root
        self.style = ttkb.Style("cosmo")
        self.root.title("LLD-Bank Data Clarification Tool")
        
        try:
            self.validator = DataValidator(FUZZY_MATCH_THRESHOLD)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal inisialisasi DataValidator: {str(e)}")
            self.root.destroy()
            return
        
        # Move recent_files initialization before create_interface
        self.recent_files = self.load_recent_files()
            
        self.create_interface()
        self.set_icon()
        self.create_menu()
        self.create_status_bar()
        self.bind_shortcuts()
        
        # Tambahkan recent files
        self.recent_files = self.load_recent_files()
        self.last_validated_file = None  # Initialize variable to store the last validated file
        self.last_validated_folder = None  # Tambahkan variabel folder

    def set_icon(self):
        """Set icon untuk window."""
        try:
            self.root.iconbitmap(ICON_PATH)
        except tk.TclError:
            db_utils.log_error(f"Gagal load icon: {ICON_PATH}")

    def create_menu(self):
        """Membuat menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open File (Ctrl+O)", command=self.process_file)
        file_menu.add_separator()
        file_menu.add_command(label="Recent Validated Folder", command=self.open_validated_folder, state="disabled")
        self.open_validated_folder_menu = file_menu  # Store reference to enable later
        file_menu.add_command(label="Exit (Alt+F4)", command=self.root.quit)

        # Database Menu
        db_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Database", menu=db_menu)
        db_menu.add_command(label="Manage Mapping (Ctrl+M)", command=self.manage_mapping)
        db_menu.add_command(label="Manage Bank Codes (Ctrl+B)", command=self.manage_bank_codes)
        db_menu.add_command(label="Manage Status (Ctrl+S)", command=self.manage_status)  # Add new menu item

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def create_status_bar(self):
        """Membuat status bar."""
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=999, column=0, sticky=(tk.W, tk.E))

    def bind_shortcuts(self):
        """Menambahkan keyboard shortcuts."""
        self.root.bind('<Control-o>', lambda e: self.process_file())
        self.root.bind('<Control-m>', lambda e: self.manage_mapping())
        self.root.bind('<Control-b>', lambda e: self.manage_bank_codes())
        self.root.bind('<Control-s>', lambda e: self.manage_status())  # Add new shortcut

    def create_interface(self):
        """Membuat interface yang lebih user-friendly."""
        # Gunakan widget dari ttkbootstrap
        main_frame = ttkb.Frame(self.root, padding=15)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        # Welcome Section
        welcome_frame = ttkb.Frame(main_frame)
        welcome_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        # Create a larger, bold title label centered in the window
        title_label = ttkb.Label(
            welcome_frame,
            text="LLD-Bank Data Clarification Tool",
            style="Primary.TLabel",
            font=("Helvetica", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(10, 20), sticky="nsew")
        welcome_frame.grid_columnconfigure(0, weight=1)

        # Quick Actions Frame
        actions_frame = ttkb.Labelframe(main_frame, text="Quick Actions", padding=15)
        actions_frame.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        
        # File Selection dengan preview area
        file_frame = ttkb.Frame(actions_frame)
        file_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.file_button = ttkb.Button(
            file_frame,
            text="Choose Excel File",
            command=self.process_file,
            style="Action.TButton"
        )
        self.file_button.grid(row=0, column=0, padx=5)
        ToolTip(self.file_button, "Klik untuk memilih file Excel yang akan divalidasi")
        
        # Recent Files
        if self.recent_files:
            recent_frame = ttkb.Labelframe(actions_frame, text="File Terakhir", padding=10)
            recent_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
            
            for idx, recent_file in enumerate(self.recent_files[-3:]):  # Show last 3 files
                btn = ttkb.Button(
                    recent_frame,
                    text=os.path.basename(recent_file),
                    command=lambda f=recent_file: self.process_file(f)
                )
                btn.grid(row=idx, column=0, sticky="ew", pady=2)
                ToolTip(btn, f"Buka file: {recent_file}")

        # Progress Section dengan informasi detail
        progress_frame = ttkb.Labelframe(main_frame, text="Processing Status", padding=15)
        progress_frame.grid(row=2, column=0, sticky="ew", pady=(0, 15))

        self.status_label = ttkb.Label(progress_frame, text="Ready to process file")
        self.status_label.grid(row=0, column=0, sticky="w", pady=(0, 5))

        progress_info_frame = ttkb.Frame(progress_frame)
        progress_info_frame.grid(row=1, column=0, sticky="ew")
        
        self.progress_bar = ttkb.Progressbar(
            progress_info_frame, 
            orient="horizontal",
            mode="determinate",
            length=300
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        # Database Management dengan grouping yang lebih baik
        db_frame = ttkb.Labelframe(main_frame, text="Data Setup", padding=15)
        db_frame.grid(row=3, column=0, sticky="ew")

        db_btn_frame = ttkb.Frame(db_frame)
        db_btn_frame.grid(row=0, column=0, pady=5)

        mapping_btn = ttkb.Button(
            db_btn_frame,
            text="Manage Categorical Mapping",
            command=self.manage_mapping,
            style="Action.TButton"
        )
        mapping_btn.grid(row=0, column=0, padx=5)
        ToolTip(mapping_btn, "Kelola data mapping kategori")

        codes_btn = ttkb.Button(
            db_btn_frame,
            text="Manage Bank Codes",
            command=self.manage_bank_codes,
            style="Action.TButton"
        )
        codes_btn.grid(row=0, column=1, padx=5)
        ToolTip(codes_btn, "Kelola data kode bank")

        # Add new button for managing status mappings
        status_btn = ttkb.Button(
            db_btn_frame,
            text="Manage Status Mapping",
            command=self.manage_status,
            style="Action.TButton"
        )
        status_btn.grid(row=0, column=2, padx=5)
        ToolTip(status_btn, "Kelola data mapping status")

        # Help section
        help_frame = ttkb.Frame(main_frame)
        help_frame.grid(row=4, column=0, sticky="ew", pady=(15, 0))
        
        help_btn = ttkb.Button(
            help_frame,
            text="Help (F1)",
            command=self.show_help,
            style="Action.TButton"
        )
        help_btn.grid(row=0, column=0)
        
        # Grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

    def load_recent_files(self):
        """Load daftar file yang terakhir digunakan."""
        try:
            with open("recent_files.json", "r") as f:
                return json.load(f)
        except:
            return []

    def save_recent_files(self, file_path):
        """Simpan file ke daftar recent files."""
        recent_files = self.recent_files
        if file_path in recent_files:
            recent_files.remove(file_path)
        recent_files.append(file_path)
        recent_files = recent_files[-5:]  # Keep last 5 files
        
        try:
            with open("recent_files.json", "w") as f:
                json.dump(recent_files, f)
        except:
            pass
        
        self.recent_files = recent_files

    def show_help(self):
        """Menampilkan panduan penggunaan."""
        help_text = """
Panduan Penggunaan Data Validation Tool

1. Memulai Validasi:
   - Klik tombol "Pilih File Excel" atau tekan Ctrl+O
   - Pilih file Excel yang akan divalidasi
   - Tunggu proses validasi selesai

2. Format Excel yang Didukung:
   - File harus memiliki kolom: nama_penerima, kategori_penerima, dll
   - Kategori yang valid: B0, C0, C9, D0, E0, F1, F2, Z9

3. Shortcuts:
   - Ctrl+O : Buka file
   - Ctrl+M : Kelola mapping
   - Ctrl+B : Kelola kode bank
   - Ctrl+S : Kelola status mapping
   - F1     : Bantuan

4. Hasil Validasi:
   - File hasil akan disimpan dengan suffix "_validated"
   - Sel yang perlu perhatian akan ditandai kuning
   - Hover pada sel untuk melihat saran perubahan
   - Klik File > Recent Validated Folder untuk membuka folder hasil validasi terakhir
        """
        help_window = tk.Toplevel(self.root)
        help_window.title("Bantuan Penggunaan")
        help_window.geometry("600x400")
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=15, pady=15)
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")
        text_widget.pack(fill="both", expand=True)
        
        close_btn = ttkb.Button(
            help_window,
            text="Tutup",
            command=help_window.destroy,
            style="Action.TButton"
        )
        close_btn.pack(pady=10)

    def show_about(self):
        """Menampilkan dialog About."""
        messagebox.showinfo(
            "About",
            "\nVersion 1.0\n\n" +
            "Shortcuts:\n" +
            "Ctrl+O: Open File\n" +
            "Ctrl+M: Manage Mapping\n" +
            "Ctrl+B: Manage Bank Codes"
        )

    def process_file(self):
        """Memproses file dengan visual feedback yang lebih baik."""
        try:
            input_file = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls")],
                initialdir="."
            )

            if not input_file:
                return

            self.status_bar.config(text="Processing file...")
            self.status_label.config(text="Processing...")
            self.progress_bar.start(10)
            self.file_button.config(state="disabled")
            self.root.config(cursor="wait")
            self.root.update()

            try:
                output_file, error_count, validation_results = self.validator.process_file(input_file)
                self.last_validated_file = output_file  # Store the validated file path
                self.last_validated_folder = os.path.dirname(output_file)  # Simpan folder
                self.open_validated_folder_menu.entryconfig("Recent Validated Folder", state="normal")  # Enable menu item
                
                detail_message = "Validation Results:\n\n"
                for result in validation_results:
                    detail_message += f"Row {result['row']}, {result['column']}:\n"
                    detail_message += f"Current: {result['current']}\n"
                    detail_message += f"Suggested: {result['suggested']}\n"
                    detail_message += f"Name: {result['name']}\n"
                    if "bank_code" in result:
                        detail_message += f"Bank code: {result['bank_code']}\n"
                    detail_message += f"Status: {result['status']}\n\n"

                message = f"Validasi selesai!\n\n"
                message += f"Ditemukan {error_count} ketidaksesuaian.\n"
                message += (
                    f"File hasil validasi tersimpan di:\n{output_file}\n\n"
                )
                message += "Lihat detail hasil validasi?"

                if messagebox.askyesno("Success", message):
                    self.show_validation_details(validation_results)

            except PermissionError:
                error_msg = (
                    "Tidak dapat menyimpan file hasil validasi karena file sedang digunakan.\n\n"
                    "Langkah penyelesaian:\n"
                    "1. Tutup file Excel yang sedang terbuka\n"
                    "2. Tutup aplikasi lain yang mungkin sedang menggunakan file\n"
                    "3. Coba proses validasi lagi"
                )
                messagebox.showerror("Permission Denied", error_msg)
                
            except Exception as e:
                messagebox.showerror("Error", f"Gagal memproses file: {str(e)}")
                
            finally:
                self.status_bar.config(text="Ready")
                self.status_label.config(text="Ready")
                self.progress_bar.stop()
                self.file_button.config(state="normal")
                self.root.config(cursor="")

        except Exception as e:
            self.status_bar.config(text="Error occurred")
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")

    def show_validation_details(self, validation_results):
        """Menampilkan detail hasil validasi di window baru."""
        detail_window = ttkb.Toplevel(self.root) 
        detail_window.title("Validation Details")

        tree = ttkb.Treeview(  
            detail_window,
            columns=(
                "Row",
                "Column",
                "Current",
                "Suggested",
                "Name",
                "Bank Code",
                "Status",
            ),
            show="headings",
        )
        tree.grid(row=0, column=0, sticky=(ttkb.W, ttkb.E, ttkb.N, ttkb.S))

        tree.heading("Row", text="Row")
        tree.heading("Column", text="Column")
        tree.heading("Current", text="Current")
        tree.heading("Suggested", text="Suggested")
        tree.heading("Name", text="Name")
        tree.heading("Bank Code", text="Bank Code")
        tree.heading("Status", text="Status")

        tree.column("Row", width=50, anchor="center")
        tree.column("Column", width=150, anchor="center")
        tree.column("Current", width=100, anchor="center")
        tree.column("Suggested", width=100, anchor="center")
        tree.column("Name", width=200)
        tree.column("Bank Code", width=100, anchor="center")
        tree.column("Status", width=100, anchor="center")

        for result in validation_results:
            tree.insert(
                "",
                "end",
                values=(
                    result["row"],
                    result["column"],
                    result["current"],
                    result["suggested"],
                    result["name"],
                    result.get("bank_code", "-"),
                    result.get("status", "-"),
                ),
            )

        scrollbar = ttkb.Scrollbar(
            detail_window, orient="vertical", command=tree.yview
        )
        scrollbar.grid(row=0, column=1, sticky=(ttkb.N, ttkb.S))
        tree.configure(yscrollcommand=scrollbar.set)

        button_frame = ttkb.Frame(detail_window)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)

        close_button = ttkb.Button(
            button_frame, text="Close", command=detail_window.destroy
        )
        close_button.pack()

        detail_window.columnconfigure(0, weight=1)
        detail_window.rowconfigure(0, weight=1)

    def manage_mapping(self, event=None):
        """Membuka window Manage Mapping."""
        try:
            mapping_window = ManageMappingWindow(self.root)
            mapping_window.transient(self.root)  # Set parent window
            mapping_window.grab_set()  # Make window modal
            
            # Center the window
            window_width = 600
            window_height = 400
            screen_width = mapping_window.winfo_screenwidth()
            screen_height = mapping_window.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            mapping_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            self.root.wait_window(mapping_window)  # Wait for window to close
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuka Manage Mapping: {str(e)}")
            db_utils.log_error(f"Error in manage_mapping: {str(e)}")

    def manage_bank_codes(self, event=None):
        """Membuka window Manage Bank Codes."""
        try:
            bank_codes_window = ManageBankCodesWindow(self.root)
            bank_codes_window.transient(self.root)  # Set parent window
            bank_codes_window.grab_set()  # Make window modal
            
            # Center the window
            window_width = 500
            window_height = 400
            screen_width = bank_codes_window.winfo_screenwidth()
            screen_height = bank_codes_window.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            bank_codes_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            self.root.wait_window(bank_codes_window)  # Wait for window to close
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuka Manage Bank Codes: {str(e)}")
            db_utils.log_error(f"Error in manage_bank_codes: {str(e)}")

    def manage_status(self, event=None):
        """Membuka window Manage Status Mapping."""
        try:
            status_window = ManageStatusMappingWindow(self.root)
            status_window.transient(self.root)  # Set parent window
            status_window.grab_set()  # Make window modal
            
            # Center the window
            window_width = 600
            window_height = 400
            screen_width = status_window.winfo_screenwidth()
            screen_height = status_window.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            status_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            self.root.wait_window(status_window)  # Wait for window to close
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuka Manage Status: {str(e)}")
            db_utils.log_error(f"Error in manage_status: {str(e)}")

    def open_validated_folder(self):
        """Buka folder hasil validasi."""
        if self.last_validated_folder and os.path.exists(self.last_validated_folder):
            try:
                os.startfile(self.last_validated_folder)  # Windows
            except AttributeError:
                subprocess.call(["xdg-open", self.last_validated_folder])  # Linux
        else:
            messagebox.showwarning("Warning", "No validated folder found or folder does not exist.")

    def add_new_category(self, keyword, category):
        """Contoh penambahan kategori baru ke DB, lalu reload."""
        success = db_utils.add_mapping_data("ref_mapping_penerima", keyword, category)
        if success:
            self.validator.reload_reference_data()
            messagebox.showinfo("Info", f"Berhasil menambah kategori {category} untuk '{keyword}'")

if __name__ == "__main__":
    root = ttkb.Window(themename="cosmo")
    app = App(root)
    root.mainloop()