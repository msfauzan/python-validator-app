import sqlite3
import json
import logging
import sys
from tkinter import messagebox
import os

# Konfigurasi logging
logging.basicConfig(
    filename="app.log",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

def load_config(config_file="config.json"):
    """
    Memuat konfigurasi dari file JSON.

    Args:
        config_file (str): Path ke file konfigurasi.

    Returns:
        dict: Konfigurasi yang dimuat dari file JSON, atau None jika gagal.
    """
    try:
        with open(config_file, "r") as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        log_error(f"File konfigurasi '{config_file}' tidak ditemukan.")
        show_error_message(f"File konfigurasi '{config_file}' tidak ditemukan.")
    except json.JSONDecodeError:
        log_error(f"File konfigurasi '{config_file}' tidak valid.")
        show_error_message(f"File konfigurasi '{config_file}' tidak valid.")
    return None

def log_error(message):
    """Mencatat pesan error ke log file dan menampilkan messagebox."""
    logging.error(message)
    messagebox.showerror("Error", message)

def show_error_message(message):
    """Menampilkan error message box."""
    messagebox.showerror("Error", message)

config = load_config()

# Inisialisasi variabel-variabel dari config
DATABASE_NAME = config.get("database", {}).get("name", "reference_data.db") if config else "reference_data.db"
CATEGORIES = config.get("ui", {}).get("categories", ["B0", "C0", "C9", "D0", "E0", "F1", "F2", "Z9"]) if config else ["B0", "C0", "C9", "D0", "E0", "F1", "F2", "Z9"]
N1_STT_CODES = config.get("validation", {}).get("n1_stt_codes", ["1NNN", "1000", "1901", "1902", "1903", "1904", "1905", "1911", "1912", "1906", "1907", "2NNN", "2000", "2901", "2902", "2903", "2904", "2905", "2911", "2912", "2906", "2907"]) if config else ["1NNN", "1000", "1901", "1902", "1903", "1904", "1905", "1911", "1912", "1906", "1907", "2NNN", "2000", "2901", "2902", "2903", "2904", "2905", "2911", "2912", "2906", "2907"]
FUZZY_MATCH_THRESHOLD = config.get("validation", {}).get("fuzzy_match_threshold", 0.9) if config else 0.9
ICON_PATH = config.get("ui", {}).get("icon_path", "icon.ico") if config else "icon.ico"

def create_database():
    """Membuat database dan tabel-tabel yang diperlukan."""
    try:
        if os.path.exists(DATABASE_NAME):
            backup_file = f"{DATABASE_NAME}.bak"
            try:
                os.rename(DATABASE_NAME, backup_file)
            except OSError:
                log_error(f"Failed to create backup of existing database")
                return False

        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            
            # Add transactions
            cursor.execute("BEGIN")
            
            try:
                # Create tables with better constraints
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS ref_mapping_penerima (
                        keyword TEXT PRIMARY KEY NOT NULL,
                        category TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)

                # Add similar improvements to other table creation statements
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS ref_mapping_pembayar (
                        keyword TEXT PRIMARY KEY NOT NULL,
                        category TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS bank_codes (
                        code TEXT PRIMARY KEY NOT NULL,
                        name TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)

                cursor.execute("COMMIT")
                return True
                
            except sqlite3.Error:
                cursor.execute("ROLLBACK")
                raise
                
    except sqlite3.Error as e:
        log_error(f"Error creating database: {e}")
        return False

def insert_initial_data():
    """Memasukkan data awal ke dalam tabel."""
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()

            # Data mapping kategori (penerima dan pembayar)
            # Data ini sebaiknya dipindahkan ke file JSON terpisah atau config
            initial_mapping = {
                "Kedutaan": "B0",
                "Kementerian": "B0",
                "DEPKEU": "B0",
                "KEMENKEU": "B0",
                "DITJEN": "B0",
                "MINISTRY": "B0",
                "AMBASADA": "B0",
                "CONSULATE GENERAL": "B0",
                "EMBASSY": "B0",
                "GOVERNMENT": "B0",
                "Bank Indonesia": "C0",
                "Federal Reserve": "C0",
                "The FED": "C0",
                "Bank Negara Malaysia": "C0",
                "Bangko Sentral ng Pilipinas": "C0",
                "Monetary Authority of Singapore": "C0",
                "Bank of Thailand": "C0",
                "Bank of Japan": "C0",
                "BPD": "C9",
                "Bank Pembangunan Daerah": "C9",
                "Indonesia exim bank": "D0",
                "insurance": "D0",
                "reinsurance": "D0",
                "asuransi": "D0",
                "reasuransi": "D0",
                "leasing": "D0",
                "broker": "D0",
                "multifinance": "D0",
                "AON": "D0",
                "sekuritas": "D0",
                "securities": "D0",
                "finance": "D0",
                "PT": "E0",
                "LTD": "E0",
                "PTE LTD": "E0",
                "Perum": "E0",
                "Pertamina": "E0",
                "Kaltim Prima Coal": "E0",
                "Wilmar Nabati": "E0",
                "Musim Mas": "E0",
                "Multimas Nabati": "E0",
                "Asian Development Bank": "F1",
                "ADB": "F1",
                "Islamic Development Bank": "F1",
                "International Monetary Fund": "F1",
                "IMF": "F1",
                "World Bank": "F1",
                "WB": "F1",
                "United Nation": "F2",
                "UN": "F2",
                "WHO": "F2",
                "UNHCR": "F2",
                "FAO": "F2",
                "Food and Agriculture Organization": "F2",
                "International Civil Aviation Organization": "F2",
                "ICAO": "F2",
                "International Atomic Energy Agency": "F2",
                "IAEA": "F2",
                "International Fund for Agricultural Development": "F2",
                "IFAD": "F2",
                "International Labour Organization": "F2",
                "ILO": "F2",
                "International Maritime Organization": "F2",
                "IMO": "F2",
                "International Telecommunication Union": "F2",
                "UNESCO": "F2",
                "UNIDO": "F2",
                "UPU": "F2",
                "World Health Organization": "F2",
                "WHO": "F2",
                "WIPO": "F2",
                "WMO": "F2",
                "UNWTO": "F2",
                "Koperasi": "Z9",
                "University": "Z9",
                "Universitas": "Z9",
                "Hospital": "Z9",
                "Rumah sakit": "Z9",
                "Sekolah": "Z9",
                "School": "Z9",
                "Institut": "Z9",
                "Institute": "Z9",
                "Yayasan": "Z9",
                "Lembaga": "Z9",
                "Perkumpulan": "Z9",
                "Gereja": "Z9",
                "Church": "Z9",
                "Organisasi": "Z9",
                "Foundation": "Z9",
            }
            for keyword, category in initial_mapping.items():
                cursor.execute(
                    "INSERT INTO ref_mapping_penerima (keyword, category) VALUES (?, ?)",
                    (keyword, category),
                )
                cursor.execute(
                    "INSERT INTO ref_mapping_pembayar (keyword, category) VALUES (?, ?)",
                    (keyword, category),
                )

            # Data kode bank
            # Data ini sebaiknya dipindahkan ke file JSON terpisah atau config
            initial_bank_codes = {
                "222": "AAA",
                "333": "BBB",
            }
            for code, name in initial_bank_codes.items():
                cursor.execute(
                    "INSERT INTO bank_codes (code, name) VALUES (?, ?)", (code, name)
                )

            conn.commit()
    except sqlite3.Error as e:
        log_error(f"Error inserting initial data: {e}")

def get_mapping_data(table_name):
    """Get mapping data with validation."""
    valid_tables = ['ref_mapping_penerima', 'ref_mapping_pembayar']
    
    if table_name not in valid_tables:
        log_error(f"Invalid table name: {table_name}")
        return {}
        
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            conn.row_factory = sqlite3.Row  # Enable dictionary access
            cursor = conn.cursor()
            
            cursor.execute(f"SELECT keyword, category FROM {table_name}")
            return {row['keyword']: row['category'] for row in cursor.fetchall()}
            
    except sqlite3.Error as e:
        log_error(f"Database error: {e}")
        return {}

def get_bank_codes():
    """
    Mengambil data kode bank.

    Returns:
        dict: Data kode bank dalam bentuk dictionary {code: name}.
    """
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT code, name FROM bank_codes")
            data = {row[0]: row[1] for row in cursor.fetchall()}
            return data
    except sqlite3.Error as e:
        log_error(f"Error getting bank codes: {e}")
        return {}

def add_mapping_data(table_name, keyword, category):
    """
    Menambahkan data mapping ke tabel.

    Args:
        table_name (str): Nama tabel.
        keyword (str): Keyword.
        category (str): Category.

    Returns:
        bool: True jika berhasil, False jika gagal.
    """
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"INSERT INTO {table_name} (keyword, category) VALUES (?, ?)",
                (keyword, category),
            )
            conn.commit()
            return True
    except sqlite3.IntegrityError:
        log_error(
            f"Error adding mapping data: Keyword '{keyword}' already exists in '{table_name}'."
        )
        return False
    except sqlite3.Error as e:
        log_error(
            f"Error adding mapping data: {e} (Keyword: {keyword}, Category: {category}, Table: {table_name})"
        )
        return False

def update_mapping_data(table_name, keyword, category):
    """
    Mengupdate data mapping di tabel.

    Args:
        table_name (str): Nama tabel.
        keyword (str): Keyword.
        category (str): Category.

    Returns:
        bool: True jika berhasil, False jika gagal.
    """
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute(
                f"UPDATE {table_name} SET category = ? WHERE keyword = ?",
                (category, keyword),
            )
            conn.commit()
            return True
    except sqlite3.Error as e:
        log_error(
            f"Error updating mapping data: {e} (Keyword: {keyword}, Category: {category}, Table: {table_name})"
        )
        return False

def delete_mapping_data(table_name, keyword):
    """
    Menghapus data mapping dari tabel.

    Args:
        table_name (str): Nama tabel.
        keyword (str): Keyword yang akan dihapus.

    Returns:
        bool: True jika berhasil, False jika gagal.
    """
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM {table_name} WHERE keyword = ?", (keyword,))
            conn.commit()
            return True
    except sqlite3.Error as e:
        log_error(
            f"Error deleting mapping data: {e} (Keyword: {keyword}, Table: {table_name})"
        )
        return False

def add_bank_code(code, name):
    """
    Menambahkan data kode bank.

    Args:
        code (str): Kode bank (harus 3 digit angka).
        name (str): Nama bank.

    Returns:
        bool: True jika berhasil, False jika gagal.
    """
    try:
        if not code.isdigit() or len(code) != 3:
            raise ValueError("Kode bank harus berupa 3 digit angka.")

        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO bank_codes (code, name) VALUES (?, ?)", (code, name)
            )
            conn.commit()
            return True
    except sqlite3.IntegrityError:
        log_error(f"Error adding bank code: Kode bank '{code}' sudah ada.")
        return False
    except ValueError as e:
        log_error(f"Invalid bank code: {e} (Code: {code}, Name: {name})")
        return False
    except sqlite3.Error as e:
        log_error(f"Error adding bank code: {e} (Code: {code}, Name: {name})")
        return False

def update_bank_code(code, name):
    """
    Mengupdate data kode bank.

    Args:
        code (str): Kode bank (harus 3 digit angka).
        name (str): Nama bank.

    Returns:
        bool: True jika berhasil, False jika gagal.
    """
    try:
        if not code.isdigit() or len(code) != 3:
            raise ValueError("Kode bank harus berupa 3 digit angka.")

        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE bank_codes SET name = ? WHERE code = ?", (name, code)
            )
            conn.commit()
            return True
    except ValueError as e:
        log_error(f"Invalid bank code: {e} (Code: {code}, Name: {name})")
        return False
    except sqlite3.Error as e:
        log_error(f"Error updating bank code: {e} (Code: {code}, Name: {name})")
        return False

def delete_bank_code(code):
    """
    Menghapus data kode bank.

    Args:
        code (str): Kode bank yang akan dihapus.

    Returns:
        bool: True jika berhasil, False jika gagal.
    """
    try:
        with sqlite3.connect(DATABASE_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM bank_codes WHERE code = ?", (code,))
            conn.commit()
            return True
    except sqlite3.Error as e:
        log_error(f"Error deleting bank code: {e} (Code: {code})")
        return False

# Panggil fungsi ini untuk membuat database (cukup sekali saja)
# create_database()

# Panggil fungsi ini untuk insert data (cukup sekali saja, setelah create_database)
# insert_initial_data()