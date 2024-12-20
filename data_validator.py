import pandas as pd
from difflib import SequenceMatcher
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.utils.exceptions import InvalidFileException
import re
import os
import db_utils

# Constants untuk nama kolom
COL_NAMA_PENERIMA = "nama_penerima"
COL_KATEGORI_PENERIMA = "kategori_penerima"
COL_NAMA_PEMBAYAR = "nama_pembayar"
COL_KATEGORI_PEMBAYAR = "kategori_pembayar"
COL_KODE_BANK = "cKdBank"
COL_STATUS_PENERIMA = "status_penerima"
COL_STATUS_PEMBAYAR = "status_pembayar"
COL_STT = "stt"
COL_TAHUN = "tahun"  # Add this near other column constants
COL_BULAN = "bulan"  # Add this near other column constants

class DataValidator:
    def __init__(self, fuzzy_match_threshold=0.9):
        self.reference_mapping = {
            "penerima": db_utils.get_mapping_data("ref_mapping_penerima"),
            "pembayar": db_utils.get_mapping_data("ref_mapping_pembayar"),
        }
        self.bank_codes = db_utils.get_bank_codes()
        self.fuzzy_match_threshold = fuzzy_match_threshold

    def validate_bank_code(self, bank_name, bank_code):
        """
        Validasi kesesuaian nama bank dengan sandi bank

        Args:
            bank_name (str): Nama bank.
            bank_code (str): Kode bank.

        Returns:
            bool: True jika sesuai, False jika tidak.
        """
        if pd.isna(bank_code) or str(bank_code).strip() == '':
            return False

        try:
            bank_code = str(bank_code).zfill(3)
        except (ValueError, AttributeError):
            return False

        if bank_code not in self.bank_codes:
            return False

        reference_bank_name = self.bank_codes[bank_code].upper()
        bank_name = str(bank_name).upper() if not pd.isna(bank_name) else ""

        def clean_bank_name(name):
            name = re.sub(r"[^\w\s]", " ", name)
            name = " ".join(name.split())
            common_words = ["PT", "BANK", "PERSERO", "TBK"]
            for word in common_words:
                name = name.replace(f" {word} ", " ")
                if name.startswith(f"{word} "):
                    name = name[len(word) :].strip()
                if name.endswith(f" {word}"):
                    name = name[: -len(word)].strip()
            return name.strip()

        clean_bank = clean_bank_name(bank_name)
        clean_reference = clean_bank_name(reference_bank_name)

        return clean_reference in clean_bank or clean_bank in clean_reference

    def is_same_bank(self, bank1, bank2):
        """
        Memeriksa apakah dua nama bank merujuk ke bank yang sama

        Args:
            bank1 (str): Nama bank pertama.
            bank2 (str): Nama bank kedua.

        Returns:
            bool: True jika sama, False jika tidak.
        """
        if pd.isna(bank1) or pd.isna(bank2) or str(bank1).strip() == '' or str(bank2).strip() == '':
            return False

        try:
            clean_bank1 = self.clean_bank_name(str(bank1))
            clean_bank2 = self.clean_bank_name(str(bank2))
        except (ValueError, AttributeError):
            return False

        def clean_bank_name(name):
            name = str(name).upper()
            name = re.sub(r"[^\w\s]", " ", name)
            name = " ".join(name.split())
            common_words = [
                "PT",
                "BANK",
                "PERSERO",
                "TBK",
                "INCORPORATION",
                "CORPORATION",
                "LTD",
                "LIMITED",
                "INCORPORATED",
            ]
            for word in common_words:
                name = name.replace(f" {word} ", " ")
                if name.startswith(f"{word} "):
                    name = name[len(word) :].strip()
                if name.endswith(f" {word}"):
                    name = name[: -len(word)].strip()
            return name

        clean_bank1 = clean_bank_name(bank1)
        clean_bank2 = clean_bank_name(bank2)

        if clean_bank1 == clean_bank2:
            return True

        def extract_bank_code(name):
            words = name.split()
            bank_codes = [
                word for word in words if len(word) in [2, 3, 4] and word.isalpha()
            ]
            return bank_codes[0] if bank_codes else None

        code1 = extract_bank_code(clean_bank1)
        code2 = extract_bank_code(clean_bank2)

        if code1 and code2 and code1 == code2:
            return True

        if (
            (clean_bank1 in clean_bank2 or clean_bank2 in clean_bank1)
            and (len(clean_bank1) == 0 or len(clean_bank2) == 0)
        ):
            return True

        return (
            SequenceMatcher(None, clean_bank1, clean_bank2).ratio()
            > self.fuzzy_match_threshold
        )

    def is_n1_category(self, stt_value):
        """
        Memeriksa apakah nilai stt termasuk dalam kategori N1.

        Args:
            stt_value (str): Nilai stt.

        Returns:
            bool: True jika termasuk kategori N1, False jika tidak.
        """
        return str(stt_value) in ["1NNN", "1000", "1901", "1902", "1903", "1904", "1905", "1911", "1912", "1906", "1907", "2NNN", "2000", "2901", "2902", "2903", "2904", "2905", "2911", "2912", "2906", "2907"]

    def process_file(self, input_file):
        """
        Memproses file Excel dan melakukan validasi.

        Args:
            input_file (str): Path ke file Excel input.

        Returns:
            tuple: (output_file, error_count, validation_results)
                - output_file (str): Path ke file Excel output.
                - error_count (int): Jumlah error yang ditemukan.
                - validation_results (list): List hasil validasi.
        """
        try:
            # Validasi file exists dan extension
            if not os.path.exists(input_file):
                raise FileNotFoundError("File tidak ditemukan")
            
            # Check if trying to validate an already validated file
            if "_validated" in input_file:
                raise ValueError(
                    "File ini merupakan hasil validasi.\n"
                    "Silakan pilih file asli (tanpa suffix '_validated')"
                )

            df = pd.read_excel(input_file)
            
            # Get year and month from dataframe
            if "tahun" not in df.columns or "bulan" not in df.columns:
                raise ValueError(f"File Excel harus memiliki kolom tahun dan bulan")
            
            # Get first non-null values for year and month
            tahun = df["tahun"].dropna().iloc[0] if not df["tahun"].isna().all() else ""
            bulan = df["bulan"].dropna().iloc[0] if not df["bulan"].isna().all() else ""
            
            # Validate year and month
            try:
                tahun = int(tahun)
                bulan = int(bulan)
                if not (2000 <= tahun <= 2100) or not (1 <= bulan <= 12):
                    raise ValueError
            except (ValueError, TypeError):
                raise ValueError("Nilai tahun atau bulan tidak valid")

            # Generate output filename with period
            output_file = os.path.splitext(input_file)[0] + f"_{tahun}_{bulan:02d}_validated.xlsx"
            
            # Check if output file is currently open
            try:
                with open(output_file, 'a+b') as f:
                    pass
            except PermissionError:
                raise PermissionError(
                    f"File hasil validasi '{os.path.basename(output_file)}' sedang terbuka.\n"
                    "Silakan tutup file tersebut terlebih dahulu."
                )

            if not input_file.lower().endswith(('.xls', '.xlsx')):
                raise ValueError("Format file harus Excel (.xls atau .xlsx)")

            df = pd.read_excel(input_file)
            
            # Validasi minimal rows
            if len(df) == 0:
                raise ValueError("File Excel kosong")
                
            required_columns = [
                "nama_penerima",
                "kategori_penerima",
                "nama_pembayar",
                "kategori_pembayar",
                "stt",
            ]
            if not all(col in df.columns for col in required_columns):
                raise Exception(
                    f"File Excel tidak memiliki kolom yang dibutuhkan: {', '.join(required_columns)}"
                )

            output_df = df.copy()
            validation_results = []

            for idx, row in df.iterrows():
                is_penerima_bank = "BANK" in str(row["nama_penerima"]).upper()
                is_pembayar_bank = "BANK" in str(row["nama_pembayar"]).upper()

                is_valid_bank_code_penerima = self.validate_bank_code(
                    row["nama_penerima"], row.get("cKdBank", "")
                )
                is_valid_bank_code_pembayar = self.validate_bank_code(
                    row["nama_pembayar"], row.get("cKdBank", "")
                )

                penerima_status = str(row.get("status_penerima", "")).upper()
                pembayar_status = str(row.get("status_pembayar", "")).upper()

                is_same_bank = self.is_same_bank(
                    row["nama_penerima"], row["nama_pembayar"]
                )

                suggested_category_penerima = None
                suggested_category_pembayar = None

                if self.is_n1_category(row.get("stt", "")):
                    suggested_category_penerima = "N1"
                    suggested_category_pembayar = "N1"
                    penerima_status = "N1"
                    pembayar_status = "N1"
                else:
                    if (
                        row["nama_penerima"] == row["nama_pembayar"]
                        and penerima_status == pembayar_status
                    ):
                        suggested_category_pembayar = "I0"
                    else:
                        for central_bank_keyword in self.reference_mapping["penerima"]:
                            if central_bank_keyword.upper() == "C0":
                                continue
                            if (
                                central_bank_keyword.upper()
                                in str(row["nama_penerima"]).upper()
                                and self.reference_mapping["penerima"][
                                    central_bank_keyword
                                ]
                                == "C0"
                            ):
                                suggested_category_penerima = "C0"
                                is_penerima_bank = False
                                break

                        for central_bank_keyword in self.reference_mapping["pembayar"]:
                            if central_bank_keyword.upper() == "C0":
                                continue
                            if (
                                central_bank_keyword.upper()
                                in str(row["nama_pembayar"]).upper()
                                and self.reference_mapping["pembayar"][
                                    central_bank_keyword
                                ]
                                == "C0"
                            ):
                                suggested_category_pembayar = "C0"
                                is_pembayar_bank = False
                                break

                        if suggested_category_penerima is None:
                            for q1_bank_keyword in self.reference_mapping["penerima"]:
                                if q1_bank_keyword.upper() == "Q1":
                                    continue
                                if (
                                    q1_bank_keyword.upper()
                                    in str(row["nama_penerima"]).upper()
                                    and self.reference_mapping["penerima"][
                                        q1_bank_keyword
                                    ]
                                    == "Q1"
                                ):
                                    suggested_category_penerima = "Q1"
                                    break

                        if suggested_category_pembayar is None:
                            for q1_bank_keyword in self.reference_mapping["pembayar"]:
                                if q1_bank_keyword.upper() == "Q1":
                                    continue
                                if (
                                    q1_bank_keyword.upper()
                                    in str(row["nama_pembayar"]).upper()
                                    and self.reference_mapping["pembayar"][
                                        q1_bank_keyword
                                    ]
                                    == "Q1"
                                ):
                                    suggested_category_pembayar = "Q1"
                                    break

                        if suggested_category_penerima is None and is_penerima_bank:
                            if is_valid_bank_code_penerima:
                                if penerima_status == "ID":
                                    suggested_category_penerima = "L1"
                                else:
                                    suggested_category_penerima = "L9"
                            else:
                                suggested_category_penerima = "L9"

                        if suggested_category_pembayar is None and is_pembayar_bank:
                            if is_valid_bank_code_pembayar:
                                if pembayar_status == "ID":
                                    suggested_category_pembayar = "L1"
                                else:
                                    suggested_category_pembayar = "L9"
                            else:
                                suggested_category_pembayar = "L9"

                        if (
                            is_penerima_bank
                            and is_pembayar_bank
                            and is_same_bank
                        ):
                            if (
                                is_valid_bank_code_penerima
                                or is_valid_bank_code_pembayar
                            ):
                                if penerima_status != "ID" or pembayar_status != "ID":
                                    if (
                                        suggested_category_penerima != "L1"
                                        and suggested_category_penerima != "Q1"
                                    ):
                                        suggested_category_penerima = "L2"
                                    if (
                                        suggested_category_pembayar != "L1"
                                        and suggested_category_pembayar != "Q1"
                                    ):
                                        suggested_category_pembayar = "L2"

                    if suggested_category_penerima is None and not is_penerima_bank:
                        for (
                            keyword,
                            expected_category,
                        ) in self.reference_mapping["penerima"].items():
                            if keyword.upper() in str(row["nama_penerima"]).upper():
                                if row["kategori_penerima"] != expected_category:
                                    suggested_category_penerima = expected_category
                                break

                    if suggested_category_pembayar is None and not is_pembayar_bank:
                        for (
                            keyword,
                            expected_category,
                        ) in self.reference_mapping["pembayar"].items():
                            if keyword.upper() in str(row["nama_pembayar"]).upper():
                                if row["kategori_pembayar"] != expected_category:
                                    suggested_category_pembayar = expected_category
                                break

                if (
                    suggested_category_penerima
                    and row["kategori_penerima"] != suggested_category_penerima
                ):
                    validation_results.append(
                        {
                            "row": idx + 2,
                            "column": "kategori_penerima",
                            "current": row["kategori_penerima"],
                            "suggested": suggested_category_penerima,
                            "name": row["nama_penerima"],
                            "bank_code": row.get("cKdBank", ""),
                            "status": penerima_status,
                        }
                    )

                if (
                    suggested_category_pembayar
                    and row["kategori_pembayar"] != suggested_category_pembayar
                ):
                    validation_results.append(
                        {
                            "row": idx + 2,
                            "column": "kategori_pembayar",
                            "current": row["kategori_pembayar"],
                            "suggested": suggested_category_pembayar,
                            "name": row["nama_pembayar"],
                            "bank_code": row.get("cKdBank", ""),
                            "status": pembayar_status,
                        }
                    )

            writer = pd.ExcelWriter(output_file, engine="openpyxl")
            output_df.to_excel(writer, index=False)

            worksheet = writer.sheets["Sheet1"]
            red_fill = PatternFill(
                start_color="FFFF00", end_color="FFFF00", fill_type="solid"
            )

            for result in validation_results:
                col_idx = df.columns.get_loc(result["column"]) + 1
                cell = worksheet.cell(row=result["row"], column=col_idx)
                cell.fill = red_fill

                comment_text = f"Suggested category: {result['suggested']}\n"
                comment_text += f"Name: {result['name']}\n"
                if "bank_code" in result:
                    comment_text += f"Bank code: {result['bank_code']}\n"
                comment_text += f"Status: {result['status']}"

                cell.comment = Comment(comment_text, "Validator")

            writer.close()
            return output_file, len(validation_results), validation_results

        except Exception as e:
            raise Exception(f"Error processing file: {str(e)}")