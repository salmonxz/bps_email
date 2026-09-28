"""
Module for parsing BPS Provincial Contact Excel files.
Supports dynamic header scanning, multi-level merged headers,
safe cell access (immune to index out of bounds), and intelligent
recipient role detection (To, Cc, Bcc) with row-level fallback.
"""

import re
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd


def clean_str(val: Any) -> str:
    """Clean string values, handling NaN, floats, newlines, and extra whitespace."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() == "nan":
        return ""
    # Replace newlines and tabs with spaces, collapse multiple spaces
    s = re.sub(r"[\r\n\t]+", " ", s)
    s = re.sub(r" +", " ", s).strip()
    return s


def safe_get(row: Any, col_idx: Optional[int]) -> str:
    """
    Safely access a cell value from a row without index out-of-bounds errors.
    Guarantees immunity against IndexError when columns are removed or modified.
    Handles pandas Series, lists, tuples, NaN, None, and returns a stripped string.
    """
    if col_idx is None:
        return ""
    try:
        idx = int(col_idx)
    except (ValueError, TypeError):
        return ""
    if idx < 0:
        return ""
    try:
        if hasattr(row, "iloc"):
            if idx < len(row):
                val = row.iloc[idx]
                if pd.notna(val):
                    s = str(val).strip()
                    return "" if s.lower() == "nan" else s
        elif hasattr(row, "__len__") and idx < len(row):
            val = row[idx]
            if pd.notna(val):
                s = str(val).strip()
                return "" if s.lower() == "nan" else s
    except Exception:
        return ""
    return ""


def clean_and_extract_emails(val: Any) -> List[str]:
    """
    Extract and clean all valid emails from a cell.
    Handles multiple emails separated by newlines, commas, semicolons,
    and removes accidental whitespace (e.g. 'muh. putra@bps.go.id' -> 'muh.putra@bps.go.id').
    """
    if val is None or pd.isna(val):
        return []
    s = str(val).strip()
    if not s or s.lower() == "nan" or s == "-":
        return []

    parts = re.split(r"[\r\n;,]+", s)
    emails = []
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    for part in parts:
        cleaned = re.sub(r"\s+", "", part).strip()
        if cleaned and re.match(pattern, cleaned):
            if cleaned not in emails:
                emails.append(cleaned)
    return emails


def clean_email(val: Any) -> str:
    """Clean email string returning comma-separated emails if multiple found, or cleaned single string."""
    emails = clean_and_extract_emails(val)
    if emails:
        return ", ".join(emails)
    s = clean_str(val)
    return re.sub(r"\s+", "", s)


def format_kode_provinsi(val: Any) -> Optional[str]:
    """Format province code to 2-digit zero-padded string (e.g. 11, 01)."""
    s = clean_str(val)
    if not s:
        return None
    # Remove decimal if float representation like '11.0'
    try:
        f = float(s)
        i = int(f)
        if 1 <= i <= 99:
            return f"{i:02d}"
    except ValueError:
        # Check if digits exist
        match = re.search(r"\b(\d{1,2})\b", s)
        if match:
            return f"{int(match.group(1)):02d}"
    return None


def format_no_cell(val: Any) -> str:
    """Format row number cell to integer string, handling floats like 1.0."""
    s = clean_str(val)
    if not s:
        return ""
    try:
        f = float(s)
        i = int(f)
        if f == i:
            return str(i)
    except ValueError:
        pass
    return s


def is_valid_email(email: str) -> bool:
    """Check if an email format is generally valid."""
    if not email:
        return False
    # Simple regex check for email
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    # Support comma/semicolon separated multiple emails
    parts = [p.strip() for p in re.split(r"[,;]", email) if p.strip()]
    return all(re.match(pattern, p) for p in parts) if parts else False


def detect_explicit_email_role(header_text: str) -> Optional[str]:
    """
    Detect explicit 'To', 'Cc', or 'Bcc' marker in header string (case-insensitive).
    Recognizes patterns such as:
    - [To], (To), To:, Kepada: -> 'to'
    - [Cc], (Cc), Cc:, Tembusan: -> 'cc'
    - [Bcc], (Bcc), Bcc:, Blind Cc: -> 'bcc'
    """
    if not header_text:
        return None
    s = header_text.strip().lower()
    # Check BCC first (before CC, since 'bcc' contains 'cc')
    if re.search(r"\[\s*bcc\s*\]|\(\s*bcc\s*\)|\bbcc\s*:|\bblind\s*cc\s*:", s):
        return "bcc"
    if re.search(r"\[\s*cc\s*\]|\(\s*cc\s*\)|\bcc\s*:|\btembusan\s*:", s):
        return "cc"
    if re.search(r"\[\s*to\s*\]|\(\s*to\s*\)|\bto\s*:|\bkepada\s*:", s):
        return "to"
    return None


class BPSExcelParser:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.raw_df: Optional[pd.DataFrame] = None
        self.email_columns: List[Dict[str, Any]] = []
        self.dynamic_tags: List[str] = []
        self.tag_col_map: Dict[str, int] = {}

    def load(self, sheet_name: Optional[Any] = 0) -> List[Dict[str, Any]]:
        """Load and parse the excel file into structured provincial records."""
        # Read raw excel without assuming header row to inspect structure
        df_raw = pd.read_excel(self.filepath, sheet_name=sheet_name, header=None)
        self.raw_df = df_raw

        # Find header rows and column mapping dynamically
        header_info = self._detect_header_structure(df_raw)
        if not header_info:
            raise ValueError(
                "Tidak dapat mendeteksi struktur kolom Excel BPS. "
                "Pastikan tabel memuat kolom 'Kode' dan kolom 'Provinsi'."
            )

        col_map, data_start_row = header_info

        records = []
        kode_idx = col_map.get("kode", -1)
        prov_idx = col_map.get("provinsi", -1)

        for idx in range(data_start_row, len(df_raw)):
            row = df_raw.iloc[idx]

            # Safe read province code
            raw_kode = safe_get(row, kode_idx)
            kode = format_kode_provinsi(raw_kode)

            # Skip row if no valid province code found (title, notes, numbering row, footer)
            if not kode:
                continue

            # Safe read province name
            provinsi = clean_str(safe_get(row, prov_idx))
            if not provinsi:
                continue

            record: Dict[str, Any] = {
                "row_number": idx + 1,
                "kode": kode,
                "provinsi": provinsi,
            }

            # Populate all detected dynamic tags for this row using safe_get
            for tag_key in self.dynamic_tags:
                c_idx = self.tag_col_map.get(tag_key, col_map.get(tag_key, -1))
                cell_val = safe_get(row, c_idx)
                if tag_key in ["no", "nomor"]:
                    record[tag_key] = format_no_cell(cell_val)
                else:
                    record[tag_key] = clean_str(cell_val)

            errors = []

            # Extract data using detected email_columns with safe_get
            for col_info in self.email_columns:
                c_idx = col_info["col_idx"]
                k = col_info["key"]
                name_k = col_info.get("name_key")
                name_idx = col_info.get("name_col_idx", -1)

                cell_val = safe_get(row, c_idx)
                emails = clean_and_extract_emails(cell_val)
                if emails:
                    val_str = ", ".join(emails)
                else:
                    raw_s = clean_str(cell_val)
                    val_str = re.sub(r"\s+", "", raw_s) if raw_s and raw_s != "-" else ""
                record[k] = val_str

                # Read paired name if exists
                if name_k:
                    if name_idx is not None and name_idx >= 0:
                        record[name_k] = clean_str(safe_get(row, name_idx))
                    elif name_k not in record:
                        record[name_k] = ""

                # Check email format validity for non-empty value
                if val_str and not is_valid_email(val_str):
                    col_label = col_info.get("role") or col_info.get("label", k)
                    errors.append(f"Format email {col_label} '{val_str}' tidak valid.")

            # Ensure backward-compatible name fields
            if "nama_kepala" not in record or not record["nama_kepala"]:
                nk_idx = col_map.get("nama_kepala", -1)
                record["nama_kepala"] = clean_str(safe_get(row, nk_idx))
            if "nama_ketua" not in record or not record["nama_ketua"]:
                record["nama_ketua"] = (
                    record.get("nama_tim_prod")
                    or record.get("nama_pjk")
                    or clean_str(safe_get(row, col_map.get("nama_pjk", -1)))
                )
            record["nama_tim_prod"] = record.get("nama_ketua", "")
            record["nama_pjk"] = record.get("nama_ketua", "")
            if "nama_koordinator" in record and not record.get("nama_koor"):
                record["nama_koor"] = record["nama_koordinator"]
            if "nama_staf" not in record or not record["nama_staf"]:
                ns_idx = col_map.get("nama_staf", -1)
                record["nama_staf"] = clean_str(safe_get(row, ns_idx))
            record["nama_provinsi"] = record.get("provinsi", "")
            record["kode_provinsi"] = record.get("kode", "")
            record["kode_provisi"] = record.get("kode", "")
            record["provisi"] = record.get("provinsi", "")
            record["nama_provisi"] = record.get("provinsi", "")
            record["ketua"] = record.get("nama_ketua", "")
            record["koordinator"] = record.get("nama_koordinator", "")
            record["koor"] = record.get("nama_koordinator", "")
            record["staf"] = record.get("nama_staf", "")
            record["staff"] = record.get("nama_staf", "")
            record["nama_staff"] = record.get("nama_staf", "")
            record["kepala"] = record.get("nama_kepala", "")
            record["kepala_prov"] = record.get("nama_kepala", "")
            record["nama_kepala_prov"] = record.get("nama_kepala", "")
            record["nomor"] = record.get("no", "")

            # Check if any email text exists on this row
            all_email_vals = [record.get(c["key"], "").strip() for c in self.email_columns if c["key"] in record]
            has_any_email_text = any(v for v in all_email_vals)

            # Allocation of To, Cc, and Bcc:
            # 1. Group columns by detected role_type ('to', 'cc', 'bcc')
            to_cols = [c for c in self.email_columns if c.get("role_type") == "to"]
            cc_cols = [c for c in self.email_columns if c.get("role_type") == "cc"]
            bcc_cols = [c for c in self.email_columns if c.get("role_type") == "bcc"]

            # If no column was marked as 'to', take the first email column as 'to'
            if not to_cols and self.email_columns:
                to_cols = [self.email_columns[0]]
                cc_cols = [c for c in self.email_columns[1:] if c.get("role_type") != "bcc"]
                bcc_cols = [c for c in self.email_columns[1:] if c.get("role_type") == "bcc"]

            to_email = ""
            to_role = ""
            to_source_key = ""

            # Step 1: Check to_cols for a valid email
            for c in to_cols:
                val = record.get(c["key"], "").strip()
                if val and is_valid_email(val):
                    to_email = val
                    to_role = c.get("role", c.get("label", "To"))
                    to_source_key = c["key"]
                    break

            # Step 2: Fallback to cc_cols if to_email is empty on this row
            if not to_email:
                for c in cc_cols:
                    val = record.get(c["key"], "").strip()
                    if val and is_valid_email(val):
                        to_email = val
                        to_role = c.get("role", c.get("label", "Cc (Fallback)"))
                        to_source_key = c["key"]
                        break

            # Step 3: Fallback to bcc_cols if still empty on this row
            if not to_email:
                for c in bcc_cols:
                    val = record.get(c["key"], "").strip()
                    if val and is_valid_email(val):
                        to_email = val
                        to_role = c.get("role", c.get("label", "Bcc (Fallback)"))
                        to_source_key = c["key"]
                        break

            if not to_email:
                if has_any_email_text and not errors:
                    errors.append("Tidak ditemukan alamat email yang valid.")
                elif not has_any_email_text:
                    errors.append("Seluruh kolom email kosong.")

            # Allocation of Cc from remaining to_cols and cc_cols (excluding whoever was promoted to To)
            to_emails_set = set(clean_and_extract_emails(to_email))
            cc_list = []
            for c in to_cols + cc_cols:
                k = c["key"]
                if k == to_source_key:
                    continue
                val = record.get(k, "").strip()
                if val and is_valid_email(val):
                    for em in clean_and_extract_emails(val):
                        if em not in cc_list and em not in to_emails_set:
                            cc_list.append(em)

            # Allocation of Bcc from bcc_cols (excluding whoever was promoted to To)
            bcc_list = []
            for c in bcc_cols:
                k = c["key"]
                if k == to_source_key:
                    continue
                val = record.get(k, "").strip()
                if val and is_valid_email(val):
                    for em in clean_and_extract_emails(val):
                        if em not in bcc_list and em not in to_emails_set and em not in cc_list:
                            bcc_list.append(em)

            record["to_email"] = to_email
            record["to_role"] = to_role
            record["to_source"] = to_source_key
            record["cc_email"] = ", ".join(cc_list)
            record["bcc_email"] = ", ".join(bcc_list)
            record["email_pjk"] = record["cc_email"]

            # Backward-compatible email fields
            if "email_kepala" not in record or not record["email_kepala"]:
                record["email_kepala"] = to_email

            record["is_valid"] = len(errors) == 0 and bool(to_email)
            record["errors"] = errors
            records.append(record)

        return records

    def _detect_header_structure(
        self, df: pd.DataFrame
    ) -> Optional[Tuple[Dict[str, Any], int]]:
        """
        Dynamically scan rows to detect table header structure and column positions.
        Zero hardcoded column indices: completely dynamic and resilient against column additions/deletions.
        """
        max_scan_rows = min(12, len(df))
        num_cols = len(df.columns)

        for r_idx in range(max_scan_rows):
            row_vals = [safe_get(df.iloc[r_idx], c).lower() for c in range(num_cols)]

            has_kode = any("kode" in v and "pos" not in v for v in row_vals)
            has_prov = any(
                ("prov" in v or "wilayah" in v or "daerah" in v) and "kode" not in v
                for v in row_vals
            )

            if not (has_kode and has_prov):
                continue

            # Check if this row or next row has sub-headers (multi-level header)
            is_multilevel = False
            lvl1_vals = [safe_get(df.iloc[r_idx], c) for c in range(num_cols)]
            lvl2_vals = [""] * num_cols

            if r_idx + 1 < len(df):
                next_row_vals = [safe_get(df.iloc[r_idx + 1], c).lower() for c in range(num_cols)]
                next_has_nama_or_mail = any(
                    "nama" in v or "name" in v or "mail" in v or "email" in v for v in next_row_vals
                )
                if next_has_nama_or_mail:
                    is_multilevel = True
                    lvl2_vals = [safe_get(df.iloc[r_idx + 1], c) for c in range(num_cols)]

            # Forward-fill merged cells in Level 1 headers
            lvl1_filled = []
            last_l1 = ""
            for v in lvl1_vals:
                if v.strip():
                    last_l1 = v.strip()
                lvl1_filled.append(last_l1)

            # Determine data start row
            header_end_row = (r_idx + 2) if is_multilevel else (r_idx + 1)
            data_start_row = header_end_row

            # Check if the row immediately after header is a numbering row (e.g. (1), (2), -1, -2, etc.)
            if data_start_row < len(df):
                sample_cell = safe_get(df.iloc[data_start_row], 1) or safe_get(df.iloc[data_start_row], 2)
                # If cell is negative or enclosed in parentheses like '(1)' or '-', skip it
                if re.match(r"^(\(-\d+\)|-\d+|\(\d+\)|-)$", sample_cell.strip()):
                    data_start_row += 1

            # Map columns dynamically using detected headers
            col_map = self._map_columns_from_headers(lvl1_filled, lvl2_vals)
            if col_map and "kode" in col_map and "provinsi" in col_map:
                return col_map, data_start_row

        return None

    def _map_columns_from_headers(
        self, lvl1_headers: List[str], lvl2_headers: List[str]
    ) -> Optional[Dict[str, Any]]:
        """
        Map logical field names to column indices dynamically from headers.
        Implements smart role detection (To, Cc, Bcc) and intelligent fallback hierarchy.
        """
        col_map: Dict[str, Any] = {}
        self.dynamic_tags = []
        self.tag_col_map = {}
        num_cols = len(lvl1_headers)

        # 1. Identify Email and Non-Email columns
        email_indices = []
        name_indices = []

        for c in range(num_cols):
            h1 = lvl1_headers[c].strip()
            h2 = lvl2_headers[c].strip()
            comb_lower = f"{h1} {h2}".lower()

            comb = f"{h1} - {h2}" if h1 and h2 else (h1 or h2)
            explicit = detect_explicit_email_role(h1) or detect_explicit_email_role(h2) or detect_explicit_email_role(comb)

            is_email = False
            if explicit is not None:
                is_email = True
            elif any(em in h2.lower() for em in ["email", "e-mail", "mail"]):
                is_email = True
            elif any(em in h1.lower() for em in ["email", "e-mail", "mail"]) and not any(
                nm in h2.lower() for nm in ["nama", "name"]
            ):
                is_email = True

            if is_email:
                email_indices.append(c)
            elif any(nm in h2.lower() for nm in ["nama", "name"]) or (
                any(nm in h1.lower() for nm in ["nama", "pejabat", "petugas"]) and "provinsi" not in comb_lower
            ):
                name_indices.append(c)

        # 2. Map standard non-email identification columns
        detected_non_email = []
        for c in range(num_cols):
            if c in email_indices:
                continue

            h1 = lvl1_headers[c].strip()
            h2 = lvl2_headers[c].strip()
            comb = f"{h1} - {h2}" if h1 and h2 else (h1 or h2)
            comb_lower = comb.lower()

            tag_key = None
            if "kode" in comb_lower and "pos" not in comb_lower and "kode" not in col_map:
                tag_key = "kode"
                col_map["kode"] = c
            elif (
                any(p in comb_lower for p in ["provinsi", "prov", "wilayah", "daerah"])
                and "kode" not in comb_lower
                and "provinsi" not in col_map
            ):
                tag_key = "provinsi"
                col_map["provinsi"] = c
            elif any(n in comb_lower.split() for n in ["no", "no.", "nomor"]) and "no" not in col_map:
                tag_key = "no"
                col_map["no"] = c
            elif "kepala" in comb_lower and "nama_kepala" not in col_map:
                tag_key = "nama_kepala"
                col_map["nama_kepala"] = c
            elif any(k in comb_lower for k in ["ketua", "tim prod", "pjk"]) and "nama_ketua" not in col_map:
                tag_key = "nama_ketua"
                col_map["nama_ketua"] = c
            elif "koor" in comb_lower and "nama_koordinator" not in col_map:
                tag_key = "nama_koordinator"
                col_map["nama_koordinator"] = c
            elif ("staf" in comb_lower or "staff" in comb_lower) and "nama_staf" not in col_map:
                tag_key = "nama_staf"
                col_map["nama_staf"] = c
            else:
                clean_k = re.sub(r"[^a-zA-Z0-9_]+", "_", comb_lower).strip("_")
                if clean_k and clean_k not in col_map:
                    tag_key = clean_k
                    col_map[clean_k] = c

            if tag_key and tag_key not in detected_non_email:
                detected_non_email.append(tag_key)
                self.tag_col_map[tag_key] = c
                if tag_key == "nama_ketua":
                    self.tag_col_map["ketua"] = c
                elif tag_key == "nama_koordinator":
                    self.tag_col_map["koordinator"] = c
                    self.tag_col_map["koor"] = c
                elif tag_key in ["nama_staf", "nama_staff"]:
                    self.tag_col_map["nama_staff"] = c
                    self.tag_col_map["staff"] = c
                    self.tag_col_map["staf"] = c
                elif tag_key == "nama_kepala":
                    self.tag_col_map["kepala"] = c
                    self.tag_col_map["kepala_prov"] = c
                    self.tag_col_map["nama_kepala_prov"] = c
                elif tag_key == "kode":
                    self.tag_col_map["kode_provinsi"] = c
                    self.tag_col_map["kode_provisi"] = c
                elif tag_key == "provinsi":
                    self.tag_col_map["nama_provinsi"] = c
                    self.tag_col_map["provisi"] = c
                    self.tag_col_map["nama_provisi"] = c

        # 3. Detect email columns with paired names and explicit/fallback roles
        detected_email_cols = []
        for c in email_indices:
            h1 = lvl1_headers[c].strip()
            h2 = lvl2_headers[c].strip()
            comb = f"{h1} - {h2}" if h1 and h2 else (h1 or h2 or f"Email {c}")

            # Check for explicit role markers in Level 1, Level 2, or combined
            explicit_role = (
                detect_explicit_email_role(h1)
                or detect_explicit_email_role(h2)
                or detect_explicit_email_role(comb)
            )

            # Find paired name column (check c - 1 first, then c + 1)
            name_c = -1
            name_label = ""
            name_key = ""

            candidates = [c - 1, c + 1]
            for cand in candidates:
                if 0 <= cand < num_cols and cand not in email_indices:
                    cand_h1 = lvl1_headers[cand].strip()
                    cand_h2 = lvl2_headers[cand].strip()
                    if (
                        "nama" in cand_h2.lower()
                        or "nama" in cand_h1.lower()
                        or cand_h1 == h1
                        or cand in name_indices
                    ):
                        name_c = cand
                        name_label = f"{cand_h1} - {cand_h2 or 'Nama'}" if cand_h1 else (cand_h2 or "Nama")
                        break

            # Determine logical key for this email column
            comb_clean = re.sub(r"(?i)\b(e-?mail|alamat)\b", "", comb).strip(" -_")
            h1_lower = h1.lower()
            if "kepala" in h1_lower:
                key = "email_kepala"
                name_key = "nama_kepala"
            elif "pjk" in h1_lower:
                key = "email_pjk"
                name_key = "nama_pjk"
            elif any(k in h1_lower for k in ["tim prod", "produksi", "ketua"]):
                key = "email_tim_prod"
                name_key = "nama_ketua"
            elif "koor" in h1_lower:
                key = "email_koor"
                name_key = "nama_koordinator"
            elif ("staf" in h1_lower or "staff" in h1_lower):
                key = "email_staf"
                name_key = "nama_staf"
            else:
                clean_name = re.sub(r"[^a-zA-Z0-9_]+", "_", comb_clean.lower()).strip("_")
                key = f"email_{clean_name}" if clean_name else f"email_{c}"
                name_key = f"nama_{clean_name}" if clean_name else f"nama_{c}"

            # If name column is found, ensure name_key and aliases are in tag_col_map
            if name_c >= 0 and name_key:
                if name_key not in self.tag_col_map:
                    self.tag_col_map[name_key] = name_c
                if name_key not in detected_non_email:
                    detected_non_email.append(name_key)
                if name_key == "nama_ketua":
                    self.tag_col_map["ketua"] = name_c
                elif name_key == "nama_koordinator":
                    self.tag_col_map["koordinator"] = name_c
                    self.tag_col_map["koor"] = name_c
                elif name_key in ["nama_staf", "nama_staff"]:
                    self.tag_col_map["nama_staff"] = name_c
                    self.tag_col_map["staff"] = name_c
                    self.tag_col_map["staf"] = name_c
                elif name_key == "nama_kepala":
                    self.tag_col_map["kepala"] = name_c

            detected_email_cols.append({
                "col_idx": c,
                "h1": h1,
                "h2": h2,
                "original_header": comb,
                "clean_header": comb_clean or comb,
                "explicit_role": explicit_role,
                "key": key,
                "name_key": name_key,
                "name_col_idx": name_c,
                "name_label": name_label,
            })

        # 4. Apply Role Assignment: Explicit Tags vs Smart Hierarchy Fallback
        has_any_explicit = any(c["explicit_role"] is not None for c in detected_email_cols)
        n_emails = len(detected_email_cols)

        for idx, col_item in enumerate(detected_email_cols):
            if has_any_explicit:
                if col_item["explicit_role"]:
                    role_type = col_item["explicit_role"]
                elif idx == 0 and not any(c["explicit_role"] == "to" for c in detected_email_cols):
                    role_type = "to"
                elif idx == n_emails - 1 and ("staf" in col_item["h1"].lower() or "arsip" in col_item["h1"].lower()):
                    role_type = "bcc"
                else:
                    role_type = "cc"
            else:
                # Smart Hierarchy Fallback:
                # - 1st valid email column on the left -> To
                # - Last valid email column on the right -> Bcc (for >= 3 cols, or if staf/arsip)
                # - All intermediate columns -> Cc
                if idx == 0:
                    role_type = "to"
                elif idx == n_emails - 1 and n_emails >= 3:
                    role_type = "bcc"
                elif idx == n_emails - 1 and n_emails == 2:
                    h_lower = col_item["h1"].lower()
                    if "staf" in h_lower or "arsip" in h_lower or "bcc" in h_lower:
                        role_type = "bcc"
                    else:
                        role_type = "cc"
                else:
                    role_type = "cc"

            col_item["role_type"] = role_type
            col_item["is_to"] = (role_type == "to")

            # Format user-friendly label without duplicate tags
            display_title = re.sub(
                r"(?i)\[\s*(to|cc|bcc)\s*\]|\(\s*(to|cc|bcc)\s*\)|\b(to|cc|bcc)\s*:",
                "",
                col_item["original_header"],
            ).strip(" -_")
            if not display_title:
                display_title = f"Email {idx + 1}"

            col_item["label"] = f"{display_title} [{role_type.upper()}]"
            col_item["role"] = f"{col_item['h1'] or display_title} ({role_type.upper()})"

        self.email_columns = detected_email_cols
        col_map["email_columns"] = detected_email_cols

        # Map dynamic tags with priority ordering
        priority_order = ["kode", "provinsi", "nama_kepala", "nama_ketua", "nama_koordinator", "nama_staf", "no"]
        ordered_tags = [t for t in priority_order if t in detected_non_email]
        ordered_tags += [t for t in detected_non_email if t not in priority_order]
        self.dynamic_tags = ordered_tags

        return col_map
