"""
Module for matching regional and national attachment files.
Supports detecting province codes from:
- Suffix / Keyword: 'Tabel Profil IMK 2025 - Provinsi 16.xlsx' -> '16'
- Suffix number: 'Data IMK 2025 - 16', 'Tabel 2025_16', 'Laporan 16' -> '16'
- Prefix: '16_Laporan', '[16] Naskah', '16 Tabel' -> '16'
- Subfolder: 'Provinsi 16/dokumen.pdf', '16_Aceh/laporan.docx' -> '16'
- National: '00_...', '..._00', '00', or containing 'nasional' / 'pusat' -> '00'
"""

import os
import re
from typing import Dict, List, Optional, Tuple, Any


def extract_code_from_string(text: str) -> Optional[str]:
    """
    Extract a 2-digit province code (or '00' for national) from a filename stem or folder name.
    Supports:
    - Suffix / Keyword: 'Tabel Profil IMK 2025 - Provinsi 16' -> '16'
    - Suffix number: 'Data IMK 2025 - 16', 'Tabel 2025_16', 'Laporan 16' -> '16'
    - Prefix: '16_Laporan', '[16] Naskah', '16 Tabel' -> '16'
    - Anywhere with keyword: 'provinsi 16', 'prov 16', 'kode 16', 'kd 16'
    - National: '00_...', '..._00', '00', or containing 'nasional' / 'pusat'
    - Avoids false positives from 4-digit years like 2024, 2025, 2026.
    """
    if not text:
        return None
    s = text.strip()

    # 1. Explicit keyword anywhere: 'provinsi 16', 'prov 16', 'kode 16', 'kd 16'
    m_kw = re.search(r"(?:provinsi|prov|kode|kd)[\s_\-\.]*(\d{1,2})(?!\d)", s, re.IGNORECASE)
    if m_kw:
        return m_kw.group(1).zfill(2)

    # 2. National keywords: 'nasional' or 'pusat' (supporting words, underscores, hyphens, brackets)
    if re.search(r"(?i)(?:^|[\s_\-\.\(\[])(?:nasional|pusat)(?:$|[\s_\-\.\)\]\d])", s):
        return "00"

    # 3. Prefix at the start: e.g. '16_...', '16 ...', '[16]...', '(16)...', '16-...'
    m_start = re.match(r"^(?:\[|\()?(\d{2})(?:\]|\))?(?:[\s_\-\.].*|[a-zA-Z].*|$)", s)
    if m_start:
        return m_start.group(1)

    # 4. Suffix at the end: e.g. '... - 16', '..._16', '...(16)', '... 16'
    # (?<!\d) ensures we do not match the trailing 2 digits of a year like 2025
    m_end = re.search(r"(?<!\d)(\d{1,2})(?:[\)\]])?$", s)
    if m_end:
        return m_end.group(1).zfill(2)

    # 5. Check isolated 2-digit number (ignoring 4-digit years like 2024, 2025, 2026)
    cleaned = re.sub(r"\b(?:19|20)\d{2}\b", "", s)
    matches = re.findall(r"(?<!\d)(\d{2})(?!\d)", cleaned)
    if len(matches) == 1:
        return matches[0]

    return None


class AttachmentEngine:
    def __init__(self, folder_path: str = "", files: Optional[List[str]] = None):
        self.folder_path = folder_path
        self.custom_files: List[str] = files or []
        self.national_files: List[str] = []
        self.provincial_files: Dict[str, List[str]] = {}
        self.unmatched_files: List[str] = []
        self.all_files: List[str] = []
        self.scan_error: Optional[str] = None

    def scan(self) -> Tuple[int, int]:
        """
        Scan directory (recursively) or custom files and categorize into national (00) and provincial ({kode}).
        Returns: (total_files, matched_files_count)
        """
        self.national_files.clear()
        self.provincial_files.clear()
        self.unmatched_files.clear()
        self.all_files.clear()
        self.scan_error = None

        candidate_files = []

        if self.custom_files:
            candidate_files = [f for f in self.custom_files if os.path.isfile(f)]
        elif self.folder_path:
            if not os.path.exists(self.folder_path):
                self.scan_error = f"Direktori tidak ditemukan: {self.folder_path}"
                return 0, 0

            if not os.path.isdir(self.folder_path):
                self.scan_error = f"Path bukan sebuah direktori: {self.folder_path}"
                return 0, 0

            try:
                # Recursive scan through subdirectories
                for root, _, filenames in os.walk(self.folder_path):
                    for fname in filenames:
                        candidate_files.append(os.path.join(root, fname))
            except Exception as e:
                self.scan_error = f"Gagal membaca direktori: {e}"
                return 0, 0
        else:
            return 0, 0

        matched_count = 0
        for full_path in candidate_files:
            if not os.path.isfile(full_path):
                continue

            self.all_files.append(full_path)
            prefix = self._extract_code(full_path)

            if prefix:
                if prefix == "00":
                    self.national_files.append(full_path)
                    matched_count += 1
                else:
                    if prefix not in self.provincial_files:
                        self.provincial_files[prefix] = []
                    self.provincial_files[prefix].append(full_path)
                    matched_count += 1
            else:
                self.unmatched_files.append(full_path)

        return len(self.all_files), matched_count

    def _extract_code(self, full_path: str) -> Optional[str]:
        """
        Extract a 2-digit province code (or '00' for national) from:
        1. Filename stem (e.g. 'Tabel Profil IMK 2025 - Provinsi 16.xlsx', '16_Laporan.pdf', '00_Pedoman.docx')
        2. Subfolder path segments (e.g. 'Provinsi 16/dokumen.pdf', '16_Aceh/data.xlsx')
        3. Root directory itself if its name has code (e.g. folder 'Provinsi 16' or '16_Aceh')
        4. Immediate parent folder if custom files
        """
        # 1. Check filename stem (without extension)
        fname_stem = os.path.splitext(os.path.basename(full_path))[0]
        code = extract_code_from_string(fname_stem)
        if code:
            return code

        # 2. Check parent directories relative to root folder (innermost to outermost)
        if self.folder_path and os.path.isdir(self.folder_path):
            try:
                rel_dir = os.path.relpath(os.path.dirname(full_path), self.folder_path)
                if rel_dir and rel_dir != ".":
                    segments = [s.strip() for s in rel_dir.replace("\\", "/").split("/") if s.strip()]
                    for seg in reversed(segments):
                        seg_code = extract_code_from_string(seg)
                        if seg_code:
                            return seg_code
            except Exception:
                pass

            # 3. Check the root folder itself
            root_name = os.path.basename(os.path.normpath(self.folder_path))
            root_code = extract_code_from_string(root_name)
            if root_code:
                return root_code

        # 4. Check immediate parent folder for custom files
        parent_name = os.path.basename(os.path.dirname(full_path))
        if parent_name:
            parent_code = extract_code_from_string(parent_name)
            if parent_code:
                return parent_code

        return None

    def get_attachments_for_province(self, kode: str) -> List[str]:
        """
        Returns all attachments for a specific province code:
        All national files (00) + all files prefixed with {kode}.
        Preserves natural order (National first, then Provincial).
        """
        kode_padded = str(kode).strip().zfill(2)
        attachments = list(self.national_files)

        if kode_padded in self.provincial_files:
            attachments.extend(self.provincial_files[kode_padded])

        return attachments

    def get_summary_for_provinces(
        self, codes: List[str]
    ) -> Dict[str, Dict[str, Any]]:
        """Return preview attachment summary for a list of province codes."""
        summary = {}
        for c in codes:
            c_str = str(c).strip().zfill(2)
            files = self.get_attachments_for_province(c_str)
            prov_only = self.provincial_files.get(c_str, [])
            summary[c_str] = {
                "total_count": len(files),
                "national_count": len(self.national_files),
                "provincial_count": len(prov_only),
                "files": files,
                "file_names": [os.path.basename(f) for f in files],
                "has_provincial_attachment": len(prov_only) > 0,
            }
        return summary


def get_file_size_formatted(filepath: str) -> str:
    """Return human readable file size."""
    try:
        size_bytes = os.path.getsize(filepath)
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
    except Exception:
        return "0 B"
