"""
Module for SMTP connection testing and email dispatching with dynamic templates and attachments.
"""

import email.policy
from email.message import EmailMessage
import mimetypes
import os
import re
import smtplib
import socket
import ssl
import time
from typing import Any, Callable, Dict, List, Optional, Tuple


def replace_template(text: str, data: Dict[str, Any]) -> str:
    """
    Replace dynamic tags in subject or body text with data from Excel record.
    Supported placeholders (case-insensitive, multi-syntax):
    - {{no}} or <<no>> or {no} or {{nomor}}
    - {{kode}} or <<kode>> or {kode} or {{kode_provinsi}} or {{kode provisi}}
    - {{provinsi}} or <<nama_provinsi>> or <<provinsi>> or {{provisi}} or {provinsi}
    - {{nama_kepala}} or <<nama_kepala>> or {{kepala}} or {nama_kepala}
    - {{nama_ketua}} or <<nama_ketua>> or {{ketua}} or {{ketua_tim}} or {nama_ketua}
    - {{nama_koordinator}} or <<nama_koordinator>> or {{koordinator}} or {{koor}}
    - {{nama_staf}} or <<nama_staf>> or {{staff}} or {{staf}} or {{nama_staff}}
    - Any other custom non-email column from Excel.

    If a value in Excel is empty or NaN, it is replaced with "" (empty string).
    """
    if not text:
        return ""
    if not data:
        return text

    def get_clean_val(v: Any) -> str:
        if v is None:
            return ""
        s = str(v).strip()
        if s.lower() == "nan":
            return ""
        return s

    mapping: Dict[str, str] = {}

    def set_alias(key_alias: str, val: str):
        k = key_alias.strip().lower()
        if not k:
            return
        mapping[k] = val
        if "_" in k:
            mapping[k.replace("_", " ")] = val
        if " " in k:
            mapping[k.replace(" ", "_")] = val

    # 1. Base mapping from all keys in data
    for k, v in data.items():
        if k in ["errors", "is_valid", "row_number"]:
            continue
        clean_v = get_clean_val(v)
        k_str = str(k).strip()
        set_alias(k_str, clean_v)

        k_lower = k_str.lower()
        # Auto-map prefix "nama_" or "nama " to short tag (e.g. nama_ketua -> ketua)
        if k_lower.startswith("nama_"):
            set_alias(k_lower[5:], clean_v)
        elif k_lower.startswith("nama "):
            set_alias(k_lower[5:], clean_v)

        # Handle staff <-> staf spelling variations
        if "staf" in k_lower and "staff" not in k_lower:
            set_alias(k_lower.replace("staf", "staff"), clean_v)
        elif "staff" in k_lower:
            set_alias(k_lower.replace("staff", "staf"), clean_v)

        # Handle provinsi <-> provisi variations
        if "provinsi" in k_lower:
            set_alias(k_lower.replace("provinsi", "provisi"), clean_v)
        elif "provisi" in k_lower:
            set_alias(k_lower.replace("provisi", "provinsi"), clean_v)

    # 2. Canonical BPS & Common User Aliases (guarantees matching regardless of tag style)
    # Province Code
    kode_val = (
        get_clean_val(data.get("kode"))
        or get_clean_val(data.get("kode_provinsi"))
        or get_clean_val(data.get("kd_prov"))
        or get_clean_val(data.get("kode_provisi"))
    )
    for a in [
        "kode", "kode_provinsi", "kode provinsi", "kode_provisi", "kode provisi",
        "kd_prov", "kd prov", "kode_prov", "kode prov", "kode_wilayah", "kode wilayah", "id_prov", "id prov"
    ]:
        set_alias(a, kode_val)

    # Province Name
    prov_val = (
        get_clean_val(data.get("provinsi"))
        or get_clean_val(data.get("nama_provinsi"))
        or get_clean_val(data.get("provisi"))
        or get_clean_val(data.get("nama_provisi"))
        or get_clean_val(data.get("wilayah"))
        or get_clean_val(data.get("nama_wilayah"))
    )
    for a in [
        "provinsi", "nama_provinsi", "nama provinsi", "provisi", "nama_provisi", "nama provisi",
        "prov", "nama_prov", "nama prov", "wilayah", "nama_wilayah", "nama wilayah", "daerah", "nama_daerah"
    ]:
        set_alias(a, prov_val)

    # Row Number
    no_val = get_clean_val(data.get("no")) or get_clean_val(data.get("nomor"))
    for a in ["no", "nomor", "no.", "urutan", "no_urut", "no urut"]:
        set_alias(a, no_val)

    # Kepala BPS / Prov
    kepala_val = (
        get_clean_val(data.get("nama_kepala"))
        or get_clean_val(data.get("kepala"))
        or get_clean_val(data.get("nama_kepala_bps"))
        or get_clean_val(data.get("kepala_bps"))
        or get_clean_val(data.get("kepala_prov"))
    )
    for a in [
        "nama_kepala", "nama kepala", "kepala", "nama_kepala_bps", "nama kepala bps",
        "kepala_bps", "kepala bps", "nama_kepala_prov", "nama kepala prov", "kepala_prov", "kepala prov"
    ]:
        set_alias(a, kepala_val)

    # Ketua Tim / PJK / Tim Produksi
    ketua_val = (
        get_clean_val(data.get("nama_ketua"))
        or get_clean_val(data.get("ketua"))
        or get_clean_val(data.get("nama_tim_prod"))
        or get_clean_val(data.get("tim_prod"))
        or get_clean_val(data.get("nama_pjk"))
        or get_clean_val(data.get("pjk"))
        or get_clean_val(data.get("ketua_tim"))
    )
    for a in [
        "nama_ketua", "nama ketua", "ketua", "nama_ketua_tim", "nama ketua tim",
        "ketua_tim", "ketua tim", "nama_tim_prod", "nama tim prod", "tim_prod", "tim prod",
        "nama_pjk", "nama pjk", "pjk"
    ]:
        set_alias(a, ketua_val)

    # Koordinator
    koor_val = (
        get_clean_val(data.get("nama_koordinator"))
        or get_clean_val(data.get("koordinator"))
        or get_clean_val(data.get("nama_koor"))
        or get_clean_val(data.get("koor"))
    )
    for a in ["nama_koordinator", "nama koordinator", "koordinator", "nama_koor", "nama koor", "koor"]:
        set_alias(a, koor_val)

    # Staf / Staff
    staf_val = (
        get_clean_val(data.get("nama_staf"))
        or get_clean_val(data.get("nama_staff"))
        or get_clean_val(data.get("staf"))
        or get_clean_val(data.get("staff"))
    )
    for a in ["nama_staf", "nama staf", "staf", "nama_staff", "nama staff", "staff"]:
        set_alias(a, staf_val)

    # 3. Replace tags in text (longer keys first to prevent collision)
    result = text
    for key in sorted(mapping.keys(), key=len, reverse=True):
        val = mapping[key]
        escaped_key = re.escape(key)
        # Matches {{ key }}, << key >>, &lt;&lt; key &gt;&gt;, [[ key ]], [% key %], and { key } case-insensitively with optional spaces
        pattern = re.compile(
            rf"(\{{\{{\s*{escaped_key}\s*\}}\}}|\<\<\s*{escaped_key}\s*\>\>|\&lt;\&lt;\s*{escaped_key}\s*\&gt;\&gt;|\[\[\s*{escaped_key}\s*\]\]|\[\%\s*{escaped_key}\s*\%\]|\{{\s*{escaped_key}\s*\}})",
            re.IGNORECASE,
        )
        result = pattern.sub(lambda _: val, result)

    return result


def extract_envelope_recipients(
    to_email: str,
    cc_email: Optional[str] = None,
    bcc_email: Optional[str] = None,
) -> List[str]:
    """
    Extract and deduplicate recipient emails for SMTP envelope (to_addrs).
    Splits comma and semicolon separated strings.
    """
    recipients = []
    for field in [to_email, cc_email, bcc_email]:
        if not field:
            continue
        parts = re.split(r"[\r\n;,]+", str(field))
        for p in parts:
            cleaned = re.sub(r"\s+", "", p).strip()
            if cleaned and cleaned not in recipients:
                recipients.append(cleaned)
    return recipients


def create_flexible_ssl_context() -> ssl.SSLContext:
    """
    Create an SSLContext that can connect to both standard public mail servers
    and internal enterprise/government mail servers with self-signed or internal CA certificates.
    """
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx
    except Exception:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx


def connect_smtp(
    host: str,
    port: int,
    username: str = "",
    password: str = "",
    timeout: int = 30,
) -> smtplib.SMTP:
    """
    Universally connect and authenticate to an SMTP server.
    Supports:
    - Port 465 (Direct SSL with flexible context)
    - Port 587 (STARTTLS with flexible context)
    - Port 25 / 2525 / other custom ports (Plain or opportunistic STARTTLS)
    - Authenticated or anonymous/unauthenticated internal relays
    """
    clean_host = str(host or "").strip()
    clean_port = int(port) if port else 587
    clean_user = str(username or "").strip()
    clean_pass = str(password or "")
    if "gmail" in clean_host.lower():
        clean_pass = clean_pass.replace(" ", "")

    ctx = create_flexible_ssl_context()

    # 1. Port 465 (Direct SSL)
    if clean_port == 465:
        server = smtplib.SMTP_SSL(clean_host, clean_port, timeout=timeout, context=ctx)
    else:
        # Ports 587, 25, 2525, etc.
        server = smtplib.SMTP(clean_host, clean_port, timeout=timeout)
        server.ehlo()
        if server.has_extn("STARTTLS"):
            try:
                server.starttls(context=ctx)
                server.ehlo()
            except Exception as tls_err:
                # If on port 25 (internal intranet relay), fallback to plain connection
                if clean_port == 25:
                    pass
                else:
                    raise tls_err

    # 2. Authenticate only if credentials are provided
    if clean_user and clean_pass:
        server.login(clean_user, clean_pass)
    elif clean_user and not clean_pass:
        try:
            server.login(clean_user, "")
        except Exception:
            pass

    return server


def test_smtp_connection(
    host: str,
    port: int,
    username: str = "",
    password: str = "",
    timeout: int = 30,
) -> Tuple[bool, str]:
    """Test SMTP authentication and connectivity for any server/port configuration."""
    if not host or not str(host).strip():
        return False, "Host SMTP tidak boleh kosong."

    clean_user = str(username or "").strip()
    clean_pass = str(password or "")

    server = None
    try:
        server = connect_smtp(host, port, clean_user, clean_pass, timeout=timeout)
        if clean_user and clean_pass:
            return True, f"Koneksi dan autentikasi ke {host}:{port} BERHASIL!"
        elif clean_user:
            return True, f"Koneksi ke {host}:{port} BERHASIL (Login sebagai {clean_user} tanpa password)."
        else:
            return True, f"Koneksi ke {host}:{port} BERHASIL! (Mode Relay Tanpa Autentikasi / Anonymous)."
    except smtplib.SMTPAuthenticationError as e:
        err_msg = e.smtp_error.decode("utf-8", errors="ignore") if isinstance(e.smtp_error, bytes) else str(e)
        return False, (
            f"Gagal Autentikasi (Username/Password salah pada {host}): {err_msg}. "
            f"Jika menggunakan Gmail, pastikan menggunakan App Password 16 digit."
        )
    except socket.timeout:
        return False, f"Koneksi timeout ke {host}:{port} (melebihi {timeout} detik). Pastikan Host & Port benar dan jaringan/VPN kantor terhubung."
    except Exception as e:
        return False, f"Gagal terhubung ke {host}:{port}: {str(e)}"
    finally:
        if server:
            try:
                server.quit()
            except Exception:
                pass


def is_html_content(content: str) -> bool:
    """Check if content string contains HTML tags."""
    if not content:
        return False
    pattern = r"(?i)<\s*(?:!doctype|html|head|body|p|div|span|b|strong|i|em|u|h[1-6]|table|thead|tbody|tr|td|th|ul|ol|li|br|hr|a\s+href)[^>]*>"
    return bool(re.search(pattern, content))


def build_email_message(
    sender: str,
    to_email: str,
    cc_email: Optional[str],
    subject: str,
    body: str,
    attachment_paths: List[str],
    is_html: Optional[bool] = None,
) -> EmailMessage:
    """Construct an EmailMessage object with UTF-8 support, HTML alternative, and attachments."""
    msg = EmailMessage(policy=email.policy.default)
    msg["From"] = sender
    msg["To"] = to_email
    if cc_email and cc_email.strip():
        msg["Cc"] = cc_email.strip()
    msg["Subject"] = subject

    has_html = is_html if is_html is not None else is_html_content(body)
    if has_html:
        # Create a clean plain-text fallback
        plain_fallback = re.sub(r"(?i)<\s*(?:br|p|div|tr|h[1-6])[^>]*>", "\n", body)
        plain_fallback = re.sub(r"<[^>]+>", "", plain_fallback)
        plain_fallback = re.sub(r"\n{3,}", "\n\n", plain_fallback).strip()

        msg.set_content(plain_fallback or body, charset="utf-8")
        msg.add_alternative(body, subtype="html", charset="utf-8")
    else:
        msg.set_content(body, charset="utf-8")

    for file_path in attachment_paths:
        if not os.path.isfile(file_path):
            continue

        filename = os.path.basename(file_path)
        ctype, encoding = mimetypes.guess_type(file_path)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)

        with open(file_path, "rb") as fp:
            file_data = fp.read()

        msg.add_attachment(
            file_data,
            maintype=maintype,
            subtype=subtype,
            filename=filename,
        )

    return msg


class EmailDispatcher:
    """Dispatches batch emails with connection reuse, error handling, and cancellation."""

    def __init__(
        self,
        host: str,
        port: int,
        username: str,
        password: str,
        delay_seconds: float = 1.0,
        sender_email: Optional[str] = None,
        timeout: int = 45,
        history_manager: Optional[Any] = None,
    ):
        self.host = host
        self.port = port
        self.username = username.strip()
        self.password = password
        self.sender_email = sender_email.strip() if sender_email and sender_email.strip() else username.strip()
        self.delay_seconds = float(delay_seconds)
        self.timeout = int(timeout)
        self.history_manager = history_manager
        self.is_cancelled = False

    def cancel(self):
        """Signal to cancel batch sending."""
        self.is_cancelled = True

    def _connect(self) -> smtplib.SMTP:
        """Create and authenticate a new SMTP connection using universal connect_smtp."""
        return connect_smtp(
            host=self.host,
            port=self.port,
            username=self.username,
            password=self.password,
            timeout=self.timeout,
        )

    def dispatch_batch(
        self,
        items: List[Dict[str, Any]],
        subject_template: str,
        body_template: str,
        attachment_engine: Any,
        status_callback: Callable[[str, str, Dict[str, Any]], None],
        progress_callback: Callable[[int, int], None],
        dry_run: bool = False,
    ) -> Dict[str, int]:
        """
        Execute dispatch across provincial records.
        status_callback(status: 'SUCCESS'|'SKIPPED'|'ERROR'|'INFO', message: str, record: dict)
        progress_callback(current: int, total: int)
        """
        self.is_cancelled = False
        total = len(items)
        stats = {"success": 0, "skipped": 0, "error": 0}

        server: Optional[smtplib.SMTP] = None

        if not dry_run:
            status_callback("INFO", f"Menghubungkan ke server SMTP {self.host}:{self.port}...", {})
            try:
                server = self._connect()
                status_callback("INFO", "Koneksi SMTP berhasil diinisialisasi.", {})
            except Exception as e:
                status_callback("ERROR", f"Koneksi awal SMTP gagal: {str(e)}", {})
                return stats

        for idx, item in enumerate(items, start=1):
            if self.is_cancelled:
                status_callback("INFO", "Proses pengiriman dibatalkan oleh pengguna.", {})
                break

            kode = item.get("kode", "")
            provinsi = item.get("provinsi", "")
            to_email = (item.get("to_email") or item.get("email_kepala") or "").strip()
            cc_email = (item.get("cc_email") or item.get("email_pjk") or "").strip()
            bcc_email = (item.get("bcc_email") or "").strip()

            # 1. Validation check
            if not item.get("is_valid", True):
                err_msg = ", ".join(item.get("errors", ["Data tidak valid"]))
                status_callback(
                    "SKIPPED",
                    f"[{kode} - {provinsi}] Dilewati karena kesalahan data: {err_msg}",
                    item,
                )
                stats["skipped"] += 1
                progress_callback(idx, total)
                continue

            if not to_email:
                status_callback(
                    "SKIPPED",
                    f"[{kode} - {provinsi}] Dilewati: Alamat email tujuan (To) kosong (seluruh kolom email kosong).",
                    item,
                )
                stats["skipped"] += 1
                progress_callback(idx, total)
                continue

            # 2. Build email content
            subject = replace_template(subject_template, item)
            body = replace_template(body_template, item)
            attachments = attachment_engine.get_attachments_for_province(kode)
            att_names = [os.path.basename(a) for a in attachments]
            envelope_recipients = extract_envelope_recipients(to_email, cc_email, bcc_email)

            # 3. Dry Run Simulation Mode
            if dry_run:
                if self.history_manager:
                    try:
                        self.history_manager.add_record(
                            sender_email=self.sender_email,
                            recipient_to=to_email,
                            recipient_cc=cc_email,
                            recipient_bcc=bcc_email,
                            provinsi_kode=kode,
                            provinsi_nama=provinsi,
                            subject=subject,
                            body_html=body,
                            body_text=body,
                            attachments=", ".join(att_names),
                            status="SUCCESS",
                            is_dry_run=True,
                        )
                    except Exception:
                        pass

                status_callback(
                    "SUCCESS",
                    f"[SIMULASI] [{kode} - {provinsi}] Ke: {to_email} | Cc: {cc_email or '-'} | Bcc: {bcc_email or '-'} | {len(attachments)} Lampiran: {', '.join(att_names) or 'Tidak ada'}",
                    item,
                )
                stats["success"] += 1
                progress_callback(idx, total)
                time.sleep(0.05)  # Fast animation for simulation
                continue

            # 4. Actual Dispatch
            try:
                msg = build_email_message(
                    sender=self.sender_email,
                    to_email=to_email,
                    cc_email=cc_email,
                    subject=subject,
                    body=body,
                    attachment_paths=attachments,
                )

                # Reconnect if connection was dropped
                # Deliver to envelope recipients (To, Cc, Bcc) without Bcc header in message
                try:
                    refused = server.send_message(msg, from_addr=self.sender_email, to_addrs=envelope_recipients)
                except (smtplib.SMTPServerDisconnected, smtplib.SMTPSenderRefused, socket.error):
                    status_callback("INFO", "Mencoba menyambungkan kembali ke server SMTP...", item)
                    try:
                        if server:
                            server.close()
                    except Exception:
                        pass
                    server = self._connect()
                    refused = server.send_message(msg, from_addr=self.sender_email, to_addrs=envelope_recipients)

                if refused:
                    refused_parts = []
                    for rcpt, err_data in refused.items():
                        err_text = err_data[1].decode("utf-8", errors="ignore") if isinstance(err_data[1], bytes) else str(err_data[1])
                        refused_parts.append(f"{rcpt}: {err_data[0]} {err_text}")
                    refused_msg = "; ".join(refused_parts)

                    status_callback(
                        "ERROR",
                        f"[{kode} - {provinsi}] Server SMTP menolak penerima ({refused_msg})",
                        item,
                    )
                    stats["error"] += 1
                    if self.history_manager:
                        try:
                            self.history_manager.add_record(
                                sender_email=self.sender_email,
                                recipient_to=to_email,
                                recipient_cc=cc_email,
                                recipient_bcc=bcc_email,
                                provinsi_kode=kode,
                                provinsi_nama=provinsi,
                                subject=subject,
                                body_html=body,
                                body_text=body,
                                attachments=", ".join(att_names),
                                status="FAILED",
                                error_message=f"Penerima ditolak: {refused_msg}",
                                is_dry_run=False,
                            )
                        except Exception:
                            pass
                else:
                    status_callback(
                        "SUCCESS",
                        f"[{kode} - {provinsi}] Email terkirim ke {to_email} (Cc: {cc_email or '-'}, Bcc: {bcc_email or '-'}), {len(attachments)} lampiran.",
                        item,
                    )
                    stats["success"] += 1

                    if self.history_manager:
                        try:
                            self.history_manager.add_record(
                                sender_email=self.sender_email,
                                recipient_to=to_email,
                                recipient_cc=cc_email,
                                recipient_bcc=bcc_email,
                                provinsi_kode=kode,
                                provinsi_nama=provinsi,
                                subject=subject,
                                body_html=body,
                                body_text=body,
                                attachments=", ".join(att_names),
                                status="SUCCESS",
                                is_dry_run=False,
                            )
                        except Exception:
                            pass

            except Exception as e:
                status_callback(
                    "ERROR",
                    f"[{kode} - {provinsi}] Gagal mengirim ke {to_email}: {str(e)}",
                    item,
                )
                stats["error"] += 1

                if self.history_manager:
                    try:
                        self.history_manager.add_record(
                            sender_email=self.sender_email,
                            recipient_to=to_email,
                            recipient_cc=cc_email,
                            recipient_bcc=bcc_email,
                            provinsi_kode=kode,
                            provinsi_nama=provinsi,
                            subject=subject,
                            body_html=body,
                            body_text=body,
                            attachments=", ".join(att_names),
                            status="FAILED",
                            error_message=str(e),
                            is_dry_run=False,
                        )
                    except Exception:
                        pass

            progress_callback(idx, total)
            if self.delay_seconds > 0 and idx < total:
                time.sleep(self.delay_seconds)

        if server:
            try:
                server.quit()
            except Exception:
                pass

        return stats


def check_smtp_sender_compatibility(
    host: str,
    username: str = "",
    sender_email: str = "",
) -> Tuple[bool, List[str], str]:
    """
    Check if SMTP host, login username, and sender email (From) are compatible.
    Returns:
        (is_compatible: bool, warnings: List[str], recommended_sender: str)
    """
    clean_host = str(host or "").strip().lower()
    clean_user = str(username or "").strip()
    clean_sender = str(sender_email or "").strip()

    warnings: List[str] = []
    recommended_sender = clean_sender or clean_user

    # 1. Sender email must not be empty
    if not clean_sender:
        warnings.append("Kolom 'Email Pengirim (From)' masih kosong. Email pengirim wajib diisi agar server dapat mengirim email.")
        if clean_user and "@" in clean_user:
            recommended_sender = clean_user
        return False, warnings, recommended_sender

    # Helper to extract domain from email string
    def get_domain(email_str: str) -> str:
        if "@" in email_str:
            return email_str.split("@")[-1].strip().lower()
        return ""

    user_domain = get_domain(clean_user)
    sender_domain = get_domain(clean_sender)

    # 2. Check BPS Server rules
    if "bps.go.id" in clean_host or "bps" in clean_host:
        if sender_domain != "bps.go.id":
            warnings.append(
                f"Host yang digunakan adalah Server BPS ('{host}'), tetapi Email Pengirim Anda adalah '{clean_sender}' (bukan @bps.go.id).\n"
                f"Server BPS memiliki proteksi Anti-Spoofing yang akan menolak/memblokir pengiriman jika alamat pengirim bukan email resmi BPS."
            )
            if user_domain == "bps.go.id":
                recommended_sender = clean_user
            elif clean_user and "@" not in clean_user:
                recommended_sender = f"{clean_user}@bps.go.id"

        elif user_domain == "bps.go.id" and clean_user.lower() != clean_sender.lower():
            warnings.append(
                f"Username login ('{clean_user}') berbeda dengan Email Pengirim ('{clean_sender}').\n"
                f"Sebagian besar server email BPS mewajibkan Email Pengirim sama persis dengan akun yang login."
            )
            recommended_sender = clean_user

    # 3. Check Gmail rules
    elif "gmail.com" in clean_host:
        if sender_domain != "gmail.com":
            warnings.append(
                f"Host yang digunakan adalah Gmail (smtp.gmail.com), tetapi Email Pengirim adalah '{clean_sender}' (bukan @gmail.com).\n"
                f"Google SMTP mewajibkan pengirim menggunakan akun @gmail.com yang valid."
            )
            if user_domain == "gmail.com":
                recommended_sender = clean_user

        elif user_domain == "gmail.com" and clean_user.lower() != clean_sender.lower():
            warnings.append(
                f"Username akun Google ('{clean_user}') berbeda dengan Email Pengirim ('{clean_sender}').\n"
                f"Google mewajibkan Email Pengirim identik dengan akun Gmail login Anda."
            )
            recommended_sender = clean_user

    # 4. Check Brevo / Sendinblue rules
    elif "brevo" in clean_host or "sendinblue" in clean_host:
        if not sender_domain:
            warnings.append("Format Email Pengirim tidak valid.")

    # 5. General check if both username and sender are emails but domains differ
    elif user_domain and sender_domain and user_domain != sender_domain:
        warnings.append(
            f"Domain akun login ('{user_domain}') berbeda dengan domain Email Pengirim ('{sender_domain}').\n"
            f"Banyak server SMTP menolak pengiriman jika domain pengirim tidak cocok dengan akun login (Error 550/553 Relaying Denied)."
        )
        if user_domain:
            recommended_sender = clean_user

    is_compatible = (len(warnings) == 0)
    return is_compatible, warnings, recommended_sender

