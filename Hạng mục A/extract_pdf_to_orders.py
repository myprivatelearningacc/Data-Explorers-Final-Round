import re
import sqlite3
import unicodedata
from pathlib import Path
from datetime import datetime
from email import policy
from email.parser import BytesParser

import pandas as pd
import pdfplumber
from tqdm import tqdm


# =========================
# CONFIG
# =========================

BASE_DIR = Path(__file__).resolve().parent

PDF_DIR = BASE_DIR / "extracted_attachments" / "email"
EMAIL_LOG_PATH = BASE_DIR / "processed_outputs_emails" / "email_log.csv"

# Code sẽ tự tìm .eml trong các folder phổ biến này.
# Nếu file .eml nằm folder khác, thêm path vào list này.
EML_SEARCH_DIRS = [
    BASE_DIR,
    BASE_DIR / "data source/email",
    BASE_DIR / "email",
    BASE_DIR / "eml",
    BASE_DIR / "raw_emails",
    BASE_DIR / "data",
    BASE_DIR / "input",
]

OUTPUT_DIR = BASE_DIR / "processed_outputs_pdf"
DB_PATH = BASE_DIR / "tnbike_orders_mar2026.db"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SALES_ORDER_CSV = OUTPUT_DIR / "sales_order.csv"
ORDER_LINE_CSV = OUTPUT_DIR / "order_line.csv"
FACT_SALES_CSV = OUTPUT_DIR / "fact_sales.csv"
PDF_ERROR_LOG_CSV = OUTPUT_DIR / "pdf_error_log.csv"

SALES_ORDER_XLSX = OUTPUT_DIR / "sales_order.xlsx"
ORDER_LINE_XLSX = OUTPUT_DIR / "order_line.xlsx"
FACT_SALES_XLSX = OUTPUT_DIR / "fact_sales.xlsx"


# =========================
# TEXT CLEANING HELPERS
# =========================

def strip_accents(text):
    if text is None or pd.isna(text):
        return ""
    s = str(text)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    return s.lower()


def clean_cell(x):
    if x is None:
        return ""
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def compact_norm(text):
    s = strip_accents(text)
    s = s.replace("□", " ").replace("�", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


COMMON_TEXT_REPLACEMENTS = {
    # Company / customer terms from broken PDF font
    "C  PH N": "CỔ PHẦN",
    "C PH N": "CỔ PHẦN",
    "Cn PHnN": "CỔ PHẦN",
    "CÔNG TY Cn PHnN": "CÔNG TY CỔ PHẦN",
    "CÔNG TY Cn PH N": "CÔNG TY CỔ PHẦN",
    "CÔNG TY C PH N": "CÔNG TY CỔ PHẦN",
    "TNHH THnnNG MnI": "TNHH THƯƠNG MẠI",
    "THnnNG MnI": "THƯƠNG MẠI",
    "THnnNG MI": "THƯƠNG MẠI",
    "THƯƠNG MnI": "THƯƠNG MẠI",
    "THnNG MnI": "THƯƠNG MẠI",
    "THnNG MI": "THƯƠNG MẠI",
    "THnnNG": "THƯƠNG",
    "MnI": "MẠI",
    "DOANH NGHInP": "DOANH NGHIỆP",
    "DOANH NGHIEP": "DOANH NGHIỆP",
    "Hn KINH DOANH": "HỘ KINH DOANH",
    "Hộ KINH DOANH": "HỘ KINH DOANH",
    "CnA HÀNG": "CỬA HÀNG",
    "CNA HÀNG": "CỬA HÀNG",
    "CỬA HÀNG XE nnp": "CỬA HÀNG XE ĐẠP",
    "CỬA HÀNG XE npp": "CỬA HÀNG XE ĐẠP",
    "CỬA HÀNG XE ĐNP": "CỬA HÀNG XE ĐẠP",
    "KHAI HOÀN": "KHẢI HOÀN",
    "KHnI HOÀN": "KHẢI HOÀN",
    "HOnN": "HOÀN",
    "TÂY BnC": "TÂY BẮC",
    "ĐnI LỢI": "ĐẠI LỢI",
    "ĐAI LỢI": "ĐẠI LỢI",
    "NAM TInN": "NAM TIẾN",
    "TOÀN THnNG": "TOÀN THẮNG",
    "Vnn LỢI": "VẠN LỢI",
    "VnN LỢI": "VẠN LỢI",
    "VInT ANH": "VIỆT ANH",

    # Product terms
    "TH NG NH T": "THỐNG NHẤT",
    "THNG NHT": "THỐNG NHẤT",
    "Thnng Nhnt": "Thống Nhất",
    "Thnng nhnt": "Thống Nhất",
    "thnng nhnt": "Thống Nhất",
    "Thng Nht": "Thống Nhất",
    "Thong Nhat": "Thống Nhất",
    "Xe nnp": "Xe đạp",
    "Xe npp": "Xe đạp",
    "Xe đnp": "Xe đạp",
    "Xe dnp": "Xe đạp",
    "Xe dap": "Xe đạp",
    "Xe dp": "Xe đạp",
    "Xe đ p": "Xe đạp",
    "nnp": "đạp",
    "npp": "đạp",
    "đnp": "đạp",
    "Chinc": "Chiếc",
    "chinc": "chiếc",

    # Colors / variants
    "Hnng": "Hồng",
    "hnng": "hồng",
    "Trnng": "Trắng",
    "trnng": "trắng",
    "Coban": "Cơ bản",
    "coban": "cơ bản",
    "Ca fé/nâu": "Cà phê/nâu",
    "Ca fe/nau": "Cà phê/nâu",
    "Cafe/nâu": "Cà phê/nâu",
    "mimt": "mint",
    "Xanh mimt": "Xanh mint",

    # Address terms
    "TP HCM": "TP. Hồ Chí Minh",
    "TP. HCM": "TP. Hồ Chí Minh",
    "Tp HCM": "TP. Hồ Chí Minh",
    "Tp. HCM": "TP. Hồ Chí Minh",
    "TP H C M": "TP. Hồ Chí Minh",
    "TP Hà Ni": "TP Hà Nội",
    "TP Ha Noi": "TP Hà Nội",
    "Hà Ni": "Hà Nội",
    "Ha Noi": "Hà Nội",
    "Qu n": "Quận",
    "Qun": "Quận",
    "quan ": "quận ",
    "phn": "phường",
    "phng": "phường",
    "ph ng": "phường",
    "Phng": "Phường",
    "Ph ng": "Phường",
    "Hàng Trng": "Hàng Trống",
    "Hang Trong": "Hàng Trống",
    "Hoàn Kim": "Hoàn Kiếm",
    "Hoan Kiem": "Hoàn Kiếm",
}


def apply_common_replacements(text):
    if text is None or pd.isna(text):
        return None
    s = str(text).strip()
    if s == "":
        return None
    s = s.replace("□", " ").replace("�", " ")
    s = re.sub(r"\s+", " ", s)
    for wrong, right in COMMON_TEXT_REPLACEMENTS.items():
        s = s.replace(wrong, right)
    s = re.sub(
        r"\bXe\s+\S{2,5}\s+(Thống\s+Nhất|Thong\s+Nhat|Thnng\s+Nhnt|Thng\s+Nht)\b",
        r"Xe đạp \1",
        s,
        flags=re.IGNORECASE,
    )
    s = s.replace("Xe đạp Thnng Nhnt", "Xe đạp Thống Nhất")
    s = s.replace("Xe đạp Thng Nht", "Xe đạp Thống Nhất")
    s = s.replace("Xe đạp Thong Nhat", "Xe đạp Thống Nhất")
    s = re.sub(r"\s*,\s*", ", ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,;:-")
    return s if s else None


def clean_customer_name(name):
    s = apply_common_replacements(name)
    if not s:
        return None
    s = re.split(
        r"\bMST\b|Địa\s*chỉ|Dia\s*chi|Số\s*chứng\s*từ|Ngày\s*đặt|Tổng\s*sản\s*phẩm|Tel\s*:",
        s,
        flags=re.IGNORECASE,
    )[0]
    s = re.sub(r"\s+", " ", s).strip(" ,;:-")
    return s or None


def clean_address(address):
    s = apply_common_replacements(address)
    if not s:
        return None
    s = re.sub(
        r"^(Địa\s*chỉ|Dia\s*chi|Address|a\s*chỉ|a\s*chi|chỉ|chi)\s*[:\-]?\s*",
        "",
        s,
        flags=re.IGNORECASE,
    )
    s = re.split(r"\bTổng\s*sản\s*phẩm\b|\bTổng\s*số\s*lượng\b|\bTổng\s*giá\s*trị\b|\bTel\s*:", s, flags=re.IGNORECASE)[0]
    s = re.sub(r"\s+", " ", s).strip(" ,;:-")
    return s or None


# =========================
# PROVINCE EXTRACT
# =========================

PROVINCES = [
    "Hà Nội", "Hải Phòng", "Quảng Ninh", "Bắc Ninh", "Bắc Giang", "Hải Dương",
    "Hưng Yên", "Thái Bình", "Nam Định", "Ninh Bình", "Hà Nam", "Vĩnh Phúc",
    "Phú Thọ", "Thái Nguyên", "Bắc Kạn", "Cao Bằng", "Lạng Sơn", "Tuyên Quang",
    "Hà Giang", "Yên Bái", "Lào Cai", "Lai Châu", "Điện Biên", "Sơn La",
    "Hòa Bình", "Thanh Hóa", "Nghệ An", "Hà Tĩnh", "Quảng Bình", "Quảng Trị",
    "Thừa Thiên Huế", "Huế", "Đà Nẵng", "Quảng Nam", "Quảng Ngãi", "Bình Định",
    "Phú Yên", "Khánh Hòa", "Ninh Thuận", "Bình Thuận", "Kon Tum", "Gia Lai",
    "Đắk Lắk", "Đắk Nông", "Lâm Đồng", "TP. Hồ Chí Minh", "Hồ Chí Minh",
    "Bình Dương", "Đồng Nai", "Bà Rịa - Vũng Tàu", "Tây Ninh", "Bình Phước",
    "Long An", "Tiền Giang", "Bến Tre", "Trà Vinh", "Vĩnh Long", "Đồng Tháp",
    "An Giang", "Kiên Giang", "Cần Thơ", "Hậu Giang", "Sóc Trăng", "Bạc Liêu",
    "Cà Mau",
]

PROVINCE_ALIASES = {
    "tp hcm": "TP. Hồ Chí Minh", "tphcm": "TP. Hồ Chí Minh",
    "tp ho chi minh": "TP. Hồ Chí Minh", "ho chi minh": "TP. Hồ Chí Minh",
    "sai gon": "TP. Hồ Chí Minh", "tp ha noi": "Hà Nội", "ha noi": "Hà Nội",
    "hanoi": "Hà Nội", "hai phong": "Hải Phòng", "da nang": "Đà Nẵng",
    "can tho": "Cần Thơ", "ba ria vung tau": "Bà Rịa - Vũng Tàu",
    "thua thien hue": "Thừa Thiên Huế", "dak lak": "Đắk Lắk",
    "dac lak": "Đắk Lắk", "dak nong": "Đắk Nông", "dac nong": "Đắk Nông",
}


def extract_province_from_address(address):
    if address is None or pd.isna(address) or str(address).strip() == "":
        return None
    norm = compact_norm(address)
    for alias, province in PROVINCE_ALIASES.items():
        if re.search(rf"\b{re.escape(compact_norm(alias))}\b", norm):
            return province
    for province in sorted(PROVINCES, key=len, reverse=True):
        p_norm = compact_norm(province)
        if re.search(rf"\b{re.escape(p_norm)}\b", norm):
            return "TP. Hồ Chí Minh" if province == "Hồ Chí Minh" else province
    return None


# =========================
# BASIC PARSERS
# =========================

def parse_number(x):
    if x is None:
        return None
    s = str(x).strip()
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s == "":
        return None
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "").replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None

def is_missing_value(x):
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    s = str(x).strip()
    return s == "" or s.lower() in ["nan", "none", "null"]


def coalesce_clean(*values):
    for v in values:
        if not is_missing_value(v):
            return v
    return None

def safe_parse_float(x, default=None):
    """Parse a number safely; return default if parsing fails or value is blank/None/NaN."""
    if x is None or pd.isna(x) or str(x).strip() == "":
        return default
    value = parse_number(x)
    return default if value is None else value


def safe_parse_int(x, default=None):
    """Parse an integer safely; return default if parsing fails or value is blank/None/NaN."""
    value = safe_parse_float(x, default=None)
    if value is None:
        return default
    return int(value)


def normalize_invoice_number(x):
    if x is None:
        return None
    s = str(x).strip().replace(".", "_").replace("-", "_")
    s = re.sub(r"\s+", "_", s)
    return s or None


def normalize_product_code(x):
    if x is None:
        return None
    s = str(x).strip().replace(" ", "").replace("\n", "")
    if s == "":
        return None
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    return s


def normalize_tax_code(x):
    if x is None:
        return None
    s = re.sub(r"[^0-9]", "", str(x))
    return s or None


def parse_order_date(x):
    if x is None or pd.isna(x) or str(x).strip() == "":
        return None
    s = str(x).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return dt.strftime("%d/%m/%Y")
    return s


def extract_sender_name(from_address):
    if from_address is None or pd.isna(from_address):
        return None
    sender = re.sub(r"<.*?>", "", str(from_address)).strip()
    return clean_customer_name(sender) if sender else None


def extract_sender_email(from_address):
    if from_address is None or pd.isna(from_address):
        return None
    match = re.search(r"<([^>]+)>", str(from_address))
    if match:
        return match.group(1).strip()
    if "@" in str(from_address):
        return str(from_address).strip()
    return None


def make_customer_code(from_address):
    email_addr = extract_sender_email(from_address)
    if email_addr is None:
        return None
    domain = email_addr.split("@")[-1].lower()
    return domain.split(".")[0].upper()


# =========================
# PRODUCT CLEANING
# =========================

PRODUCT_NAME_BY_CODE = {}


def clean_product_name(product_code, product_name_pdf):
    product_code = normalize_product_code(product_code)
    if product_code in PRODUCT_NAME_BY_CODE:
        return PRODUCT_NAME_BY_CODE[product_code], "CODE_MAPPING"
    original = str(product_name_pdf).strip() if product_name_pdf is not None and not pd.isna(product_name_pdf) else ""
    if original == "":
        return None, "MISSING"
    name = apply_common_replacements(original)
    if not name:
        return None, "MISSING"
    name = re.sub(r"^Xe\s+\S{2,5}\s+Thống\s+Nhất", "Xe đạp Thống Nhất", name, flags=re.IGNORECASE)
    name = re.sub(r"^Xe\s+(?=Thống\s+Nhất)", "Xe đạp ", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()
    return (name, "AUTO_CLEAN") if name != original else (name, "PDF_RAW")


# =========================
# PDF READERS
# =========================

def read_pdf_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += "\n" + (page.extract_text() or "")
    return text


def read_pdf_tables(pdf_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                for row in table:
                    row = [clean_cell(cell) for cell in row]
                    if any(cell != "" for cell in row):
                        all_rows.append(row)
    return all_rows


# =========================
# EML READER + EMAIL BODY PARSER
# =========================

def find_eml_path(eml_file):
    if eml_file is None or pd.isna(eml_file) or str(eml_file).strip() == "":
        return None
    eml_file = str(eml_file).strip()
    candidate = Path(eml_file)
    if candidate.is_absolute() and candidate.exists():
        return candidate
    for d in EML_SEARCH_DIRS:
        if d.exists():
            p = d / eml_file
            if p.exists():
                return p
    for d in EML_SEARCH_DIRS:
        if d.exists():
            try:
                matches = list(d.rglob(eml_file))
                if matches:
                    return matches[0]
            except Exception:
                pass
    return None


def read_eml_body(eml_path):
    if eml_path is None:
        return None
    try:
        with open(eml_path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        bodies = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    continue
                if part.get_content_type() == "text/plain":
                    try:
                        bodies.append(part.get_content())
                    except Exception:
                        payload = part.get_payload(decode=True)
                        if payload:
                            bodies.append(payload.decode("utf-8", errors="replace"))
            if not bodies:
                for part in msg.walk():
                    if part.get_content_type() == "text/html" and part.get_content_disposition() != "attachment":
                        try:
                            html = part.get_content()
                        except Exception:
                            payload = part.get_payload(decode=True)
                            html = payload.decode("utf-8", errors="replace") if payload else ""
                        html = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
                        html = re.sub(r"</p\s*>", "\n", html, flags=re.IGNORECASE)
                        html = re.sub(r"<[^>]+>", " ", html)
                        bodies.append(html)
        else:
            try:
                bodies.append(msg.get_content())
            except Exception:
                payload = msg.get_payload(decode=True)
                if payload:
                    bodies.append(payload.decode("utf-8", errors="replace"))
        body = "\n".join(str(x) for x in bodies if x)
        body = re.sub(r"\r\n?", "\n", body)
        body = re.sub(r"[ \t]+", " ", body)
        return body.strip() or None
    except Exception:
        return None


def extract_line_value_from_email(body, labels):
    if body is None:
        return None
    lines = [line.strip() for line in str(body).splitlines()]
    label_norms = [compact_norm(x) for x in labels]
    for line in lines:
        if ":" not in line:
            continue
        left, right = line.split(":", 1)
        left_norm = compact_norm(left)
        for lab_norm in label_norms:
            if left_norm == lab_norm or left_norm.endswith(lab_norm):
                value = right.strip()
                return value if value else None
    joined = "\n".join(lines)
    for label in labels:
        pattern = rf"{re.escape(label)}\s*:\s*(.+?)(?:\n|$)"
        m = re.search(pattern, joined, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def parse_email_body_order_info(body):
    info = {
        "invoice_number": None,
        "order_date": None,
        "customer_name": None,
        "tax_code": None,
        "address": None,
        "province_name": None,
        "total_amount_pdf": None,
        "total_quantity_email": None,
        "num_lines_email": None,
    }
    if body is None or str(body).strip() == "":
        return info
    invoice = extract_line_value_from_email(body, ["Số chứng từ", "Số đơn hàng", "Mã đơn hàng", "Số hóa đơn"])
    date = extract_line_value_from_email(body, ["Ngày đặt", "Ngày chứng từ", "Ngày lập", "Ngày"])
    customer = extract_line_value_from_email(body, ["Khách hàng", "Đại lý", "Tên khách hàng", "Tên đại lý"])
    tax = extract_line_value_from_email(body, ["MST", "Mã số thuế", "Tax code"])
    address = extract_line_value_from_email(body, ["Địa chỉ", "Dia chi", "Address"])
    total_qty = extract_line_value_from_email(body, ["Tổng số lượng"])
    num_lines = extract_line_value_from_email(body, ["Tổng sản phẩm"])
    total_amount = extract_line_value_from_email(body, ["Tổng giá trị", "Tổng tiền", "Tổng cộng"])
    if invoice:
        info["invoice_number"] = normalize_invoice_number(invoice)
    if date:
        info["order_date"] = parse_order_date(date)
    if customer:
        info["customer_name"] = clean_customer_name(customer)
    if tax:
        info["tax_code"] = normalize_tax_code(tax)
    if address:
        info["address"] = clean_address(address)
        info["province_name"] = extract_province_from_address(info["address"])
    if total_amount:
        info["total_amount_pdf"] = parse_number(total_amount)
    if total_qty:
        info["total_quantity_email"] = parse_number(total_qty)
    if num_lines:
        info["num_lines_email"] = parse_number(num_lines)
    return info


# =========================
# EMAIL LOG
# =========================

def load_email_log():
    if not EMAIL_LOG_PATH.exists():
        print(f"Warning: email_log.csv not found at {EMAIL_LOG_PATH}")
        return pd.DataFrame()
    email_log = pd.read_csv(EMAIL_LOG_PATH, dtype=str)
    required_cols = [
        "eml_file", "message_id", "from_address", "to_address", "subject",
        "received_at", "attachment_name", "attachment_path",
        "processing_status", "processed_at", "error_message"
    ]
    for col in required_cols:
        if col not in email_log.columns:
            email_log[col] = None
    email_log["source_pdf_file"] = email_log["attachment_name"].apply(
        lambda x: Path(str(x)).name if pd.notna(x) else None
    )
    email_log["customer_name_from_email_sender"] = email_log["from_address"].apply(extract_sender_name)
    email_log["customer_code_from_email"] = email_log["from_address"].apply(make_customer_code)
    parsed_rows = []
    found_eml = 0
    for _, row in email_log.iterrows():
        eml_path = find_eml_path(row.get("eml_file"))
        if eml_path is not None:
            found_eml += 1
        body = read_eml_body(eml_path)
        parsed = parse_email_body_order_info(body)
        parsed["eml_body_found"] = body is not None
        parsed_rows.append(parsed)
    parsed_df = pd.DataFrame(parsed_rows).add_prefix("email_body_")
    email_log = pd.concat([email_log.reset_index(drop=True), parsed_df.reset_index(drop=True)], axis=1)
    print(f"EML files found/readable: {found_eml}/{len(email_log)}")
    return email_log


# =========================
# ORDER HEADER EXTRACTION FROM PDF
# =========================

def is_product_table_row(row):
    joined = compact_norm(" ".join(row))
    if re.match(r"^\d+\s", joined):
        return True
    product_header_terms = ["stt", "ma hang", "ten san pham", "dvt", "sl", "don gia", "thanh tien"]
    return sum(term in joined for term in product_header_terms) >= 2


def looks_like_address(value):
    if value is None:
        return False
    s = compact_norm(value)
    if s == "":
        return False
    terms = [
        "tp", "ha noi", "hcm", "ho chi minh", "quan", "phuong", "huyen", "thi xa",
        "thi tran", "tinh", "duong", "pho", "hang", "hoan kiem", "thanh pho",
        "thanh hoa", "nguyen trai"
    ]
    return sum(term in s for term in terms) >= 2 or any(p in s for p in ["ha noi", "ho chi minh", "hcm", "thanh hoa"])


def label_is_invoice(label):
    s = compact_norm(label)
    return any(k in s for k in ["so don hang", "don hang", "so chung tu", "so hoa don", "invoice"])


def label_is_date(label):
    return "ngay" in compact_norm(label)


def label_is_customer(label):
    s = compact_norm(label)
    return any(k in s for k in ["dai ly", "khach hang", "ten dai ly", "ten khach hang"])


def label_is_tax(label):
    s = compact_norm(label)
    return "mst" in s or "ma so thue" in s or "tax" in s


def label_is_address(label):
    s = compact_norm(label)
    return any(k in s for k in ["dia chi", "ia chi", "a chi", "address"])


def extract_order_header_from_tables(table_rows, pdf_file):
    header = {
        "invoice_number": normalize_invoice_number(Path(pdf_file).stem),
        "order_date": None,
        "customer_code": None,
        "customer_name": None,
        "tax_code": None,
        "address": None,
        "province_name": None,
        "total_amount_pdf": None,
    }
    for raw_row in table_rows:
        row = [clean_cell(x) for x in raw_row]
        if not row or is_product_table_row(row):
            continue
        cells = row + [""] * max(0, 6 - len(row))
        for cell in cells:
            if not header["invoice_number"]:
                m = re.search(r"BH\d{2}[._\-]\d+", cell, flags=re.IGNORECASE)
                if m:
                    header["invoice_number"] = normalize_invoice_number(m.group(0))
            if not header["order_date"]:
                m = re.search(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b", cell)
                if m:
                    header["order_date"] = parse_order_date(m.group(0))
            if not header["address"] and looks_like_address(cell):
                header["address"] = clean_address(cell)
        for i in range(0, len(cells) - 1, 2):
            label = cells[i]
            value = cells[i + 1]
            if value is None or str(value).strip() == "":
                continue
            if label_is_invoice(label):
                header["invoice_number"] = normalize_invoice_number(value)
            elif label_is_date(label):
                header["order_date"] = parse_order_date(value)
            elif label_is_customer(label):
                header["customer_name"] = clean_customer_name(value)
            elif label_is_tax(label):
                header["tax_code"] = normalize_tax_code(value)
            elif label_is_address(label):
                header["address"] = clean_address(value)
        if len(cells) >= 4:
            left_label = compact_norm(cells[0])
            right_label = compact_norm(cells[2])
            if not header["customer_name"] and ("ly" in left_label or "khach" in left_label) and cells[1]:
                header["customer_name"] = clean_customer_name(cells[1])
            if not header["tax_code"] and ("mst" in right_label or "thue" in right_label) and cells[3]:
                header["tax_code"] = normalize_tax_code(cells[3])
            if not header["address"] and cells[1] and looks_like_address(cells[1]):
                header["address"] = clean_address(cells[1])
    header["province_name"] = extract_province_from_address(header.get("address"))
    return header


def extract_order_header_from_text(text, pdf_file):
    header = {
        "invoice_number": normalize_invoice_number(Path(pdf_file).stem),
        "order_date": None,
        "customer_code": None,
        "customer_name": None,
        "tax_code": None,
        "address": None,
        "province_name": None,
        "total_amount_pdf": None,
    }
    patterns = {
        "invoice_number": [
            r"Số\s*chứng\s*từ[:\s]+([A-Za-z0-9_\-./]+)",
            r"Số\s*đơn\s*hàng[:\s]+([A-Za-z0-9_\-./]+)",
            r"Số\s*hóa\s*đơn[:\s]+([A-Za-z0-9_\-./]+)",
            r"(BH\d{2}[._\-]\d+)",
        ],
        "order_date": [
            r"Ngày\s*đặt[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
            r"Ngày\s*chứng\s*từ[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
            r"Ngày[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
        ],
        "customer_name": [
            r"Đại\s*lý[:\s]+(.+?)(?:\s+MST|\s+Địa\s*chỉ|$)",
            r"Khách\s*hàng[:\s]+(.+?)(?:\s+MST|\s+Địa\s*chỉ|$)",
        ],
        "tax_code": [
            r"MST[:\s]*([0-9\-. ]+)",
            r"Mã\s*số\s*thuế[:\s]*([0-9\-. ]+)",
        ],
        "address": [
            r"Địa\s*chỉ[:\s]+(.+?)(?:\n|$)",
            r"Dia\s*chi[:\s]+(.+?)(?:\n|$)",
        ],
        "total_amount_pdf": [
            r"Tổng\s*giá\s*trị\s*đơn\s*hàng.*?([\d.,]+)",
            r"Tổng\s*cộng.*?([\d.,]+)",
            r"Tổng\s*tiền.*?([\d.,]+)",
        ],
    }
    for field, regex_list in patterns.items():
        for pattern in regex_list:
            match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
            if match:
                value = match.group(1).strip()
                if field == "invoice_number":
                    value = normalize_invoice_number(value)
                elif field == "order_date":
                    value = parse_order_date(value)
                elif field == "customer_name":
                    value = clean_customer_name(value)
                elif field == "tax_code":
                    value = normalize_tax_code(value)
                elif field == "address":
                    value = clean_address(value)
                elif field == "total_amount_pdf":
                    value = parse_number(value)
                header[field] = value
                break
    header["province_name"] = extract_province_from_address(header.get("address"))
    return header


def merge_headers(preferred, *fallbacks):
    out = dict(preferred)

    for fallback in fallbacks:
        for key, value in fallback.items():
            if is_missing_value(out.get(key)) and not is_missing_value(value):
                out[key] = value

    if not is_missing_value(out.get("address")) and is_missing_value(out.get("province_name")):
        out["province_name"] = extract_province_from_address(out.get("address"))

    return out


# =========================
# EXTRACT ORDER LINES
# =========================

def looks_like_product_row(row):
    if len(row) < 6:
        return False
    row = [clean_cell(x) for x in row]
    joined = compact_norm(" ".join(row))
    header_keywords = ["stt", "ma hang", "ten san pham", "dvt", "so luong", "don gia", "thanh tien", "tong", "cong"]
    if any(k in joined for k in header_keywords):
        return False
    first = row[0].strip()
    return bool(re.match(r"^\d+$", first))


def extract_order_lines_from_tables(rows):
    order_lines = []
    for row in rows:
        row = [clean_cell(x) for x in row]
        if not looks_like_product_row(row):
            continue
        row = row + [""] * max(0, 7 - len(row))
        stt = row[0]
        product_code = normalize_product_code(row[1])
        product_name_pdf = row[2]
        unit_pdf = row[3]
        quantity = parse_number(row[4])
        unit_price = parse_number(row[5])
        line_total = parse_number(row[6])
        if line_total is None and quantity is not None and unit_price is not None:
            line_total = quantity * unit_price
        product_name_clean, product_name_source = clean_product_name(product_code, product_name_pdf)
        unit_clean = apply_common_replacements(unit_pdf)
        if unit_clean is None:
            unit_clean = unit_pdf
        if compact_norm(unit_clean) in ["chiec", "chinc"]:
            unit_clean = "Chiếc"
        order_lines.append({
            "stt": stt,
            "product_code": product_code,
            "product_name_pdf": product_name_pdf,
            "product_name_clean": product_name_clean,
            "product_name_source": product_name_source,
            "unit": unit_clean,
            "unit_pdf": unit_pdf,
            "quantity": quantity,
            "unit_price": unit_price,
            "line_total": line_total,
        })
    return order_lines


# =========================
# VALIDATION
# =========================

def validate_order(header, lines, calculated_total_amount):
    errors = []
    if not header.get("invoice_number"):
        errors.append("Missing invoice_number")
    if not header.get("order_date"):
        errors.append("Missing order_date")
    if not header.get("customer_name"):
        errors.append("Missing customer_name")
    if not header.get("tax_code"):
        errors.append("Missing tax_code")
    if not header.get("address"):
        errors.append("Missing customer address")
    if header.get("address") and not header.get("province_name"):
        errors.append("Cannot infer province from address")
    if len(lines) == 0:
        errors.append("No product lines extracted")
    for idx, line in enumerate(lines, start=1):
        quantity = line.get("quantity")
        unit_price = line.get("unit_price")
        line_total = line.get("line_total")
        if not line.get("product_code"):
            errors.append(f"Missing product_code at line {idx}")
        if quantity is None or quantity <= 0:
            errors.append(f"Invalid quantity at line {idx}")
        if unit_price is None or unit_price < 0:
            errors.append(f"Invalid unit_price at line {idx}")
        if line_total is None or line_total < 0:
            errors.append(f"Invalid line_total at line {idx}")
        if quantity is not None and unit_price is not None and line_total is not None:
            expected = quantity * unit_price
            tolerance = max(1000, expected * 0.02)
            if abs(expected - line_total) > tolerance:
                errors.append(f"Line total mismatch at line {idx}: expected={expected}, actual={line_total}")
    total_amount_pdf = header.get("total_amount_pdf")
    if total_amount_pdf is not None and calculated_total_amount is not None:
        tolerance = max(1000, total_amount_pdf * 0.02)
        if abs(calculated_total_amount - total_amount_pdf) > tolerance:
            errors.append(f"Order total mismatch: calculated={calculated_total_amount}, pdf_total={total_amount_pdf}")
    return errors


# =========================
# SQLITE OUTPUT DB
# =========================

def init_database():
    if DB_PATH.exists():
        DB_PATH.unlink()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE email_log (
        email_id INTEGER PRIMARY KEY AUTOINCREMENT,
        eml_file TEXT,
        message_id TEXT,
        from_address TEXT,
        to_address TEXT,
        subject TEXT,
        received_at TEXT,
        attachment_name TEXT,
        attachment_path TEXT,
        processing_status TEXT,
        error_message TEXT,
        processed_at TEXT
    );
    """)
    cur.execute("""
    CREATE TABLE sales_order (
        order_id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_number TEXT,
        order_date TEXT,
        customer_code TEXT,
        customer_name TEXT,
        tax_code TEXT,
        address TEXT,
        province_name TEXT,
        total_amount REAL,
        total_quantity REAL,
        num_lines INTEGER,
        message_id TEXT,
        received_at TEXT,
        source_from_address TEXT,
        validation_status TEXT,
        validation_errors TEXT,
        created_at TEXT
    );
    """)
    cur.execute("""
    CREATE TABLE order_line (
        line_id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER,
        invoice_number TEXT,
        stt TEXT,
        product_code TEXT,
        product_name_pdf TEXT,
        product_name_clean TEXT,
        product_name_source TEXT,
        unit TEXT,
        unit_pdf TEXT,
        quantity REAL,
        unit_price REAL,
        line_total REAL,
        FOREIGN KEY (order_id) REFERENCES sales_order(order_id)
    );
    """)
    cur.execute("""
    CREATE TABLE fact_sales (
        fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER,
        line_id INTEGER,
        invoice_number TEXT,
        order_date TEXT,
        customer_code TEXT,
        customer_name TEXT,
        tax_code TEXT,
        address TEXT,
        province_name TEXT,
        product_code TEXT,
        product_name_pdf TEXT,
        product_name_clean TEXT,
        product_name_source TEXT,
        unit TEXT,
        quantity REAL,
        unit_price REAL,
        line_total REAL,
        message_id TEXT,
        received_at TEXT
    );
    """)
    conn.commit()
    conn.close()


def insert_email_log_to_db(email_log_df):
    if email_log_df.empty:
        return
    conn = sqlite3.connect(DB_PATH)
    cols = [
        "eml_file", "message_id", "from_address", "to_address", "subject",
        "received_at", "attachment_name", "attachment_path",
        "processing_status", "error_message", "processed_at"
    ]
    for col in cols:
        if col not in email_log_df.columns:
            email_log_df[col] = None
    email_log_df[cols].to_sql("email_log", conn, if_exists="append", index=False)
    conn.close()


def insert_order_to_db(order_row, line_rows):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO sales_order (
        invoice_number, order_date,
        customer_code, customer_name, tax_code,
        address, province_name,
        total_amount, total_quantity, num_lines,
        message_id, received_at, source_from_address,
        validation_status, validation_errors, created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        order_row["invoice_number"], order_row["order_date"], order_row["customer_code"],
        order_row["customer_name"], order_row["tax_code"], order_row["address"],
        order_row["province_name"], order_row["total_amount"], order_row["total_quantity"],
        order_row["num_lines"], order_row["message_id"], order_row["received_at"],
        order_row["source_from_address"], order_row["validation_status"],
        order_row["validation_errors"], order_row["created_at"],
    ))
    order_id = cur.lastrowid
    for line in line_rows:
        cur.execute("""
        INSERT INTO order_line (
            order_id, invoice_number, stt,
            product_code, product_name_pdf,
            product_name_clean, product_name_source,
            unit, unit_pdf,
            quantity, unit_price, line_total
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            order_id, order_row["invoice_number"], line["stt"], line["product_code"],
            line["product_name_pdf"], line["product_name_clean"], line["product_name_source"],
            line["unit"], line["unit_pdf"], line["quantity"], line["unit_price"], line["line_total"],
        ))
    conn.commit()
    conn.close()
    return order_id


def build_fact_sales_in_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("DELETE FROM fact_sales;")
    cur.execute("""
    INSERT INTO fact_sales (
        order_id, line_id, invoice_number, order_date,
        customer_code, customer_name, tax_code,
        address, province_name,
        product_code, product_name_pdf,
        product_name_clean, product_name_source, unit,
        quantity, unit_price, line_total,
        message_id, received_at
    )
    SELECT
        so.order_id,
        ol.line_id,
        so.invoice_number,
        so.order_date,
        so.customer_code,
        so.customer_name,
        so.tax_code,
        so.address,
        so.province_name,
        ol.product_code,
        ol.product_name_pdf,
        ol.product_name_clean,
        ol.product_name_source,
        ol.unit,
        ol.quantity,
        ol.unit_price,
        ol.line_total,
        so.message_id,
        so.received_at
    FROM sales_order so
    JOIN order_line ol ON so.order_id = ol.order_id;
    """)
    conn.commit()
    conn.close()


# =========================
# MAIN
# =========================

def process_all_pdfs():
    print("BASE_DIR:", BASE_DIR)
    print("PDF_DIR:", PDF_DIR)
    print("PDF_DIR exists:", PDF_DIR.exists())
    print("EMAIL_LOG_PATH:", EMAIL_LOG_PATH)
    print("EMAIL_LOG_PATH exists:", EMAIL_LOG_PATH.exists())
    init_database()
    email_log_df = load_email_log()
    insert_email_log_to_db(email_log_df)
    # IMPORTANT: recursive search because PDFs may be stored in subfolders.
    direct_pdf_files = sorted(PDF_DIR.glob("*.pdf"))
    recursive_pdf_files = sorted(PDF_DIR.rglob("*.pdf"))

    # De-duplicate by normalized invoice/PDF stem, keeping first occurrence.
    seen_pdf_keys = set()
    pdf_files = []
    for p in recursive_pdf_files:
        key = normalize_invoice_number(p.stem)
        if key and key not in seen_pdf_keys:
            seen_pdf_keys.add(key)
            pdf_files.append(p)

    print("Direct PDFs:", len(direct_pdf_files))
    print("Recursive PDFs raw:", len(recursive_pdf_files))
    print("Recursive PDFs after de-dup:", len(pdf_files))
    print("Found PDFs:", len(pdf_files))

    if len(pdf_files) == 0:
        print("No PDF files found. Check PDF_DIR.")
        return None

    email_lookup = pd.DataFrame()
    if not email_log_df.empty:
        # Create normalized key so BH26.0935.pdf and BH26_0935.pdf match.
        email_log_df["source_pdf_key"] = email_log_df["source_pdf_file"].apply(
            lambda x: normalize_invoice_number(Path(str(x)).stem)
            if not is_missing_value(x) else None
        )

        lookup_cols = [
            "source_pdf_file", "source_pdf_key", "message_id", "from_address", "received_at",
            "customer_name_from_email_sender", "customer_code_from_email",
            "email_body_invoice_number", "email_body_order_date",
            "email_body_customer_name", "email_body_tax_code",
            "email_body_address", "email_body_province_name",
            "email_body_total_amount_pdf",
            "email_body_total_quantity_email",
            "email_body_num_lines_email",
            "email_body_eml_body_found",
        ]
        for col in lookup_cols:
            if col not in email_log_df.columns:
                email_log_df[col] = None
        email_lookup = email_log_df[lookup_cols].drop_duplicates(subset=["source_pdf_key"])
    sales_order_rows = []
    order_line_rows = []
    error_rows = []
    for pdf_path in tqdm(pdf_files, desc="Processing PDFs"):
        try:
            text = read_pdf_text(pdf_path)
            table_rows = read_pdf_tables(pdf_path)
            header_pdf_table = extract_order_header_from_tables(table_rows, pdf_path.name)
            header_pdf_text = extract_order_header_from_text(text, pdf_path.name)
            message_id = None
            received_at = None
            source_from_address = None
            customer_code_from_email = None
            customer_name_from_email_sender = None
            header_email_body = {
                "invoice_number": None, "order_date": None, "customer_name": None,
                "tax_code": None, "address": None, "province_name": None,
                "total_amount_pdf": None, "total_quantity_email": None, "num_lines_email": None,
            }
            if not email_lookup.empty:
                current_pdf_key = normalize_invoice_number(pdf_path.stem)
                matched = email_lookup[email_lookup["source_pdf_key"] == current_pdf_key]

                # Fallback exact filename match if needed.
                if len(matched) == 0:
                    matched = email_lookup[email_lookup["source_pdf_file"] == pdf_path.name]

                if len(matched) > 0:
                    row_email = matched.iloc[0]
                    message_id = row_email.get("message_id")
                    received_at = row_email.get("received_at")
                    source_from_address = row_email.get("from_address")
                    customer_code_from_email = row_email.get("customer_code_from_email")
                    customer_name_from_email_sender = row_email.get("customer_name_from_email_sender")
                    header_email_body = {
                        "invoice_number": row_email.get("email_body_invoice_number"),
                        "order_date": row_email.get("email_body_order_date"),
                        "customer_name": row_email.get("email_body_customer_name"),
                        "tax_code": row_email.get("email_body_tax_code"),
                        "address": row_email.get("email_body_address"),
                        "province_name": row_email.get("email_body_province_name"),
                        "total_amount_pdf": row_email.get("email_body_total_amount_pdf"),
                        "total_quantity_email": row_email.get("email_body_total_quantity_email"),
                        "num_lines_email": row_email.get("email_body_num_lines_email"),
                    }
            # Ưu tiên email body vì customer_name/address/MST trong PDF bị lỗi font.
            header = merge_headers(header_email_body, header_pdf_table, header_pdf_text)
            lines = extract_order_lines_from_tables(table_rows)
            calculated_total_quantity = sum(line["quantity"] for line in lines if line["quantity"] is not None)
            calculated_total_amount = sum(line["line_total"] for line in lines if line["line_total"] is not None)
            total_amount = safe_parse_float(header.get("total_amount_pdf"), default=calculated_total_amount)
            total_quantity = safe_parse_float(header.get("total_quantity_email"), default=calculated_total_quantity)
            num_lines = safe_parse_int(header.get("num_lines_email"), default=len(lines))
            customer_name = clean_customer_name(header.get("customer_name"))
            if not customer_name:
                customer_name = clean_customer_name(customer_name_from_email_sender)
            customer_code = customer_code_from_email
            invoice_number = coalesce_clean(
                normalize_invoice_number(header.get("invoice_number")),
                normalize_invoice_number(Path(pdf_path).stem))
            address = clean_address(header.get("address"))
            province_name = header.get("province_name") or extract_province_from_address(address)
            validation_header = dict(header)
            validation_header["invoice_number"] = invoice_number
            validation_header["customer_name"] = customer_name
            validation_header["address"] = address
            validation_header["province_name"] = province_name
            validation_errors = validate_order(validation_header, lines, calculated_total_amount)
            validation_status = "VALID" if len(validation_errors) == 0 else "CHECK"
            order_row = {
                "invoice_number": invoice_number,
                "order_date": header.get("order_date"),
                "customer_code": customer_code,
                "customer_name": customer_name,
                "tax_code": normalize_tax_code(header.get("tax_code")),
                "address": address,
                "province_name": province_name,
                "total_amount": total_amount,
                "total_quantity": total_quantity,
                "num_lines": num_lines,
                "message_id": message_id,
                "received_at": received_at,
                "source_from_address": source_from_address,
                "validation_status": validation_status,
                "validation_errors": "; ".join(validation_errors),
                "created_at": datetime.now().isoformat(),
            }
            db_order_id = insert_order_to_db(order_row, lines)
            order_row["db_order_id"] = db_order_id
            sales_order_rows.append(order_row)
            for line in lines:
                order_line_rows.append({
                    "db_order_id": db_order_id,
                    "invoice_number": order_row["invoice_number"],
                    "stt": line["stt"],
                    "product_code": line["product_code"],
                    "product_name_pdf": line["product_name_pdf"],
                    "product_name_clean": line["product_name_clean"],
                    "product_name_source": line["product_name_source"],
                    "unit": line["unit"],
                    "unit_pdf": line["unit_pdf"],
                    "quantity": line["quantity"],
                    "unit_price": line["unit_price"],
                    "line_total": line["line_total"],
                })
            if validation_errors:
                error_rows.append({
                    "invoice_number": order_row["invoice_number"],
                    "source_pdf_file": pdf_path.name,
                    "error_type": "VALIDATION_ERROR",
                    "error_message": "; ".join(validation_errors),
                })
        except Exception as e:
            error_rows.append({
                "invoice_number": normalize_invoice_number(Path(pdf_path).stem),
                "source_pdf_file": pdf_path.name,
                "error_type": "PROCESSING_ERROR",
                "error_message": str(e),
            })
    sales_order_df = pd.DataFrame(sales_order_rows)
    order_line_df = pd.DataFrame(order_line_rows)
    error_df = pd.DataFrame(error_rows)
    sales_order_cols = [
        "db_order_id", "invoice_number", "order_date", "customer_code", "customer_name", "tax_code",
        "address", "province_name", "total_amount", "total_quantity", "num_lines",
        "message_id", "received_at", "source_from_address", "validation_status", "validation_errors", "created_at"
    ]
    if not sales_order_df.empty:
        sales_order_df = sales_order_df.reindex(columns=sales_order_cols)
    order_line_cols = [
        "db_order_id", "invoice_number", "stt", "product_code", "product_name_pdf",
        "product_name_clean", "product_name_source", "unit", "unit_pdf",
        "quantity", "unit_price", "line_total"
    ]
    if not order_line_df.empty:
        order_line_df = order_line_df.reindex(columns=order_line_cols)
        order_line_df["product_code"] = order_line_df["product_code"].astype("string")
    build_fact_sales_in_db()
    conn = sqlite3.connect(DB_PATH)
    fact_sales_df = pd.read_sql_query("SELECT * FROM fact_sales", conn)
    conn.close()
    if not fact_sales_df.empty:
        fact_sales_df = fact_sales_df.rename(columns={"order_id": "db_order_id", "line_id": "db_line_id"})
        fact_sales_cols = [
            "fact_id", "db_order_id", "db_line_id", "invoice_number", "order_date",
            "customer_code", "customer_name", "tax_code", "address", "province_name",
            "product_code", "product_name_pdf", "product_name_clean", "product_name_source", "unit",
            "quantity", "unit_price", "line_total", "message_id", "received_at"
        ]
        fact_sales_df = fact_sales_df.reindex(columns=fact_sales_cols)
        fact_sales_df["product_code"] = fact_sales_df["product_code"].astype("string")
    sales_order_df.to_csv(SALES_ORDER_CSV, index=False, encoding="utf-8-sig")
    order_line_df.to_csv(ORDER_LINE_CSV, index=False, encoding="utf-8-sig")
    fact_sales_df.to_csv(FACT_SALES_CSV, index=False, encoding="utf-8-sig")
    error_df.to_csv(PDF_ERROR_LOG_CSV, index=False, encoding="utf-8-sig")
    try:
        sales_order_df.to_excel(SALES_ORDER_XLSX, index=False)
        order_line_df.to_excel(ORDER_LINE_XLSX, index=False)
        fact_sales_df.to_excel(FACT_SALES_XLSX, index=False)
    except Exception as e:
        print("Warning: Could not export xlsx files. Install openpyxl if needed.")
        print("XLSX export error:", e)
    print("\nDone.")
    print("Saved sales_order CSV:", SALES_ORDER_CSV)
    print("Saved order_line CSV:", ORDER_LINE_CSV)
    print("Saved fact_sales CSV:", FACT_SALES_CSV)
    print("Saved PDF error log:", PDF_ERROR_LOG_CSV)
    print("Saved SQLite database:", DB_PATH)
    print("\nSummary:")
    print("email_log rows:", len(email_log_df))
    print("processed PDF files:", len(pdf_files))
    print("sales_order rows:", len(sales_order_df))
    print("order_line rows:", len(order_line_df))
    print("fact_sales rows:", len(fact_sales_df))
    print("error/check rows:", len(error_df))
    if len(sales_order_df) > 0:
        print("\nValidation status:")
        print(sales_order_df["validation_status"].value_counts(dropna=False))
        print("\nMissing values in sales_order:")
        print(sales_order_df.isna().sum())
        print("\nAddress/province check:")
        print(sales_order_df[[
            "invoice_number", "customer_name", "address", "province_name", "validation_status", "validation_errors"
        ]].head(15))
    if len(order_line_df) > 0:
        print("\nMissing values in order_line:")
        print(order_line_df.isna().sum())
        print("\nProduct name source:")
        print(order_line_df["product_name_source"].value_counts(dropna=False))
        print("\nSample cleaned product names:")
        print(order_line_df[["product_code", "product_name_pdf", "product_name_clean", "product_name_source"]].head(15))
    return sales_order_df, order_line_df, fact_sales_df, error_df


if __name__ == "__main__":
    process_all_pdfs()
