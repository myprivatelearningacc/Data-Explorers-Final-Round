import os
import email
from email import policy
from email.parser import BytesParser
from pathlib import Path
from datetime import datetime
from tqdm import tqdm
import pandas as pd

# =========================
# CONFIG
# =========================

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

RAW_EMAIL_DIR = BASE_DIR / "data source/email/tnbike_emails_mar2026"
ATTACHMENT_DIR = BASE_DIR / "extracted_attachments" / "email"
OUTPUT_DIR = BASE_DIR / "processed_outputs_emails"

ATTACHMENT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print("BASE_DIR:", BASE_DIR)
print("RAW_EMAIL_DIR:", RAW_EMAIL_DIR)
print("RAW_EMAIL_DIR exists:", RAW_EMAIL_DIR.exists())

if RAW_EMAIL_DIR.exists():
    eml_preview = list(RAW_EMAIL_DIR.glob("*.eml"))
    print("Number of .eml files:", len(eml_preview))
    print("First 5 files:", eml_preview[:5])
else:
    print("Folder does not exist. Check folder name, especially spaces.")

# =========================
# HELPER FUNCTIONS
# =========================

def safe_filename(name):
    """
    Làm sạch tên file để tránh lỗi khi lưu attachment.
    """
    if name is None:
        return "unknown.pdf"

    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid_chars:
        name = name.replace(ch, "_")

    return name.strip()


def parse_eml_file(eml_path):
    """
    Đọc 1 file .eml, lấy metadata email và tách PDF attachment.
    """
    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    message_id = msg.get("Message-ID")
    from_address = msg.get("From")
    to_address = msg.get("To")
    subject = msg.get("Subject")
    received_at = msg.get("Date")

    attachments = []

    for part in msg.walk():
        content_disposition = part.get_content_disposition()
        filename = part.get_filename()
        content_type = part.get_content_type()

        is_attachment = content_disposition == "attachment"
        is_pdf = (
            filename is not None and filename.lower().endswith(".pdf")
        ) or content_type == "application/pdf"

        if is_attachment and is_pdf:
            payload = part.get_payload(decode=True)

            if payload is None:
                continue

            if filename is None:
                filename = f"{eml_path.stem}.pdf"

            filename = safe_filename(filename)

            if Path(filename).stem.startswith(eml_path.stem):
                saved_filename = filename
            else:
                saved_filename = f"{eml_path.stem}_{filename}"

            attachment_path = ATTACHMENT_DIR / saved_filename

            with open(attachment_path, "wb") as f:
                f.write(payload)

            attachments.append({
                "attachment_name": saved_filename,
                "attachment_path": str(attachment_path),
                "content_type": content_type
            })

    return {
        "eml_file": eml_path.name,
        "message_id": message_id,
        "from_address": from_address,
        "to_address": to_address,
        "subject": subject,
        "received_at": received_at,
        "num_pdf_attachments": len(attachments),
        "attachments": attachments
    }


# =========================
# MAIN PIPELINE
# =========================

def process_all_emails():
    eml_files = sorted(RAW_EMAIL_DIR.glob("*.eml"))

    print(f"Found {len(eml_files)} .eml files in: {RAW_EMAIL_DIR}")

    if len(eml_files) == 0:
        print("No .eml files found. Please check RAW_EMAIL_DIR.")
        return None

    email_log_rows = []
    processing_rows = []

    for eml_path in tqdm(eml_files, desc="Processing .eml files"):
        try:
            result = parse_eml_file(eml_path)

            if result["num_pdf_attachments"] == 0:
                email_log_rows.append({
                    "eml_file": result["eml_file"],
                    "message_id": result["message_id"],
                    "from_address": result["from_address"],
                    "to_address": result["to_address"],
                    "subject": result["subject"],
                    "received_at": result["received_at"],
                    "attachment_name": None,
                    "attachment_path": None,
                    "processing_status": "NO_PDF_ATTACHMENT",
                    "error_message": None,
                    "processed_at": datetime.now().isoformat()
                })

                processing_rows.append({
                    "eml_file": result["eml_file"],
                    "status": "NO_PDF_ATTACHMENT",
                    "num_pdf_attachments": 0,
                    "error_message": None
                })

            else:
                for attachment in result["attachments"]:
                    email_log_rows.append({
                        "eml_file": result["eml_file"],
                        "message_id": result["message_id"],
                        "from_address": result["from_address"],
                        "to_address": result["to_address"],
                        "subject": result["subject"],
                        "received_at": result["received_at"],
                        "attachment_name": attachment["attachment_name"],
                        "attachment_path": attachment["attachment_path"],
                        "processing_status": "PDF_EXTRACTED",
                        "error_message": None,
                        "processed_at": datetime.now().isoformat()
                    })

                processing_rows.append({
                    "eml_file": result["eml_file"],
                    "status": "SUCCESS",
                    "num_pdf_attachments": result["num_pdf_attachments"],
                    "error_message": None
                })

        except Exception as e:
            email_log_rows.append({
                "eml_file": eml_path.name,
                "message_id": None,
                "from_address": None,
                "to_address": None,
                "subject": None,
                "received_at": None,
                "attachment_name": None,
                "attachment_path": None,
                "processing_status": "FAILED",
                "error_message": str(e),
                "processed_at": datetime.now().isoformat()
            })

            processing_rows.append({
                "eml_file": eml_path.name,
                "status": "FAILED",
                "num_pdf_attachments": 0,
                "error_message": str(e)
            })

    email_log_df = pd.DataFrame(email_log_rows)
    processing_report_df = pd.DataFrame(processing_rows)

    email_log_path = OUTPUT_DIR / "email_log.csv"
    processing_report_path = OUTPUT_DIR / "processing_report.csv"

    email_log_df.to_csv(email_log_path, index=False, encoding="utf-8-sig")
    processing_report_df.to_csv(processing_report_path, index=False, encoding="utf-8-sig")

    print("\nDone.")
    print(f"Email log saved to: {email_log_path}")
    print(f"Processing report saved to: {processing_report_path}")
    print(f"PDF attachments saved to: {ATTACHMENT_DIR}")

    print("\nSummary:")
    print(processing_report_df["status"].value_counts())

    print("\nTotal extracted PDF files:", email_log_df["attachment_name"].notna().sum())

    return email_log_df, processing_report_df

# =========================
# RUN
# =========================

result = process_all_emails()

if result is not None:
    email_log_df, processing_report_df = result
else:
    print("Pipeline stopped because no .eml files were found.")