import os
import shutil
import smtplib
import tempfile
import time
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import openpyxl
from dotenv import load_dotenv

load_dotenv()

# ── CONFIGURATION ─────────────────────────────────────────────────────────────

SMTP_SERVER     = "smtp.gmail.com"
SMTP_PORT       = 587
SENDER_EMAIL    = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")  # Use an App Password for Gmail

EXCEL_FILE     = "Restaurant_Owners.xlsx"
EMAILS_TO_SEND = 5  # ← Change this to control how many emails to send per run

EMAIL_SUBJECT = "Quick question about {business_name}"

# Use {owner_name} and {business_name} as placeholders.
# Lines that contain {owner_name} are automatically removed when owner name is missing.
EMAIL_TEMPLATE = """\
Hi,
{owner_name}, I know you're worried about missing out on customer orders and reservations by solely relying on staff to cover them all, instead of utilizing the power of AI to strengthen your business communication 24/7.

With that in mind I created a tool that will help {business_name} take more orders, reservations and answer customer queries simultaneously. This could be a game changer for your business.
Mind if I send over a demo of how it works?

Warm regards,
Ameer Hamza Malik
AI Automation for Restaurants 

"""

# ─────────────────────────────────────────────────────────────────────────────


def build_email_body(template: str, owner_name: str, business_name: str) -> str:
    """Inject variables into the template. Removes owner_name lines when empty."""
    if not owner_name:
        template = template.replace("{owner_name}, ", "")
        return template.format(business_name=business_name)

    return template.format(owner_name=owner_name, business_name=business_name)


def send_email(to_address: str, subject: str, body: str) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = to_address
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_address, msg.as_string())


def main():
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        raise ValueError("SENDER_EMAIL and SENDER_PASSWORD must be set in .env")

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active

    # Locate columns by header name so column order doesn't matter
    headers = {cell.value: cell.column for cell in ws[1]}
    col_business = headers.get("Business Name")
    col_owner    = headers.get("Owner Name")
    col_email    = headers.get("Email")
    col_sent     = headers.get("Sent")

    if not all([col_business, col_email, col_sent]):
        raise ValueError(
            "Required columns (Business Name, Email, Sent) not found in the sheet."
        )

    # Collect rows eligible to send: Sent column is empty
    eligible_rows = []
    for row in ws.iter_rows(min_row=2):
        sent_val = row[col_sent - 1].value
        if sent_val and str(sent_val).strip().lower() == "yes":
            continue  # already sent

        business_name = row[col_business - 1].value
        email_addr    = row[col_email - 1].value

        # Skip rows missing required fields
        if not business_name or not email_addr:
            continue

        eligible_rows.append(row)

    available = len(eligible_rows)
    print(f"Eligible rows (not yet sent): {available}")

    if available == 0:
        print("No emails to send.")
        return

    if EMAILS_TO_SEND > available:
        print(
            f"Warning: EMAILS_TO_SEND ({EMAILS_TO_SEND}) exceeds available rows "
            f"({available}). Sending {available} email(s)."
        )

    to_send = min(EMAILS_TO_SEND, available)
    sent_count = 0

    for row in eligible_rows[:to_send]:
        business_name = str(row[col_business - 1].value).strip()
        owner_name = (
            str(row[col_owner - 1].value).strip()
            if col_owner and row[col_owner - 1].value
            else ""
        )
        email_addr = str(row[col_email - 1].value).strip()

        subject = EMAIL_SUBJECT.format(business_name=business_name)
        body    = build_email_body(EMAIL_TEMPLATE, owner_name, business_name)

        try:
            send_email(email_addr, subject, body)
            row[col_sent - 1].value = "Yes"
            sent_count += 1
            print(f"  [OK] Sent to {email_addr} ({business_name})")
            if sent_count < to_send:
                delay = random.randint(30, 60)
                print(f"  Waiting {delay} seconds before next email...")
                time.sleep(delay)
        except Exception as exc:
            print(f"  [FAIL] {email_addr} ({business_name}): {exc}")

    # Save to a temp file first, then replace the original.
    # This avoids PermissionError when the Excel file is open in another program.
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(os.path.abspath(EXCEL_FILE)))
    os.close(tmp_fd)
    try:
        wb.save(tmp_path)
        shutil.move(tmp_path, EXCEL_FILE)
    except Exception:
        os.remove(tmp_path)
        raise
    print(f"\nDone. {sent_count}/{to_send} email(s) sent. Excel file updated.")


if __name__ == "__main__":
    main()
