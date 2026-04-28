# Email Sender

A Python script that sends personalised cold emails to business_owners owners in bulk, sourced from an Excel spreadsheet. Tracks which emails have already been sent to avoid duplicates.

## Features

- Reads recipients from an Excel file (`Restaurant_Owners.xlsx`)
- Personalises each email with the owner's name and business name
- Gracefully handles missing owner names by removing that line from the email
- Marks rows as **Sent** in the spreadsheet after each successful send
- Adds a random delay (30–60 seconds) between emails to reduce spam-filter risk
- Saves the Excel file atomically (temp-file swap) to avoid corruption if the file is open

## Requirements

- Python 3.8+
- A Gmail account with an **App Password** enabled (2FA required)

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/your-username/email-sender.git
cd email-sender
```

### 2. Create and activate a virtual environment

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install openpyxl python-dotenv
```

### 4. Configure environment variables

Create a `.env` file in the project root:

```env
SENDER_EMAIL=you@gmail.com
SENDER_PASSWORD=your_app_password_here
```

> **Important:** Use a [Gmail App Password](https://support.google.com/accounts/answer/185833), not your regular Gmail password.

### 5. Prepare the Excel file

Create `Restaurant_Owners.xlsx` with the following columns (order does not matter):

| Business Name | Owner Name | Email            | Sent |
| ------------- | ---------- | ---------------- | ---- |
| Pizza Palace  | John Smith | john@example.com |      |
| Burger Barn   |            | info@burgerb.com |      |

- **Business Name** — required
- **Owner Name** — optional; if blank, the greeting line is removed automatically
- **Email** — required
- **Sent** — leave blank; the script writes `Yes` here after a successful send
- **important** close excel file when running script so it can write in the file after sending email, it marks **Sent** column to "Yes"

## Usage

Open `send_emails.py` and adjust the configuration at the top of the file:

```python
EMAILS_TO_SEND = 5  # Number of emails to send per run
```

Then run the script:

```bash
python send_emails.py
```

### Example output

```
Eligible rows (not yet sent): 12
  [OK] Sent to john@example.com (Pizza Palace)
  Waiting 47 seconds before next email...
  [OK] Sent to info@burgerb.com (Burger Barn)
  ...
Done. 5/5 email(s) sent. Excel file updated.
```

## Configuration Reference

| Variable          | Description                                                              |
| ----------------- | ------------------------------------------------------------------------ |
| `SMTP_SERVER`     | SMTP host (default: `smtp.gmail.com`)                                    |
| `SMTP_PORT`       | SMTP port (default: `587`)                                               |
| `SENDER_EMAIL`    | Loaded from `.env`                                                       |
| `SENDER_PASSWORD` | Loaded from `.env` (use an App Password)                                 |
| `EXCEL_FILE`      | Path to the spreadsheet (default: `Restaurant_Owners.xlsx`)              |
| `EMAILS_TO_SEND`  | Max emails to send per run                                               |
| `EMAIL_SUBJECT`   | Subject line template (`{business_name}` placeholder supported)          |
| `EMAIL_TEMPLATE`  | Body template (`{owner_name}`, `{business_name}` placeholders supported) |

## Security Notes

- Never commit `.env` to version control — add it to `.gitignore`
- Always use an App Password, not your main Gmail password
- The script connects over TLS (STARTTLS on port 587)

## License

MIT
