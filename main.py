import os
import time
import smtplib
from datetime import datetime
import logging
import subprocess
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ---------------- Logging Setup ---------------- #
logging.basicConfig(
    filename='smtp_mail.log',
    filemode='a',
    level=logging.DEBUG,
    format='%(asctime)s | %(levelname)s | %(message)s',
)

line_break = "---"

# ---------------- Config ---------------- #
SMTP_SERVER = "smtp.email.ap-mumbai-1.oci.oraclecloud.com"
SMTP_PORT = 587
USERNAME = "ocid1.user.oc1..aaaaaaaa2c3ixyurrci5w2nq5ghop7yamjogmf4drfttpr5vlk4x4heont7q@ocid1.tenancy.oc1..aaaaaaaayjeh3xfidtcwqz4juu6wrhve5xnakxt5mw3fdlzqmyhvh3vnfi6a.m8.com"
PASSWORD = "4$DDNvkYQzfg]uKc<k05"
FROM_ADDRESS = "kimbal-alerts@apdclintelli.in"
TO_ADDRESS = [
    "harojyoti.borah@kimbal.io",
    "Yash.jain@kimbal.io",
    "akhtar.nawaz@kimbal.io",
    "saiful.khan@kimbal.io",
    "naman.kumar@kimbal.io"
    ]

MAX_ATTACHMENT_SIZE_MB = 10  # Set max size per attachment
MAX_ATTACHMENT_SIZE_BYTES = MAX_ATTACHMENT_SIZE_MB * 1024 * 1024

# ---------------- Run PowerShell Script ---------------- #
def run_powershell_script(script_path):
    try:
        logging.info(f"Running PowerShell script: {script_path}")
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path],
            capture_output=True,
            text=True
        )
        if result.returncode != 0:
            logging.error(f"PowerShell script failed:\n{result.stderr}")
            print("❌ PowerShell script failed:", result.stderr)
            logging.info(f"{50 * line_break}")
            return False
        else:
            logging.info("PowerShell script executed successfully.")
            print("✅ PowerShell script executed.")
            logging.info(f"{50 * line_break}")
            return True
    except Exception as e:
        logging.exception("Error running PowerShell script")
        print("❌ Failed to run PowerShell script:", str(e))
        logging.info(f"{50 * line_break}")
        return False

# ---------------- Get All Today's ZIP Parts ---------------- #
def get_all_zip_parts(directory, base_filename=""):
    date_string = datetime.now().strftime("%Y-%m-%d")
    zip_files = []
    for file in os.listdir(directory):
        if file.startswith(f"{base_filename}_{date_string}") and file.endswith(".7z"):
            zip_files.append(os.path.join(directory, file))
    zip_files.sort()
    return zip_files

# ---------------- Construct and Send Email ---------------- #
def send_email_with_attachment(attachment_path):
    file_size = os.path.getsize(attachment_path)
    file_size_mb = file_size / (1024 * 1024)

    if file_size > MAX_ATTACHMENT_SIZE_BYTES:
        message = (
            f"❌ Skipped file (exceeds size limit): {attachment_path} "
            f"({file_size_mb:.2f} MB > {MAX_ATTACHMENT_SIZE_MB} MB limit)"
        )
        logging.warning(message)
        print(message)
        logging.info(f"{50 * line_break}")
        return
    else:
        logging.info(f"✅ File within size limit: {attachment_path} ({file_size_mb:.2f} MB)")

    basename = os.path.basename(attachment_path)
    part_number = basename.split("Part")[-1].replace(".7z", "") if "Part" in basename else basename.replace(".7z", "")
    subject = f"Single Phase Billing Data Dump || APDCL Pkg 7 || File_Number: {part_number}"
    body = (
        "Hi Team,\n\n"
        "Please find attached the validated and unique Single Phase Billing Data Dump for today.\n\n"
        "Details:\n"
        "- Report Name: SinglePhaseBillingProfile\n"
        f"- Date: {datetime.now().strftime('%d-%b-%Y')}\n"
        "- Format: .7z (Extract this using WinRar etc.)\n\n"
        "If you have any queries or require further assistance, feel free to reach out to me at naman.kumar@kimbal.io.\n\n"
        "Regards,\n"
        "Kimbal Alerts\n\n"
        "-----------------------------\n"
        "Note: This is a system-generated email. Please do not reply."
    )

    msg = MIMEMultipart()
    msg['From'] = FROM_ADDRESS
    msg['To'] = ', '.join(TO_ADDRESS)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with open(attachment_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=basename)
        part['Content-Disposition'] = f'attachment; filename="{basename}"'
        msg.attach(part)
        logging.info(f"📎 Attached file: {basename}")
    except Exception as e:
        logging.exception("⚠️ Error attaching file")
        print("⚠️ Could not attach file:", str(e))
        logging.info(f"{50 * line_break}")
        return

    try:
        logging.info(f"📡 Connecting to SMTP server {SMTP_SERVER}:{SMTP_PORT}")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            logging.debug("🔐 TLS started successfully")
            server.login(USERNAME, PASSWORD)
            logging.debug(f"🔓 Logged in as {USERNAME}")
            server.sendmail(FROM_ADDRESS, TO_ADDRESS, msg.as_string())
            logging.info(f"✅ Email sent for: {basename}\n{50 * line_break}")
            print(f"✅ Email sent for: {basename}")
            time.sleep(10)

    except smtplib.SMTPAuthenticationError as e:
        error_msg = e.smtp_error.decode() if hasattr(e.smtp_error, 'decode') else str(e)
        logging.error(f"❌ SMTP Authentication failed: {error_msg}\n{50 * line_break}")
        print("❌ Authentication failed:", error_msg)
        time.sleep(10)

    except Exception as e:
        logging.exception(f"❌ Failed to send email\n{50 * line_break}")
        print(f"❌ Failed to send email: \n{str(e)}\n{50 * line_break}")
        time.sleep(10)


# ---------------- Main ---------------- #
if __name__ == "__main__":
    powershell_script_path = "generate_sql_report.ps1"
    zip_directory = "D:\\BillingProfileDataEmailer"
    base_zip_name = "SinglePhaseBillingProfile"

    # Run PowerShell script and exit if it fails
    if not run_powershell_script(powershell_script_path):
        logging.info(f"{50 * line_break}")
        exit(1)

    zip_files = get_all_zip_parts(zip_directory, base_zip_name)

    if not zip_files:
        print(f"❌ No ZIP files found for today. Email not sent.\n{50 * line_break}")
        logging.info(f"❌ No ZIP files found for today. Email not sent.\n{50 * line_break}")
    else:
        sent_count = 0
        now = time.time()
        one_hour_seconds = 3600

        for zip_file in zip_files:
            file_mod_time = os.path.getmtime(zip_file)
            age_seconds = now - file_mod_time

            if age_seconds > one_hour_seconds:
                logging.info(f"⏳ Skipping stale file older than 1 hour: {zip_file}")
                print(f"⏳ Skipping stale file older than 1 hour: {zip_file}")
                logging.info(f"{50 * line_break}")
                continue

            print(f"📤 Sending email for: {zip_file}")
            send_email_with_attachment(zip_file)
            sent_count += 1

        print(f"✅ Total emails sent: {sent_count}")
        logging.info(f"✅ Total emails sent: {sent_count}\n{50 * line_break}")
