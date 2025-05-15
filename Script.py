import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import smtplib
import time
from datetime import datetime
import csv
from email.message import EmailMessage

GMAIL_ADDRESS = "rohansuri.prime@gmail.com"   # Your Gmail
APP_PASSWORD = "mgzy buig ycsm cmib"     # App password from Gmail
REPLY_TO = "rohan.suri@yale.edu"             # Where replies should go
SHEET_NAME = "Chem Contact Info"
CREDENTIALS_FILE = "email-automation-459819-c5fd66abf44b.json"
SIGNATURE = """
Best regards,  
Rohan Suri  
B.S. Candidate, Chemistry  
Yale University  
rohan.suri@yale.edu
"""


def load_sheet(sheet_name: str, creds_path: str):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    client = gspread.authorize(creds)

    sheet = client.open(sheet_name).sheet1 #grabs first tab in sheet
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    return df

def clean_data(df):
	for i, item in df["Institution"].items():
		if item.split()[0] == "University":
			df.loc[i, "Institution"] = "the " + item
		elif "University" in item:
			df.loc[i, "Institution"] = item.rsplit(' ', 1)[0]
	return df

def load_template(filepath="TEMPLATE.txt"):
	with open(filepath, "r") as file:
		return file.read()

def generate_emails(df, template):
	emails = []
	for _, row in df.iterrows():
		email_body = template.format(
			Name=row["Name"],
			Institution=row["Institution"],
			Custom_Note=row["Custom Note"],
			Salutation=row["Salutation"]
			) + "\n" + SIGNATURE
		emails.append({
			"to": row["Email"],
			"body": email_body
			})
	return emails

def build_email_message(email_dict):
	msg = EmailMessage()
	msg["From"] = f"Rohan Suri <{GMAIL_ADDRESS}>"
	msg["To"] = email_dict["to"]
	msg["Subject"] = "Exploring AI in Chemistry - Quick Chat?"
	msg["Reply-To"] = REPLY_TO
	msg.set_content(email_dict["body"])
	msg.add_alternative(f"""
<html>
  <body>
    <p>{email_dict["body"].replace('\n', '<br>')}</p>
  </body>
</html>
""", subtype="html")
	return msg

def log_email_sent(email_dict):
	with open("sent_log.csv", "a", newline='') as logfile:
		writer = csv.writer(logfile)
		writer.writerow([
			 email_dict["to"],
			 email_dict.get("salutation", ""),  # Optional
			 datetime.now().isoformat()
			 ])

def load_sent_emails(log_path="sent_log.csv"):
	sent = set()
	try:
		with open(log_path, "r") as f:
			reader = csv.reader(f)
			for row in reader:
				sent.add(row[0])  # assumes first column is email
	except FileNotFoundError:
		pass  # no log yet
	return sent

df = load_sheet(SHEET_NAME, CREDENTIALS_FILE)
df = clean_data(df)

template = load_template()
emails_list = generate_emails(df, template)
sent_emails = load_sent_emails()


with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
	smtp.login(GMAIL_ADDRESS, APP_PASSWORD)
	for email in emails_list:
		if email["to"] in sent_emails:
			print(f"Skipping {email['to']} (already sent)")
			continue

		msg = build_email_message(email)
		print(f"Sending to {msg['To']}...")
		smtp.send_message(msg)
		log_email_sent(email)
		time.sleep(30)  # Delay to avoid triggering spam filters
