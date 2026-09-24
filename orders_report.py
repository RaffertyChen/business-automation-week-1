import csv
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook

#all the variables you can change
folder = Path(__file__).parent 
wb = load_workbook(folder / "orders.xlsx")
ws = wb.active
PRICE_LIMIT = 50
output = folder / "expensive_orders.csv"

#list
orders = []
expensive = []

# Read every order in the file and skipping rows with no date 
for row in ws.iter_rows(min_row=2, values_only=True):
    date = row[3]
    if date is None:
        continue
    orders.append(row)


#to find the rows with the price over the PRICE_LIMIT
for row in orders:
    price = row[2]
    if price >= PRICE_LIMIT:
        expensive.append(row)

#Save the expensive orders to a CSV file
with open(output, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["name", "item", "price", "date"])
    writer.writerows(expensive)

#to load the proper login details stored in the .env file
load_dotenv(folder / ".env")

#login details
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_RECIPIENT = os.getenv("GMAIL_RECIPIENT")
GMAIL_PASSWORD = os.getenv("GMAIL_PASSWORD")

#the email part
msg = EmailMessage()
msg["From"] = GMAIL_USER
msg["To"] = GMAIL_RECIPIENT
msg["Subject"] = f"{len(expensive)} orders over {PRICE_LIMIT}"
msg.set_content("This is a test email sent from Python.")

#Attach the CSV
with open(output, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="text",
        subtype="csv",
        filename=output.name,
    )

#Log in to Gmail and send
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(GMAIL_USER, GMAIL_PASSWORD)
    smtp.send_message(msg)

print("Email sent")