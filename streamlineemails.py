import os
import email
import imaplib
import json
import pandas as pd
from groq import Groq
from email import policy
from email.parser import BytesParser
from PIL import Image
import cv2

# et up your Gmail IMAP credentials

IMAP_SERVER = "imap.gmail.com"
EMAIL_ACCOUNT = "xxxxx@gmail.com"
EMAIL_PASSWORD = "xxxxxxxxxxx"
IMAP_PORT = 993

# Initialize Groq Client
os.environ["GROQ_API_KEY"] = "xxxxxxxxxxxxxxxxx"
groq_client = Groq(api_key=os.getenv("GROQ_API_KEY"))

# Create DataFrame for extracted data
df = pd.DataFrame(columns=[
    "Abhasi Name", "Abhasi ID", "Your Centre",
    "Recipient Name", "Recipient Phone", "Date of Sharing"
])

# Connect to Gmail IMAP
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
mail.select("inbox")  # Select the inbox folder

# Search for unread emails
status, email_ids = mail.search(None, "UNSEEN")
email_ids = email_ids[0].split()

for email_id in email_ids:
    # Fetch the email by ID
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = BytesParser(policy=policy.default).parsebytes(raw_email)

    # Extract the email content
    email_content = ""
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() == "text/plain":
                email_content = part.get_payload(decode=True).decode()
                break
    else:
        email_content = msg.get_payload(decode=True).decode()

    # Use Groq to extract structured data from email content
    prompt = f"""
    Extract the following structured information from the email text and return it strictly as a JSON object:

    - Abhasi Name
    - Abhasi ID
    - Your Centre
    - Recipient Name(s)
    - Recipient Phone(s)
    - Date(s) of Sharing

    Email Content:
    \"\"\"{email_content}\"\"\"

    ### IMPORTANT INSTRUCTIONS:
    - ONLY return a JSON object.
    - Do NOT include explanations, comments, or extra text.
    - Format your response **EXACTLY** like this:
    {{
        "Abhasi Name": "John Doe",
        "Abhasi ID": "INPSAG019",
        "Your Centre": "Jaipur",
        "Recipients": [
            {{"Recipient Name": "Prashant", "Recipient Phone": "9384759642", "Date of Sharing": "2/16/2025"}}
        ]
    }}

    ONLY return JSON. No other text.
    """


    response = groq_client.chat.completions.create(
        model="llama3-70b-8192",
        messages=[{"role": "system", "content": "You are an AI that extracts structured data from text."},
                    {"role": "user", "content": prompt}],
        temperature=0.3
    )
    # Debugging: Print response
    print("Raw Groq Response:", response.choices[0].message.content)
    # Parse the Groq response and append to DataFrame
    try:
        if response.choices and response.choices[0].message.content:
            groq_data = json.loads(response.choices[0].message.content)
        else:
            print("Error: Empty response from Groq.")
            exit()
    except json.JSONDecodeError:
        print("JSON decoding failed. Groq Response:", response.choices[0].message.content)
        exit()
    recipient_rows = []
    for recipient in groq_data["Recipients"]:
        row = {
            "Abhasi Name": groq_data["Abhasi Name"],
            "Abhasi ID": groq_data["Abhasi ID"],
            "Your Centre": groq_data["Your Centre"],
            "Recipient Name": recipient["Recipient Name"],
            "Recipient Phone": recipient["Recipient Phone"],
            "Date of Sharing": recipient["Date of Sharing"]
        }
        recipient_rows.append(row)
    df = pd.concat([df, pd.DataFrame(recipient_rows)], ignore_index=True)

    # Save to Excel
df.to_excel('extracted_emails.xlsx', index=False)

print("Data saved successfully!")
