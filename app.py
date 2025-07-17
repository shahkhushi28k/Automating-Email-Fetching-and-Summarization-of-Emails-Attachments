import os
import re
import time
import pandas as pd
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
import openai
from dotenv import load_dotenv
import pytesseract
from PIL import Image
import pdfplumber
import docx
import io
import base64
from PIL import Image
from pdf2image import convert_from_path 

# Loading env vars from the .env file
load_dotenv()

# Getting the OpenAI API key from the env
openai.api_key = os.getenv("OPENAI_API_KEY")

# SCOPE
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Timestamp storage file
TIMESTAMP_FILE = "last_timestamp.txt"
CSV_FILE = "data.csv"

# Authenticate Gmail API
def authenticate_gmail_api():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=8080)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

# Reading the last saved timestamp
def get_last_timestamp():
    if os.path.exists(TIMESTAMP_FILE):
        with open(TIMESTAMP_FILE, "r") as f:
            return f.read().strip()
    return None

# Saving the last email timestamp for the next execution
def save_last_timestamp(timestamp):
    with open(TIMESTAMP_FILE, "w") as f:
        f.write(timestamp)

# Extract sender email and name
def extract_sender(sender_info):
    match = re.match(r"(.*)<(.*)>", sender_info)
    if match:
        return match.group(1).strip(), match.group(2).strip()  
    return sender_info, sender_info  

# Summarization 
def summarize_email_with_openai(body):
    if body:
        response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "system", "content": "Summarize the following text concisely, preserving key information without adding explanations or commentary."},
                  {"role": "user", "content": body}],
        max_tokens=100
        )

        return response.choices[0].message["content"].strip()
    return "No content"

def summarize_email(body):
    return summarize_email_with_openai(body)


# Converting email date to timestamp 
def parse_email_date(date):
    cleaned_date = re.sub(r"\s*\(UTC\)", "", date)  
    try:
        return datetime.strptime(cleaned_date, "%a, %d %b %Y %H:%M:%S %z").timestamp()
    except ValueError:
        try:
            return datetime.strptime(cleaned_date, "%a, %d %b %Y %H:%M:%S %Z").timestamp()
        except ValueError:
            print(f"Error parsing date: {date}")
            return None  

# Fetch emails from Gmail
def get_attachment_type(msg_payload):
    if 'parts' not in msg_payload:
        return "NILL"
    
    for part in msg_payload['parts']:
        if 'filename' in part and part['filename']:
            filename = part['filename'].lower()
            
            if filename.endswith(('.doc', '.docx')):
                return "Word Document"
            elif filename.endswith('.pdf'):
                return "PDF"
            elif filename.endswith(('.txt', '.csv')):
                return "Text File"
            elif filename.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp', '.heic')):
                return "Image"
            else:
                return "Other"
    
    return "NILL"

# Tesseract path 
pytesseract.pytesseract.tesseract_cmd = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
def extract_text_from_image(image_path):
    """Extract text from an image using OCR."""
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        return text.strip()
    except Exception as e:
        return f"Error extracting text from image: {e}"

#Extract text from documents (Word/PDF/TextFile)
def extract_text_from_file(file_path):
    """Extract text from PDFs (including scanned images), Word, and Text files."""
    try:
        extracted_text = ""

        if file_path.lower().endswith('.pdf'):
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                   
                    page_text = page.extract_text() or ''
                    extracted_text += page_text + "\n"

                # For an image-based PDF
                if not extracted_text.strip():
                    print("No text found in PDF, extracting text from images...")

                    # Converting entire PDF to images for OCR
                    images = convert_from_path(file_path, dpi=300)
                    for img in images:
                        img_text = pytesseract.image_to_string(img, lang="eng")
                        extracted_text += img_text + "\n"

        elif file_path.lower().endswith(('.doc', '.docx')):
            doc = docx.Document(file_path)
            extracted_text = '\n'.join([para.text for para in doc.paragraphs])

        elif file_path.lower().endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as f:
                extracted_text = f.read()

        else:
            return "Unsupported file format"

        return extracted_text.strip() if extracted_text.strip() else "No text extracted."

    except Exception as e:
        return f"Error extracting text from file: {e}"

#Fetching Emails
def fetch_emails(service, after_timestamp=None, max_results=100):
    query = ""
    if after_timestamp:
        query = f"after:{int(after_timestamp)}"
    
    try:
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query, maxResults=max_results).execute()
        messages = results.get('messages', [])
        
        email_data = []
        if not messages:
            print("No new emails found.")
        else:
            print(f"Fetching {len(messages)} new emails...")
            for message in messages:
                msg = service.users().messages().get(userId='me', id=message['id']).execute()
                
                headers = msg['payload']['headers']
                sender = None
                subject = None
                date = None
                
                for header in headers:
                    if header['name'] == 'From':
                        sender = header['value']
                    elif header['name'] == 'Subject':
                        subject = header['value']
                    elif header['name'] == 'Date':
                        date = header['value']
                
                if not date:
                    continue
                
                timestamp = parse_email_date(date)
                if timestamp is None:
                    continue  
                
                sender_name, sender_email = extract_sender(sender)
                
                if subject:
                    subject = subject.strip()
                
                has_attachments = False
                attachment_type = "NILL"
                extracted_text = "NILL"
                summarized_extracted_text = "NILL"  

                if 'parts' in msg['payload']:
                    for part in msg['payload']['parts']:
                        if 'filename' in part and part['filename']:
                            has_attachments = True
                            attachment_type = get_attachment_type(msg['payload'])

                            # Extracting and processing attachment content
                            if 'body' in part and 'attachmentId' in part['body']:
                                attachment_id = part['body']['attachmentId']
                                attachment = service.users().messages().attachments().get(
                                    userId='me', messageId=message['id'], id=attachment_id
                                ).execute()

                                data = base64.urlsafe_b64decode(attachment['data'])

                                file_path = f"temp_{part['filename']}"
                                with open(file_path, "wb") as f:
                                    f.write(data)

                                # Processing extracted text
                                try:
                                    if attachment_type in ["Word Document", "PDF", "Text File"]:
                                        extracted_text = extract_text_from_file(file_path)
                                    elif attachment_type == "Image":
                                        extracted_text = extract_text_from_image(file_path)
                                    else:
                                        extracted_text = "Unsupported file format"
                                except Exception as e:
                                    extracted_text = f"Error extracting text: {e}"

                                # Delete the attachment after extracting text
                                try:
                                    os.remove(file_path)
                                    print(f"Deleted attachment: {file_path}")
                                except Exception as e:
                                    print(f"Error deleting file {file_path}: {e}")

                                # Summarize extracted text
                                if extracted_text and extracted_text != "NILL":
                                    summarized_extracted_text = summarize_email_with_openai(extracted_text)

                if not has_attachments:
                    attachment_type = "NILL"
                    extracted_text = "NILL"
                    summarized_extracted_text = "NILL"

                body = msg.get('snippet', '').replace('â€Œ', '').replace('â€™', '')  
                summarization = summarize_email(body)
                
                email_info = {
                    "ID": msg['id'],
                    "Timestamp": timestamp,
                    "Date": date,
                    "From": f"{sender_name} <{sender_email}>",
                    "Subject": subject,
                    "Body": body,  
                    "Summarization": summarization, 
                    "Has_Attachments": has_attachments,
                    "Attachment_Type": attachment_type,
                    "Extracted_Text": extracted_text,
                    "Summarized_Extracted_Text_From_Attachments": summarized_extracted_text
                }
                
                email_data.append(email_info)

        return email_data
    except HttpError as error:
        print(f'An error occurred: {error}')
        return []


# Save emails to CSV
def save_to_csv(email_data, filename=CSV_FILE):
    new_data = pd.DataFrame(email_data)

    # Check if file exists
    if os.path.exists(filename):
        try:
            existing_data = pd.read_csv(filename)
            combined_data = pd.concat([existing_data, new_data]).drop_duplicates(subset=['ID'])
        except PermissionError:
            print(f"Error: Permission denied while accessing {filename}. Ensure it's not open in another application.")
            return
    else:
        combined_data = new_data 

    # Sort emails by timestamp in ascending order (oldest to newest)
    combined_data = combined_data.sort_values(by=['Timestamp'], ascending=True)

    try:
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            combined_data.to_csv(f, index=False)
            print(f"Emails saved to {filename} .")
    except PermissionError:
        print(f"Error: Permission denied while saving {filename}. Make sure it's not open in any application.")

# Main function 
def main():
    service = authenticate_gmail_api()
    if not service:
        return

    while True:
        last_timestamp = get_last_timestamp()

        if last_timestamp:
            print(f"Fetching emails after {last_timestamp}...")
            email_data = fetch_emails(service, after_timestamp=last_timestamp)
        else:
            print("No previous timestamp found. Fetching latest n emails.")
            email_data = fetch_emails(service, max_results=100)

        if email_data:
            save_to_csv(email_data)

            # Store the latest timestamp for the next execution
            latest_timestamp = max(email["Timestamp"] for email in email_data)
            save_last_timestamp(str(int(latest_timestamp)))

        # Wait before checking again
        time.sleep(1) 

if __name__ == '__main__':
    main()