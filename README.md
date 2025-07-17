# Gmail Email Processing Script

## Overview
This script automates email extraction from Gmail using the Gmail API. It fetches new emails, extracts attachments (PDFs, Word documents, text files, images), performs OCR using Tesseract, and summarizes the email content using OpenAI's GPT-3.5. The extracted data is saved in CSV format for further analysis.

## Features
- Authenticate with Gmail API
- Fetch emails from Gmail inbox
- Extract email metadata (sender, subject, timestamp, body, attachments)
- Process attachments (PDFs, Word documents, text files, images)
- Perform OCR on image-based attachments
- Summarize email body and extracted text using OpenAI's GPT-3.5
- Save processed data into a CSV file
- Maintain a timestamp of the last processed email to avoid duplicate processing

## Setup

### 1. Clone the Repository
```sh
$ git clone https://github.com/shahkhushi028/Automating-Email-Fetching-and-Summarization-of-Emails-Attachments.git
$ cd Automating-Email-Fetching-and-Summarization-of-Emails-Attachments  
```

### 2. Install Dependencies
```sh
$ pip install -r requirements.txt
```

### 3. Configure Google API
1. Obtain OAuth credentials from Google Cloud Console.
2. Download `credentials.json` and place it in the project directory.
3. Run the script to authenticate and generate `token.json`.

### 4. Set Up Environment Variables
Create a `.env` file in the project directory and add:
```sh
OPENAI_API_KEY=your_openai_api_key
```

### 5. Install and Configure Tesseract OCR
1. Download and install Tesseract OCR from [Tesseract OCR](https://github.com/tesseract-ocr/tesseract).
2. Update the path in the script:
```python
pytesseract.pytesseract.tesseract_cmd = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
```

## Running the Script
Run the script using:
```sh
$ python script.py
```
The script will fetch new emails, process attachments, extract relevant text, summarize content, and store the output in `data.csv`.

## Output
The extracted email data is stored in `data.csv` with the following fields:

- **ID**
- **Timestamp**
- **Date**
- **From**
- **Subject**
- **Body**
- **Summarization**
- **Has Attachments**
- **Attachment Type**
- **Extracted Text**
- **Summarized Extracted Text from Attachments**

## Notes
- Ensure that `credentials.json` and `token.json` are stored securely.
- Tesseract OCR should be correctly installed and configured for image processing.
- If running in a production environment, implement proper error handling and logging.

## Author
Khushi Shah