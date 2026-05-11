# Employee Document Automator

A professional Windows-based automation suite for processing multi-page PDF documents and dispatching them to recipients via Email (Outlook) and WhatsApp Web.

## Workflow Overview
The system operates in a three-step cycle:
1. **Process:** A master PDF is stamped, split into individual files, and named automatically.
2. **Review:** An Excel template is generated for you to provide contact details.
3. **Dispatch:** The system automates Outlook and browser interactions to deliver the files.

## Prerequisites
- **OS:** Windows (Required for Outlook and PowerShell integration).
- **Python:** 3.7+
- **Software:** - Microsoft Outlook (must be open in the background).
  - WhatsApp Web (must be logged in on your default browser).
- **Libraries:** Install via `pip install -r requirements.txt`.

## Usage Instructions

### Step 1: Processing the Documents
Run `document_processor.py`.
1. Select your source PDF (e.g., a 60-page payroll document).
2. Select a stamp image (PNG/JPG).
3. Select an output folder.
The script will split the PDF, add the stamp, and generate a `Dispatch_List_Template.xlsx` file.

### Step 2: Preparing the Data
Open the generated Excel file.
- Fill in the **Phone** and **Email Address** columns for each recipient.
- Ensure the phone number is in digits only (e.g., 0501234567).
- Save and close the file.

### Step 3: Dispatching
Run `document_dispatcher.py`.
1. Select the Excel file you just edited.
2. **Important:** Do not touch your mouse or keyboard once the process starts. The script uses `pyautogui` to simulate keystrokes in WhatsApp Web.

## Technical Details
- **PDF Handling:** Powered by `PyMuPDF` (fitz) for high-performance text extraction and image insertion.
- **WhatsApp Automation:** Uses a combination of `webbrowser` and `pyautogui` to navigate WhatsApp Web.
- **Email Delivery:** Utilizes `pywin32` to interface directly with the Outlook MAPI engine.

---
*Note: This tool is intended for professional administrative use. Ensure compliance with privacy and anti-spam regulations in your jurisdiction.*
