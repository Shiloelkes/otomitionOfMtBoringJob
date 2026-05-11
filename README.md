# Employee Document Automator

## About This Project

**The Challenge: The Document Distribution Bottleneck**
Processing and distributing bulk, personalized documents is traditionally a severe administrative bottleneck. When tasked with taking a massive, multi-page PDF, applying a signature to every single page, and routing each specific page to a different recipient, the standard methods quickly become impractical. Executing this physically (printing, signing, scanning, handing out) or even attempting to manually digitize the workflow (stamping, saving, and messaging one by one) is incredibly tedious, highly repetitive, and prone to human error. 

**The Solution: End-to-End Python Automation**
To solve this, I engineered a Python automation suite that eliminates the friction entirely. This tool transforms hours of manual data entry into an instant, seamless workflow by handling the entire lifecycle of document processing and distribution:

* **Automated Processing:** Ingests the master PDF, auto-applies a designated digital stamp to every page, and splits the master document into individual files.
* **Intelligent Extraction:** Scans the text of each page to dynamically extract the recipient's name and automatically renames the split files accordingly.
* **Multi-Channel Dispatch:** Seamlessly interfaces with Microsoft Outlook and WhatsApp Web to automatically route and deliver each personalized document to the correct individual.

By replacing a tedious manual process with a fully automated pipeline, this project ensures rapid, secure, and error-free document distribution at scale.

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
