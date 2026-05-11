import pandas as pd
import win32com.client as win32
import os
import time
import webbrowser
import pyautogui
import tkinter as tk
import subprocess
from tkinter import filedialog, messagebox
from datetime import datetime

def send_automated_notifications():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 1. Select the filled Excel file
    excel_path = filedialog.askopenfilename(title="Select the Dispatch Excel File", filetypes=[("Excel", "*.xlsx")])
    if not excel_path: return

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"Cannot read Excel: {e}")
        return

    # Initialize Outlook
    try:
        outlook = win32.Dispatch('outlook.application')
    except:
        outlook = None
        print("Outlook not detected. Skipping email delivery.")

    report_lines = [f"Dispatch Report - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n", "-"*50 + "\n"]
    print("Starting process... Please do not use the mouse or keyboard during automation.")

    for index, row in df.iterrows():
        try:
            name = str(row.iloc[0]).strip()
            # Clean phone number formatting
            phone_raw = str(row.iloc[1]).strip().split('.')[0].replace("+", "").replace("-", "").replace(" ", "")
            email = str(row.iloc[2]).strip()
            file_path = str(row.iloc[3]).strip()
            
            if not os.path.exists(file_path) or file_path.lower() == 'nan' or len(phone_raw) < 8:
                continue

            # --- 1. Email Delivery via Outlook ---
            status_mail = "Not Sent"
            if outlook and "@" in email:
                try:
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.Subject = f"Document Update - {name}"
                    mail.Body = f"Hello {name},\n\nPlease find your attached document for the current period.\n\nBest regards,\n[AUTHORIZED_SENDER]"
                    mail.Attachments.Add(file_path)
                    mail.Send()
                    status_mail = "Success"
                except: status_mail = "Error"

            # --- 2. WhatsApp Delivery (Automated Browser Control) ---
            status_whatsapp = "Not Sent"
            try:
                # Format phone for international WhatsApp link
                full_phone = f"972{phone_raw[-9:]}"
                url = f"https://web.whatsapp.com/send?phone={full_phone}"
                webbrowser.open(url)
                
                # Wait for WhatsApp Web to load
                time.sleep(25) 
                
                # Focus browser
                pyautogui.click(pyautogui.size().width / 2, pyautogui.size().height / 2)
                time.sleep(2)

                # Prepare and paste message
                message_text = f"Hello {name}, your updated document is attached below."
                cmd_text = f'powershell.exe -command "Set-Clipboard -Value \'{message_text}\'"'
                subprocess.run(cmd_text, shell=True)
                
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(2)

                # Prepare and paste file
                cmd_file = f'powershell.exe -command "Set-Clipboard -Path \'{file_path}\'"'
                subprocess.run(cmd_file, shell=True)
                
                time.sleep(3)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(5) 
                pyautogui.press('enter')
                time.sleep(3)
                
                status_whatsapp = "Sent"
            except Exception as e:
                status_whatsapp = f"Error: {e}"

            log_entry = f"Recipient: {name} | Email: {status_mail} | WhatsApp: {status_whatsapp}\n"
            print(log_entry.strip())
            report_lines.append(log_entry)

        except Exception as e:
            print(f"Error processing row {index}: {e}")

    # Final Report Generation
    report_path = os.path.join(os.path.dirname(excel_path), "Final_Dispatch_Report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.writelines(report_lines)
    
    messagebox.showinfo("Complete", f"Process finished!\n\nReport saved to: {report_path}")

if __name__ == "__main__":
    send_automated_notifications()
