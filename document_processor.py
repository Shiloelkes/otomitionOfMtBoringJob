import fitz
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def process_and_prepare_excel():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 1. Selection: Source PDF, Stamp Image, and Output Folder
    input_pdf = filedialog.askopenfilename(title="Select Source PDF", filetypes=[("PDF", "*.pdf")])
    if not input_pdf: return

    stamp_img = filedialog.askopenfilename(title="Select Stamp Image", filetypes=[("Images", "*.png;*.jpg;*.jpeg")])
    if not stamp_img: return

    output_folder = filedialog.askdirectory(title="Select Folder to Save Split Files")
    if not output_folder: return

    try:
        doc = fitz.open(input_pdf)
        data_for_excel = []
        
        for i, page in enumerate(doc):
            # A. Add Stamp (Position: 100, 650)
            # Coordinates and size (120x120) based on your original request
            rect = fitz.Rect(100, 650, 100 + 120, 650 + 120)
            page.insert_image(rect, filename=stamp_img)

            # B. Identify Recipient Name for file naming
            # Searching for the Hebrew anchor "לכבוד" to locate the name below it
            text_instances = page.search_for("לכבוד")
            clean_name = ""
            if text_instances:
                inst = text_instances[0]
                name_rect = fitz.Rect(inst.x0 - 100, inst.y1, inst.x1 + 200, inst.y1 + 100)
                page_dict = page.get_text("dict", clip=name_rect)
                lines = []
                for block in page_dict["blocks"]:
                    if "lines" in block:
                        for line in block["lines"]:
                            line_text = "".join([span["text"] for span in line["spans"]]).strip()
                            if line_text: lines.append((line["bbox"][1], line_text))
                if lines:
                    lines.sort(key=lambda x: x[0])
                    # Remove special characters to ensure valid filename
                    clean_name = "".join(c for c in lines[0][1] if c.isalnum() or c in (' ', '-', '_')).strip()

            if not clean_name: 
                clean_name = f"Recipient_{i+1}"

            # C. Create single-page document
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
            
            file_name = f"{clean_name}.pdf"
            output_filename = os.path.join(output_folder, file_name)
            
            # Save with compression (garbage collection and deflation)
            new_doc.save(output_filename, garbage=4, deflate=True, clean=True)
            new_doc.close()

            # D. Collect data for Excel
            data_for_excel.append({
                "Recipient Name": clean_name,
                "Phone (Format: 972501234567)": "",
                "Email Address": "",
                "File Path": output_filename
            })

        doc.close()

        # 4. Export to Excel Template
        df = pd.DataFrame(data_for_excel)
        excel_output = os.path.join(output_folder, "Dispatch_List_Template.xlsx")
        df.to_excel(excel_output, index=False)

        messagebox.showinfo("Success", f"Processing Complete!\n\nFiles saved to: {output_folder}\nExcel template created: {excel_output}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    process_and_prepare_excel()
