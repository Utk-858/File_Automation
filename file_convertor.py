import os
import pandas as pd
from docx2pdf import convert
from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox

# Path to converted folder on desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "converted")
os.makedirs(desktop_path, exist_ok=True)  # Create folder if it doesn't exist

# Conversion functions
def convert_csv_to_excel(csv_file):
    output_xlsx = os.path.join(desktop_path, os.path.splitext(os.path.basename(csv_file))[0] + ".xlsx")
    df = pd.read_csv(csv_file)
    df.to_excel(output_xlsx, index=False)
    messagebox.showinfo("Success", f"Converted {csv_file} to {output_xlsx}")

def convert_docx_to_pdf(docx_file):
    output_pdf = os.path.join('/Users/utkarshbansal/Desktop/converted/', os.path.splitext(os.path.basename(docx_file))[0] + ".pdf")
    try:
        convert(docx_file, output_pdf)  # Convert DOCX to PDF
        if os.path.exists(output_pdf):
            print(f"Successfully converted {docx_file} to {output_pdf}.")
        else:
            print("PDF file was not created successfully.")
    except Exception as e:
        print(f"Error occurred during conversion: {e}")




def convert_pdf_to_docx(pdf_file):
    output_docx = os.path.join(desktop_path, os.path.splitext(os.path.basename(pdf_file))[0] + ".docx")
    cv = Converter(pdf_file)
    cv.convert(output_docx, start=0, end=None)
    cv.close()
    messagebox.showinfo("Success", f"Converted {pdf_file} to {output_docx}")

# General conversion handler
def convert_file():
    input_file = input_path_var.get()
    ext = os.path.splitext(input_file)[1].lower()
    target_format = format_var.get()

    if not input_file:
        messagebox.showerror("Error", "Please select an input file.")
        return

    if ext == '.csv' and target_format == "Excel":
        convert_csv_to_excel(input_file)
    elif ext == '.docx' and target_format == "pdf":
        convert_docx_to_pdf(input_file)
    elif ext == '.pdf' and target_format == "docx":
        convert_pdf_to_docx(input_file)
    else:
        messagebox.showerror("Error", f"Unsupported conversion for file type: {ext} to {target_format}")

# GUI setup
root = tk.Tk()
root.title("File Converter Tool")
root.geometry("400x300")

input_path_var = tk.StringVar()
format_var = tk.StringVar(value="Select Format")

def browse_input():
    file_path = filedialog.askopenfilename(title="Select Input File",
                                           filetypes=[("All files", "*.*"),
                                                      ("CSV files", "*.csv"),
                                                      ("Word files", "*.docx"),
                                                      ("PDF files", "*.pdf")])
    input_path_var.set(file_path)

# GUI Layout
tk.Label(root, text="Input File").pack(pady=5)
tk.Entry(root, textvariable=input_path_var, width=40).pack()
tk.Button(root, text="Browse", command=browse_input).pack(pady=5)

# Dropdown for format selection
tk.Label(root, text="Select Output Format").pack(pady=5)
format_options = ["Excel", "pdf", "docx"]
format_menu = tk.OptionMenu(root, format_var, *format_options)
format_menu.pack(pady=5)

# Convert button
tk.Button(root, text="Convert", command=convert_file, bg="blue", fg="black").pack(pady=20)

root.mainloop()
