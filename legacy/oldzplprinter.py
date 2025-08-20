import os
import requests
import win32print
import win32api
import tkinter as tk
from tkinter import filedialog, messagebox
import chardet

# Labelary API endpoint (using GET)
LABELARY_URL = "https://api.labelary.com/v1/printers/6dpmm/labels/4x6/0/"
headers = {
    "Accept": "application/pdf",
    "Content-Type": "application/x-www-form-urlencoded"
}
PRINTER_NAME = "Microsoft Print to PDF"  # Replace with your Zebra printer name if needed

def print_label(pdf_path):
    printer_name = win32print.GetDefaultPrinter()
    win32api.ShellExecute(0, "open", pdf_path, None, ".", 0)

def process_zpl_file(file_path):
    """Convert ZPL to an image using Labelary API and prompt to print."""
    with open(file_path, "rb") as file:
        raw_data = file.read()
        detected_encoding = chardet.detect(raw_data)["encoding"] or "latin-1"

    zpl_code = raw_data.decode(detected_encoding, errors="replace")

    # Use multipart/form-data to send large ZPL files
    files = {"file": (file_path, zpl_code)}

    response = requests.post(LABELARY_URL, files=files, headers=headers)

    if response.status_code == 200:
        print(response.headers)
        print(response.content)
        pdf_path = file_path.replace(".zpl", ".pdf")
        with open(pdf_path, "wb") as pdf_file:
            pdf_file.write(response.content)


        messagebox.showinfo("Label Generated", f"Label saved: {pdf_path}\nPress OK to print.")
        print_label(pdf_path)
        print(f"Printed: {pdf_path}")
    else:
        messagebox.showerror("Error", f"Failed to process ZPL: {response.status_code}")
        print(f"Error: {response.status_code}")


def select_file():
    """Open file dialog for user to select ZPL file."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="Select ZPL File", filetypes=[("ZPL Files", "*.zpl")])
    
    if file_path:
        print(f"Selected File: {file_path}")
        process_zpl_file(file_path)
    else:
        print("No file selected.")

if __name__ == "__main__":
    print("Select a ZPL file to open and print.")
    select_file()