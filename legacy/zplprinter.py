import os
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
import win32print
import chardet
import sys
from PyPDF2 import PdfReader, PdfWriter
import subprocess
import time

def get_available_printers():
    printers = []
    for printer_info in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
        printers.append(printer_info[2])
    return printers

def print_with_adobe(pdf_path, printer_name):
    import sys
    try:
        adobe_reader_path = "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe"
        if not os.path.exists(adobe_reader_path):
            raise FileNotFoundError("Adobe Acrobat Reader is not installed at the expected location.")

        # Close Acrobat if already open to avoid conflict with /t flag
        subprocess.run(["taskkill", "/f", "/im", "Acrobat.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1)  # Give it time to fully close

        # Launch Acrobat to print the file
        subprocess.Popen([
            adobe_reader_path,
            "/t",
            pdf_path,
            printer_name
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"Successfully sent {pdf_path} to the printer using Adobe Acrobat Reader.")
        return
    except Exception as e:
        print(f"Error printing with Adobe Acrobat Reader: {e}")
        messagebox.showerror("Printing Error", f"Error printing the label: {e}\nThe label will now be opened in the PDF viewer.")
        open_pdf(pdf_path)

def open_pdf(pdf_path):
    try:
        subprocess.run(["start", pdf_path], shell=True)
        print(f"Opened {pdf_path} in the default PDF viewer.")
    except Exception as e:
        print(f"Error opening PDF viewer: {e}")
        messagebox.showerror("Error", f"Failed to open PDF: {e}")

def process_zpl_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    is_zpl = ext == ".zpl"
    is_zplii = ext == ".zplii"

    if not (is_zpl or is_zplii):
        messagebox.showerror("Invalid File", "Please select a .zpl or .zplii file.")
        return

    with open(file_path, "rb") as file:
        raw_data = file.read()
        detected_encoding = chardet.detect(raw_data)["encoding"] or "latin-1"

    zpl_code = raw_data.decode(detected_encoding, errors="replace")
    files = {"file": (file_path, zpl_code)}

    labelary_url = "https://api.labelary.com/v1/printers/8dpmm/labels/4x6.21/0/" if is_zpl else "https://api.labelary.com/v1/printers/8dpmm/labels/4x7.25/0/"

    headers = {
        "Accept": "application/pdf",
        "Content-Type": "application/x-www-form-urlencoded"
    }

    response = requests.post(labelary_url, files=files, headers=headers)

    if response.status_code == 200:
        pdf_path = file_path.rsplit(".", 1)[0] + ".pdf"
        with open(pdf_path, "wb") as pdf_file:
            pdf_file.write(response.content)
            os.fsync(pdf_file.fileno())

        if is_zpl:
            rotate_pdf(pdf_path)
            crop_pdf_top(pdf_path, points_to_trim=12)

        final_pdf_path = pdf_path

        action = messagebox.askyesnocancel(
            "Label Generated",
            f"Label saved as PDF: {final_pdf_path}\nDo you want to print it?",
            icon='question'
        )

        if action is None:
            print("Action canceled.")
            sys.exit(0)
        elif action:
            available_printers = get_available_printers()
            if available_printers:
                printer_selection_window(available_printers, final_pdf_path)
            else:
                messagebox.showerror("No Printers Found", "No printers available. Please install a printer.")
        else:
            print(f"Saved as PDF: {final_pdf_path}")
            messagebox.showinfo("File Saved", f"Label saved as PDF: {final_pdf_path}")
            sys.exit(0)
    else:
        messagebox.showerror("Error", f"Failed to process ZPL: {response.status_code}")
        print(f"Error: {response.status_code}")

def printer_selection_window(available_printers, pdf_path):
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    def on_print_button_click():
        selected_printer = printer_var.get()
        if selected_printer != "":
            printer_window.destroy()
            print_with_adobe(pdf_path, selected_printer)
            sys.exit(0)
            printer_window.destroy()
        else:
            messagebox.showerror("Printer Selection", "Please select a printer before printing.")

    def on_cancel_button_click():
        printer_window.destroy()
        print("Printing operation canceled.")
        sys.exit(0)

    def on_window_close():
        print("Window closed using X button.")
        printer_window.destroy()
        sys.exit(0)

    root = tk.Tk()
    root.withdraw()

    printer_window = tk.Toplevel()
    printer_window.lift()
    printer_window.attributes('-topmost', True)
    printer_window.after_idle(printer_window.attributes, '-topmost', False)
    printer_window.title("Select Printer")
    printer_window.geometry('500x300')
    printer_window.protocol("WM_DELETE_WINDOW", on_window_close)

    label = tk.Label(printer_window, text="Choose a printer:", font=('Arial', 12))
    label.pack(pady=20)

    printer_var = tk.StringVar()
    printer_menu = tk.OptionMenu(printer_window, printer_var, *available_printers)
    printer_menu.config(width=30)
    printer_menu.pack(pady=10)

    info_label = tk.Label(printer_window, text="Please select a printer and click Print.", font=('Arial', 10), fg="gray")
    info_label.pack(pady=5)

    print_button = tk.Button(printer_window, text="Print", command=on_print_button_click, bg='green', fg='white', font=('Arial', 12))
    print_button.pack(pady=10)

    cancel_button = tk.Button(printer_window, text="Cancel", command=on_cancel_button_click, bg='red', fg='white', font=('Arial', 12))
    cancel_button.pack(pady=10)

    printer_window.config(padx=20, pady=20)
    printer_window.mainloop()

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select ZPL File",
        filetypes=[("All ZPL Files", "*.zpl *.ZPL *.zplii *.ZPLII")]
    )

    if file_path:
        print(f"Selected File: {file_path}")
        process_zpl_file(file_path)
    else:
        print("No file selected.")
        sys.exit(0)

def rotate_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(180)
        writer.add_page(page)
    with open(pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)
    print(f"PDF rotated and overwritten: {pdf_path}")

def crop_pdf_top(pdf_path, points_to_trim):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        media_box = page.mediabox
        media_box.upper_right = (
            media_box.upper_right[0],
            media_box.upper_right[1] - points_to_trim
        )
        page.mediabox = media_box
        writer.add_page(page)
    with open(pdf_path, "wb") as out_file:
        writer.write(out_file)
    print(f"PDF cropped and overwritten: {pdf_path}")

def main():
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        print(f"File path provided: {file_path}")
        process_zpl_file(file_path)
    else:
        print("No file provided. Please select a ZPL file.")
        select_file()

if __name__ == "__main__":
    main()