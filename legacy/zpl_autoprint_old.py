import os
import requests
import win32print
import chardet
import sys
from PyPDF2 import PdfReader, PdfWriter
import subprocess
import time

def printer_selection_window(pdf_path):
    import tkinter as tk
    from tkinter import messagebox
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    def on_print_button_click():
        selected_printer = printer_var.get()
        if selected_printer:
            printer_window.destroy()
            print_with_sumatra(pdf_path, selected_printer)
            sys.exit(0)
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
    available_printers = [printer_info[2] for printer_info in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
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

def print_with_sumatra(pdf_path, printer_name):
    try:
        sumatra_path = "C:\\Users\\Owner\\Documents\\Dev\\SumatraPDF-3.5.2-64.exe"  # corrected escape
        if not os.path.exists(sumatra_path):
            raise FileNotFoundError("SumatraPDF not found.")

        # Use "fit,center,paper=Letter" for test printers only. In production, let the printer use native settings.
        print_settings = "fit,center"
        subprocess.Popen([
            sumatra_path,
            "-print-to", printer_name,
            "-print-settings", print_settings,
            pdf_path
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        print(f"Sent {pdf_path} to {printer_name} using SumatraPDF.")
    except Exception as e:
        print(f"Error printing with SumatraPDF: {e}")

def process_zpl_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    is_zpl = ext == ".zpl"
    is_zplii = ext == ".zplii"

    if not (is_zpl or is_zplii):
        print("Invalid file type. Must be .zpl or .zplii")
        sys.exit(1)

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

        printer_name = "Zebra  ZP 450-200 dpi (2) ZP230D" if is_zpl else "Zebra  ZP 450-200 dpi"
        # printer_name = "HPCD1896 (HP LaserJet Pro M404-M405)" #test env
        available_printers = [p[2].lower() for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        if printer_name.lower() in available_printers:
            print_with_sumatra(pdf_path, printer_name)
        else:
            print(printer_name + " not found. Opening manual selection window.")
            printer_selection_window(pdf_path)
        sys.exit(0)
    else:
        print(f"Failed to process ZPL: {response.status_code}")
        sys.exit(1)

def rotate_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(180)
        writer.add_page(page)
    with open(pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)
    print(f"PDF rotated: {pdf_path}")

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
    print(f"PDF cropped: {pdf_path}")

def main():
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        print(f"Received file: {file_path}")
        process_zpl_file(file_path)
    else:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="Select ZPL File",
            filetypes=[("ZPL and ZPLII Files", "*.zpl *.zplii")]
        )

        if file_path:
            print(f"Manually selected file: {file_path}")
            process_zpl_file(file_path)
        else:
            print("No file selected.")
            sys.exit(0)

if __name__ == "__main__":
    main()
