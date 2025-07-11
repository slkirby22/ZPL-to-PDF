# New version of the app with dynamic label settings (DPMM, dimensions, rotate, crop, scaleopts)
# Based on existing ZPL-to-PDF print app

import os
import sys
import time
import requests
import win32print
import chardet
from PyPDF2 import PdfReader, PdfWriter
import subprocess
import tkinter as tk
import shutil
from tkinter import filedialog, messagebox

# --- Paths ---
BASE_DIR = os.path.dirname(sys.argv[0])
CONFIG_PATH = os.path.join(BASE_DIR, 'config.txt')
LABEL_CONFIG_PATH = os.path.join(BASE_DIR, 'label_settings.txt')

# --- Config Handling ---
def read_config():
    zpl = zplii = None
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r') as f:
            for line in f:
                if line.startswith('zpl='):
                    zpl = line.strip().split('=', 1)[1]
                elif line.startswith('zplii='):
                    zplii = line.strip().split('=', 1)[1]
    return zpl, zplii

def write_config(zpl_printer, zplii_printer):
    with open(CONFIG_PATH, 'w') as f:
        f.write(f"zpl={zpl_printer}\n")
        f.write(f"zplii={zplii_printer}\n")

def read_label_settings(label_type):
    settings = {
        'dpmm': '8', 'width': '4', 'height': '6',
        'rotate': '0', 'crop': '0', 'scaleopts': 'fit'
    }
    if os.path.exists(LABEL_CONFIG_PATH):
        with open(LABEL_CONFIG_PATH, 'r') as f:
            for line in f:
                if line.lower().startswith(label_type.lower() + '='):
                    parts = line.strip().split('=')[1].split(',')
                    if len(parts) == 6:
                        settings = dict(zip(['dpmm', 'width', 'height', 'rotate', 'crop', 'scaleopts'], parts))
    return settings

def write_label_settings(label_type, dpmm, width, height, rotate, crop, scaleopts):
    lines = []
    updated = False
    if os.path.exists(LABEL_CONFIG_PATH):
        with open(LABEL_CONFIG_PATH, 'r') as f:
            lines = f.readlines()
    with open(LABEL_CONFIG_PATH, 'w') as f:
        for line in lines:
            if line.lower().startswith(label_type.lower() + '='):
                f.write(f"{label_type}={dpmm},{width},{height},{rotate},{crop},{scaleopts}\n")
                updated = True
            else:
                f.write(line)
        if not updated:
            f.write(f"{label_type}={dpmm},{width},{height},{rotate},{crop},{scaleopts}\n")

# --- Label Config GUI ---
def configure_label_settings():
    root = tk.Tk()
    root.title('Label Settings')
    root.geometry('400x400')

    tk.Label(root, text='Label Type (zpl or zplii):', font=('Arial', 12)).pack()
    label_type_var = tk.StringVar(value='zpl')
    tk.Entry(root, textvariable=label_type_var, width=40).pack(pady=5)

    fields = {}
    for label in ['DPMM (6/8/10/12)', 'Width (in)', 'Height (in)', 'Rotate (deg)', 'Crop (pts)', 'Scale Opts']:
        tk.Label(root, text=label + ':', font=('Arial', 12)).pack()
        var = tk.StringVar()
        tk.Entry(root, textvariable=var, width=40).pack(pady=5)
        fields[label] = var

    def save():
        write_label_settings(
            label_type_var.get(),
            fields['DPMM (6/8/10/12)'].get(),
            fields['Width (in)'].get(),
            fields['Height (in)'].get(),
            fields['Rotate (deg)'].get(),
            fields['Crop (pts)'].get(),
            fields['Scale Opts'].get()
        )
        root.destroy()

    tk.Button(root, text='Save', command=save, bg='green', fg='white', font=('Arial', 12)).pack(pady=10)
    root.mainloop()

# --- URL Builder ---
def build_labelary_url(dpmm, width, height, rotate):
    return f"https://api.labelary.com/v1/printers/{dpmm}dpmm/labels/{width}x{height}/{rotate}/"

# --- Rotate / Crop ---
def rotate_pdf(pdf_path, degrees):
    if int(degrees) == 0:
        return
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for p in reader.pages:
        p.rotate(int(degrees))
        writer.add_page(p)
    with open(pdf_path, 'wb') as f:
        writer.write(f)

def crop_pdf_top(pdf_path, points_to_trim):
    if int(points_to_trim) == 0:
        return
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for p in reader.pages:
        mb = p.mediabox
        mb.upper_right = (mb.upper_right[0], mb.upper_right[1] - int(points_to_trim))
        p.mediabox = mb
        writer.add_page(p)
    with open(pdf_path, 'wb') as f:
        writer.write(f)

# --- Print with Sumatra ---
def print_with_sumatra(pdf_path, printer_name, scale_opts):
    if printer_name == 'Microsoft Print to PDF':
        out_dir = os.path.join(os.path.dirname(pdf_path), 'printed_pdfs')
        os.makedirs(out_dir, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        dest = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}.pdf")
        shutil.copy(pdf_path, dest)
        return

    exe = 'SumatraPDF-3.5.2-64.exe'
    path = os.path.join(BASE_DIR, exe)
    if not os.path.exists(path):
        messagebox.showerror('Print Error', f'Missing: {exe}')
        sys.exit(1)

    proc = subprocess.Popen([
        path, '-print-to', printer_name,
        '-print-settings', scale_opts,
        pdf_path
    ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = proc.communicate()
    if proc.returncode != 0:
        messagebox.showerror('Print Error', stderr.decode().strip())
        sys.exit(1)

# --- Process ZPL File ---
def process_zpl_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    label_type = 'zplii' if 'zplii' in ext else 'zpl'
    zpl_prn, zplii_prn = read_config()

    try:
        raw = open(file_path, 'rb').read()
    except:
        messagebox.showerror('File Error', f'Missing file: {file_path}')
        return
    enc = chardet.detect(raw)['encoding'] or 'latin-1'
    code = raw.decode(enc, errors='replace')

    s = read_label_settings(label_type)
    url = build_labelary_url(s['dpmm'], s['width'], s['height'], s['rotate'])
    resp = requests.post(url, files={'file': (file_path, code)}, headers={'Accept': 'application/pdf'})
    if resp.status_code != 200:
        print(f'Failed to convert ZPL: {resp.status_code}')
        sys.exit(1)
    pdf_path = file_path.rsplit('.', 1)[0] + '.pdf'
    with open(pdf_path, 'wb') as f:
        f.write(resp.content)

    rotate_pdf(pdf_path, s['rotate'])
    crop_pdf_top(pdf_path, s['crop'])
    printer = zpl_prn if label_type == 'zpl' else zplii_prn

    if printer and printer in [info[2] for info in win32print.EnumPrinters(2)]:
        print_with_sumatra(pdf_path, printer, s['scaleopts'])
    else:
        printer_selection_window(pdf_path, s['scaleopts'])

# --- Printer Selection Window ---
def printer_selection_window(pdf_path, scale_opts):
    def on_print():
        selected = printer_var.get()
        if selected:
            window.destroy()
            print_with_sumatra(pdf_path, selected, scale_opts)
            sys.exit(0)

    def on_cancel():
        window.destroy()
        sys.exit(0)

    root = tk.Tk()
    root.withdraw()
    window = tk.Toplevel()
    window.title("Select Printer")
    window.geometry('500x300')
    window.protocol("WM_DELETE_WINDOW", on_cancel)
    tk.Label(window, text="Choose a printer:", font=('Arial', 12)).pack(pady=20)

    printer_var = tk.StringVar()
    printers = [p[2] for p in win32print.EnumPrinters(2)]
    tk.OptionMenu(window, printer_var, *printers).pack(pady=10)
    tk.Button(window, text="Print", command=on_print, bg='green', fg='white').pack(pady=10)
    tk.Button(window, text="Cancel", command=on_cancel, bg='red', fg='white').pack(pady=10)
    window.mainloop()

# --- Main Window ---
def main():
    if len(sys.argv) > 1:
        process_zpl_file(sys.argv[1])
    else:
        root = tk.Tk()
        root.title('ZPL Label Printer')
        root.geometry('300x200')
        tk.Button(root, text='Print Label', command=lambda: process_zpl_file(filedialog.askopenfilename(filetypes=[("ZPL Files", "*.zpl *.zplii")]))).pack(pady=10)
        tk.Button(root, text='Configure Printers', command=lambda: configure_printers()).pack(pady=10)
        tk.Button(root, text='Label Settings', command=lambda: configure_label_settings()).pack(pady=10)
        root.mainloop()

def configure_printers():
    zpl, zplii = read_config()
    printers = [p[2] for p in win32print.EnumPrinters(2)]
    root = tk.Tk()
    root.title('Configure Printers')
    root.geometry('350x200')

    tk.Label(root, text='UPS (.zpl) printer:', font=('Arial', 12)).pack()
    zpl_var = tk.StringVar(value=zpl if zpl in printers else printers[0])
    tk.OptionMenu(root, zpl_var, *printers).pack(pady=5)

    tk.Label(root, text='FedEx (.zplii) printer:', font=('Arial', 12)).pack()
    zplii_var = tk.StringVar(value=zplii if zplii in printers else printers[0])
    tk.OptionMenu(root, zplii_var, *printers).pack(pady=5)

    def save():
        write_config(zpl_var.get(), zplii_var.get())
        root.destroy()

    tk.Button(root, text='Save', command=save, bg='green', fg='white').pack(pady=10)
    root.mainloop()

if __name__ == '__main__':
    main()
