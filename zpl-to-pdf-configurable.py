import os
import sys
import time
import requests
import win32print
import chardet
from PyPDF2 import PdfReader, PdfWriter
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.txt')
LABEL_CONFIG_PATH = os.path.join(BASE_DIR, 'label_settings.txt')
LABEL_TYPES = ['zpl', 'zplii']
DPMM_OPTIONS = ['6', '8', '10', '12']
SCALE_OPTIONS = ['fit', 'shrink', 'noscale', 'center']

# --- Config Read/Write ---
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
        'rotate': '0', 'crop': '0', 'scaleopts': 'fit', 'paper': ''
    }
    if os.path.exists(LABEL_CONFIG_PATH):
        with open(LABEL_CONFIG_PATH, 'r') as f:
            for line in f:
                if line.lower().startswith(label_type.lower() + '='):
                    parts = line.strip().split('=')[1].split(',')
                    if len(parts) >= 7:
                        settings = dict(zip(['dpmm', 'width', 'height', 'rotate', 'crop', 'scaleopts', 'paper'], parts))
    return settings

def write_label_settings(label_type, dpmm, width, height, rotate, crop, scaleopts, paper):
    config = {}
    if os.path.exists(LABEL_CONFIG_PATH):
        with open(LABEL_CONFIG_PATH, 'r') as f:
            for line in f:
                if '=' in line:
                    k, v = line.strip().split('=', 1)
                    config[k.lower()] = v
    config[label_type.lower()] = f"{dpmm},{width},{height},{rotate},{crop},{scaleopts},{paper}"
    with open(LABEL_CONFIG_PATH, 'w') as f:
        for k, v in config.items():
            f.write(f"{k}={v}\n")

def configure_label_settings():
    root = tk.Tk()
    root.title('Label Settings')
    root.geometry('420x480')

    notebook = ttk.Notebook(root)
    fields = {}

    for label_type in LABEL_TYPES:
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=label_type.upper())
        settings = read_label_settings(label_type)
        label_fields = {}

        # DPMM OptionMenu
        tk.Label(frame, text='DPMM:', font=('Arial', 12)).pack()
        initial_dpmm = settings['dpmm'] if settings['dpmm'] in DPMM_OPTIONS else DPMM_OPTIONS[0]
        dpmm_var = tk.StringVar(root, value=initial_dpmm)
        dpmm_menu = tk.OptionMenu(frame, dpmm_var, *DPMM_OPTIONS)
        dpmm_menu.config(width=10)
        dpmm_menu.pack(pady=5)
        dpmm_var.set(initial_dpmm)
        label_fields['dpmm'] = dpmm_var

        # Entry fields
        for key in ['width', 'height', 'rotate', 'crop', 'paper']:
            tk.Label(frame, text=f'{key.capitalize()}:', font=('Arial', 12)).pack()
            entry = tk.Entry(frame, width=30)
            entry.insert(0, settings[key])
            entry.pack(pady=5)
            label_fields[key] = entry

        # Scale Options OptionMenu
        tk.Label(frame, text='Scale Options:', font=('Arial', 12)).pack()
        initial_scale = settings['scaleopts'] if settings['scaleopts'] in SCALE_OPTIONS else SCALE_OPTIONS[0]
        scale_var = tk.StringVar(root, value=initial_scale)
        scale_menu = tk.OptionMenu(frame, scale_var, *SCALE_OPTIONS)
        scale_menu.config(width=10)
        scale_menu.pack(pady=5)
        scale_var.set(initial_scale)
        label_fields['scaleopts'] = scale_var

        fields[label_type] = label_fields

    notebook.pack(expand=1, fill='both')

    def save_all():
        for label_type in LABEL_TYPES:
            lf = fields[label_type]
            write_label_settings(
                label_type,
                lf['dpmm'].get(),
                lf['width'].get(),
                lf['height'].get(),
                lf['rotate'].get(),
                lf['crop'].get(),
                lf['scaleopts'].get(),
                lf['paper'].get()
            )
        root.destroy()

    tk.Button(root, text='Save', command=save_all, bg='green', fg='white', font=('Arial', 12), height=2).pack(pady=10)
    root.mainloop()

# --- URL Builder ---
def build_labelary_url(dpmm, width, height, rotate):
    return f"https://api.labelary.com/v1/printers/{dpmm}dpmm/labels/{width}x{height}/"

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
def print_with_sumatra(pdf_path, printer_name, scale_opts, paper):
    setting_string = scale_opts
    if paper:
        setting_string += f",paper={paper}"

    if printer_name == 'Microsoft Print to PDF':
        out_dir = os.path.join(os.path.dirname(pdf_path), 'printed_pdfs')
        os.makedirs(out_dir, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        dest = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_{timestamp}.pdf")
        shutil.copy(pdf_path, dest)
        sys.exit() #???? is this appropriate here? Just buffers and sits in limbo without it.
        return

    exe = 'SumatraPDF-3.5.2-64.exe'
    path = os.path.join(BASE_DIR, exe)
    if not os.path.exists(path):
        messagebox.showerror('Print Error', f'Missing: {exe}')
        sys.exit(1)

    proc = subprocess.Popen([
        path, '-print-to', printer_name,
        '-print-settings', setting_string,
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
        print_with_sumatra(pdf_path, printer, s['scaleopts'], s['paper'])
    else:
        printer_selection_window(pdf_path, s['scaleopts'], s['paper'])

# --- Printer Selection Window ---
def printer_selection_window(pdf_path, scale_opts, paper):
    def on_print():
        selected = printer_var.get()
        if selected:
            window.destroy()
            print_with_sumatra(pdf_path, selected, scale_opts, paper)
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

    available_printers = [printer_info[2] for printer_info in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    printer_var = tk.StringVar(value=available_printers[0])
    printer_menu = ttk.Combobox(window, textvariable=printer_var, values=available_printers, state='readonly', width=40)
    printer_menu.pack(pady=10)

    info_label = tk.Label(window, text="Please select a printer and click Print.", font=('Arial', 10), fg="gray")
    info_label.pack(pady=5)

    tk.Button(window, text="Print", command=on_print, bg='green', fg='white', font=('Arial', 12), height=2).pack(pady=10)
    tk.Button(window, text="Cancel", command=on_cancel, bg='red', fg='white', font=('Arial', 12), height=2).pack(pady=10)
    window.mainloop()

# --- Main Window ---
def main():
    if len(sys.argv) > 1:
        process_zpl_file(sys.argv[1])
    else:
        root = tk.Tk()
        root.title('ZPL Label Printer')
        root.geometry('300x225')
        tk.Button(root, text='Print Label', command=lambda: process_zpl_file(filedialog.askopenfilename(filetypes=[("ZPL Files", "*.zpl *.zplii")])), width=20, height=2).pack(pady=15)
        tk.Button(root, text='Configure Printers', command=configure_printers, width=20, height=2).pack(pady=15)
        tk.Button(root, text='Label Settings', command=configure_label_settings, width=20, height=2).pack(pady=15)
        root.mainloop()

# --- Configure Printers ---
def configure_printers():
    # Load existing defaults
    zpl_default, zplii_default = read_config()
    available = [info[2] for info in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    root = tk.Tk()
    root.title('Configure Printers')
    root.geometry('350x200')
    
    tk.Label(root, text='UPS (.zpl) printer:', font=('Arial', 16)).pack(pady=5)
    # Prepare initial selection
    initial_zpl = zpl_default if zpl_default in available else (available[0] if available else '')
    zpl_var = tk.StringVar(root, value=initial_zpl)
    zpl_menu = tk.OptionMenu(root, zpl_var, *available)
    zpl_menu.config(width=40)
    zpl_menu.pack(pady=5)
    # show current selection explicitly
    zpl_var.set(initial_zpl)
    
    tk.Label(root, text='FedEx (.zplii) printer:', font=('Arial', 16)).pack(pady=5)
    initial_zplii = zplii_default if zplii_default in available else (available[0] if available else '')
    zplii_var = tk.StringVar(root, value=initial_zplii)
    zplii_menu = tk.OptionMenu(root, zplii_var, *available)
    zplii_menu.config(width=40)
    zplii_menu.pack(pady=5)
    # show current selection explicitly
    zplii_var.set(initial_zplii)
    
    def save():
        write_config(zpl_var.get(), zplii_var.get())
        root.destroy()
    
    save_button = tk.Button(root, text='Save', command=save, bg='green', fg='white', font=('Arial', 12))
    save_button.config(width=10)
    save_button.pack(pady=5)
    
    root.mainloop()

if __name__ == '__main__':
    main()
