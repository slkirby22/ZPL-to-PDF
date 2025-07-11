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
from PyPDF2 import PdfReader, PdfWriter

# -- Configuration Functions --------------------------------------------------
def read_config():
    """Read default printer names from config.txt."""
    config_path = os.path.join(os.path.dirname(sys.argv[0]), 'config.txt')
    if not os.path.exists(config_path):
        return None, None
    zpl = None
    zplii = None
    with open(config_path, 'r') as f:
        for line in f:
            if line.startswith('zpl='):
                zpl = line.strip().split('=',1)[1]
            elif line.startswith('zplii='):
                zplii = line.strip().split('=',1)[1]
    return zpl, zplii

def write_config(zpl_printer, zplii_printer):
    """Save default printer names to config.txt."""
    config_path = os.path.join(os.path.dirname(sys.argv[0]), 'config.txt')
    with open(config_path, 'w') as f:
        f.write(f"zpl={zpl_printer}\n")
        f.write(f"zplii={zplii_printer}\n")

# -- Printer Selection Window ------------------------------------------------
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
    printer_menu.config(width=40)
    printer_menu.pack(pady=10)

    info_label = tk.Label(printer_window, text="Please select a printer and click Print.", font=('Arial', 10), fg="gray")
    info_label.pack(pady=5)

    print_button = tk.Button(printer_window, text="Print", command=on_print_button_click, bg='green', fg='white', font=('Arial', 12))
    print_button.pack(pady=10)

    cancel_button = tk.Button(printer_window, text="Cancel", command=on_cancel_button_click, bg='red', fg='white', font=('Arial', 12))
    cancel_button.pack(pady=10)

    printer_window.config(padx=20, pady=20)
    printer_window.mainloop()

# -- SumatraPDF Print ---------------------------------------------------------
def print_with_sumatra(pdf_path, printer_name, scale_opts):
    # Special-case the exact Microsoft driver:
    if printer_name == "Microsoft Print to PDF":
        # ensure output folder exists
        out_dir = os.path.join(os.path.dirname(pdf_path), 'printed_pdfs')
        os.makedirs(out_dir, exist_ok=True)
        # build timestamped filename
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        dest = os.path.join(out_dir, f"{base}_{timestamp}.pdf")
        # copy the file
        try:
            shutil.copy(pdf_path, dest)
        except Exception as e:
            messagebox.showerror('Print Error', f'Failed to save PDF: {e}')
            sys.exit(1)
        return

    # otherwise, send to a real printer via SumatraPDF
    try:
        # locate Sumatra next to script or exe
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        sumatra = os.path.join(base, 'SumatraPDF-3.5.2-64.exe')
        if not os.path.exists(sumatra):
            messagebox.showerror('Print Error', f'SumatraPDF not found at {sumatra}')
            sys.exit(1)

        proc = subprocess.Popen([
            sumatra,
            '-print-to', printer_name,
            '-print-settings', scale_opts,
            pdf_path
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # wait indefinitely (no timeout)
        stdout, stderr = proc.communicate()

        if proc.returncode != 0:
            messagebox.showerror('Print Error', stderr.decode().strip())
            sys.exit(1)

    except Exception as e:
        messagebox.showerror('Print Error', str(e))
        sys.exit(1)

# -- ZPL Processing -----------------------------------------------------------
def process_zpl_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    is_zpl = ext == '.zpl'
    is_zplii = ext == '.zplii'
    if not (is_zpl or is_zplii):
        print('Invalid file type.')
        sys.exit(1)
    # Load config
    zpl_prn, zplii_prn = read_config()
    # Convert ZPL to PDF
    try:    
        raw = open(file_path,'rb').read()
    except Exception as e:
        messagebox.showerror('File Path Error', f"Failed to find file: {file_path}\nPlease manually select the file.")
        fp = filedialog.askopenfilename(title='Select ZPL File', filetypes=[("ZPL and ZPLII Files", "*.zpl *.zplii")])
        if fp:
            process_zpl_file(fp)
        else:
            sys.exit()

    enc = chardet.detect(raw)['encoding'] or 'latin-1'
    zpl_code = raw.decode(enc, errors='replace')

    url = 'https://api.labelary.com/v1/printers/8dpmm/labels/4x6.21/0/' if is_zpl else 'https://api.labelary.com/v1/printers/8dpmm/labels/4x6.75/0/'
    resp = requests.post(url, files={'file': (file_path, zpl_code)}, headers={'Accept':'application/pdf','Content-Type':'application/x-www-form-urlencoded'})
    if resp.status_code != 200:
        print(f'ZPL conversion failed: {resp.status_code}')
        sys.exit(1)
    pdf_path = file_path.rsplit('.',1)[0] + '.pdf'
    with open(pdf_path,'wb') as f:
        f.write(resp.content)
    if is_zpl:
        rotate_pdf(pdf_path)
        crop_pdf_top(pdf_path,12)
    # Determine printer and set scale
    if is_zpl:
        printer_name = zpl_prn
        scale_opts = 'fit'
    elif is_zplii:
        printer_name = zplii_prn
        scale_opts = 'fit,center'

    # Use config or fallback
    installed = [info[2] for info in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    if printer_name and printer_name in installed:
        print_with_sumatra(pdf_path, printer_name, scale_opts)
    else:
        printer_selection_window(pdf_path)
    sys.exit(0)

# -- PDF Rotate & Crop --------------------------------------------------------
def rotate_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for p in reader.pages:
        p.rotate(180)
        writer.add_page(p)
    with open(pdf_path,'wb') as f:
        writer.write(f)

def crop_pdf_top(pdf_path, points_to_trim):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for p in reader.pages:
        mb = p.mediabox
        mb.upper_right = (mb.upper_right[0], mb.upper_right[1] - points_to_trim)
        p.mediabox = mb
        writer.add_page(p)
    with open(pdf_path,'wb') as f:
        writer.write(f)

# -- Main UI ------------------------------------------------------------------
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

def main():
    if len(sys.argv) > 1:
        process_zpl_file(sys.argv[1])
    else:
        root = tk.Tk()
        root.title('ZPL Label Printer')
        root.geometry('300x150')
        def print_label():
            fp = filedialog.askopenfilename(title='Select ZPL File', filetypes=[("ZPL and ZPLII Files", "*.zpl *.zplii")])
            if fp:
                process_zpl_file(fp)
        def config():
            configure_printers()
        tk.Button(root, text='Print Label', command=print_label, width=20, height=2).pack(pady=10)
        tk.Button(root, text='Configure Printers', command=config, width=20, height=2).pack(pady=10)
        root.mainloop()

if __name__ == '__main__':
    main()