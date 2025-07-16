import os
import tkinter as tk
from tkinter import ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LABEL_CONFIG_PATH = os.path.join(BASE_DIR, 'label_settings.txt')
LABEL_TYPES = ['zpl', 'zplii']
DPMM_OPTIONS = ['6', '8', '10', '12']
SCALE_OPTIONS = ['fit', 'shrink', 'noscale', 'center']

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
                    while len(parts) < 7:
                        parts.append('')
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

        # DPMM Combobox
        tk.Label(frame, text='DPMM:', font=('Arial', 12)).pack()
        dpmm_var = tk.StringVar(value=settings['dpmm'])
        dpmm_menu = ttk.Combobox(frame, textvariable=dpmm_var, values=DPMM_OPTIONS, state='readonly')
        dpmm_menu.pack(pady=5)
        label_fields['dpmm'] = dpmm_var

        for key in ['width', 'height', 'rotate', 'crop', 'paper']:
            tk.Label(frame, text=f'{key.capitalize()}:', font=('Arial', 12)).pack()
            entry = tk.Entry(frame, width=30)
            entry.insert(0, settings[key])
            entry.pack(pady=5)
            label_fields[key] = entry

        tk.Label(frame, text='Scale Options:', font=('Arial', 12)).pack()
        scale_var = tk.StringVar(value=settings['scaleopts'])
        scale_menu = ttk.Combobox(frame, textvariable=scale_var, values=SCALE_OPTIONS, state='readonly')
        scale_menu.pack(pady=5)
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

if __name__ == "__main__":
    configure_label_settings()
