# ZPL-to-PDF

Simple Python script to convert ZPL and ZPLII files into printable PDFs using the Labelary API.

## Repository Layout

- `app/` – contains the application and its required files (`zpl-to-pdf-configurable.py`, `config.txt`, `SumatraPDF-3.5.2-64.exe`, `SumatraPDF-settings.txt`). `label_settings.txt` will be generated in this folder if it does not exist.
- `examples/` – sample ZPL files for testing.
- `legacy/` – older scripts kept for reference.

Run the tool:

```bash
python app/zpl-to-pdf-configurable.py
```
