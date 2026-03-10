# Outlook_PDF_Processor

A small Python utility that automates a repetitive team task: scans a specified folder in **Outlook (classic)**, downloads PDF attachments, extracts the **city name** and **packing list number** from each document, renames files to a standardized `<city>_<packing-list>.pdf` format, and opens the destination folder for immediate review.

---

## Features

- Connects to Outlook (classic desktop) via MAPI using `pywin32`.
- Downloads PDF attachments from a specified Outlook folder.
- Extracts packing-list numbers and city names from PDF text using `PyPDF`.
- Renames files to a consistent pattern: `{city}_{packing}.pdf`.
- Opens the destination folder after processing so the user can access files immediately.

---

## Requirements

- **Platform:** Windows (tested on Windows 10/11)
- **Email client:** Outlook (classic / desktop)
- **Python:** 3.10+

## Quickstart

1. **Clone the repository**

```bash
git clone https://github.com/AntonisTerzo/outlook_pdf_processor.git
cd outlook_PDF_Processor/src
```

2. **Install dependencies**

```bash
pip install pywin32 PyPDF2
```

3. **Run the script**

```bash
python outlook_pdf_processor.py
```

After processing, the script will open the configured download folder automatically so you can verify the renamed files.

Or you can use the outlook_pdf_processor.exe file.

## How it works (high level)

1. Connects to Outlook using `pywin32` MAPPI interfaces and opens the specified folder.
2. Iterates over messages and downloads PDF attachments to ``.
3. Extracts text from each PDF using `PyPDF2`.
4. Uses pattern matching (regex) or simple parsing to locate a **city name** and a **packing list number** inside the extracted text.
5. Renames the PDF using the configured pattern (for example: `Athens_12345678.pdf`).
6. Opens the destination folder for the user to continue working.

---
