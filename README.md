# Outlook_PDF_Processor

A small Python utility that automates a repetitive team task: scans a specified folder in **Outlook (classic)**, downloads PDF attachments, extracts the **city name** and **packing list number** from each document, renames files to a standardized `<city>_<packing-list>.pdf` format, and opens the destination folder for immediate review.

---

## Features

**Task 1 — Rename by city + packing-list number**
- Connects to Outlook (classic desktop) via MAPI using `pywin32`.
- Downloads PDF attachments from the **Task_1** Outlook folder.
- Extracts packing-list numbers and city names from PDF text using `pypdf`.
- Renames files to a consistent pattern: `{city}_{packing}.pdf`.
- Unmatched files are moved to a **MANUAL REVIEW** subfolder.
- Opens the destination folder after processing so the user can access files immediately.

**Task 2 — Rename by city from filename + extract dimensions**
- Extracts the **city name directly from the PDF filename** using a configurable city list.
- Renames files to `{city} {n}.pdf` with a per-email counter (e.g. `BRASILIEN 1.pdf`).
- Extracts **package dimensions** from the `Abmessung(MM)` column, converts mm → cm with custom rounding, and warns when a section has no dimensions.
- Generates a per-email **Excel report** (`Dimensions_Report.xlsx`) with one sheet per city.
- Each email is saved in its own subfolder named after the (sanitized) email subject.
- Unmatched files are moved to a **MANUAL REVIEW** subfolder.

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
pip install pywin32 pypdf openpyxl
```

3. **Run the script**

```bash
python outlook_pdf_processor.py
```

After processing, the script will open the configured download folder automatically so you can verify the renamed files.

Or you can use the outlook_pdf_processor.exe file.

## How it works (high level)

### Task 1
1. Connects to Outlook using `pywin32` MAPI and opens the **Task_1** folder.
2. Iterates over emails and downloads PDF attachments (including PDFs nested inside `.msg` files).
3. Extracts text from each PDF using `pypdf`.
4. Searches the `Warenempfänger:` section for a city name and locates the 8-digit packing-list number.
5. Renames the PDF to `{city}_{number}.pdf` (e.g. `Athens_12345678.pdf`); unmatched files go to **MANUAL REVIEW**.
6. Opens the destination folder for the user to continue working.

### Task 2
1. Connects to Outlook using `pywin32` MAPI and opens the **Task_2** folder.
2. Creates a subfolder per email named after the sanitized email subject.
3. Matches each PDF's filename against the configured city list to determine the city.
4. Extracts package dimensions from the `Abmessung(MM)` column, converts mm → cm with custom rounding, and warns if a section has no dimensions.
5. Renames the PDF to `{city} {n}.pdf` (e.g. `BRASILIEN 1.pdf`); unmatched files go to **MANUAL REVIEW**.
6. Generates a per-email Excel report (`Dimensions_Report.xlsx`) with one sheet per city.
7. Opens the destination folder for the user to continue working.

---
