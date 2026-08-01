# Outlook_PDF_Processor

A small Python utility that automates repetitive team tasks around PDF attachments in **Outlook (classic)**. It scans dedicated Outlook folders, processes the PDF attachments found there, and produces renamed files and/or Excel reports depending on the task. A simple GUI with one button per task shows live progress and a summary when done.

---

## Features

**Task 1 — Rename by city + packing-list number**
- Connects to Outlook (classic desktop) via MAPI using `pywin32`.
- Downloads PDF attachments from the **Task_1** Outlook folder (including PDFs nested inside `.msg` files).
- Extracts packing-list numbers and city names from PDF text using `pypdf`.
- Renames files to a consistent pattern: `{city}_{packing}.pdf`.
- Special handling:
  - **TIANJIN** — checks the `Versand` field for specific values (e.g. `China TI ASS`) and includes them in the filename: `TIANJIN_China TI ASS_12345678.pdf`.
  - **SUZHOU** — checks the `Warenempfänger` section for the keyword `Motors` and renames to `SUZHOU_Motors_87654321.pdf` when found.
- Each email gets its own subfolder named after the (sanitized) email subject.
- Unmatched files are moved to a **MANUAL REVIEW** subfolder.
- Opens the destination folder after processing so the user can access files immediately.

**Task 2 — Rename by city from filename + extract dimensions**
- Extracts the **city name directly from the PDF filename** using a configurable city list.
- Renames files to `{city} {n}.pdf` with a per-email counter (e.g. `BRASILIEN 1.pdf`).
- Special handling:
  - **BRASILIEN** — checks inside the PDF for the subcities `INDAIATUBA` / `RIO CLARO` and includes them in the name: `BRASILIEN INDAIATUBA 1.pdf`.
  - **LYMAN ELECTRONICS** — checks inside the PDF for `MAXOLUTION` and uses it as the name: `MAXOLUTION 1.pdf`.
  - Files headed for manual review are first checked for fallback cities (e.g. `TAPUKARA`) inside the PDF.
- Extracts **package dimensions** from the `Abmessung(MM)` column, converts mm → cm with custom rounding, and warns when a section has no dimensions (`variofix` entries use the fixed dimension `36x23x19`).
- Generates a per-email **Excel report** (`Dimensions_Report.xlsx`) with one sheet per city.
- Each email is saved in its own subfolder named after the (sanitized) email subject.
- Unmatched files are moved to a **MANUAL REVIEW** subfolder.

**Task 3 — Extract shipment data into a consolidated Excel report**
- Reads PDF attachments from the **Task_3** Outlook folder **in memory** — nothing is downloaded except the final report.
- Each email contains one or more **combos** of shipment documents that share a number in their filenames (e.g. `...651515...`). A combo consists of:
  - `LISTA DE CONTENIDO / PACKING LIST`
  - `PARTIDAS ESTADÍSTICAS / CUSTOMS CODE`
  - (`FACTURA / INVOICE` files are ignored)
- File types are detected from the **content** of each PDF, and tables are parsed with `pdfplumber` for reliable cell-level extraction.
- Produces one consolidated `Task_3_Report.xlsx` in Downloads with **one row per combo** and 9 columns:

  | Column | Source |
  |---|---|
  | Cargo Description (Mail) | Email subject — the text between dashes before `CHINA` |
  | Cargo Description | Packing list — `Descripción` column (multiple values comma-separated) |
  | IV | Packing list — value after `Factura/Invoice` **or** `Proforma` |
  | SRN | Packing list — value after `Nº Envío/Shipment Nr.` |
  | PCS | Packing list summary — `Total Nª Parcels` |
  | KG | Packing list summary — `Total Gross Weight (Kgs)` |
  | M3 | Packing list summary — `Total Volume (m3)` |
  | DIMS | Packing list — dimensions converted to cm, grouped with counts, e.g. `2-80x50x15 1-43x35x31` |
  | HS Code | Customs code file — `Cod. Arancel/Customs code`, dots removed, multiple codes joined with `;` (e.g. `64039996;73181590`) |

- **Missing-data handling:**
  - If the customs-code file is missing, the row is still written; the HS Code cell reads `NOT FOUND - check PDF` and is shaded **red**.
  - If any packing-list field can't be extracted, the cell reads `NOT FOUND - check PDF` and the **entire row** is shaded red.
  - Combos with no packing list are reported as incomplete in the log and final summary.
- Extraction diagnostics (ambiguous or missing values) are written to the live log.

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
pip install pywin32 pypdf pdfplumber openpyxl
```

3. **Run the app**

```bash
python main.py
```

This opens the GUI with a button per task (`Start Task_1`, `Start Task_2`, `Start Task_3`) and a live log. Tasks can also be run from the command line:

```bash
python main.py --task1
python main.py --task2
python main.py --task3
```

Or use the prebuilt `outlook_pdf_processor.exe` (built automatically by GitHub Actions on pushes to the build branch).

## How it works (high level)

### Task 1
1. Connects to Outlook using `pywin32` MAPI and opens the **Task_1** folder.
2. Iterates over emails and downloads PDF attachments (including PDFs nested inside `.msg` files).
3. Extracts text from each PDF using `pypdf`.
4. Searches the `Warenempfänger:` section for a city name (takes the **last** match) and locates the 8-digit packing-list number after `Pack- und Gewichtsliste Nr.`.
5. Applies TIANJIN/SUZHOU special handling, then renames the PDF to `{city}_{number}.pdf`; unmatched files go to **MANUAL REVIEW**.
6. Opens the destination folder for the user to continue working.

### Task 2
1. Connects to Outlook using `pywin32` MAPI and opens the **Task_2** folder.
2. Creates a subfolder per email named after the sanitized email subject.
3. Matches each PDF's filename against the configured city list (longest names first) to determine the city, with BRASILIEN / LYMAN ELECTRONICS / fallback-city checks inside the PDF.
4. Extracts package dimensions from the `Abmessung(MM)` column (reading until the `Bestellung` section), converts mm → cm with custom rounding, and warns if a section has no dimensions.
5. Renames the PDF to `{city} {n}.pdf`; unmatched files go to **MANUAL REVIEW**.
6. Generates a per-email Excel report (`Dimensions_Report.xlsx`) with one sheet per city.
7. Opens the destination folder for the user to continue working.

### Task 3
1. Connects to Outlook using `pywin32` MAPI and opens the **Task_3** folder (notifies the user if it doesn't exist or is empty).
2. Reads every PDF attachment into memory (including PDFs nested inside `.msg` files) — the PDFs themselves are never saved to disk.
3. Groups PDFs into **combos** by the longest digit run in their filenames, and detects each file's type from its content.
4. Parses the packing list with `pdfplumber` (table-aware extraction): descriptions from the `Descripción` column, summary values (PCS/KG/M3) line-by-line from the `RESUMEN / SUMMARY` block, dimensions with duplicate counting, and the IV/SRN reference numbers.
5. Extracts HS codes from the customs-code file and strips the dots.
6. Writes one consolidated `Task_3_Report.xlsx` to a unique folder in Downloads, with red shading for any missing data, and opens the folder.

## Configuration

City lists and Task 3 file-type identifiers live in `src/config.py`:
- `TASK1_CITIES`, `TASK2_CITIES` — recognized city names.
- `BRASILIEN_SUBCITIES`, `LYMAN_ELECTRONICS_SUBCITIES`, `MANUAL_REVIEW_FALLBACK_CITIES` — Task 2 special cases.
- `TASK3_FILE_TYPES` — the header text used to identify packing-list and customs-code PDFs.

## Project structure

```
repo/
├── .github/
│   └── workflows/
│       └── build.yml          # Builds the EXE on push (PyInstaller)
└── src/
    ├── main.py                # GUI application (3 task buttons + live log)
    ├── task_1.py              # Task 1 logic
    ├── task_2.py              # Task 2 logic
    ├── task_3.py              # Task 3 logic
    ├── config.py              # City lists and configuration
    ├── pdf_utils.py           # PDF extraction (pypdf for Tasks 1-2, pdfplumber for Task 3)
    └── outlook_utils.py       # Outlook connection utilities
```

---
