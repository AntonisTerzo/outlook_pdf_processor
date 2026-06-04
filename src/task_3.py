import re
from pathlib import Path
from collections import defaultdict
from config import TASK3_FILE_TYPES
from pdf_utils import (
    detect_file_type_task3,
    extract_combo_id_from_filename,
    extract_packing_list_data,
    extract_hs_codes,
    extract_cargo_description_from_subject,
)
from outlook_utils import (
    connect_to_outlook, find_outlook_folder,
    initialize_com, uninitialize_com,
    create_unique_folder
)
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill


def _read_attachment_bytes(attachment):
    """
    Read an Outlook attachment's content into memory as bytes WITHOUT saving
    a permanent copy. Outlook's COM API has no direct in-memory read, so we
    write to a short-lived temp file, read it back, and delete it immediately.
    """
    import tempfile
    import os

    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        attachment.SaveAsFile(tmp_path)
        with open(tmp_path, "rb") as f:
            return f.read()
    except Exception as e:
        print(f"Error reading attachment bytes: {e}")
        return None
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def _collect_pdfs_from_message(message, outlook, log_func):
    """
    Walk an email's attachments and return a list of (filename, pdf_bytes).
    PDFs nested inside .msg attachments are also collected. Nothing is saved
    to a permanent location.
    """
    collected = []

    if not (hasattr(message, 'Attachments') and message.Attachments.Count > 0):
        return collected

    for attachment in message.Attachments:
        filename = attachment.FileName

        if filename.lower().endswith('.pdf'):
            data = _read_attachment_bytes(attachment)
            if data:
                collected.append((filename, data))
                log_func(f"  Read PDF: {filename}")

        elif filename.lower().endswith('.msg'):
            log_func(f"  Opening .msg file: {filename}")
            import tempfile
            import os
            tmp_msg = None
            try:
                fd, tmp_msg = tempfile.mkstemp(suffix=".msg")
                os.close(fd)
                attachment.SaveAsFile(tmp_msg)
                namespace = outlook.GetNamespace("MAPI")
                msg = namespace.OpenSharedItem(tmp_msg)
                if msg.Attachments.Count > 0:
                    for inner in msg.Attachments:
                        if inner.FileName.lower().endswith('.pdf'):
                            data = _read_attachment_bytes(inner)
                            if data:
                                collected.append((inner.FileName, data))
                                log_func(f"    Read PDF from .msg: {inner.FileName}")
            except Exception as e:
                log_func(f"    Error opening .msg: {e}")
            finally:
                if tmp_msg:
                    try:
                        os.remove(tmp_msg)
                    except Exception:
                        pass

    return collected


def _group_combos(pdfs, log_func):
    """
    Group (filename, bytes) PDFs by combo id (longest digit run in filename),
    detecting each file's type. The invoice file maps to None and is skipped.

    Returns dict:
        {combo_id: {"PACKING_LIST": bytes|None, "CUSTOMS_CODE": bytes|None}}
    """
    combos = defaultdict(lambda: {"PACKING_LIST": None, "CUSTOMS_CODE": None})

    for filename, data in pdfs:
        combo_id = extract_combo_id_from_filename(filename)
        if not combo_id:
            log_func(f"  ! Could not extract combo id from: {filename}")
            continue

        file_type = detect_file_type_task3(data, TASK3_FILE_TYPES)
        if file_type is None:
            # Most likely the invoice file, which we intentionally ignore.
            log_func(f"  (skipping non-target file: {filename})")
            continue

        if combos[combo_id][file_type] is None:
            combos[combo_id][file_type] = data
            log_func(f"  {filename} -> combo {combo_id} / {file_type}")
        else:
            log_func(f"  ! Duplicate {file_type} in combo {combo_id}: {filename} (keeping first)")

    return combos


def _build_combo_row(combo_id, combo_files, mail_description, log_func):
    """
    Build the 9-column row for a combo. Requires both PACKING_LIST and
    CUSTOMS_CODE. Returns (row_dict, notes_list); row is None if incomplete.
    """
    notes = []
    packing = combo_files.get("PACKING_LIST")
    customs = combo_files.get("CUSTOMS_CODE")

    if packing is None:
        notes.append("missing PACKING_LIST")
    if customs is None:
        notes.append("missing CUSTOMS_CODE")

    if packing is None or customs is None:
        return None, notes

    pl = extract_packing_list_data(packing)
    hs_codes = extract_hs_codes(customs)

    row = {
        "Cargo Description (Mail)": mail_description or "",
        "Cargo Description": ", ".join(pl["descriptions"]) if pl["descriptions"] else "",
        "IV": pl["iv"] or "",
        "SRN": pl["srn"] or "",
        "PCS": pl["pcs"] or "",
        "KG": pl["kg"] or "",
        "M3": pl["m3"] or "",
        "DIMS": pl["dims"] or "",
        "HS Code": ";".join(hs_codes) if hs_codes else "",
    }
    return row, notes

def _create_excel_report(rows, output_folder):
    """Create the consolidated Excel report - one sheet, one row per combo."""
    if not rows:
        return None

    excel_path = output_folder / "Task_3_Report.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Combos"

    headers = [
        "Cargo Description (Mail)",
        "Cargo Description",
        "IV",
        "SRN",
        "PCS",
        "KG",
        "M3",
        "DIMS",
        "HS Code",
    ]

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row.get(header, ""))

    widths = {"A": 25, "B": 40, "C": 18, "D": 18, "E": 10, "F": 12, "G": 12, "H": 18, "I": 30}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A2"
    wb.save(excel_path)
    return excel_path

def run_task_3(log_func=print):
    """
    Task 3:
    - Connect to Outlook, find Inbox/Task_3 folder (notify if missing).
    - Read all PDF attachments into memory (PDFs are NOT downloaded/saved).
    - Group PDFs into combos by the longest digit-run in their filename.
    - Each combo needs PACKING_LIST + CUSTOMS_CODE (invoice is ignored).
    - Extract data and write one consolidated Task_3_Report.xlsx to Downloads.
    Returns: (processed_count, incomplete_combos_list, excel_folder_path)
    """
    initialize_com()

    try:
        log_func("Task 3: Attempting to connect to Outlook...")
        outlook, namespace = connect_to_outlook()
        log_func("Successfully connected to Outlook")

        task_folder = find_outlook_folder(namespace, "Task_3")
        if not task_folder:
            log_func("Error: Task_3 folder not found in Inbox")
            return 0, [], None

        log_func(f"Found folder: {task_folder.Name}")

        if task_folder.Items.Count == 0:
            log_func("\nThere were no emails to process inside Task_3 folder.")
            return 0, [], None

        log_func(f"Processing {task_folder.Items.Count} emails\n")

        all_rows = []
        incomplete_combos = []  # list of (email_subject, combo_id, notes)
        total_processed = 0

        for message in task_folder.Items:
            if not (hasattr(message, 'Attachments') and message.Attachments.Count > 0):
                continue

            subject_raw = message.Subject if message.Subject else "No_Subject"
            mail_description = extract_cargo_description_from_subject(subject_raw)

            log_func(f"\nProcessing email: {subject_raw}")
            if mail_description:
                log_func(f"  Cargo Description (Mail): {mail_description}")
            else:
                log_func("  ! Could not find description (- word - CHINA) in subject")

            pdfs = _collect_pdfs_from_message(message, outlook, log_func)
            if not pdfs:
                log_func("  No PDF attachments found in this email")
                continue

            combos = _group_combos(pdfs, log_func)

            for combo_id, files in combos.items():
                row, notes = _build_combo_row(combo_id, files, mail_description, log_func)
                if row is None:
                    log_func(f"  ! Combo {combo_id} incomplete: {', '.join(notes)}")
                    incomplete_combos.append((subject_raw, combo_id, notes))
                else:
                    log_func(f"  Combo {combo_id}: complete")
                    all_rows.append(row)
                    total_processed += 1

        # Write the consolidated Excel to Downloads
        excel_folder = None
        if all_rows:
            downloads = Path.home() / "Downloads"
            excel_folder = create_unique_folder(downloads, "outlook_pdf_processor_task_3")
            log_func("\nCreating Excel report...")
            excel_path = _create_excel_report(all_rows, excel_folder)
            if excel_path:
                log_func(f"Excel report created: {excel_path}")

        log_func("\n" + "=" * 60)
        log_func("Task 3 Processing complete")
        log_func(f"{total_processed} combo(s) written to Excel")
        if excel_folder:
            log_func(f"Location: {excel_folder}")
        if incomplete_combos:
            log_func(f"\nWARNING: {len(incomplete_combos)} incomplete combo(s):")
            for subj, cid, notes in incomplete_combos:
                log_func(f"  - {subj} / combo {cid}: {', '.join(notes)}")
        log_func("=" * 60)

        return total_processed, incomplete_combos, excel_folder

    except Exception as e:
        log_func(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        return 0, [], None

    finally:
        uninitialize_com()
