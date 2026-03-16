from pathlib import Path
from collections import defaultdict, Counter
from config import TASK2_CITIES, BRASILIEN_SUBCITIES, MANUAL_REVIEW_FALLBACK_CITIES
from pdf_utils import extract_city_from_filename_task2, extract_dimensions_from_pdf, check_city_inside_pdf
from outlook_utils import (
    connect_to_outlook, find_outlook_folder,
    process_msg_file, initialize_com, uninitialize_com,
    create_unique_folder
)
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill


# Track city counts for numbering
city_counter = defaultdict(int)

# Store dimensions data: {city: {file_number: [dimensions]}}
dimensions_data = defaultdict(lambda: defaultdict(list))


def process_pdf_task2(temp_pdf_path, output_folder, original_filename):
    """
    Process a single PDF for Task 2.
    Extracts city from filename, extracts dimensions, and renames with numbering.
    Special handling:
    - BRASILIEN (not JOINVILLE): Check inside PDF for INDAIATUBA or RIO CLARO
    - Manual review files: Check inside PDF for TAPUKARA
    Returns dict with 'success' and 'message' keys.
    """
    city = extract_city_from_filename_task2(original_filename, TASK2_CITIES)

    if city:
        # Special handling for BRASILIEN (not BRASILIEN JOINVILLE)
        if city == "BRASILIEN":
            subcity = check_city_inside_pdf(temp_pdf_path, BRASILIEN_SUBCITIES)
            if subcity:
                city = f"BRASILIEN {subcity}"

        # Increment counter for this city
        city_counter[city] += 1
        count = city_counter[city]

        # Extract dimensions from PDF
        dimensions = extract_dimensions_from_pdf(temp_pdf_path)

        # Store dimensions for this file
        dimensions_data[city][count] = dimensions

        # Create filename: "BRASILIEN INDAIATUBA 1.pdf", etc.
        new_filename = f"{city} {count}.pdf"
        final_path = output_folder / new_filename
        temp_pdf_path.rename(final_path)

        return {'success': True, 'message': f"Saved as: {new_filename} ({len(dimensions)} dimensions found)"}
    else:
        # Manual review - check for fallback cities inside PDF
        fallback_city = check_city_inside_pdf(
            temp_pdf_path, MANUAL_REVIEW_FALLBACK_CITIES)

        if fallback_city:
            # Found a fallback city (e.g., TAPUKARA)
            city_counter[fallback_city] += 1
            count = city_counter[fallback_city]

            # Extract dimensions
            dimensions = extract_dimensions_from_pdf(temp_pdf_path)
            dimensions_data[fallback_city][count] = dimensions

            new_filename = f"{fallback_city} {count}.pdf"
            final_path = output_folder / new_filename
            temp_pdf_path.rename(final_path)

            return {'success': True, 'message': f"Saved as: {new_filename} (found via PDF scan)"}
        else:
            # Move to MANUAL REVIEW folder with unique name
            manual_folder = output_folder / "MANUAL REVIEW"
            manual_folder.mkdir(exist_ok=True)

            # Create unique filename if file already exists
            counter = 1
            final_path = manual_folder / original_filename

            while final_path.exists():
                name_without_ext = original_filename.replace(
                    '.pdf', '').replace('.PDF', '')
                final_path = manual_folder / \
                    f"{name_without_ext}({counter}).pdf"
                counter += 1

            temp_pdf_path.rename(final_path)
            return {'success': False, 'message': f"Moved to MANUAL REVIEW: {final_path.name}"}


def create_excel_report(output_folder):
    """
    Create Excel file with dimensions grouped by city.
    Each city gets its own sheet with files separated.
    """
    if not dimensions_data:
        return None

    excel_path = output_folder / "Dimensions_Report.xlsx"
    wb = openpyxl.Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Create a sheet for each city
    for city in sorted(dimensions_data.keys()):
        ws = wb.create_sheet(title=city[:31])  # Excel sheet names max 31 chars

        # Header styling
        header_fill = PatternFill(
            start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        # Add headers
        ws['A1'] = 'File'
        ws['B1'] = 'Dimension'
        ws['C1'] = 'Count'

        for cell in ['A1', 'B1', 'C1']:
            ws[cell].fill = header_fill
            ws[cell].font = header_font
            ws[cell].alignment = Alignment(
                horizontal='center', vertical='center')

        row = 2

        # Process each file for this city
        for file_num in sorted(dimensions_data[city].keys()):
            dimensions_list = dimensions_data[city][file_num]

            if not dimensions_list:
                # No dimensions found
                ws[f'A{row}'] = f"{city} {file_num}"
                ws[f'B{row}'] = "No dimensions found"
                ws[f'C{row}'] = 0
                row += 1
            else:
                # Group and count dimensions
                dimension_counts = Counter(dimensions_list)

                # Sort by dimension for consistency
                for dimension, count in sorted(dimension_counts.items()):
                    ws[f'A{row}'] = f"{city} {file_num}"
                    ws[f'B{row}'] = dimension
                    ws[f'C{row}'] = count
                    row += 1

            # Add separator row
            row += 1

        # Auto-adjust column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 10

    wb.save(excel_path)
    return excel_path


def run_task_2(log_func=print):
    """
    Main Task 2 logic: Download PDFs from Task_2 folder, rename by city with numbering,
    extract dimensions and create Excel report.
    """
    # Reset counters and data for this run
    city_counter.clear()
    dimensions_data.clear()

    initialize_com()

    try:
        # Create unique task folder
        downloads = Path.home() / "Downloads"
        pdf_folder = create_unique_folder(
            downloads, "outlook_pdf_processor_task_2")
        log_func(f"Created folder: {pdf_folder.name}")

        temp_folder = pdf_folder / "temp"
        temp_folder.mkdir(exist_ok=True)

        log_func("Task 2: Attempting to connect to Outlook...")
        outlook, namespace = connect_to_outlook()
        log_func(" Successfully connected to Outlook!")

        # Find Task_2 folder
        task_folder = find_outlook_folder(namespace, "Task_2")

        if not task_folder:
            log_func(" Error: 'Task_2' folder not found!")
            return 0, 0

        log_func(f" Found folder: {task_folder.Name}")
        log_func(f"Processing {task_folder.Items.Count} emails...\n")

        processed_count = 0
        manual_review_count = 0

        for message in task_folder.Items:
            if hasattr(message, 'Attachments'):
                for attachment in message.Attachments:
                    filename = attachment.FileName

                    # Handle .msg files
                    if filename.lower().endswith('.msg'):
                        log_func(f"\nFound .msg file: {filename}")
                        temp_msg_path = temp_folder / filename
                        attachment.SaveAsFile(str(temp_msg_path))

                        msg_processed, msg_manual = process_msg_file(
                            temp_msg_path, temp_folder, pdf_folder, outlook,
                            process_pdf_task2, log_func
                        )
                        processed_count += msg_processed
                        manual_review_count += msg_manual

                        temp_msg_path.unlink()

                    # Handle PDF files
                    elif filename.lower().endswith('.pdf'):
                        temp_path = temp_folder / filename
                        attachment.SaveAsFile(str(temp_path))
                        log_func(f"\nDownloaded PDF: {filename}")

                        result = process_pdf_task2(
                            temp_path, pdf_folder, filename)

                        if result['success']:
                            log_func(f" {result['message']}")
                            processed_count += 1
                        else:
                            log_func(f" {result['message']}")
                            manual_review_count += 1

        # Create Excel report with dimensions
        if processed_count > 0:
            log_func("\nCreating Excel report with dimensions...")
            excel_path = create_excel_report(pdf_folder)
            if excel_path:
                log_func(f" Excel report created: {excel_path.name}")

        # Final summary
        log_func("\n" + "="*60)
        log_func("Task 2 Processing complete!")
        log_func(f" {processed_count} files renamed and saved.")
        log_func(f"  Location: {pdf_folder}")

        if manual_review_count > 0:
            manual_review_folder = pdf_folder / "MANUAL REVIEW"
            log_func(
                f"\n WARNING: {manual_review_count} file(s) need manual review")
            log_func(f"  Location: {manual_review_folder}")

        log_func("="*60)

        # Clean up temp folder if empty
        if temp_folder.exists():
            remaining_files = list(temp_folder.glob('*'))
            if not remaining_files:
                temp_folder.rmdir()

        return processed_count, manual_review_count, pdf_folder

    except Exception as e:
        log_func(f"\n Error: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0, None

    finally:
        uninitialize_com()
