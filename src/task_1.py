from pathlib import Path
from config import TASK1_CITIES
from pdf_utils import extract_info_from_pdf_task1
from outlook_utils import (
    connect_to_outlook, find_outlook_folder,
    process_msg_file, initialize_com, uninitialize_com,
    create_unique_folder
)


def process_pdf_task1(temp_pdf_path, output_folder, original_filename):
    """
    Process a single PDF for Task 1.
    Returns dict with 'success' and 'message' keys.
    """
    city, doc_number = extract_info_from_pdf_task1(temp_pdf_path, TASK1_CITIES)

    if city and doc_number:
        new_filename = f"{city}_{doc_number}.pdf"
        final_path = output_folder / new_filename
        temp_pdf_path.rename(final_path)
        return {'success': True, 'message': f"Saved as: {new_filename}"}
    else:
        # Move to MANUAL REVIEW folder with unique name
        manual_folder = output_folder / "MANUAL REVIEW"

        # Create unique filename if file already exists
        counter = 1
        final_path = manual_folder / original_filename

        while final_path.exists():
            name_without_ext = original_filename.replace(
                '.pdf', '').replace('.PDF', '')
            final_path = manual_folder / f"{name_without_ext}({counter}).pdf"
            counter += 1

        temp_pdf_path.rename(final_path)
        return {'success': False, 'message': f"Moved to MANUAL REVIEW: {final_path.name}"}


def run_task_1(log_func=print):
    """
    Main Task 1 logic: Download PDFs from Task_1 folder and rename them.
    """
    initialize_com()

    try:
        # Create unique task folder
        downloads = Path.home() / "Downloads"
        pdf_folder = create_unique_folder(
            downloads, "outlook_pdf_processor_task_1")
        log_func(f"Created folder: {pdf_folder.name}")

        temp_folder = pdf_folder / "temp"
        temp_folder.mkdir(exist_ok=True)

        manual_review_folder = pdf_folder / "MANUAL REVIEW"
        manual_review_folder.mkdir(exist_ok=True)

        log_func("Task 1: Attempting to connect to Outlook...")
        outlook, namespace = connect_to_outlook()
        log_func(" Successfully connected to Outlook!")

        # Find Task_1 folder
        task_folder = find_outlook_folder(namespace, "Task_1")

        if not task_folder:
            log_func(" Error: 'Task_1' folder not found!")
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
                            process_pdf_task1, log_func
                        )
                        processed_count += msg_processed
                        manual_review_count += msg_manual

                        temp_msg_path.unlink()

                    # Handle PDF files
                    elif filename.lower().endswith('.pdf'):
                        temp_path = temp_folder / filename
                        attachment.SaveAsFile(str(temp_path))
                        log_func(f"\nDownloaded PDF: {filename}")

                        result = process_pdf_task1(
                            temp_path, pdf_folder, filename)

                        if result['success']:
                            log_func(f" {result['message']}")
                            processed_count += 1
                        else:
                            log_func(f" {result['message']}")
                            manual_review_count += 1

        # Final summary
        log_func("\n" + "="*60)
        log_func("Task 1 Processing complete!")
        log_func(f" {processed_count} files renamed and saved.")
        log_func(f"  Location: {pdf_folder}")

        if manual_review_count > 0:
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
        return 0, 0, None

    finally:
        uninitialize_com()
