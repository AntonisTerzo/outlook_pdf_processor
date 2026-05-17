from pathlib import Path
import re
from config import TASK1_CITIES
from pdf_utils import extract_info_from_pdf_task1, extract_versand_for_tianjin, check_motors_in_warenempfanger
from outlook_utils import (
    connect_to_outlook, find_outlook_folder,
    process_msg_file, initialize_com, uninitialize_com,
    create_unique_folder
)


def clean_email_subject(subject):
    
    cleaned = subject.strip()
    
    # Replace all Windows invalid characters with underscore
    invalid_chars = r'[<>:"/\\|?*.]'
    cleaned = re.sub(invalid_chars, '_', cleaned)
    
    return cleaned


def process_pdf_task1(temp_pdf_path, output_folder, original_filename):
    """
    Process a single PDF for Task 1.
    Special handling:
    - TIANJIN: checks Versand field for specific values
    - SUZHOU: checks Warenempfänger for "Motors"
    Returns dict with 'success' and 'message' keys.
    """
    city, doc_number = extract_info_from_pdf_task1(temp_pdf_path, TASK1_CITIES)

    if city and doc_number:
        # Special handling for TIANJIN - check Versand field
        if city.upper() == "TIANJIN":
            versand_value = extract_versand_for_tianjin(temp_pdf_path)
            if versand_value:
                # Include Versand value in filename: TIANJIN_China TI ASS_12345678.pdf
                new_filename = f"{city}_{versand_value}_{doc_number}.pdf"
            else:
                # No Versand value found, just use TIANJIN
                new_filename = f"{city}_{doc_number}.pdf"
        # Special handling for SUZHOU - check for Motors
        elif city.upper() == "SUZHOU":
            has_motors = check_motors_in_warenempfanger(temp_pdf_path)
            if has_motors:
                # Include Motors in filename: SUZHOU_Motors_12345678.pdf
                new_filename = f"{city}_Motors_{doc_number}.pdf"
            else:
                # No Motors found, just use SUZHOU
                new_filename = f"{city}_{doc_number}.pdf"
        else:
            # Regular filename for other cities
            new_filename = f"{city}_{doc_number}.pdf"

        final_path = output_folder / new_filename
        temp_pdf_path.rename(final_path)
        return {'success': True, 'message': f"Saved as: {new_filename}"}
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
            final_path = manual_folder / f"{name_without_ext}({counter}).pdf"
            counter += 1

        temp_pdf_path.rename(final_path)
        return {'success': False, 'message': f"Moved to MANUAL REVIEW: {final_path.name}"}


def run_task_1(log_func=print):
    """
    Main Task 1 logic: Download PDFs from Task_1 folder and rename them.
    Each email gets its own subfolder named after the email subject.
    """
    initialize_com()

    try:
        log_func("Task 1: Attempting to connect to Outlook...")
        outlook, namespace = connect_to_outlook()
        log_func("Successfully connected to Outlook")

        # Find Task_1 folder
        task_folder = find_outlook_folder(namespace, "Task_1")

        if not task_folder:
            log_func("Error: Task_1 folder not found in Inbox")
            return 0, 0, None

        log_func(f"Found folder: {task_folder.Name}")

        # Check if there are any emails to process
        if task_folder.Items.Count == 0:
            log_func("\nThere were no emails to process inside Task_1 folder.")
            return 0, 0, None

        log_func(f"Processing {task_folder.Items.Count} emails\n")

        # Only create folders if there are emails to process
        downloads = Path.home() / "Downloads"
        pdf_folder = create_unique_folder(
            downloads, "outlook_pdf_processor_task_1")
        log_func(f"Created folder: {pdf_folder.name}")

        temp_folder = pdf_folder / "temp"
        temp_folder.mkdir(exist_ok=True)

        total_processed_count = 0
        total_manual_review_count = 0

        # Process each email - each gets its own subfolder
        for message in task_folder.Items:
            if hasattr(message, 'Attachments') and message.Attachments.Count > 0:
                email_subject = message.Subject if message.Subject else "No_Subject"
                email_subject = clean_email_subject(email_subject)
                email_subject = email_subject[:300]

                # Make folder name unique if it already exists
                base_email_folder = pdf_folder / email_subject
                email_folder = base_email_folder
                counter = 1
                while email_folder.exists():
                    email_folder = pdf_folder / f"{email_subject}({counter})"
                    counter += 1

                email_folder.mkdir(exist_ok=True)

                log_func(f"\nProcessing email: {email_subject}")

                processed_count = 0
                manual_review_count = 0

                for attachment in message.Attachments:
                    filename = attachment.FileName

                    # Handle .msg files
                    if filename.lower().endswith('.msg'):
                        log_func(f"\nFound .msg file: {filename}")
                        temp_msg_path = temp_folder / filename
                        attachment.SaveAsFile(str(temp_msg_path))

                        msg_processed, msg_manual = process_msg_file(
                            temp_msg_path, temp_folder, email_folder, outlook,
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
                            temp_path, email_folder, filename)

                        if result['success']:
                            log_func(result['message'])
                            processed_count += 1
                        else:
                            log_func(result['message'])
                            manual_review_count += 1

                log_func(
                    f"Email '{email_subject}': {processed_count} processed, {manual_review_count} need review")
                total_processed_count += processed_count
                total_manual_review_count += manual_review_count

        # Final summary
        log_func("\n" + "="*60)
        log_func("Task 1 Processing complete")
        log_func(f"{total_processed_count} files renamed and saved")
        log_func(f"Location: {pdf_folder}")

        if total_manual_review_count > 0:
            log_func(
                f"\nWARNING: {total_manual_review_count} files need manual review")

        log_func("="*60)

        # Clean up temp folder if empty
        if temp_folder.exists():
            remaining_files = list(temp_folder.glob('*'))
            if not remaining_files:
                temp_folder.rmdir()

        return total_processed_count, total_manual_review_count, pdf_folder

    except Exception as e:
        log_func(f"\nError: {e}")
        return 0, 0, None

    finally:
        uninitialize_com()