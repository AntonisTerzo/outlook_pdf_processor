from pathlib import Path
from collections import defaultdict
from config import TASK2_CITIES
from pdf_utils import extract_city_from_filename_task2
from outlook_utils import (
    connect_to_outlook, find_outlook_folder,
    process_msg_file, initialize_com, uninitialize_com
)


# Track city counts for numbering
city_counter = defaultdict(int)


def process_pdf_task2(temp_pdf_path, output_folder, original_filename):
    """
    Process a single PDF for Task 2.
    Extracts city from filename and renames with numbering if duplicates.
    Returns dict with 'success' and 'message' keys.
    """
    city = extract_city_from_filename_task2(original_filename, TASK2_CITIES)

    if city:
        # Increment counter for this city
        city_counter[city] += 1
        count = city_counter[city]

        new_filename = f"{city} {count}.pdf"
        final_path = output_folder / new_filename
        temp_pdf_path.rename(final_path)

        return {'success': True, 'message': f"Saved as: {new_filename}"}
    else:
        return {'success': False, 'message': f"Could not find city in filename: {original_filename}"}


def run_task_2(log_func=print):
    """
    Main Task 2 logic: Download PDFs from Task_2 folder and rename by city with numbering.
    """
    # Reset city counter for this run
    city_counter.clear()

    initialize_com()

    try:
        # Create PDF extraction folder
        downloads = Path.home() / "Downloads"
        pdf_folder = downloads / "pdf extraction"
        pdf_folder.mkdir(exist_ok=True)

        temp_folder = pdf_folder / "temp_task2"
        temp_folder.mkdir(exist_ok=True)

        log_func("Task 2: Attempting to connect to Outlook...")
        outlook, namespace = connect_to_outlook()
        log_func("✓ Successfully connected to Outlook!")

        # Find Task_2 folder
        task_folder = find_outlook_folder(namespace, "Task_2")

        if not task_folder:
            log_func("✗ Error: 'Task_2' folder not found!")
            return 0, 0

        log_func(f"✓ Found folder: {task_folder.Name}")
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
                            log_func(f"✓ {result['message']}")
                            processed_count += 1
                        else:
                            log_func(f"⚠ {result['message']}")
                            log_func(
                                f"  Keeping in temp folder for manual review")
                            manual_review_count += 1

        # Final summary
        log_func("\n" + "="*60)
        log_func("Task 2 Processing complete!")
        log_func(f"✓ {processed_count} files renamed and saved.")

        if manual_review_count > 0:
            log_func(
                f"\n⚠ WARNING: {manual_review_count} file(s) need manual review")
            log_func(f"  Location: {temp_folder}")

        log_func("="*60)

        # Clean up temp folder if empty
        if temp_folder.exists():
            remaining_files = list(temp_folder.glob('*'))
            if not remaining_files:
                temp_folder.rmdir()

        return processed_count, manual_review_count

    except Exception as e:
        log_func(f"\n✗ Error: {e}")
        return 0, 0

    finally:
        uninitialize_com()
