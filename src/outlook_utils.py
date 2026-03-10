import win32com.client
import pythoncom


def connect_to_outlook():
    """
    Connect to Outlook application.
    Returns outlook object and namespace.
    """
    try:
        outlook = win32com.client.gencache.EnsureDispatch(
            "Outlook.Application")
    except:
        outlook = win32com.client.Dispatch("Outlook.Application")

    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace


def find_outlook_folder(namespace, folder_name):
    """
    Find a folder in Outlook by name.
    Returns the folder object if found, None otherwise.
    """
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

    # Check in Inbox folders first
    for folder in inbox.Folders:
        if folder.Name.lower() == folder_name.lower():
            return folder

    # Search in all root folders
    for folder in namespace.Folders:
        for subfolder in folder.Folders:
            if subfolder.Name.lower() == folder_name.lower():
                return subfolder

    return None


def process_msg_file(msg_path, temp_folder, output_folder, outlook, process_pdf_func, log_func=print):
    """
    Open a .msg file and extract PDFs from it.
    Uses the provided process_pdf_func to handle each PDF.
    Returns: (processed_count, manual_review_count)
    """
    try:
        namespace = outlook.GetNamespace("MAPI")
        msg = namespace.OpenSharedItem(str(msg_path))

        log_func(f"  Opening .msg file: {msg_path.name}")
        processed = 0
        manual_review = 0

        if msg.Attachments.Count > 0:
            log_func(
                f"    Found {msg.Attachments.Count} attachment(s) in .msg file")
            for attachment in msg.Attachments:
                if attachment.FileName.lower().endswith('.pdf'):
                    temp_pdf_path = temp_folder / attachment.FileName
                    attachment.SaveAsFile(str(temp_pdf_path))
                    log_func(
                        f"    Found PDF inside .msg: {attachment.FileName}")

                    # Use the provided processing function
                    result = process_pdf_func(
                        temp_pdf_path, output_folder, attachment.FileName)

                    if result['success']:
                        log_func(f"    ✓ {result['message']}")
                        processed += 1
                    else:
                        log_func(f"    ⚠ {result['message']}")
                        manual_review += 1
                else:
                    log_func(f"    Skipping non-PDF: {attachment.FileName}")
        else:
            log_func(f"    No attachments found in .msg file")

        return processed, manual_review

    except Exception as e:
        log_func(f"  Error processing .msg file: {e}")
        return 0, 0


def initialize_com():
    """Initialize COM for the current thread."""
    pythoncom.CoInitialize()


def uninitialize_com():
    """Uninitialize COM for the current thread."""
    pythoncom.CoUninitialize()
