import tkinter as tk
from tkinter import scrolledtext, messagebox
import win32com.client
import pythoncom
import os
import re
from pathlib import Path
import PyPDF2
import threading
import ctypes

# List of cities to search for
CITIES = [
    "TIANJIN", "SHENYANG", "AMMAN", "WUHAN", "XIAN", "SUZHOU",
    "CAIRO", "HOLON", "ABIDJAN", "ANSAN-SI", "BRAMALEA", "BRIDGEPORT",
    "DELTA", "DORVAL", "EAST TAMAKI", "FOSHAN", "GUANGZHOU", "HAYWARD",
    "KOWLOON", "MONTEVIDEO", "MUANG CHONBURI", "SIHEUNG", "SINGAPUR",
    "TROY", "TULLAMARINE"
]


def extract_info_from_pdf(pdf_path):
    """
    Extract city name and document number from PDF.
    City: Search for any city from the CITIES list after 'Warenempfänger:' (takes LAST match)
    Number: 8-digit number after 'Pack- und Gewichtsliste'
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Find the section after "Warenempfänger:"
            warenempfanger_match = re.search(
                r'Warenempfänger:(.*?)(?=Lieferkondition:|$)', text, re.DOTALL)

            city = None
            if warenempfanger_match:
                warenempfanger_section = warenempfanger_match.group(1)

                # Find ALL cities in the section, then take the LAST one
                found_cities = []
                for city_name in CITIES:
                    # Find all matches for this city
                    for match in re.finditer(r'\b' + re.escape(city_name) + r'\b',
                                             warenempfanger_section, re.IGNORECASE):
                        found_cities.append((match.start(), city_name))

                # Sort by position and take the last one (closest to Pack- und Gewichtsliste)
                if found_cities:
                    found_cities.sort(key=lambda x: x[0])
                    city = found_cities[-1][1]  # Get the last city found

            # Find 8-digit number after "Pack- und Gewichtsliste"
            number_pattern = r'Pack- und Gewichtsliste\s+Nr\.\s+(\d{8})'
            number_match = re.search(number_pattern, text)

            if city and number_match:
                doc_number = number_match.group(1)
                return city, doc_number

            return None, None
    except Exception as e:
        print(f"Error extracting info from PDF: {e}")
        return None, None


def process_msg_file(msg_path, temp_folder, desktop, outlook, log_func=print):
    """
    Open a .msg file and extract PDFs from it.
    Returns: (processed_count, manual_review_count)
    """
    try:
        # Use Session.OpenSharedItem to open the .msg file
        namespace = outlook.GetNamespace("MAPI")
        msg = namespace.OpenSharedItem(str(msg_path))

        log_func(f"  Opening .msg file: {msg_path.name}")
        processed = 0
        manual_review = 0

        # Process attachments inside the .msg file
        if msg.Attachments.Count > 0:
            log_func(
                f"    Found {msg.Attachments.Count} attachment(s) in .msg file")
            for attachment in msg.Attachments:
                if attachment.FileName.lower().endswith('.pdf'):
                    # Save PDF to temp folder
                    temp_pdf_path = temp_folder / attachment.FileName
                    attachment.SaveAsFile(str(temp_pdf_path))
                    log_func(
                        f"    Found PDF inside .msg: {attachment.FileName}")

                    # Extract info and rename
                    city, doc_number = extract_info_from_pdf(temp_pdf_path)

                    if city and doc_number:
                        new_filename = f"{city}_{doc_number}.pdf"
                        final_path = desktop / new_filename
                        temp_pdf_path.rename(final_path)
                        log_func(f"    ✓ Saved as: {new_filename}")
                        processed += 1
                    else:
                        log_func(f"    ⚠ Could not find city from list")
                        log_func(
                            f"      Keeping in temp folder for manual review")
                        manual_review += 1
                else:
                    log_func(f"    Skipping non-PDF: {attachment.FileName}")
        else:
            log_func(f"    No attachments found in .msg file")

        return processed, manual_review

    except Exception as e:
        log_func(f"  Error processing .msg file: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0


def run_pdf_processor(log_func=print):
    """
    Main PDF processing logic
    """
    # Initialize COM for this thread
    pythoncom.CoInitialize()

    try:
        desktop = Path.home() / "Desktop"
        temp_folder = desktop / "temp_outlook_downloads"
        temp_folder.mkdir(exist_ok=True)

        log_func("Attempting to connect to Outlook...")

        try:
            outlook = win32com.client.gencache.EnsureDispatch(
                "Outlook.Application")
        except:
            outlook = win32com.client.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        log_func("✓ Successfully connected to Outlook!")

        # Find test folder
        inbox = namespace.GetDefaultFolder(6)
        test_folder = None

        for folder in inbox.Folders:
            if folder.Name.lower() == "test":
                test_folder = folder
                break

        if not test_folder:
            log_func("Could not find 'test' folder in Inbox.")
            log_func("Searching in all folders...")
            for folder in namespace.Folders:
                for subfolder in folder.Folders:
                    if subfolder.Name.lower() == "test":
                        test_folder = subfolder
                        break
                if test_folder:
                    break

        if not test_folder:
            log_func("✗ Error: 'test' folder not found!")
            return 0, 0

        log_func(f"✓ Found folder: {test_folder.Name}")
        log_func(f"Processing {test_folder.Items.Count} emails...\n")

        processed_count = 0
        manual_review_count = 0

        for message in test_folder.Items:
            if hasattr(message, 'Attachments'):
                for attachment in message.Attachments:
                    filename = attachment.FileName

                    # Handle .msg files
                    if filename.lower().endswith('.msg'):
                        log_func(f"\nFound .msg file: {filename}")
                        temp_msg_path = temp_folder / filename
                        attachment.SaveAsFile(str(temp_msg_path))

                        msg_processed, msg_manual = process_msg_file(
                            temp_msg_path, temp_folder, desktop, outlook, log_func)
                        processed_count += msg_processed
                        manual_review_count += msg_manual

                        temp_msg_path.unlink()

                    # Handle PDF files
                    elif filename.lower().endswith('.pdf'):
                        temp_path = temp_folder / filename
                        attachment.SaveAsFile(str(temp_path))
                        log_func(f"\nDownloaded PDF: {filename}")

                        city, doc_number = extract_info_from_pdf(temp_path)

                        if city and doc_number:
                            new_filename = f"{city}_{doc_number}.pdf"
                            final_path = desktop / new_filename
                            temp_path.rename(final_path)
                            log_func(f"✓ Saved as: {new_filename}")
                            processed_count += 1
                        else:
                            log_func(
                                f"⚠ Could not find city from list in {filename}")
                            log_func(
                                f"  Keeping in temp folder for manual review")
                            manual_review_count += 1

        # Final summary
        log_func("\n" + "="*60)
        log_func("Processing complete!")
        log_func(f"✓ {processed_count} files renamed and saved to Desktop.")

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
        # Uninitialize COM
        pythoncom.CoUninitialize()

# ============================================================================
# GUI CLASS
# ============================================================================


class OutlookProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook PDF Processor")
        self.root.geometry("600x400")

        # Get user's display name from Windows
        try:
            GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
            NameDisplay = 3

            size = ctypes.pointer(ctypes.c_ulong(0))
            GetUserNameEx(NameDisplay, None, size)

            nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
            GetUserNameEx(NameDisplay, nameBuffer, size)

            username = nameBuffer.value

            # Trim company name: "Last, First / Company" -> "Last, First"
            if username and '/' in username:
                username = username.split('/')[0].strip()

            # If full name not available, fall back to USERNAME
            if not username:
                username = os.environ.get('USERNAME', 'User')
        except Exception:
            # Fallback to regular username if display name fails
            username = os.environ.get('USERNAME', 'User')

        # Welcome Message
        welcome_label = tk.Label(
            root,
            text=f"Welcome back {username}",
            font=("Arial", 18, "bold"),
            fg="#2C3E50"
        )
        welcome_label.pack(pady=20)

        # Start Button - Light Blue with border (smaller size)
        self.start_button = tk.Button(
            root,
            text="Start Task_1",
            command=self.start_task,
            font=("Arial", 12, "bold"),
            bg="#5DADE2",  # Light blue
            fg="white",
            width=15,
            height=1,
            relief="raised",  # Gives it a raised 3D effect
            borderwidth=3,    # Makes the border thicker
            cursor="hand2"    # Changes cursor to hand on hover
        )
        self.start_button.pack(pady=10)

        # Log Output Area
        log_label = tk.Label(root, text="Log Output:",
                             font=("Arial", 10, "bold"))
        log_label.pack(pady=(20, 5))

        self.log_text = scrolledtext.ScrolledText(
            root, width=70, height=15, state='disabled')
        self.log_text.pack(pady=10)

    def log(self, message):
        """Add message to log output"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def start_task(self):
        """Start the PDF processing task in a separate thread"""
        self.start_button.config(state='disabled', text="Processing...")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        # Run in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.run_task)
        thread.start()

    def run_task(self):
        """Run the processor and show results"""
        try:
            processed, manual = run_pdf_processor(self.log)

            if processed > 0 or manual > 0:
                # Open Desktop folder where files are saved
                desktop = Path.home() / "Desktop"
                os.startfile(desktop)

                messagebox.showinfo(
                    "Success", f"Processing complete!\n{processed} files processed.\n{manual} need manual review.\n\nFolder opened for you to work with the files.")
            else:
                messagebox.showwarning("No Files", "No files were processed.")

        except Exception as e:
            self.log(f"\n✗ Error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            self.start_button.config(state='normal', text="Start Task_1")

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================


if __name__ == "__main__":
    # Check if running with GUI
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "--nogui":
        # Command line mode
        print("Starting Outlook PDF processor...")
        processed, manual = run_pdf_processor()
        print("Done!")
    else:
        # GUI mode (default)
        root = tk.Tk()
        app = OutlookProcessorGUI(root)
        root.mainloop()
