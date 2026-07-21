import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import ctypes
from pathlib import Path
import threading
import sys

# Import tasks
from task_1 import run_task_1
from task_2 import run_task_2
from task_3 import run_task_3


class OutlookProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook PDF Processor v3.0")
        self.root.geometry("700x520")

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

        # Buttons Frame
        buttons_frame = tk.Frame(root)
        buttons_frame.pack(pady=10)

        # Task 1 Button
        self.task1_button = tk.Button(
            buttons_frame,
            text="Start Task_1",
            command=self.start_task_1,
            font=("Arial", 12, "bold"),
            bg="#5DADE2",  # Light blue
            fg="white",
            width=15,
            height=1,
            relief="raised",
            borderwidth=3,
            cursor="hand2"
        )
        self.task1_button.grid(row=0, column=0, padx=10)

        # Task 2 Button
        self.task2_button = tk.Button(
            buttons_frame,
            text="Start Task_2",
            command=self.start_task_2,
            font=("Arial", 12, "bold"),
            bg="#5DADE2",
            fg="white",
            width=15,
            height=1,
            relief="raised",
            borderwidth=3,
            cursor="hand2"
        )
        self.task2_button.grid(row=0, column=1, padx=10)

        # Task 3 Button
        self.task3_button = tk.Button(
            buttons_frame,
            text="Start Task_3",
            command=self.start_task_3,
            font=("Arial", 12, "bold"),
            bg="#5DADE2",
            fg="white",
            width=15,
            height=1,
            relief="raised",
            borderwidth=3,
            cursor="hand2"
        )
        self.task3_button.grid(row=0, column=2, padx=10)

        # Log Output Area
        log_label = tk.Label(root, text="Log Output:",
                             font=("Arial", 10, "bold"))
        log_label.pack(pady=(20, 5))

        self.log_text = scrolledtext.ScrolledText(
            root, width=75, height=18, state='disabled')
        self.log_text.pack(pady=10)

    def log(self, message):
        """Add message to log output"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def clear_log(self):
        """Clear the log output"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def disable_buttons(self):
        """Disable all buttons during processing"""
        self.task1_button.config(state='disabled')
        self.task2_button.config(state='disabled')
        self.task3_button.config(state='disabled')

    def enable_buttons(self):
        """Enable all buttons after processing"""
        self.task1_button.config(state='normal')
        self.task2_button.config(state='normal')
        self.task3_button.config(state='normal')

    def start_task_1(self):
        """Start Task 1 in a separate thread"""
        self.disable_buttons()
        self.task1_button.config(text="Processing...")
        self.clear_log()

        thread = threading.Thread(target=self.run_task_1)
        thread.start()

    def run_task_1(self):
        """Run Task 1 processor"""
        try:
            processed, manual, folder_path, variofix_files = run_task_1(self.log)

            if processed > 0 or manual > 0:
                # Open the task folder
                if folder_path:
                    os.startfile(folder_path)

                message = (f"Processing complete!\n{processed} files processed."
                           f"\n{manual} need manual review.")
                if variofix_files:
                    message += (f"\n\n⚠️ ATTENTION VARIOFIX DETECTED "
                                f"in {len(variofix_files)} file(s):")
                    for filename in variofix_files:
                        message += f"\n  - {filename}"
                message += "\n\nFolder opened."

                messagebox.showinfo("Task 1 Complete", message)
            else:
                messagebox.showwarning("Task 1", "No files were processed.")

        except Exception as e:
            self.log(f"\n✗ Error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            self.enable_buttons()
            self.task1_button.config(text="Start Task_1")

    def start_task_2(self):
        """Start Task 2 in a separate thread"""
        self.disable_buttons()
        self.task2_button.config(text="Processing...")
        self.clear_log()

        thread = threading.Thread(target=self.run_task_2)
        thread.start()

    def run_task_2(self):
        """Run Task 2 processor"""
        try:
            (processed, manual, folder_path,
             files_with_warnings, variofix_files) = run_task_2(self.log)

            if processed > 0 or manual > 0:
                # Open the task folder
                if folder_path:
                    os.startfile(folder_path)

                # Build message with warning if present
                message = f"Processing complete!\n{processed} files processed.\n{manual} need manual review."
                if len(files_with_warnings) > 0:
                    message += f"\n\n⚠️ {len(files_with_warnings)} file(s) had Abmessung sections with no dimensions:"
                    for filename in files_with_warnings:
                        message += f"\n  - {filename}"
                if variofix_files:
                    message += (f"\n\n⚠️ ATTENTION VARIOFIX DETECTED "
                                f"in {len(variofix_files)} file(s):")
                    for filename in variofix_files:
                        message += f"\n  - {filename}"
                message += "\n\nFolder opened."

                messagebox.showinfo("Task 2 Complete", message)
            else:
                messagebox.showwarning("Task 2", "No files were processed.")

        except Exception as e:
            self.log(f"\n✗ Error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            self.enable_buttons()
            self.task2_button.config(text="Start Task_2")

    def start_task_3(self):
        """Start Task 3 in a separate thread"""
        self.disable_buttons()
        self.task3_button.config(text="Processing...")
        self.clear_log()

        thread = threading.Thread(target=self.run_task_3)
        thread.start()

    def run_task_3(self):
        """Run Task 3 processor"""
        try:
            processed, incomplete_combos, missing_field_combos, folder_path = run_task_3(self.log)

            if folder_path is None and processed == 0 and not incomplete_combos and not missing_field_combos:
                # Task_3 folder missing OR no emails - run_task_3 already logged details
                messagebox.showwarning(
                    "Task 3",
                    "No files were processed.\n\n"
                    "Either the Task_3 folder is missing from your Inbox or it contains no emails."
                )
                return

            if folder_path:
                os.startfile(folder_path)

            message = f"Processing complete!\n{processed} combo(s) written to the Excel report."

            if missing_field_combos:
                message += f"\n\n⚠️ {len(missing_field_combos)} combo(s) written with missing fields (see the red rows / cells in the Excel)."

            if incomplete_combos:
                message += f"\n\n⚠️ {len(incomplete_combos)} combo(s) skipped (no packing list):"
                for subj, cid, notes in incomplete_combos[:10]:
                    short_subj = subj if len(subj) <= 40 else subj[:40] + "..."
                    message += f"\n  - {short_subj} / {cid}"
                if len(incomplete_combos) > 10:
                    message += f"\n  ...and {len(incomplete_combos) - 10} more (see log)"

            message += "\n\nFolder opened."
            messagebox.showinfo("Task 3 Complete", message)

        except Exception as e:
            self.log(f"\n✗ Error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            self.enable_buttons()
            self.task3_button.config(text="Start Task_3")


if __name__ == "__main__":
    # Check if running with GUI
    if len(sys.argv) > 1 and sys.argv[1] == "--task1":
        # Command line mode - Task 1
        print("Starting Task 1...")
        run_task_1()
        print("Done!")
    elif len(sys.argv) > 1 and sys.argv[1] == "--task2":
        # Command line mode - Task 2
        print("Starting Task 2...")
        run_task_2()
        print("Done!")
    elif len(sys.argv) > 1 and sys.argv[1] == "--task3":
        # Command line mode - Task 3
        print("Starting Task 3...")
        run_task_3()
        print("Done!")
    else:
        # GUI mode (default)
        root = tk.Tk()
        app = OutlookProcessorGUI(root)
        root.mainloop()
