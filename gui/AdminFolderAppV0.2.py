import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
from eml_processor import eml_get_date, eml_to_pdf, eml_extract_attachments
from msg_processor import msg_get_date, msg_to_pdf, msg_extract_attachments
from pathlib import Path

class EmailProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Processor")
        self.root.geometry("700x400")

        self.input_files = []
        self.output_dir = ""
        self.start_number = 1

        self.setup_ui()

    def setup_ui(self):
        self.input_label = tk.Label(self.root, text="Step 1: Select Email Files (.msg or .eml)")
        self.input_label.pack(pady=5)
        self.select_button = tk.Button(self.root, text="Select Files", command=self.select_files)
        self.select_button.pack(pady=5)

        self.output_label = tk.Label(self.root, text="Step 2: Select Output Folder")
        self.output_label.pack(pady=5)
        self.output_button = tk.Button(self.root, text="Select Output Folder", command=self.select_output_folder)
        self.output_button.pack(pady=5)

        self.start_number_label = tk.Label(self.root, text="Step 3: Enter Start Number")
        self.start_number_label.pack(pady=5)
        self.start_number_entry = tk.Entry(self.root)
        self.start_number_entry.insert(0, "1")
        self.start_number_entry.pack(pady=5)

        self.start_button = tk.Button(self.root, text="Start Processing", command=self.start_processing)
        self.start_button.pack(pady=10)

        self.progress_label = tk.Label(self.root, text="Overall Progress")
        self.progress_label.pack()
        self.overall_progress = ttk.Progressbar(self.root, length=600, mode='determinate')
        self.overall_progress.pack(pady=5)

        # self.file_progress_label = tk.Label(self.root, text="Current File Progress")
        # self.file_progress_label.pack()
        # self.file_progress = ttk.Progressbar(self.root, length=600, mode='determinate')
        # self.file_progress.pack(pady=5)

        self.log_text = tk.Text(self.root, height=10, state='disabled')
        self.log_text.pack(pady=10, fill='both', expand=True)

    def select_files(self):
        self.input_files = filedialog.askopenfilenames(filetypes=[("Email files", "*.msg *.eml")])
        self.log_message(f"Selected {len(self.input_files)} files")

    def select_output_folder(self):
        self.output_dir = filedialog.askdirectory()
        self.log_message(f"Selected output folder: {self.output_dir}")

    def start_processing(self):
        try:
            self.start_number = int(self.start_number_entry.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Start number must be an integer.")
            return

        if not self.input_files:
            messagebox.showerror("No Input", "Please select email files to process.")
            return

        if not self.output_dir:
            messagebox.showerror("No Output Folder", "Please select an output folder.")
            return

        threading.Thread(target=self.process_emails).start()

    def process_emails(self):
        self.log_message("Beginning processing...")
        
        total = len(self.input_files)
        self.overall_progress["maximum"] = total

        email_data = []

        # Step 1: Sort emails by date
        for file in self.input_files:
            try:
                if file.lower().endswith(".eml"):
                    date = eml_get_date(file)
                elif file.lower().endswith(".msg"):
                    date = msg_get_date(file)
                else:
                    self.log_message(f"Unsupported file type: {file}")
                    continue

                email_data.append({"path": file, "date": date})
            except Exception as e:
                self.log_message(f"Failed to get date for {file}: {e}")

        email_data.sort(key=lambda x: x["date"])  # sort chronologically

        # Step 2: Process each file
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)

        wkPath = os.path.join(base_path, "bin", "wkhtmltopdf.exe")

        start_num = int(self.start_number_entry.get())

        for idx, entry in enumerate(email_data, start=start_num):
            filepath = entry["path"]
            try:
                if filepath.lower().endswith(".eml"):
                    eml_to_pdf(filepath, self.output_dir, idx, wkPath, self.log_message)
                    eml_extract_attachments(filepath, self.output_dir, idx, self.log_message)
                elif filepath.lower().endswith(".msg"):
                    msg_to_pdf(filepath, self.output_dir, idx, wkPath, self.log_message)
                    msg_extract_attachments(filepath, self.output_dir, idx, self.log_message)
                else:
                    self.log_message(f"Skipped unsupported file: {filepath}")
                    continue

                self.log_message(f"Processed: {Path(filepath).name}")
            except Exception as e:
                self.log_message(f"Error processing {filepath}: {e}")

            self.overall_progress["value"] = idx - start_num + 1
            self.root.update_idletasks()

        self.log_message("All files processed.")

    def log_message(self, message):
        self.log_text["state"] = 'normal'
        self.log_text.insert('end', message + "\n")
        self.log_text.see('end')
        self.log_text["state"] = 'disabled'

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailProcessorApp(root)
    root.mainloop()
