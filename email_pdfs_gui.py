import os
import re
import fitz  # PyMuPDF
import smtplib
import time
import random
from email.message import EmailMessage
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import json

# === CONFIGURATION === #
DEFAULT_FOLDER = "test_pdfs"
DELAY_RANGE = (2, 4)  # Seconds between emails


SETTINGS_FILE = "settings.json"

DEFAULTS = {
    "folder": "pdfs",
    "sender_email": "youremail@gmail.com",
    "sender_password": "your_app_password",
    "subject": "subject",
    "body": "body"
}

# === LOGIC FUNCTIONS === #

def extract_recipient_email(pdf_path, ignored_emails):
    try:
        doc = fitz.open(pdf_path)
        text = doc[0].get_text()
        emails = re.findall(r'\b[\w.-]+@[\w.-]+\.\w+\b', text)
        emails = [email for email in emails if email not in ignored_emails]
        return emails[0] if len(emails) == 1 else None
    except Exception:
        return None


def get_sorted_pdf_files(folder_path):
    return sorted(
        [file for file in os.listdir(folder_path) if file.lower().endswith(".pdf")],
        key=str.lower
    )


def send_email(receiver_email, pdf_path, sender_email, sender_password, subject, body):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.set_content(body)

    with open(pdf_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        os.remove(pdf_path)  # DELETE FILE AFTER SENDING
        return True
    except Exception:
        return False


# === GUI APPLICATION === #

class PDFMailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📧 Gmail PDF Mailer")
        # Load saved defaults
        self.settings = self.load_settings()

        # PDF Folder
        tk.Label(root, text="📂 PDF Folder:").grid(row=0, column=0, sticky="e")
        self.folder_entry = tk.Entry(root, width=40)
        default_folder = DEFAULT_FOLDER if os.path.exists(DEFAULT_FOLDER) else ""
        self.folder_entry.insert(0, self.settings.get("folder", ""))
        self.folder_entry.grid(row=0, column=1)
        tk.Button(root, text="Browse", command=self.select_folder).grid(row=0, column=2)

        # Gmail + App Password
        tk.Label(root, text="📧 Gmail:").grid(row=1, column=0, sticky="e")
        self.email_entry = tk.Entry(root, width=40)
        self.email_entry.insert(0, self.settings.get("sender_email", ""))
        self.email_entry.grid(row=1, column=1)

        tk.Label(root, text="🔑 App Password:").grid(row=2, column=0, sticky="e")
        self.pass_entry = tk.Entry(root, width=40)
        self.pass_entry.insert(0, self.settings.get("sender_password", ""))
        self.pass_entry.grid(row=2, column=1)

        # Subject
        tk.Label(root, text="📝 Subject:").grid(row=3, column=0, sticky="ne")
        self.subject_entry = tk.Entry(root, width=40)
        self.subject_entry.insert(0, self.settings.get("subject", ""))
        self.subject_entry.grid(row=3, column=1)

        # Body
        tk.Label(root, text="📄 Body:").grid(row=4, column=0, sticky="ne")
        self.body_text = scrolledtext.ScrolledText(root, width=50, height=10)
        self.body_text.insert("1.0", self.settings.get("body", ""))
        self.body_text.grid(row=4, column=1, columnspan=2, sticky="nsew")

        # Send Button
        tk.Button(root, text="▶️ Send Emails", command=self.start_sending).grid(row=5, column=1, pady=10)
        tk.Button(root, text="💾 Save as Defaults", command=self.save_settings).grid(row=5, column=2, pady=10)

        # Log Box
        self.log_box = scrolledtext.ScrolledText(root, width=70, height=15, state='disabled')
        self.log_box.grid(row=6, column=0, columnspan=3, pady=10, sticky="nsew")

        # Make the layout resizable
        for i in range(7):
            self.root.grid_rowconfigure(i, weight=1)
        for j in range(3):
            self.root.grid_columnconfigure(j, weight=1)
    def log(self, message):
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)

    def start_sending(self):
        folder = self.folder_entry.get().strip()
        sender_email = self.email_entry.get().strip()
        sender_password = self.pass_entry.get().strip()
        subject = self.subject_entry.get().strip()
        body = self.body_text.get("1.0", tk.END).strip()

        if not all([folder, sender_email, sender_password, subject, body]):
            messagebox.showwarning("Missing Info", "Please fill in all fields.")
            return

        # ⚠️ Warn about deletion
        proceed = messagebox.askyesno(
            "Warning",
            "⚠️ This will DELETE all PDFs after sending.\n\nDo you want to continue?"
        )
        if not proceed:
            return

        pdf_files = get_sorted_pdf_files(folder)
        ignored_emails = {}  # Add known sender emails here
        email_jobs = []

        for file in pdf_files:
            path = os.path.join(folder, file)
            recipient = extract_recipient_email(path, ignored_emails)
            if recipient:
                email_jobs.append((recipient, path))
            else:
                self.log(f"⚠️ Skipped: couldn't extract valid email from {file}")

        for i, (email, path) in enumerate(email_jobs, 1):
            self.log(f"📨 Sending to {email} ({os.path.basename(path)})")
            success = send_email(email, path, sender_email, sender_password, subject, body)
            if success:
                self.log("✅ Sent successfully.\n")
            else:
                self.log("❌ Failed to send.\n")

            delay = random.uniform(*DELAY_RANGE)
            self.log(f"🕒 Waiting {delay:.1f}s before next email...\n")
            self.root.update()
            time.sleep(delay)

        self.log("✅ All emails processed.")

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    return json.load(f)
            except Exception:
                return DEFAULTS.copy()
        return DEFAULTS.copy()

    def save_settings(self):
        settings = {
            "folder": self.folder_entry.get().strip(),
            "sender_email": self.email_entry.get().strip(),
            "sender_password": self.pass_entry.get().strip(),
            "subject": self.subject_entry.get().strip(),
            "body": self.body_text.get("1.0", tk.END).strip()
        }
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(settings, f, indent=2)
            messagebox.showinfo("Saved", "✅ Default values saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"❌ Failed to save settings: {e}")

# === RUN APP === #
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("800x700")  # Larger window to start
    root.resizable(True, True)
    app = PDFMailerApp(root)
    root.mainloop()