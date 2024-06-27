import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from bs4 import BeautifulSoup
import re
import os
import threading
import time
import queue
import random


class ProfMailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ProfMailer Application")
        self.file_path = tk.StringVar()
        self.folder_path = tk.StringVar()
        self.email = tk.StringVar()
        self.password = tk.StringVar()
        self.resume = tk.StringVar()
        self.progress_queue = queue.Queue()
        self.setup_gui()

    def setup_gui(self):
        """
        setting GUI
        :return:
        """
        tk.Label(self.root, text="Select CSV File:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.file_path, width=50).pack(pady=5)
        tk.Button(self.root, text="Browse", command=self.select_file).pack(pady=5)

        tk.Label(self.root, text="Select Folder with Templates & Resumes:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.folder_path, width=50).pack(pady=5)
        tk.Button(self.root, text="Browse", command=self.select_folder).pack(pady=5)

        tk.Label(self.root, text="Your Email:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.email, width=50).pack(pady=5)

        tk.Label(self.root, text="App Password:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.password, show='*', width=50).pack(pady=5)

        tk.Label(self.root, text="Resume file name Prefix:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.resume, width=50).pack(pady=5)

        tk.Button(self.root, text="Send Emails", command=self.start_sending_emails).pack(pady=20)

        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(pady=10)

        self.status_label = tk.Label(self.root, text="", relief="sunken", anchor="w")
        self.status_label.pack(fill=tk.X, padx=5, pady=5)

        self.check_queue()

    def select_file(self):
        # loading the list as CSV file
        self.file_path.set(filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")]))
        if not self.file_path.get().endswith('.csv'):
            messagebox.showerror("Error", "Please select a CSV file (.csv).")
            self.file_path.set("")

    def select_folder(self):
        self.folder_path.set(filedialog.askdirectory())

    def start_sending_emails(self):
        if not (self.file_path.get() and self.folder_path.get() and self.email.get() and self.password.get()):
            messagebox.showerror("Error", "Please fill all fields and select paths.")
            return

        threading.Thread(target=self.send_emails).start()

    def send_emails(self):
        try:
            # loading the CSV file
            df = pd.read_csv(self.file_path.get())
        except Exception as e:
            self.progress_queue.put(("error", f"Failed to read CSV file: {e}"))
            return

        total_emails = len(df)
        sent_count = 0

        for index, row in df.iterrows():
            # skip if we have already sent an email
            if row['send_status'] != 0:
                continue

            send_status = self.send_email(row)
            if send_status:
                df.loc[index, 'send_status'] = 1
                sent_count += 1
                self.progress_queue.put(("update", (sent_count, total_emails)))
                df.to_csv(self.file_path.get(), index=False)
                name = df.loc[index, 'name']
                sleep_time = random.randint(100, 300)
                print(f"Sent email to : {name} _ waiting for {sleep_time} seconds to respect limit rate")
                time.sleep(sleep_time)  # Wait a random time between emails

        self.progress_queue.put(("complete", "Emails sent successfully."))

    def send_email(self, row):
        try:
            message = EmailMessage()
            message['From'] = self.email.get()
            message['To'] = row['email']
            # setting the subject
            message['Subject'] = row['subject'] if row['subject'] != "none" else "Prospective Graduate Student"

            # loading template file : template_fieldName
            template_file = os.path.join(self.folder_path.get(), f"template_{row['field']}.html")
            with open(template_file, "r", encoding='utf-8') as f:
                html = f.read()
                soup = BeautifulSoup(html, "html.parser")

                # Replace placeholders
                placeholders = {
                    '{name}': row['name'],
                    '{university}': row['university'],
                    '{group}': '' if row['group'] == "none" else ' at ' + row['group']
                }

                for placeholder, value in placeholders.items():
                    target = soup.find_all(string=re.compile(re.escape(placeholder)))
                    for v in target:
                        v.replace_with(v.replace(placeholder, value))

                message.set_content(str(soup), subtype='html')

            # load specific resume
            resume_file = os.path.join(self.folder_path.get(), f"{self.resume.get()}_{row['field']}.pdf")
            self.attach_file(message, resume_file)

            # load transcript if needed
            if row['transcript']:
                transcript_file = os.path.join(self.folder_path.get(), 'transcript.pdf')
                self.attach_file(message, transcript_file)

            mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
            mail_server.login(self.email.get(), self.password.get())
            mail_server.send_message(message)
            mail_server.quit()
            return True
        except Exception as e:
            self.progress_queue.put(("error", f"Failed to send email to {row['email']}: {e}"))
            return False

    def attach_file(self, message, file_path):
        with open(file_path, 'rb') as file:
            mime_type, _ = mimetypes.guess_type(file_path)
            mime_type, mime_subtype = mime_type.split('/')
            message.add_attachment(file.read(),
                                   maintype=mime_type,
                                   subtype=mime_subtype,
                                   filename=os.path.basename(file_path))

    def check_queue(self):
        try:
            while True:
                message_type, message = self.progress_queue.get_nowait()
                if message_type == "update":
                    sent, total = message
                    self.progress['value'] = (sent / total) * 100
                    self.status_label.config(text=f"Sent {sent}/{total} emails")
                elif message_type == "complete":
                    messagebox.showinfo("Success", message)
                    self.progress['value'] = 0
                    self.status_label.config(text="")
                elif message_type == "error":
                    messagebox.showerror("Error", message)
                    self.status_label.config(text=message)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_queue)

if __name__ == "__main__":
    root = tk.Tk()
    app = ProfMailerApp(root)
    root.mainloop()
