import random
import tkinter as tk
from tkinter import filedialog, messagebox
import smtplib
import mimetypes
from email.message import EmailMessage
import pandas as pd
from bs4 import BeautifulSoup
import re
import os
import time


# GUI Setup
def setup_gui():
    root = tk.Tk()
    root.title("ProfMailer Application")

    def select_file():
        file_path.set(filedialog.askopenfilename())
        if not file_path.get().endswith('.csv'):
            messagebox.showerror("Error", "Please select a CSV file (.csv).")
            file_path.set("")

    def select_folder():
        folder_path.set(filedialog.askdirectory())

    def send_emails():
        if not (file_path.get() and folder_path.get() and email.get() and password.get()):
            messagebox.showerror("Error", "Please fill all fields and select paths.")
            return

        try:
            df = pd.read_csv(file_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return

        for index, row in df.iterrows():
            if row['send_status'] != 0:
                continue

            send_status = send_email(row, email.get(), password.get(), folder_path.get(), resume.get())
            if send_status:
                df.loc[index, 'send_status'] = 1
                df.to_csv(file_path.get(), index=False)
                n = random.randint(180, 300)
                time.sleep(300)  # Wait a random time between emails

        messagebox.showinfo("Success", "Emails sent successfully.")

    tk.Label(root, text="Select CSV File:").pack(pady=5)
    file_path = tk.StringVar()
    tk.Entry(root, textvariable=file_path, width=50).pack(pady=5)
    tk.Button(root, text="Browse", command=select_file).pack(pady=5)

    tk.Label(root, text="Select Folder with Templates & Resumes:").pack(pady=5)
    folder_path = tk.StringVar()
    tk.Entry(root, textvariable=folder_path, width=50).pack(pady=5)
    tk.Button(root, text="Browse", command=select_folder).pack(pady=5)

    tk.Label(root, text="Your Email:").pack(pady=5)
    email = tk.StringVar()
    tk.Entry(root, textvariable=email, width=50).pack(pady=5)

    tk.Label(root, text="App Password:").pack(pady=5)
    password = tk.StringVar()
    tk.Entry(root, textvariable=password, show='*', width=50).pack(pady=5)

    tk.Label(root, text="Resume file name Prefix:").pack(pady=5)
    resume = tk.StringVar()
    tk.Entry(root, textvariable=resume, width=50).pack(pady=5)

    tk.Button(root, text="Send Emails", command=send_emails).pack(pady=20)

    root.mainloop()


def send_email(row, sender, password, folder, resume):
    try:
        message = EmailMessage()
        message['From'] = sender
        message['To'] = row['email']
        message['Subject'] = row['subject'] if row['subject'] != "none" else "Prospective Graduate Student"

        template_file = os.path.join(folder, f"template_{row['field']}.html")
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
                target = soup.find_all(text=re.compile(re.escape(placeholder)))
                for v in target:
                    v.replace_with(v.replace(placeholder, value))

            message.set_content(str(soup), subtype='html')

        resume_file = os.path.join(folder, f"{resume}_{row['field']}.pdf")
        attach_file(message, resume_file)

        if row['transcript']:
            transcript_file = os.path.join(folder, 'transcript.pdf')
            attach_file(message, transcript_file)

        mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
        mail_server.login(sender, password)
        mail_server.send_message(message)
        mail_server.quit()
        return True
    except Exception as e:
        print(f"Failed to send email to {row['email']}: {e}")
        return False


def attach_file(message, file_path):
    with open(file_path, 'rb') as file:
        mime_type, _ = mimetypes.guess_type(file_path)
        mime_type, mime_subtype = mime_type.split('/')
        message.add_attachment(file.read(),
                               maintype=mime_type,
                               subtype=mime_subtype,
                               filename=os.path.basename(file_path))


if __name__ == "__main__":
    setup_gui()
