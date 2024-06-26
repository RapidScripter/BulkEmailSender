import tkinter as tk
from ttkbootstrap import Style
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.style = Style(theme="darkly")
        self.attachments = []
        self.excel_file = None
        self.master.title("Bulk Email Sender")
        self.master.configure(bg="#282C34")
        self.create_widgets()

    def create_widgets(self):
        # Configure grid layout
        self.grid(row=0, column=0, sticky="nsew")
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

        # Application Heading
        header_font = ("TkHeadingFont", 20, "bold")
        self.header_label = tk.Label(self, text="Bulk Email Sender", font=header_font, bg="#282C34", fg="white")
        self.header_label.grid(row=0, column=0, columnspan=2, pady=20)

        # Sender Email
        self.sender_email_label = tk.Label(
            self, text="Sender Email:", font=("Helvetica", 14, "bold"), bg="#282C34", fg="white")
        self.sender_email_label.grid(row=1, column=0, sticky="w", padx=10, pady=10)

        self.sender_email_entry = tk.Entry(
            self, width=50, font=("Helvetica", 10), bg="#282C34", fg="white")
        self.sender_email_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

        # Sender Password
        self.sender_password_label = tk.Label(
            self, text="Sender Password:", font=("Helvetica", 14, "bold"), bg="#282C34", fg="white")
        self.sender_password_label.grid(row=2, column=0, sticky="w", padx=10, pady=10)

        self.sender_password_entry = tk.Entry(
            self, width=50, font=("Helvetica", 10), bg="#282C34", fg="white", show="*")
        self.sender_password_entry.grid(row=2, column=1, sticky="ew", padx=10, pady=10)

        # Email Subject
        self.subject_label = tk.Label(
            self, text="Email Subject:", font=("Helvetica", 14, "bold"), bg="#282C34", fg="white")
        self.subject_label.grid(row=3, column=0, sticky="w", padx=10, pady=10)

        self.subject_entry = tk.Entry(
            self, width=50, font=("Helvetica", 10), bg="#282C34", fg="white")
        self.subject_entry.grid(row=3, column=1, sticky="ew", padx=10, pady=10)

        # Email Body
        self.body_label = tk.Label(
            self, text="Email Body:", font=("Helvetica", 14, "bold"), bg="#282C34", fg="white")
        self.body_label.grid(row=4, column=0, sticky="w", padx=10, pady=10)

        self.body_entry = tk.Text(
            self, height=10, font=("Helvetica", 10), bg="#282C34", fg="white")
        self.body_entry.grid(row=4, column=1, sticky="nsew", padx=10, pady=10)

        # Select Excel File
        self.excel_button = tk.Button(
            self, text="Select Excel File", command=self.select_excel_file, font=("Helvetica", 10, "bold"))
        self.excel_button.grid(row=5, column=0, sticky="w", padx=10, pady=10)

        self.excel_label = tk.Label(
            self, text="No file selected", font=("Helvetica", 8), bg="#282C34", fg="white")
        self.excel_label.grid(row=5, column=1, sticky="w", padx=10, pady=10)

        # Select Attachments
        self.select_button = tk.Button(
            self, text="Select Attachments", command=self.select_attachments, font=("Helvetica", 10, "bold"))
        self.select_button.grid(row=6, column=0, sticky="w", padx=10, pady=10)

        self.attachments_label = tk.Label(
            self, text="Attachments: ", font=("Helvetica", 8), bg="#282C34", fg="white")
        self.attachments_label.grid(row=6, column=1, sticky="w", padx=10, pady=10)

        # Send Emails
        self.send_button = tk.Button(
            self, text="Send Emails", command=self.send_emails, font=("Helvetica", 11, "bold"), bg="#007BFF", fg="white")
        self.send_button.grid(row=7, column=1, sticky="e", padx=10, pady=10)

        # Configure weight for rows and columns to make them expandable
        for i in range(8):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(initialdir="/", title="Select Excel File",
                                               filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_file = file_path
            self.excel_label.configure(text=os.path.basename(file_path))
        else:
            self.excel_label.configure(text="No file selected")

    def select_attachments(self):
        file_paths = filedialog.askopenfilenames(initialdir="/", title="Select Files",
                                                 filetypes=[("All Files", "*.*")])
        self.attachments = list(file_paths)
        self.attachments_label.configure(text="Attachments: " + ", ".join(os.path.basename(f) for f in self.attachments))

    def send_emails(self):
        # Define email parameters
        email_sender = self.sender_email_entry.get()
        email_password = self.sender_password_entry.get()
        email_subject = self.subject_entry.get()
        email_body = self.body_entry.get("1.0", tk.END).strip()

        if not self.excel_file:
            messagebox.showerror("Error", "Please select an Excel file with email addresses.")
            return

        try:
            # Read the email addresses from the Excel file
            df = pd.read_excel(self.excel_file)
            email_recipients = df['Email'].tolist()

            # Set up the SMTP server
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587
            smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
            smtp_connection.ehlo()
            smtp_connection.starttls()
            smtp_connection.login(email_sender, email_password)

            # Loop through the list of recipients and send the email
            for recipient in email_recipients:
                # Create the email message
                message = MIMEMultipart()
                message['From'] = email_sender
                message['To'] = recipient
                message['Subject'] = email_subject
                message.attach(MIMEText(email_body, 'plain'))

                # Attach the files
                for attachment_file in self.attachments:
                    try:
                        attachment = open(attachment_file, 'rb')
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload((attachment).read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', "attachment; filename= " + os.path.basename(attachment_file))
                        message.attach(part)
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to attach file {os.path.basename(attachment_file)}. Error: {str(e)}")
                        return

                # Send the email
                smtp_connection.sendmail(email_sender, recipient, message.as_string())

            # Close the SMTP connection
            smtp_connection.quit()

            # Reset the UI
            self.sender_email_entry.delete(0, tk.END)
            self.sender_password_entry.delete(0, tk.END)
            self.subject_entry.delete(0, tk.END)
            self.body_entry.delete("1.0", tk.END)
            self.attachments = []
            self.attachments_label.configure(text="Attachments: ")
            self.excel_file = None
            self.excel_label.configure(text="No file selected")

            # Show success message
            messagebox.showinfo("Success", "Emails sent successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails. Error: {str(e)}")

root = tk.Tk()
root.geometry("700x700")  # set initial size of the window
root.wm_state("zoomed")   # maximize the window on Windows, use "fullscreen" on MacOS
app = Application(master=root)
app.mainloop()
