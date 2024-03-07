import imaplib
import email
import os
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from ttkthemes import ThemedStyle
import getpass

class TekOxExtractor:
    def __init__(self, email_address, password, server):
        self.email_address = email_address
        self.password = password
        self.server = server
        self.valid_password = False

    def connect_to_email_server(self):
        try:
            if self.server == "Gmail":
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
            elif self.server == "Office 365":
                mail = imaplib.IMAP4_SSL("outlook.office365.com")
            elif self.server == "Yahoo Mail":
                mail = imaplib.IMAP4_SSL("imap.mail.yahoo.com")
            else:
                raise ValueError("Unsupported email server")

            mail.login(self.email_address, self.password)
            self.valid_password = True
            return mail
        except Exception as e:
            self.valid_password = False
            return None

    def verify_password(self):
        mail = self.connect_to_email_server()
        if mail:
            mail.logout()
            return True
        return False

    def extract_hyperlinks(self, email_content):
        soup = BeautifulSoup(email_content, 'html.parser')
        links = [a['href'] for a in soup.find_all('a', href=True)]
        return links

    def fetch_emails(self, folder='inbox', criteria='ALL'):
        mail = self.connect_to_email_server()
        if not mail:
            return []

        mail.select(folder)
        result, data = mail.search(None, criteria)
        email_links = []

        for num in data[0].split():
            _, msg_data = mail.fetch(num, '(RFC822)')
            _, msg = msg_data[0]

            email_message = email.message_from_bytes(msg)
            for part in email_message.walk():
                if part.get_content_type() == "text/html":
                    email_content = part.get_payload(decode=True).decode('utf-8')
                    email_links.extend(self.extract_hyperlinks(email_content))

        mail.logout()
        return email_links

    def save_links_as_html(self, links, output_folder='TekOx_Output'):
        os.makedirs(output_folder, exist_ok=True)
        output_file = os.path.join(output_folder, 'output.html')
        with open(output_file, 'w') as file:
            file.write('<html><body>\n')
            for link in links:
                file.write(f'<a href="{link}">{link}</a><br>\n')
            file.write('</body></html>')

class TekOxGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("TekOx Email Link Extractor")

        style = ThemedStyle(root)
        style.set_theme("arc")  # You can choose a theme you prefer

        # Use a larger font size for readability
        style.configure("TLabel", font=('Helvetica', 14))
        style.configure("TEntry", font=('Helvetica', 14))
        style.configure("TButton", font=('Helvetica', 14))

        self.header_label = ttk.Label(root, text="TekOx Email Link Extractor", font=('Helvetica', 18, 'bold'))
        self.header_label.grid(row=0, column=0, columnspan=3, pady=15)

        self.input_frame = ttk.Frame(root)
        self.input_frame.grid(row=1, column=0, columnspan=3, pady=15)

        self.email_label = ttk.Label(self.input_frame, text="Email Address:")
        self.email_entry = ttk.Entry(self.input_frame, width=40)

        self.password_label = ttk.Label(self.input_frame, text="Password:")
        self.password_entry = ttk.Entry(self.input_frame, show="*", width=40)

        self.server_label = ttk.Label(self.input_frame, text="Email Server:")
        self.server_combobox = ttk.Combobox(self.input_frame, values=["Gmail", "Office 365", "Yahoo Mail"], width=36)
        self.server_combobox.set("Gmail")

        self.email_label.grid(row=0, column=0, sticky=tk.W, pady=5, padx=10)
        self.email_entry.grid(row=0, column=1, pady=5, padx=10)
        self.password_label.grid(row=1, column=0, sticky=tk.W, pady=5, padx=10)
        self.password_entry.grid(row=1, column=1, pady=5, padx=10)
        self.server_label.grid(row=2, column=0, sticky=tk.W, pady=5, padx=10)
        self.server_combobox.grid(row=2, column=1, pady=5, padx=10)

        self.extract_button = ttk.Button(root, text="Extract Links", command=self.extract_links, style='TButton')
        self.extract_button.grid(row=2, column=2, pady=15)

        self.result_frame = ttk.Frame(root)
        self.result_frame.grid(row=3, column=0, columnspan=3, pady=15)

        self.result_label = ttk.Label(self.result_frame, text="", font=('Helvetica', 14))
        self.result_label.grid(row=0, column=0)

        self.link_counter_label = ttk.Label(self.result_frame, text="", font=('Helvetica', 14))
        self.link_counter_label.grid(row=1, column=0)

    def extract_links(self):
        email_address = self.email_entry.get()
        password = self.password_entry.get()
        server = self.server_combobox.get()

        if not email_address or not password:
            self.result_label.config(text="Please provide both email address and password.", foreground='#cb2431')
            return

        self.result_label.config(text="Verifying password...", foreground='#24292e')
        self.root.after(10, lambda: self.verify_and_extract_links(email_address, password, server))

    def verify_and_extract_links(self, email_address, password, server):
        tekox = TekOxExtractor(email_address, password, server)

        if not tekox.verify_password():
            self.result_label.config(text="Incorrect password. Please enter the correct password.", foreground='#cb2431')
            return

        self.result_label.config(text="Fetching emails...", foreground='#24292e')
        email_links = tekox.fetch_emails()

        if email_links:
            tekox.save_links_as_html(email_links)
            num_links = len(email_links)
            self.result_label.config(text=f"Hyperlinks extracted and saved as HTML file.", foreground='#2cbe4e')
            self.link_counter_label.config(text=f"Number of hyperlinks found: {num_links}", foreground='#24292e')
        else:
            self.result_label.config(text="No hyperlinks found in the emails.", foreground='#cb2431')
            self.link_counter_label.config(text="", foreground='#24292e')

if __name__ == "__main__":
    root = tk.Tk()
    gui = TekOxGUI(root)
    root.mainloop()
