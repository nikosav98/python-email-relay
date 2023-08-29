import imaplib
import email
from email.header import decode_header
import os
from time import sleep
import subprocess
import platform
import credentials
import constants

# use your email provider's IMAP server, you can look for your provider's IMAP server on Google
# or check this page: https://www.systoolsgroup.com/imap/
imap_server = "imap-mail.outlook.com"
# while flag is set to True the program will run
flag = True

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

imap = imaplib.IMAP4_SSL(imap_server)
# authenticate
try:
    imap.login(credentials.username, credentials.password)
    status, messages = imap.select("INBOX")
    print("Authentication successful")
except Exception as e:
    print("Authentication failed")
    flag = False

def process_mail(messages):
    for i in range(messages, messages-constants.N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject:", subject)
                print("From:", From)
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            print(body)
                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            if filename:
                                # Check if the attachment is a PDF or PNG file
                                if filename.lower().endswith((".pdf", ".png")):
                                    folder_name = clean(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(part.get_payload(decode=True))

                                    # Print the file if it's a PDF or PNG
                                    if filename.lower().endswith(".pdf") or filename.lower().endswith(".png"):
                                        print_file(filepath)
                                else:
                                    print(f"Skipping attachment {filename} as it's not a PDF or PNG.")
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)
                if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)
                print("="*100)
            return
    # close the connection and logout

def print_file(file_path):
    file_path = os.path.abspath(file_path)
    system_platform = platform.system()

    if system_platform == "Windows":
        import win32print
        import win32api

        printer_name = win32print.GetDefaultPrinter()
        if not printer_name:
            print("No default printer found.")
            return

        try:
            win32api.ShellExecute(
                0, "print", file_path, f'"{printer_name}"', ".", 0
            )
            print("Printing...")
        except Exception as e:
            print(f"Error printing: {e}")

    elif system_platform in ["Linux", "Darwin"]:
        try:
            subprocess.run(["lp", file_path], check=True)
            print("Printing...")
        except subprocess.CalledProcessError as e:
            print(f"Error printing: {e}")
    else:
        print("Unsupported operating system.")

while flag:
    # Select the mailbox and get the number of messages
    status, messages = imap.select("INBOX")
    if status == "OK":
        total_messages = int(messages[0])

        if total_messages > 0:
            process_mail(total_messages)  # Process the most recent email
            for i in range(total_messages, total_messages - constants.N, -1):
                # Mark email as deleted
                imap.store(str(i), '+FLAGS', '\\Deleted')
            imap.expunge()  # Permanently remove deleted emails

            total_messages = total_messages - constants.N  # Update the total count
            print("Remaining unread emails:", total_messages)
        else:
            print("No new emails.")
    else:
        print("Failed to select the mailbox.")

    sleep(constants.sleep_for)
    
imap.logout()
if flag:
    imap.close()
