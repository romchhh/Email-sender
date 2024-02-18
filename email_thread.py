from typing import List
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from itertools import cycle
import logging
import smtplib
from datetime import datetime
from PyQt5.QtCore import QThread, pyqtSignal
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from asyncio import sleep
from accounts_serializer import Account, get_accounts


class Signal():
    def __init__(self) -> None:
        self.account_request_counts = None
        self.sent_label = None
        self.log = None
        self.last_log = False

class EmailSenderThread(QThread):

    update_log = pyqtSignal(Signal)
    def __init__(self, receiver_emails, subject, message, accounts: List[Account], attachments, parent=None):
        super(EmailSenderThread, self).__init__(parent)
        self.signal = Signal()
        self.is_running = True

        self.attachments = attachments
        self.accounts = accounts
        self.reciever_emails = receiver_emails
        self.subject = subject
        self.message = message
        
    def run(self):
        try:
            receiver_emails = self.reciever_emails
            subject = self.subject
            message = self.message

            start_time = datetime.now()
            self.account_request_counts = defaultdict(int)

            with ThreadPoolExecutor(max_workers=len(self.accounts)) as executor:
                account_cycle = cycle(self.accounts)
                index_cycle = cycle(range(0, len(self.accounts)))
                
                


                futures = []

                for receiver_email in receiver_emails:
                    account = next(account_cycle)
                    account_index = next(index_cycle)
                    print(account_index)
                    print
                                
                    if self.account_request_counts[account] < 500:
                        future = executor.submit(
                            self.send_email,
                            receiver_email,
                            subject,
                            message,
                            self.attachments,
                            account, 
                            account_index               
                        )
                        futures.append(future)  

            for future in as_completed(futures):
                    future.result()
                    
            end_time = datetime.now()
            time_taken = end_time - start_time


            self.signal.log = f"Time taken: {time_taken}\nTotal sent {sum(self.account_request_counts.values())} mails"
            self.signal.sent_label = f'Надіслані листи: {sum(self.account_request_counts.values())}'
            self.signal.account_request_counts = self.account_request_counts
            self.update_log.emit(self.signal)
                    
        except Exception as e:
            log_message = f"Error sending email to {receiver_email}: {e}"
            self.signal.log = log_message
            self.update_log.emit(self.signal)
    

    

    def send_email(self, receiver_email, subject, message, attachments, account: Account, account_index):
        try:
            if self.accounts[account_index].banned:
                return
            
            if '@hotmail' in account.mail:
                server = smtplib.SMTP("smtp-mail.outlook.com", 587)
            if '@gmail' in account.mail:
                server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(account.mail, account.key)
            

            msg = MIMEMultipart()
            msg['From'] = account.mail
            msg['To'] = receiver_email
            msg['Subject'] = subject


            for attachment in attachments:
                if attachment.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', 'txt', )):
                    with open(attachment, 'rb') as file:
                        img = MIMEImage(file.read(), name=f"image_{attachments.index(attachment)}.png")
                        msg.attach(img)
                elif attachment.lower().endswith(('.html')) :
                    with open(attachment, 'rb') as file:
                        html_content = file.read().decode('utf-8')
                        msg.attach(MIMEText(html_content, 'html'))
                else:
                    with open(attachment, 'rb') as file:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(file.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f"attachment; filename={attachment}")
                        msg.attach(part)

            text = msg.as_string()
            

            server.sendmail(account.mail, receiver_email, text)   
            self.account_request_counts[account] += 1        
            
     
            logging_message = f"Email sent from {account.mail} to: {receiver_email}, Subject: {subject}, Message: {message}, Timestamp: {datetime.now()}"
            logging.info(logging_message)

            self.signal.sent_label = f'Надіслані листи: {sum(self.account_request_counts.values())}'
            self.signal.log = logging_message
            self.update_log.emit(self.signal)

            server.quit()

        except Exception as e:
            self.accounts[account_index].banned = True
            print(f'{self.accounts[account_index].mail} has been banned')
            print(e)
            log_message = f"Error sending email to {receiver_email}: {e}"
            self.signal.log = log_message
            self.update_log.emit(self.signal)
        
    
