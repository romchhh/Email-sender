from itertools import cycle
import sys
import time
import pandas as pd
import logging
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QVBoxLayout, QTextEdit, QHBoxLayout, QMessageBox, QDialog, QTextBrowser,
    QSizePolicy, QListWidget
)
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QTextBrowser, QLabel, QTextEdit, QListWidget, QListWidgetItem, QHBoxLayout, QSizePolicy, QPushButton, QFrame
)
from PyQt5.QtGui import QPixmap, QImageReader
from PyQt5.QtCore import QSize
from PyQt5.QtGui import QIcon
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from PyQt5.QtCore import Qt
from info import html_guide
import datetime
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from accounts_serializer import Account, get_accounts
from collections import defaultdict


class UTF8LogFormatter(logging.Formatter):
    def format(self, record):
        result = super().format(record)
        return result.encode('utf-8', 'replace').decode('utf-8')

logging.basicConfig(filename='email_sender.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logging.getLogger().handlers[0].setFormatter(UTF8LogFormatter())

MIN_MESSAGE_HEIGHT = 200
MIN_DIALOG_SIZE = (800, 600)


def send_email(receiver_email, subject, message, attachments, account: Account):
        try:
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
            print(f"Email sent successfully to {receiver_email}")

            logging.info(f"Email sent from {account.mail} to: {receiver_email}, Subject: {subject}, Message: {message}, Timestamp: {datetime.now()}")

            server.quit()

        except Exception as e:
            print(f"Error: {e}")

class EmailSenderWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setGeometry(100, 100, 1350, 900)
        self.setWindowTitle('Програма для розсилки')

        self.setStyleSheet(
            "QWidget { background-color: #f5f5f5; color: #333; font-family: 'Segoe UI', sans-serif; }"
            "QLabel { font-weight: bold; font-size: 16px; color: #333; }"
            "QLineEdit, QTextEdit { background-color: #fff; border: 2px solid #aaa; padding: 8px; font-size: 18px; "
            "border-radius: 5px; color: #333; }"
            "QPushButton { background-color: #3498db; color: white; border: none; padding: 10px 20px; border-radius: 5px; "
            "font-size: 16px; }"
            "QPushButton:hover { background-color: #2980b9; }"
            "QDialog { background-color: #f5f5f5; color: #333; font-family: 'Segoe UI', sans-serif; }"
            "QTextEdit { background-color: #fff; border: 2px solid #aaa; padding: 8px; font-size: 18px; border-radius: 5px; "
            "color: #333; }"
            "QFileDialog { background-color: #fff; }"
        )


        self.download_button = QPushButton('Виберіть файл Excel', self)
        self.download_button.setToolTip('Натисніть, щоб вибрати файл Excel з адресами електронної пошти.')
        self.download_button.clicked.connect(self.choose_excel_file)

        self.accounts_button = QPushButton('Прикріпити аккаунти', self)
        self.accounts_button.setToolTip('Натисніть, щоб пприкріпити файли')
        self.accounts_button.clicked.connect(self.get_accounts_file)

        self.attach_button = QPushButton('Прикріпити файли', self)
        self.attach_button.setToolTip('Натисніть, щоб прикріпити файли.')
        self.attach_button.clicked.connect(self.attach_files)

        self.subject_label = QLabel('Тема листа:', self)
        self.subject_input = QLineEdit(self)

        self.message_label = QLabel('Повідомлення:', self)
        self.message_input = QTextEdit(self)
        self.message_input.setMinimumHeight(200)

        self.review_button = QPushButton('Перегляд листа', self)
        self.review_button.setToolTip('Натисніть, щоб переглянути вміст електронного листа перед надсиланням.')
        self.review_button.clicked.connect(self.review_emails)

        self.send_button = QPushButton('Надіслати листи', self)
        self.send_button.setToolTip('Натисніть, щоб надіслати електронні листи вибраним одержувачам.')
        self.send_button.clicked.connect(self.send_emails)

        self.html_guide_button = QPushButton('?', self)
        self.html_guide_button.setToolTip('Натисніть, щоб переглянути вказівки.')
        self.html_guide_button.clicked.connect(self.show_html_guide)

        self.sent_label = QLabel(self)

        self.status_label = QLabel(self)
        self.status_label.setStyleSheet("color: #d9534f; font-size: 12px;")

        self.attachments = []

        layout = QVBoxLayout()

        header_label = QLabel("Програма для розсилки", self)
        header_label.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 20px; color: #3498db;")
        layout.addWidget(header_label, alignment=Qt.AlignCenter)


        button_layout = QHBoxLayout()
        button_layout.addWidget(self.download_button)
        button_layout.addWidget(self.accounts_button)
        button_layout.addWidget(self.attach_button)
        layout.addLayout(button_layout)

        layout.addWidget(self.subject_label)
        layout.addWidget(self.subject_input)
        layout.addWidget(self.message_label)
        layout.addWidget(self.message_input)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.review_button, alignment=Qt.AlignLeft)
        button_layout.addWidget(self.send_button, alignment=Qt.AlignRight)
        layout.addLayout(button_layout)

        layout.addWidget(self.html_guide_button, alignment=Qt.AlignLeft)
        layout.addWidget(self.status_label, alignment=Qt.AlignLeft)
        layout.addWidget(self.sent_label, alignment=Qt.AlignLeft)

        self.setLayout(layout)

    def choose_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(None, 'Виберіть файл Excel', '', 'Excel Files (*.xlsx)')
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.status_label.setText('Файл Excel успішно завантажено.')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Error loading Excel file: {str(e)}')
                self.status_label.setText('Error loading Excel file.')

    def attach_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(None, 'Прикріпити файли', '', 'All Files (*)')
        if file_paths:
            self.attachments.extend(file_paths)
            self.status_label.setText(f'{len(self.attachments)} file(s) attached.')

    def get_accounts_file(self):
        file_path, _ = QFileDialog.getOpenFileNames(None, 'Прикріпити файли', '', 'All Files (*)')
        if file_path:
            try:
                self.accounts = get_accounts(file_path[0])
                self.status_label.setText(f'Файл аккаунтів успішно завантажено.')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Error loading accounts file: {str(e)}')
                self.status_label.setText('Error loading accounts file.')

    def send_emails(self):
        if hasattr(self, 'df'):
            receiver_emails = self.df.iloc[:, 0].tolist()
            subject = self.subject_input.text()
            message = self.message_input.toPlainText()

            if not (receiver_emails):
                self.status_label.setText('Будь ласка, заповніть усі поля.')
            else:
                try:
                    start_time = datetime.now()

                    account_request_counts = defaultdict(int)

                    with ThreadPoolExecutor(max_workers=len(self.accounts)) as executor:
                        account_cycle = cycle(self.accounts)
                        futures = []

                        for receiver_email in receiver_emails:
                            account = next(account_cycle)
                            
                            if account_request_counts[account] < 500:
                                future = executor.submit(
                                    send_email,
                                    receiver_email,
                                    subject,
                                    message,
                                    self.attachments,
                                    account
                                )
                                futures.append(future)
                                account_request_counts[account] += 1  

                    
                    for future in as_completed(futures):
                        future.result()

                    print(f'Total sended {sum(account_request_counts.values())} mails ')


                    end_time = datetime.now()
                    time_taken = end_time - start_time
                    print(f"Time taken: {time_taken}")

                    self.status_label.setText('Листи успішно надіслано.')
                    self.sent_label.setText(f'Надіслані листи: {len(receiver_emails)}')
                except Exception as e:
                    QMessageBox.critical(self, 'Error', f'Error sending emails: {str(e)}')
                    self.status_label.setText('Помилка надсилання електронних листів.')
        else:
            self.status_label.setText('Спочатку виберіть файл Excel.')

    def review_emails(self):
        subject = self.subject_input.text()
        message = self.message_input.toPlainText()

        if not (subject and message):
            self.status_label.setText('Будь ласка, введіть тему електронного листа та повідомлення.')
        else:
            review_dialog = EmailReviewDialog(subject, message, self.attachments)
            review_dialog.exec_()

    def show_html_guide(self):
        msg_box = ResizableMessageBox(self)
        msg_box.setText(html_guide)
        msg_box.setWindowTitle('?')

        msg_box.exec_()


class ResizableMessageBox(QMessageBox):
    def __init__(self, parent=None):
        super(ResizableMessageBox, self).__init__(parent)

        self.setMinimumSize(800, 600)

    def resizeEvent(self, event):
        result = super(ResizableMessageBox, self).resizeEvent(event)
        if self.width() < 800 or self.height() < 600:
            self.setFixedSize(800, 600)
        return result




class EmailReviewDialog(QDialog):
    def __init__(self, subject, message, attachments):
        super().__init__()

        self.setWindowTitle('Перегляд листа')
        self.setGeometry(200, 200, 600, 500)

        self.setStyleSheet(
            "QDialog { background-color: #f8f8f8; border: 1px solid #ccc; }"
            "QWidget { background-color: #f8f8f8; color: #333; font-family: 'Segoe UI', sans-serif; }"
            "QTextEdit { background-color: #fff; border: 2px solid #ccc; padding: 8px; font-size: 12px; "
            "border-radius: 5px; }"
            "QLabel { font-weight: bold; color: #555; }"
            "QListWidget { border: 2px solid #ccc; padding: 8px; font-size: 12px; border-radius: 5px; }"
            "QListWidget::item { padding: 4px; }"
            "QPushButton { background-color: #4CAF50; color: white; border: none; padding: 8px 12px; "
            "text-align: center; text-decoration: none; display: inline-block; font-size: 12px; "
            "margin: 4px 2px; cursor: pointer; border-radius: 4px; }"
        )

        self.review_text = QTextBrowser(self)
        self.subject_label = QLabel(f"<font color='#007acc'>Subject:</font> {subject}")
        self.message_text = QTextEdit(message)
        self.message_text.setReadOnly(True)

        self.attachment_label = QLabel("<font color='#007acc'>Attachments:</font>")
        self.attachment_list = QListWidget()
        self.attachment_list.setSelectionMode(QListWidget.SingleSelection)
        self.attachment_list.setViewMode(QListWidget.IconMode)
        self.attachment_list.setIconSize(QSize(100, 100))  # Set the desired icon size

        if attachments:
            for attachment in attachments:
                if attachment.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    item = QListWidgetItem(QIcon(attachment), attachment)
                    item.setData(Qt.UserRole, attachment)
                    self.attachment_list.addItem(item)

        self.attachment_list.itemSelectionChanged.connect(self.show_selected_image)

        layout = QVBoxLayout()
        layout.addWidget(self.subject_label)
        layout.addWidget(self.message_text)
        layout.addWidget(self.attachment_label)
        layout.addWidget(self.attachment_list)

        self.image_viewer = QLabel()
        self.image_viewer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.image_viewer.setAlignment(Qt.AlignCenter)
        self.image_viewer.setScaledContents(True)

        layout.addWidget(self.image_viewer)

        button_layout = QHBoxLayout()
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def show_selected_image(self):
        selected_item = self.attachment_list.currentItem()
        if selected_item:
            file_path = selected_item.data(Qt.UserRole)
            pixmap = QPixmap(file_path)
            self.image_viewer.setPixmap(pixmap)
            self.image_viewer.setHidden(False)


if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        email_window = EmailSenderWindow()
        email_window.show()

        logging.info('Email Sender application started.')

        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error in main: {e}")
        logging.error(f"Error in main: {e}")
