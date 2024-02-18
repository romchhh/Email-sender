import sys
import pandas as pd
import logging
from typing import Dict
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QFileDialog, QVBoxLayout, QTextEdit, QHBoxLayout, 
                             QMessageBox, QDialog, QListWidget, QListWidget, QListWidgetItem, QHBoxLayout, QPushButton
)
from PyQt5.QtCore import QSize
from PyQt5.QtCore import Qt
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QFile, QTextStream
from PyQt5.QtWebEngineWidgets import QWebEngineView
from accounts_serializer import get_accounts
from email_thread import EmailSenderThread, Signal
from PyQt5.QtCore import Qt
from info import html_guide


class UTF8LogFormatter(logging.Formatter):
    def format(self, record):
        result = super().format(record)
        return result.encode('utf-8', 'replace').decode('utf-8')

logging.basicConfig(filename='email_sender.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logging.getLogger().handlers[0].setFormatter(UTF8LogFormatter())

MIN_MESSAGE_HEIGHT = 200
MIN_DIALOG_SIZE = (800, 600)
EMAIL_SEND_LIMIT = 5

class EmailSenderWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.thread: Dict[EmailSenderThread] = {}
        self.init_ui()

    def init_ui(self):
        self.setGeometry(100, 100, 1350, 900)
        self.setWindowTitle('Програма для розсилки')

        style_file = QFile("styles/design.qss")
        style_file.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(style_file)
        self.setStyleSheet(stream.readAll())
        style_file.close()


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
        self.send_button.clicked.connect(self.start_email_thread)

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
        
        self.print_log = QTextEdit(self)
        self.print_log.setReadOnly(True)
        self.print_log.setMinimumHeight(100)
        self.print_log.setStyleSheet("background-color: #fff; border: 2px solid #aaa; padding: 8px; font-size: 16px; border-radius: 5px;")

        layout.addWidget(self.print_log)
        
        

    def print_to_log(self, message):
        current_text = self.print_log.toPlainText()
        self.print_log.setPlainText(current_text + message + '\n')
        self.print_log.verticalScrollBar().setValue(self.print_log.verticalScrollBar().maximum())

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

    def start_email_thread(self):
        if hasattr(self, 'df'):
            receiver_emails = self.df.iloc[:, 0].tolist()
            subject = self.subject_input.text()
            message = self.message_input.toPlainText()
            
            if not (receiver_emails and subject and message):
                self.status_label.setText('Будь ласка, заповніть усі поля.')

            else:
                try:

                    self.thread['email'] = EmailSenderThread(receiver_emails, subject, message, self.accounts, self.attachments)
                    self.thread['email'].start()
                    self.thread['email'].update_log.connect(self.send_emails)

                except Exception as e:
                    QMessageBox.critical(self, 'Error', f'Error sending emails: {str(e)}')
                    self.status_label.setText('Помилка надсилання електронних листів.')
        else:
            self.status_label.setText('Спочатку виберіть файл Excel.')

    def send_emails(self, signal: Signal):
        try:
            
            self.print_to_log(signal.log)

            print(signal.sent_label)
            # self.sent_label.setText(signal.sent_label)
            self.status_label.setText(signal.sent_label)            
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error sending emails: {str(e)}')
            self.status_label.setText('Помилка надсилання електронних листів.')


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
        self.setGeometry(200, 200, 800, 600)

        style_file = QFile("styles/design_preview.qss")
        style_file.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(style_file)
        stylesheet = stream.readAll()
        style_file.close()

        # Apply the stylesheet
        self.setStyleSheet(stylesheet)

        self.subject_label = QLabel(f"<font color='#007acc'>Subject:</font> {subject}")
        self.message_text = QTextEdit(message)
        self.message_text.setReadOnly(True)

        self.attachment_label = QLabel("<font color='#007acc'>Attachments:</font>")
        self.attachment_list = QListWidget()
        self.attachment_list.setSelectionMode(QListWidget.SingleSelection)
        self.attachment_list.setViewMode(QListWidget.IconMode)
        self.attachment_list.setIconSize(QSize(100, 100))

        if attachments:
            for attachment in attachments:
                item = QListWidgetItem(attachment)
                item.setData(Qt.UserRole, attachment)
                self.attachment_list.addItem(item)

        self.attachment_list.itemSelectionChanged.connect(self.show_selected_attachment)

        layout = QVBoxLayout()
        layout.addWidget(self.subject_label)
        layout.addWidget(self.message_text)
        layout.addWidget(self.attachment_label)
        layout.addWidget(self.attachment_list)

        self.attachment_content_viewer = QWebEngineView()
        layout.addWidget(self.attachment_content_viewer)

        button_layout = QHBoxLayout()
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def show_selected_attachment(self):
        selected_item = self.attachment_list.currentItem()
        if selected_item:
            file_path = selected_item.data(Qt.UserRole)
            if file_path.lower().endswith(('.html', '.htm')):
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                self.attachment_content_viewer.setHtml(html_content)
            else:
                # Clear the viewer if the selected attachment is not HTML
                self.attachment_content_viewer.setHtml('')


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
