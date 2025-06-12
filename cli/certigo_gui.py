import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QFileDialog, QVBoxLayout, QLabel, QPushButton, QLineEdit,
    QComboBox, QTextEdit, QCheckBox, QHBoxLayout
)
from PyQt5.QtCore import Qt
from certigo import create_certificate, digitally_sign, send_email, CERT_FOLDER

class CertigoGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Certigo Certificate Generator")
        self.setGeometry(100, 100, 600, 500)
        self.layout = QVBoxLayout()

        # File selectors
        self.excel_input = self.add_file_input("Excel File (.xlsx)", "*.xlsx")
        self.bg_input = self.add_file_input("Background Image", "*.png *.jpg")
        self.config_input = self.add_file_input("Layout Config (JSON)", "*.json")

        # Options
        self.paper_size = self.add_dropdown("Paper Size", ["A4", "LETTER"])
        self.orientation = self.add_dropdown("Orientation", ["landscape", "portrait"])

        # Sign options
        self.sign_checkbox = QCheckBox("Digitally Sign Certificates")
        self.layout.addWidget(self.sign_checkbox)
        self.cert_input = self.add_file_input("Certificate (.pem)", "*.pem")
        self.key_input = self.add_file_input("Private Key (.pem)", "*.pem")
        self.pass_input = self.add_text_input("Private Key Password")

        # Email options
        self.email_checkbox = QCheckBox("Send via Email")
        self.layout.addWidget(self.email_checkbox)
        self.sender_input = self.add_text_input("Sender Gmail")
        self.app_pass_input = self.add_text_input("Gmail App Password", echo=True)

        # Run button
        self.run_button = QPushButton("Generate Certificates")
        self.run_button.clicked.connect(self.run_certigo)
        self.layout.addWidget(self.run_button)

        # Output log
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.layout.addWidget(self.log_output)

        self.setLayout(self.layout)

    def add_file_input(self, label, file_filter):
        lbl = QLabel(label)
        self.layout.addWidget(lbl)
        hbox = QHBoxLayout()
        le = QLineEdit()
        btn = QPushButton("Browse")
        def browse():
            path, _ = QFileDialog.getOpenFileName(self, f"Select {label}", "", file_filter)
            if path:
                le.setText(path)
        btn.clicked.connect(browse)
        hbox.addWidget(le)
        hbox.addWidget(btn)
        self.layout.addLayout(hbox)
        return le

    def add_text_input(self, label, echo=False):
        lbl = QLabel(label)
        self.layout.addWidget(lbl)
        le = QLineEdit()
        if echo:
            le.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(le)
        return le

    def add_dropdown(self, label, options):
        lbl = QLabel(label)
        self.layout.addWidget(lbl)
        cb = QComboBox()
        cb.addItems(options)
        self.layout.addWidget(cb)
        return cb

    def log(self, message):
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def run_certigo(self):
        try:
            df = pd.read_excel(self.excel_input.text())
            config = json.load(open(self.config_input.text()))
            bg = self.bg_input.text()
            paper_size = self.paper_size.currentText()
            orientation = self.orientation.currentText()

            os.makedirs(CERT_FOLDER, exist_ok=True)

            sign = self.sign_checkbox.isChecked()
            email = self.email_checkbox.isChecked()

            for _, row in df.iterrows():
                name = str(row['name']).strip()
                cert_no = str(row['cert_no']).strip()
                to_email = row['email']
                filename = f"{name.replace(' ', '_')}_{cert_no}.pdf"
                pdf_path = os.path.join(CERT_FOLDER, filename)

                create_certificate(name, cert_no, bg, pdf_path, paper_size, orientation, config)
                self.log(f"✔ Created certificate for {name}")

                final_path = pdf_path
                if sign:
                    final_path = digitally_sign(
                        self.cert_input.text(), 
                        self.key_input.text(),
                        self.pass_input.text(),
                        pdf_path
                    )
                    self.log(f"🔏 Signed certificate for {name}")

                if email:
                    send_email(
                        self.sender_input.text(),
                        self.app_pass_input.text(),
                        to_email,
                        subject="Your Certificate",
                        body=f"Dear {name},\n\nPlease find your certificate attached.\n\nCertificate Code: {cert_no}\n\nRegards,\nTeam",
                        attachment=final_path
                    )

        except Exception as e:
            self.log(f"❌ Error: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CertigoGUI()
    window.show()
    sys.exit(app.exec_())
