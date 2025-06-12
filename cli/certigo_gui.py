import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QFileDialog, QVBoxLayout, QLabel, QPushButton, QLineEdit,
    QComboBox, QTextEdit, QCheckBox, QHBoxLayout, QGroupBox, QProgressBar, QMessageBox
)
from PyQt5.QtCore import Qt
from certigo import create_certificate, digitally_sign, send_email

class CertigoGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Certigo Certificate Generator")
        self.setGeometry(100, 100, 700, 750)

        self.layout = QVBoxLayout()

        self.build_input_section()
        self.build_sign_section()
        self.build_email_section()
        self.build_controls()

        self.setLayout(self.layout)

    def build_input_section(self):
        group = QGroupBox("1. Input Files")
        vbox = QVBoxLayout()
        self.excel_input = self.add_file_input("Excel File (.xlsx)", "*.xlsx", vbox)
        self.bg_input = self.add_file_input("Background Image", "*.png *.jpg", vbox)
        self.config_input = self.add_file_input("Layout Config (JSON)", "*.json", vbox)
        self.output_dir_input = self.add_folder_input("Output Folder", vbox)
        self.paper_size = self.add_dropdown("Paper Size", ["A4", "LETTER"], vbox)
        self.orientation = self.add_dropdown("Orientation", ["landscape", "portrait"], vbox)
        group.setLayout(vbox)
        self.layout.addWidget(group)

    def build_sign_section(self):
        self.sign_checkbox = QCheckBox("Enable Digital Signing")
        self.sign_checkbox.stateChanged.connect(self.toggle_sign_section)
        self.layout.addWidget(self.sign_checkbox)

        self.sign_group = QGroupBox("2. Digital Sign Settings")
        self.sign_group.setVisible(False)
        vbox = QVBoxLayout()
        self.cert_input = self.add_file_input("Certificate File (.pem)", "*.pem", vbox)
        self.key_input = self.add_file_input("Private Key File (.pem)", "*.pem", vbox)
        self.pass_input = self.add_text_input("Private Key Password", vbox, echo=True)
        self.sign_group.setLayout(vbox)
        self.layout.addWidget(self.sign_group)

    def build_email_section(self):
        self.email_checkbox = QCheckBox("Enable Emailing Certificates")
        self.email_checkbox.stateChanged.connect(self.toggle_email_section)
        self.layout.addWidget(self.email_checkbox)

        self.email_group = QGroupBox("3. Email Settings")
        self.email_group.setVisible(False)
        vbox = QVBoxLayout()

        self.sender_input = self.add_text_input("Sender Gmail", vbox)

        vbox.addWidget(QLabel("Gmail App Password"))
        hbox = QHBoxLayout()
        self.app_pass_input = QLineEdit()
        self.app_pass_input.setEchoMode(QLineEdit.Password)
        hbox.addWidget(self.app_pass_input)

        help_btn = QPushButton("❓")
        help_btn.setFixedWidth(30)
        help_btn.clicked.connect(self.show_gmail_help)
        hbox.addWidget(help_btn)

        vbox.addLayout(hbox)
        self.email_group.setLayout(vbox)
        self.layout.addWidget(self.email_group)

    def build_controls(self):
        self.run_button = QPushButton("Generate Certificates")
        self.run_button.clicked.connect(self.run_certigo)
        self.layout.addWidget(self.run_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.layout.addWidget(self.log_output)

    def add_file_input(self, label, file_filter, layout):
        layout.addWidget(QLabel(label))
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
        layout.addLayout(hbox)
        return le

    def add_folder_input(self, label, layout):
        layout.addWidget(QLabel(label))
        hbox = QHBoxLayout()
        le = QLineEdit()
        btn = QPushButton("Browse")
        def browse():
            folder = QFileDialog.getExistingDirectory(self, f"Select {label}")
            if folder:
                le.setText(folder)
        btn.clicked.connect(browse)
        hbox.addWidget(le)
        hbox.addWidget(btn)
        layout.addLayout(hbox)
        return le

    def add_text_input(self, label, layout, echo=False):
        layout.addWidget(QLabel(label))
        le = QLineEdit()
        if echo:
            le.setEchoMode(QLineEdit.Password)
        layout.addWidget(le)
        return le

    def add_dropdown(self, label, options, layout):
        layout.addWidget(QLabel(label))
        cb = QComboBox()
        cb.addItems(options)
        layout.addWidget(cb)
        return cb

    def toggle_sign_section(self):
        self.sign_group.setVisible(self.sign_checkbox.isChecked())

    def toggle_email_section(self):
        self.email_group.setVisible(self.email_checkbox.isChecked())

    def log(self, message):
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def show_gmail_help(self):
        help_text = (
            "How to Create a Gmail App Password:\n\n"
            "1. Go to your Google Account → Security tab.\n"
            "2. Enable 2-Step Verification (if not already enabled).\n"
            "3. After enabling, go to 'App passwords'.\n"
            "4. Select app: 'Mail', device: 'Other', and name it (e.g., Certigo).\n"
            "5. Click 'Generate' – copy the 16-character password.\n"
            "6. Paste it here in 'Gmail App Password'."
        )
        QMessageBox.information(self, "Gmail App Password Instructions", help_text)

    def run_certigo(self):
        try:
            df = pd.read_excel(self.excel_input.text())
            config = json.load(open(self.config_input.text()))
            bg = self.bg_input.text()
            paper_size = self.paper_size.currentText()
            orientation = self.orientation.currentText()
            output_folder = self.output_dir_input.text().strip()

            if not output_folder:
                QMessageBox.warning(self, "Missing Output Folder", "Please select an output folder.")
                return

            os.makedirs(output_folder, exist_ok=True)

            sign = self.sign_checkbox.isChecked()
            email = self.email_checkbox.isChecked()

            self.progress_bar.setVisible(True)
            self.progress_bar.setMaximum(len(df))
            self.progress_bar.setValue(0)

            for i, (_, row) in enumerate(df.iterrows()):
                name = str(row['name']).strip()
                cert_no = str(row['cert_no']).strip()
                to_email = row['email']
                filename = f"{name.replace(' ', '_')}_{cert_no}.pdf"
                pdf_path = os.path.join(output_folder, filename)

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

                self.progress_bar.setValue(i + 1)

            QMessageBox.information(self, "Done", "All certificates processed successfully.")

        except Exception as e:
            self.log(f"❌ Error: {e}")
            QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CertigoGUI()
    window.show()
    sys.exit(app.exec_())
