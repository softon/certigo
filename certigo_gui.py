import sys, os, json, fitz
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QFileDialog, QVBoxLayout, QLabel, QPushButton, QLineEdit,
    QComboBox, QTextEdit, QCheckBox, QHBoxLayout, QGroupBox, QProgressBar, QMessageBox,
    QTabWidget, QSpinBox, QColorDialog, QScrollArea
)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt
from certigo import create_certificate, digitally_sign, send_email

CONFIG_PATH = "config.json"
PREVIEW_PDF = "__preview__.pdf"

class CertigoGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Certigo Certificate Generator")
        self.setGeometry(100, 100, 800, 900)

        self.tabs = QTabWidget()
        self.main_tab = QWidget()
        self.settings_tab = QWidget()
        self.preview_tab = QWidget()

        self.tabs.addTab(self.main_tab, "Main")
        self.tabs.addTab(self.settings_tab, "Settings")
        self.tabs.addTab(self.preview_tab, "Preview")

        self.build_main_tab()
        self.build_settings_tab()
        self.build_preview_tab()

        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    # ========== MAIN TAB ==========
    def build_main_tab(self):
        layout = QVBoxLayout()

        self.excel_input = self.add_file_input("Excel File (.xlsx)", "*.xlsx", layout)
        self.bg_input = self.add_file_input("Background Image", "*.png *.jpg", layout)
        self.output_dir_input = self.add_folder_input("Output Folder", layout)
        self.paper_size = self.add_dropdown("Paper Size", ["A4", "LETTER"], layout)
        self.orientation = self.add_dropdown("Orientation", ["landscape", "portrait"], layout)

        self.sign_checkbox = QCheckBox("Enable Digital Signing")
        self.sign_checkbox.stateChanged.connect(self.toggle_sign_section)
        layout.addWidget(self.sign_checkbox)

        self.sign_group = QGroupBox("Digital Sign Settings")
        self.sign_group.setVisible(False)
        sign_layout = QVBoxLayout()
        self.cert_input = self.add_file_input("Certificate File (.pem)", "*.pem", sign_layout)
        self.key_input = self.add_file_input("Private Key File (.pem)", "*.pem", sign_layout)
        self.pass_input = self.add_text_input("Private Key Password", sign_layout, echo=True)
        self.sign_group.setLayout(sign_layout)
        layout.addWidget(self.sign_group)

        self.email_checkbox = QCheckBox("Enable Emailing Certificates")
        self.email_checkbox.stateChanged.connect(self.toggle_email_section)
        layout.addWidget(self.email_checkbox)

        self.email_group = QGroupBox("Email Settings")
        self.email_group.setVisible(False)
        email_layout = QVBoxLayout()
        self.sender_input = self.add_text_input("Sender Gmail", email_layout)
        email_layout.addWidget(QLabel("Gmail App Password"))
        email_hbox = QHBoxLayout()
        self.app_pass_input = QLineEdit()
        self.app_pass_input.setEchoMode(QLineEdit.Password)
        help_btn = QPushButton("❓")
        help_btn.setFixedWidth(30)
        help_btn.clicked.connect(self.show_gmail_help)
        email_hbox.addWidget(self.app_pass_input)
        email_hbox.addWidget(help_btn)
        email_layout.addLayout(email_hbox)
        self.email_group.setLayout(email_layout)
        layout.addWidget(self.email_group)

        self.run_button = QPushButton("Generate Certificates")
        self.run_button.clicked.connect(self.run_certigo)
        layout.addWidget(self.run_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        self.main_tab.setLayout(layout)

    # ========== SETTINGS TAB ==========
    def build_settings_tab(self):
        layout = QVBoxLayout()

        self.config = self.load_config()
        self.setting_fields = {}

        for key in ["name", "cert_no"]:
            group = QGroupBox(f"{key.upper()} Settings")
            vbox = QVBoxLayout()
            self.setting_fields[key] = {}

            for field in ["x", "y", "font", "size", "align"]:
                hbox = QHBoxLayout()
                label = QLabel(field)
                hbox.addWidget(label)
                if field in ["x", "y", "size"]:
                    spin = QSpinBox()
                    spin.setMaximum(5000)
                    spin.setValue(self.config[key].get(field, 0))
                    hbox.addWidget(spin)
                    self.setting_fields[key][field] = spin
                elif field == "align":
                    cb = QComboBox()
                    cb.addItems(["left", "center"])
                    cb.setCurrentText(self.config[key].get(field, "left"))
                    hbox.addWidget(cb)
                    self.setting_fields[key][field] = cb
                else:
                    txt = QLineEdit(self.config[key].get(field, ""))
                    hbox.addWidget(txt)
                    self.setting_fields[key][field] = txt
                vbox.addLayout(hbox)

            # Color picker
            color_btn = QPushButton("Set Color")
            color_btn.clicked.connect(lambda _, k=key: self.pick_color(k))
            vbox.addWidget(color_btn)

            color_lbl = QLabel("Current Color: " + str(self.config[key]["color"]))
            self.setting_fields[key]["color_label"] = color_lbl
            vbox.addWidget(color_lbl)

            group.setLayout(vbox)
            layout.addWidget(group)

        hbox = QHBoxLayout()
        save_btn = QPushButton("Save Settings")
        save_btn.clicked.connect(self.save_config)
        hbox.addWidget(save_btn)

        preview_btn = QPushButton("Preview")
        preview_btn.clicked.connect(self.generate_preview)
        hbox.addWidget(preview_btn)

        layout.addLayout(hbox)
        self.settings_tab.setLayout(layout)

    def build_preview_tab(self):
        layout = QVBoxLayout()
        self.preview_label = QLabel("No preview yet.")
        self.preview_label.setAlignment(Qt.AlignCenter)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.preview_label)
        layout.addWidget(scroll)

        self.preview_tab.setLayout(layout)

    def load_config(self):
        if not os.path.exists(CONFIG_PATH):
            QMessageBox.critical(self, "Missing Config", "config.json not found!")
            sys.exit(1)
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)

    def pick_color(self, key):
        color = QColorDialog.getColor()
        if color.isValid():
            rgb = [color.red(), color.green(), color.blue()]
            self.config[key]["color"] = rgb
            self.setting_fields[key]["color_label"].setText("Current Color: " + str(rgb))

    def save_config(self):
        for key in self.setting_fields:
            for field in ["x", "y", "font", "size", "align"]:
                widget = self.setting_fields[key][field]
                self.config[key][field] = widget.currentText() if isinstance(widget, QComboBox) else widget.text() if isinstance(widget, QLineEdit) else widget.value()
        with open(CONFIG_PATH, 'w') as f:
            json.dump(self.config, f, indent=2)
        QMessageBox.information(self, "Saved", "Configuration updated successfully.")

    def generate_preview(self):
        bg = self.bg_input.text()
        if not os.path.exists(bg):
            QMessageBox.warning(self, "Missing Background", "Please select a background image.")
            return
        orientation = self.orientation.currentText()
        paper_size = self.paper_size.currentText()
        create_certificate("John Doe", "PREVIEW123", bg, PREVIEW_PDF, paper_size, orientation, self.config)

        try:
            doc = fitz.open(PREVIEW_PDF)
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            self.preview_label.setPixmap(pixmap)
            self.tabs.setCurrentWidget(self.preview_tab)
        except Exception as e:
            QMessageBox.critical(self, "Preview Error", f"Could not render preview: {e}")

    # ========== HELPERS ==========
    def add_file_input(self, label, file_filter, layout):
        layout.addWidget(QLabel(label))
        hbox = QHBoxLayout()
        le = QLineEdit()
        btn = QPushButton("Browse")
        btn.clicked.connect(lambda: self.browse_file(le, file_filter))
        hbox.addWidget(le)
        hbox.addWidget(btn)
        layout.addLayout(hbox)
        return le

    def browse_file(self, line_edit, file_filter):
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filter)
        if path:
            line_edit.setText(path)

    def add_folder_input(self, label, layout):
        layout.addWidget(QLabel(label))
        hbox = QHBoxLayout()
        le = QLineEdit()
        btn = QPushButton("Browse")
        btn.clicked.connect(lambda: self.browse_folder(le))
        hbox.addWidget(le)
        hbox.addWidget(btn)
        layout.addLayout(hbox)
        return le

    def browse_folder(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            line_edit.setText(folder)

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

    def show_gmail_help(self):
        text = (
            "How to Create a Gmail App Password:\n\n"
            "1. Go to your Google Account → Security tab.\n"
            "2. Enable 2-Step Verification (if not already enabled).\n"
            "3. After enabling, go to 'App passwords'.\n"
            "4. Select app: 'Mail', device: 'Other', and name it (e.g., Certigo).\n"
            "5. Click 'Generate' – copy the 16-character password.\n"
            "6. Paste it here in 'Gmail App Password'."
        )
        QMessageBox.information(self, "Gmail App Password Instructions", text)

    def log(self, message):
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def run_certigo(self):
        try:
            df = pd.read_excel(self.excel_input.text())
            with open(CONFIG_PATH) as f:
                config = json.load(f)
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
