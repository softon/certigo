# CertiGo
A simple cli and gui tool to generate customised certificates and email it using your gmail account.

## Setup
1. Start a new `venv` if not already present
```
py -3.13 -m venv venv
```
1. Activate the `venv`
```
venv\Scripts\activate
```
1. Install required packages
```
pip install pandas yagmail reportlab pyhanko click openpyxl pyopenssl PyQt5 PyMuPDF pyinstaller
```

# Generate Self Signed Certificate for Digital Signing of Documents
1. Run `certigen.py`
```
python certigen.py
```
2. Complete the wizard and two files will be created `cert.pem` and `key.pem`

# Run the Program
Run the main program to start the generation
```
python certigo.py
```

### Options of certigo
```
python certigo.py --help
```

```
Options:
  --excel TEXT                    Excel file with cert_no,name,email
                                  [required]
  --bg-image TEXT                 Background PNG/JPG image  [required]
  --config TEXT                   JSON file with font/position/color settings
                                  [required]
  --orientation [portrait|landscape]
  --paper-size [A4|LETTER]
  --sign / --no-sign
  --cert TEXT
  --key TEXT
  --password TEXT                 Password for the private key. If not set,
                                  Default password is password.
  --email / --no-email
  --sender TEXT                   Gmail sender email
  --app-pass TEXT                 Gmail app password
  --help                          Show this message and exit.
```

# How to Generate an App Password for Gmail

## Prerequisites
- ✅ **2-Step Verification** must be enabled on your Google account  
- ⚠️ For applications that don't support OAuth (Outlook, Thunderbird, scripts, etc.)

---

## Step-by-Step Guide

### 1. Access Google Account Settings
🔗 [myaccount.google.com](https://myaccount.google.com)  
→ Click **"Security"** in left sidebar

### 2. Enable 2-Step Verification (If Needed)

Skip if already enabled
Under "Signing in to Google" → "2-Step Verification"

Follow prompts (requires phone verification)

### 3. Generate App Password
Return to Security page

Find "App passwords" under "Signing in to Google"
(If missing, verify 2-Step Verification is active)

Configure settings:

Select App: Choose Mail

Select Device: Pick your device or Other (Custom Name)

Click "Generate"

### 4. Save Your App Password
📋 You'll receive a 16-character code formatted like:
xxxx xxxx xxxx xxxx
(Remove spaces when using)

# Sample Commands

```
python
```


# CLI Guide

```
pip install pyinstaller

pyinstaller --onefile --name certigo --add-data "config.json;data.xlsx;template.png" certigo.py
pyinstaller --onefile --name certigo certigo.py
pyinstaller --onefile --name certigen certigen.py
pyinstaller splash.py --noconfirm --onefile --windowed --icon=icon.ico --name=Certigo --clean

```


# Use UPX Compression
UPX (Ultimate Packer for eXecutables) compresses .exe files losslessly.

🔧 How to use:
Install UPX:
Download from https://upx.github.io/ and add it to your PATH.

Rebuild with UPX enabled:

```
pyinstaller yourscript.py --onefile --windowed --noconfirm --clean --upx-dir=C:\path\to\upx
```
🎯 Compression rate: 20%–50% smaller