import os
import click
import pandas as pd
import json
from pyhanko.sign import signers, sign_pdf
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, LETTER, landscape, portrait
from reportlab.lib.utils import ImageReader
import yagmail

CERT_FOLDER = "certificates"
PAPER_SIZES = {'A4': A4, 'LETTER': LETTER}

# Draw certificate
def create_certificate(name, cert_no, bg_image_path, output_path, paper_size, orientation, layout_cfg):
    size = PAPER_SIZES.get(paper_size.upper(), A4)
    page_size = landscape(size) if orientation == 'landscape' else portrait(size)
    width, height = page_size
    c = canvas.Canvas(output_path, pagesize=page_size)

    if os.path.exists(bg_image_path):
        bg = ImageReader(bg_image_path)
        c.drawImage(bg, 0, 0, width=width, height=height)

    # Draw name
    draw_text(c, layout_cfg['name'], name)

    # Draw cert_no
    draw_text(c, layout_cfg['cert_no'], "Certificate No: " + cert_no)

    c.save()

def draw_text(c, cfg, text):
    c.setFont(cfg['font'], cfg['size'])
    c.setFillColorRGB(*(v / 255 for v in cfg['color']))
    align = cfg.get('align', 'left')
    x, y = cfg['x'], cfg['y']
    if align == 'center':
        c.drawCentredString(x, y, text)
    else:
        c.drawString(x, y, text)

# Sign PDF
def digitally_sign(cert_path, key_path, password, pdf_path):
    signer = signers.SimpleSigner.load(
        cert_file=cert_path,
        key_file=key_path,
        key_passphrase=password.encode()
    )
    with open(pdf_path, 'rb') as inf:
        w = IncrementalPdfFileWriter(inf)
        out = sign_pdf(
            w,
            signature_meta=signers.PdfSignatureMetadata(field_name="Signature1"),
            signer=signer,
            existing_fields_only=False
        )
        signed_path = pdf_path.replace(".pdf", "_signed.pdf")
        with open(signed_path, 'wb') as outf:
            outf.write(out.getbuffer())
        return signed_path

# Send email
def send_email(sender_email, app_password, to_email, subject, body, attachment):
    try:
        yag = yagmail.SMTP(
            user=sender_email, 
            password=app_password,
            host="smtp.gmail.com",
            port=587,
            smtp_starttls=True,
            smtp_ssl=False
        )
        yag.send(to=to_email, subject=subject, contents=body, attachments=attachment)
        print(f"📧 Email sent to {to_email}")
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")

# CLI
@click.command()
@click.option('--excel', required=True, help='Excel file with cert_no,name,email', default='data.xlsx')
@click.option('--bg-image', required=True, help='Background PNG/JPG image', default='template.png')
@click.option('--config', required=True, help='JSON file with font/position/color settings', default='config.json')
@click.option('--orientation', type=click.Choice(['portrait', 'landscape']), default='landscape')
@click.option('--paper-size', type=click.Choice(['A4', 'LETTER']), default='A4')
@click.option('--sign/--no-sign', default=False)
@click.option('--cert', default='cert.pem')
@click.option('--key', default='key.pem')
@click.option('--password', default='password', help='Password for the private key. If not set, Default password is password.')
@click.option('--email/--no-email', default=False)
@click.option('--sender', help='Gmail sender email')
@click.option('--app-pass', help='Gmail app password')
def main(excel, bg_image, config, orientation, paper_size, sign, cert, key, password, email, sender, app_pass):
    df = pd.read_excel(excel)
    layout_cfg = json.load(open(config))
    os.makedirs(CERT_FOLDER, exist_ok=True)

    for _, row in df.iterrows():
        name = str(row['name']).strip()
        cert_no = str(row['cert_no']).strip()
        to_email = row['email']

        filename = f"{name.replace(' ', '_')}_{cert_no}.pdf"
        pdf_path = os.path.join(CERT_FOLDER, filename)

        create_certificate(name, cert_no, bg_image, pdf_path, paper_size, orientation, layout_cfg)

        final_path = digitally_sign(cert, key, password, pdf_path) if sign else pdf_path

        if email:
            send_email(sender, app_pass, to_email,
                       subject="Your Certificate",
                       body=f"Dear {name},\n\nPlease find your certificate attached.\n\nCertificate Code: {cert_no}\n\nRegards,\nTeam",
                       attachment=final_path)

        print(f"✔ Certificate for {name} (Code: {cert_no}) generated.")

if __name__ == '__main__':
    main()
