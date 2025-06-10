from OpenSSL import crypto
import getpass
from datetime import datetime, timedelta

def get_user_details():
    print("\n" + "="*50)
    print("Document Signing Certificate Wizard")
    print("="*50 + "\n")
    
    details = {
        'common_name': input("1. Common Name (e.g., your name/company): ").strip(),
        'email': input("2. Email address: ").strip(),
        'country': input("3. Country Code (2 letters, e.g., US): ").strip().upper(),
        'state': input("4. State/Province: ").strip(),
        'city': input("5. City/Locality: ").strip(),
        'organization': input("6. Organization: ").strip(),
        'valid_days': int(input("7. Validity period (days, default 365): ") or 365),
        'key_password': getpass.getpass("8. Enter password for private key: "),
        'key_password_confirm': getpass.getpass("   Confirm password: ")
    }
    
    while details['key_password'] != details['key_password_confirm']:
        print("\nPasswords don't match! Please try again.")
        details['key_password'] = getpass.getpass("Enter password for private key: ")
        details['key_password_confirm'] = getpass.getpass("Confirm password: ")
    
    return details

def generate_certificate(user_details):
    # Generate key pair
    key = crypto.PKey()
    key.generate_key(crypto.TYPE_RSA, 4096)
    
    # Create certificate
    cert = crypto.X509()
    
    # Set certificate details
    cert.get_subject().CN = user_details['common_name']
    cert.get_subject().emailAddress = user_details['email']
    cert.get_subject().C = user_details['country']
    cert.get_subject().ST = user_details['state']
    cert.get_subject().L = user_details['city']
    cert.get_subject().O = user_details['organization']
    
    # Set validity period
    cert.set_serial_number(int(datetime.now().timestamp()))
    cert.gmtime_adj_notBefore(0)
    cert.gmtime_adj_notAfter(user_details['valid_days'] * 24 * 60 * 60)
    
    # Self-sign
    cert.set_issuer(cert.get_subject())
    cert.set_pubkey(key)
    cert.sign(key, 'sha512')
    
    # Save password-protected private key
    with open("key.pem", "wb") as f:
        f.write(crypto.dump_privatekey(
            crypto.FILETYPE_PEM,
            key,
            cipher='aes256',
            passphrase=user_details['key_password'].encode()
        ))
    
    # Save certificate
    with open("cert.pem", "wb") as f:
        f.write(crypto.dump_certificate(crypto.FILETYPE_PEM, cert))
    
    print("\n" + "="*50)
    print("Certificate Generation Complete!")
    print("="*50)
    print(f"\nFiles created:")
    print(f"- Private Key: key.pem (password protected)")
    print(f"- Certificate: cert.pem")
    print("\nKeep these files secure, especially the private key!")

if __name__ == "__main__":
    try:
        user_details = get_user_details()
        generate_certificate(user_details)
    except Exception as e:
        print(f"\nError: {e}")
        print("Certificate generation failed.")