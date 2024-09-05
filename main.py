import imaplib
import email
import time
import os
import xml.etree.ElementTree as ET


from email.header import decode_header
from dotenv import load_dotenv

load_dotenv()

# Configuración
CHECK_INTERVAL = 30  # Intervalo de chequeo en segundos
IMAP_SERVER = "imap.gmail.com"
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
PASSWORD = os.getenv("PASSWORD")
XLS_FILE = os.getenv("XLS_FILE")


def check_for_new_emails():
    # Conectarse al servidor IMAP
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    mail.select("inbox")

    # Buscar correos no leídos
    _, data = mail.search(None, "UNSEEN")
    email_ids = data[0].split()

    for email_id in email_ids:
        _, msg_data = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        subject = msg["subject"]
        print(f"Nuevo correo: {subject}")

        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        print(f"Archivo adjunto encontrado: {filename}")
                        # with open(filename, "wb") as f:
                        # f.write(part.get_payload(decode=True))
                        if filename.endswith(".xml"):
                            xml_content = (
                                part.get_payload(decode=True)
                                .decode("utf-8")
                                .replace("&lt;", "<")
                                .replace("&gt;", ">")
                            )

        mail.store(email_id, "-FLAGS", "\\Seen")  # Marcar como no leído

    mail.logout()


def parse_xml_content(xml_content, root_tag, child_tag):
    # Parsear el contenido XML
    root = ET.fromstring(xml_content)
    fact = root.find(".//%s" % root_tag)
    if fact:
        result = fact.find(".//%s" % child_tag)
        return result.text
    return None


if __name__ == "__main__":
    while True:
        check_for_new_emails()
        time.sleep(CHECK_INTERVAL)
