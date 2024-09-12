import re
import os
import email
import imaplib
import pandas as pd
import xml.etree.ElementTree as ET

from ast import parse
from email.header import decode_header
from dotenv import load_dotenv

load_dotenv()

# Configuración
CHECK_INTERVAL = 10  # Intervalo de chequeo en segundos
IMAP_SERVER = "imap.gmail.com"
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
PASSWORD = os.getenv("PASSWORD")
XLS_FILE = os.getenv("XLS_FILE")


def check_for_new_emails():
    # Conexión al servidor IMAP
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    mail.select("inbox")

    # Buscar correos no leídos
    _, data = mail.search(None, "UNSEEN")
    email_ids = data[0].split()

    for email_id in email_ids:
        _, msg_data = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        subject = msg["subject"] if msg["subject"] else "Sin Asunto"
        print(f"Nuevo correo: {subject}")

        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename and filename.endswith(".xml"):
                        # with open(filename, "wb") as f:
                        # f.write(part.get_payload(decode=True))
                        xml_content = (
                            part.get_payload(decode=True)
                            .decode("utf-8")
                            .replace("&lt;", "<")
                            .replace("&gt;", ">")
                        )
                        # TODO: Parsear el contenido XML
                        mail.store(email_id, "+FLAGS", "\\Seen")  # Marcar como leído
                    else:
                        mail.store(email_id, "-FLAGS", "\\Seen")  # Marcar como no leído
    mail.logout()


def parse_xml(root, parent_tag, child_tags):
    # Si child_tags es una cadena, lo convertimos en una lista
    if isinstance(child_tags, str):
        child_tags = [child_tags]

    parent = root.find(parent_tag).text
    parent_root = ET.fromstring(parent)

    # Extraemos los resultados para cada etiqueta secundaria
    results = [parent_root.find(".//%s" % tag).text for tag in child_tags]
    return results if len(results) > 1 else results[0]


def extract_child_tags(root, parent_tag, child_tags):
    # Obtener el contenido del CDATA
    factura_xml_cdata = root.find(parent_tag).text
    factura_xml = re.sub(r"<!\[CDATA\[|\]\]>", "", factura_xml_cdata).strip()

    try:
        factura_root = ET.fromstring(factura_xml)
    except ET.ParseError as e:
        print(f"Error parsing factura XML: {e}")
        return []

    # Encontrar la sección <detalles>
    detalles = factura_root.find(child_tags[0])

    # Lista para almacenar los resultados
    detalles_list = []

    # Iterar sobre cada <detalle>
    for detalle in detalles.findall(child_tags[1]):
        detalle_dict = {}
        for field in detalle:
            detalle_dict[field.tag] = field.text
        detalles_list.append(detalle_dict)

    return detalles_list


def extract_block(root, parent_tag, child_tag, block_name=None):
    factura_xml_cdata = root.find(parent_tag).text
    factura_xml = re.sub(r"<!\[CDATA\[|\]\]>", "", factura_xml_cdata).strip()
    try:
        factura_root = ET.fromstring(factura_xml)
    except ET.ParseError as e:
        print(f"Error parsing factura XML: {e}")

    info_factura = factura_root.find(child_tag)

    # create a map to store the results

    if block_name:
        for child in info_factura:
            if child.tag == block_name:
                return child.text
    results = {}

    if info_factura is not None:
        for child in info_factura:
            results[child.tag] = child.text

    return results


if __name__ == "__main__":
    # open file fact_0103183026001_001-003-000094309.xml and convert it to string
    # with open("FA001613000005158.xml", "r") as f:
    # with open("fact_0103183026001_001-003-000094309.xml", "r") as f:
    with open("0509202401010284147500120010020000036261234567811.xml", "r") as f:
        xml_content = f.read()

    root = ET.fromstring(xml_content)

    razonSocial = parse_xml(root, "comprobante", "razonSocial")
    nombreComercial = parse_xml(root, "comprobante", "nombreComercial")
    ruc = parse_xml(root, "comprobante", "ruc")

    numFactura_parts = parse_xml(root, "comprobante", ["estab", "ptoEmi", "secuencial"])
    numFactura = "".join(numFactura_parts)
    detalles = extract_child_tags(root, "comprobante", ["detalles", "detalle"])
    
    # fechaEmision = extract_block(root, "comprobante", "infoFactura", "fechaEmision")
    infofactura = (extract_block(root, "comprobante", "infoFactura"))

    file_path = "FACTURAS.xls"
    if os.path.exists(file_path):
        # Load existing file
        df = pd.read_excel(file_path, sheet_name=None)
    else:
        # Create a new file with the required headers
        df = {
            "Factura": pd.DataFrame(
                columns=[
                    "Num Factura",
                    "Razón Social",
                    "Nombre Comercial",
                    "RUC",
                    "Fecha de emisión",
                ]
            ),
            "Detalle": pd.DataFrame(
                columns=["Descripción", "Cantidad", "Precio Unitario", "Descuento"]
            ),
        }


# while True:
# check_for_new_emails()
# time.sleep(CHECK_INTERVAL)
