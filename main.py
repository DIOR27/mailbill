import re
import os
import time
import xlrd
import xlwt
import email
import imaplib
import pandas as pd
import xml.etree.ElementTree as ET

from ast import parse
from xlutils.copy import copy
from dotenv import load_dotenv
from email.header import decode_header

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
                xml_content = None
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename and filename.endswith(".xml"):
                        xml_content = (
                            part.get_payload(decode=True)
                            .decode("utf-8")
                            .replace("&lt;", "<")
                            .replace("&gt;", ">")
                        )

                        process_xml(xml_content)

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
    factura_xml_cdata = root.find(parent_tag).text
    factura_xml = re.sub(r"<!\[CDATA\[|\]\]>", "", factura_xml_cdata).strip()
    try:
        factura_root = ET.fromstring(factura_xml)
    except ET.ParseError as e:
        print(f"Error parsing factura XML: {e}")
        return []

    # Encontrar la sección <detalles>
    detalles = factura_root.find(child_tags[0])

    # Crear una lista para guardar los detalles
    detalles_list = []

    # Iterar sobre cada <detalle>
    for detalle in detalles.findall(child_tags[1]):
        detalle_dict = {}
        for field in detalle:
            # Guardar solo los campos necesarios
            if field.tag in [
                "descripcion",
                "cantidad",
                "precioUnitario",
                "precioTotalSinImpuesto",
            ]:
                detalle_dict[field.tag] = field.text
        if detalle_dict:  # Agregar solo si se encontraron los campos relevantes
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


def write_to_excel(file_path, main_data, bill_data, details_data):
    # Verifica si el archivo ya existe
    try:
        # Abrir archivo existente
        workbook_rd = xlrd.open_workbook(file_path, formatting_info=True)
        sheet_rd = workbook_rd.sheet_by_index(0)
        row_start = (
            sheet_rd.nrows
        )  # Número de filas existentes para saber dónde escribir
        workbook = copy(workbook_rd)  # Crear una copia para escribir
        sheet = workbook.get_sheet(0)  # Seleccionar la hoja para escribir
    except FileNotFoundError:
        # Crear un nuevo libro de trabajo si no existe
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Factura")
        row_start = 0

    # Agregar cabeceras y datos
    MAIN_HEADER = ["Razon Social", "Nombre Comercial", "RUC"]
    BILL_HEADER = ["Num. Factura", "Fecha Emision", "Total", "Total IVA"]
    DETAILS_HEADER = [
        "Descripción",
        "Cantidad",
        "Precio Unitario",
        "Precio Total sin IVA",
    ]

    # Si hay datos existentes, deja dos líneas en blanco antes de agregar nuevos datos
    if row_start > 0:
        row_start += 2

    # Escribir la cabecera MAIN_HEADER
    for col, header in enumerate(MAIN_HEADER):
        sheet.write(row_start, col, header)
    # Escribir los datos debajo de MAIN_HEADER
    for col, value in enumerate(main_data):
        sheet.write(row_start + 1, col, value)

    # Escribir la cabecera BILL_HEADER
    for col, header in enumerate(BILL_HEADER):
        sheet.write(row_start + 3, col, header)
    # Escribir los datos debajo de BILL_HEADER
    for col, value in enumerate(bill_data):
        sheet.write(row_start + 4, col, value)

    # Escribir la cabecera DETAILS_HEADER
    for col, header in enumerate(DETAILS_HEADER):
        sheet.write(row_start + 6, col, header)
    # Escribir los detalles debajo de DETAILS_HEADER
    for row, detalle in enumerate(details_data, start=row_start + 7):
        sheet.write(row, 0, detalle.get("descripcion", ""))
        sheet.write(row, 1, detalle.get("cantidad", ""))
        sheet.write(row, 2, detalle.get("precioUnitario", ""))
        sheet.write(row, 3, detalle.get("precioTotalSinImpuesto", ""))

    # Guardar el archivo de Excel
    workbook.save(file_path)
    print(f"Archivo '{file_path}' actualizado exitosamente.")


def process_xml(xml_content):
    """Función para procesar el contenido XML y escribir en Excel."""

    # Parsear el XML
    root = ET.fromstring(xml_content)

    razonSocial = parse_xml(root, "comprobante", "razonSocial")
    nombreComercial = parse_xml(root, "comprobante", "nombreComercial")
    ruc = parse_xml(root, "comprobante", "ruc")

    numFactura_parts = parse_xml(root, "comprobante", ["estab", "ptoEmi", "secuencial"])
    numFactura = "".join(numFactura_parts)
    detalles = extract_child_tags(root, "comprobante", ["detalles", "detalle"])
    fechaEmision = extract_block(root, "comprobante", "infoFactura", "fechaEmision")
    total = extract_block(root, "comprobante", "infoFactura", "totalSinImpuestos")
    totalIVA = extract_block(root, "comprobante", "infoFactura", "importeTotal")

    main_data = [razonSocial, nombreComercial, ruc]
    bill_data = [numFactura, fechaEmision, total, totalIVA]
    details_data = detalles

    # Escribir los datos en el archivo Excel
    write_to_excel(XLS_FILE, main_data, bill_data, details_data)


if __name__ == "__main__":
    while True:
        check_for_new_emails()
        time.sleep(CHECK_INTERVAL)
