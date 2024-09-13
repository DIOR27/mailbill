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
CHECK_INTERVAL = 5  # Intervalo de chequeo en segundos
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
        sender = msg["from"]
        print(f"Nuevo correo: {subject} de {sender}")

        if msg.is_multipart():
            for part in msg.walk():
                xml_content = None
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename and filename.endswith(".xml"):
                        print(f"Analizando archivo adjunto: {filename}")
                        xml_content = part.get_payload(decode=True).decode("utf-8")

                        try:
                            process_xml(xml_content)
                        except Exception as e:
                            print(f"Error procesando XML: {e}")

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


def extract_child_tags(root, parent_tag, child_tags, main_data=None):
    factura_xml_cdata = root.find(parent_tag).text
    factura_xml = re.sub(r"<!\[CDATA\[|\]\]>", "", factura_xml_cdata).strip()
    try:
        factura_root = ET.fromstring(factura_xml)
    except ET.ParseError as e:
        print(f"Error parsing factura XML: {e}")
        return []

    # Encontrar la sección <detalles>
    detalles = factura_root.find(child_tags[0])

    detalles_list = []

    # Iterar sobre cada <detalle>
    for detalle in detalles.findall(child_tags[1]):
        # Lista temporal para almacenar cada detalle con los datos principales
        detalle_data = main_data.copy() if main_data else []

        # Extraer campos específicos de cada detalle
        valor = 0
        for field in detalle:
            if field.tag in [
                "descripcion",
                "cantidad",
                "precioUnitario",
                "precioTotalSinImpuesto",
            ]:
                detalle_data.append(field.text)
            if field.tag == "precioTotalSinImpuesto":
                valor = float(field.text)

        # Extraer tarifa de impuestos
        impuestos = detalle.find("impuestos")
        if impuestos is not None:
            for impuesto in impuestos.findall("impuesto"):
                tarifa = impuesto.find("tarifa")
                if tarifa is not None:
                    valorIVA = (float(tarifa.text) / 100) * valor + valor
                    detalle_data.append(str(round(valorIVA, 2)))
                    detalle_data.append(tarifa.text)
                else:
                    detalle_data.append("")

        # Agregar la combinación de main_data y los detalles específicos a la lista de detalles
        detalles_list.append(detalle_data)

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


def write_to_excel(file_path, data):
    """Escribe los datos en un archivo Excel, línea por línea."""

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
    HEADERS = [
        "Razon Social",
        "Nombre Comercial",
        "RUC",
        "Num. Factura",
        "Fecha Emision",
        "Total",
        "Total IVA",
        "Descripción",
        "Cantidad",
        "Precio Unitario",
        "Precio Total",
        "Precio Total con IVA",
        "Impuesto",
    ]

    # Definir estilos
    bold_style = xlwt.easyxf("font: bold 1")  # Estilo de negrita

    # Escribir la cabecera MAIN_HEADER solo si es un archivo nuevo
    if row_start == 0:
        for col, header in enumerate(HEADERS):
            sheet.write(row_start, col, header, bold_style)
        row_start += 1  # Mueve el inicio de los datos a la siguiente fila

    # Escribir los datos debajo de MAIN_HEADER
    row = row_start  # Empieza en la fila siguiente a la cabecera

    for entry in data:
        if isinstance(entry[0], list):
            # Si los datos están en formato de lista de listas
            for row_data in entry:
                for col, value in enumerate(row_data):
                    sheet.write(row, col, value)
                row += 1
        else:
            # Si los datos están en una sola lista
            for col, value in enumerate(entry):
                sheet.write(row, col, value)
            row += 1

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
    fechaEmision = extract_block(root, "comprobante", "infoFactura", "fechaEmision")
    total = extract_block(root, "comprobante", "infoFactura", "totalSinImpuestos")
    totalIVA = extract_block(root, "comprobante", "infoFactura", "importeTotal")

    main_data = [
        razonSocial,
        nombreComercial,
        ruc,
        numFactura,
        fechaEmision,
        total,
        totalIVA,
    ]

    details_data = extract_child_tags(
        root, "comprobante", ["detalles", "detalle"], main_data
    )

    # Escribir los datos en el archivo Excel
    write_to_excel(XLS_FILE, details_data)


if __name__ == "__main__":
    while True:
        check_for_new_emails()
        time.sleep(CHECK_INTERVAL)
