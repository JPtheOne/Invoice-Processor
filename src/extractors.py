"""
Module containing functions to parse different CFDI types
and export the extracted data to Excel.
"""

import xml.etree.ElementTree as ET
from openpyxl import Workbook, load_workbook

# -------------------------
# Global XML namespaces
# -------------------------
NAMESPACES = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'pago20': 'http://www.sat.gob.mx/Pagos20',
    'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital',
    'nomina12': 'http://www.sat.gob.mx/nomina12'
}


# -------------------------
# Utility functions
# -------------------------
def strip_namespace(tag):
    """Removes the namespace from an XML tag."""
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def find_all_tags(root, tag_name):
    """Finds all elements with a specific tag name, ignoring namespaces."""
    return [elem for elem in root.iter() if strip_namespace(elem.tag) == tag_name]


def get_or_create_sheet(wb, sheet_name, headers):
    """
    Returns an Excel sheet, creating it with headers if it doesn't exist.

    Args:
        wb (Workbook): OpenPyXL workbook object.
        sheet_name (str): Name of the sheet.
        headers (list): List of column names for the first row.

    Returns:
        Worksheet: The sheet ready for writing.
    """
    if sheet_name not in wb.sheetnames:
        sheet = wb.create_sheet(sheet_name)
        sheet.append(headers)
    else:
        sheet = wb[sheet_name]
    return sheet


# -------------------------
# Parsing and saving Pago CFDI
# -------------------------
def parse_P(xml_file):
    """
    Parses a CFDI of type 'Pago' and extracts relevant information.
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    emisor = root.find('cfdi:Emisor', NAMESPACES).attrib
    receptor = root.find('cfdi:Receptor', NAMESPACES).attrib
    timbre = root.find('cfdi:Complemento/tfd:TimbreFiscalDigital', NAMESPACES).attrib

    pagos = []
    for pago in root.findall('cfdi:Complemento/pago20:Pagos/pago20:Pago', NAMESPACES):
        doctos_relacionados = []
        for docto in pago.findall('pago20:DoctoRelacionado', NAMESPACES):
            doctos_relacionados.append({
                'IdDocumento': docto.attrib.get('IdDocumento'),
                'Serie': docto.attrib.get('Serie'),
                'Folio': docto.attrib.get('Folio'),
                'MonedaDR': docto.attrib.get('MonedaDR'),
                'EquivalenciaDR': docto.attrib.get('EquivalenciaDR'),
                'NumParcialidad': docto.attrib.get('NumParcialidad'),
                'ImpSaldoAnt': docto.attrib.get('ImpSaldoAnt'),
                'ImpPagado': docto.attrib.get('ImpPagado'),
                'ImpSaldoInsoluto': docto.attrib.get('ImpSaldoInsoluto'),
                'ObjetoImpDR': docto.attrib.get('ObjetoImpDR')
            })
        pagos.append({
            'FechaPago': pago.attrib.get('FechaPago'),
            'FormaDePagoP': pago.attrib.get('FormaDePagoP'),
            'MonedaP': pago.attrib.get('MonedaP'),
            'TipoCambioP': pago.attrib.get('TipoCambioP'),
            'Monto': pago.attrib.get('Monto'),
            'DoctosRelacionados': doctos_relacionados
        })

    return {
        'Emisor': emisor,
        'Receptor': receptor,
        'TimbreFiscal': timbre,
        'Pagos': pagos
    }


def writeP_to_excel(data, output_file):
    """
    Writes Pago CFDI data to Excel.
    """
    try:
        wb = load_workbook(output_file)
    except FileNotFoundError:
        wb = Workbook()

    headers = [
        "UUID Timbre", "Fecha Timbrado",
        "RFC Emisor", "Nombre Emisor", "Régimen Fiscal Emisor",
        "RFC Receptor", "Nombre Receptor",
        "Fecha Pago", "Forma De Pago P", "Moneda P", "Tipo Cambio P", "Monto",
        "Id Documento", "Serie", "Folio", "Moneda DR", "Equivalencia DR",
        "Num Parcialidad", "Imp Saldo Ant", "Imp Pagado", "Imp Saldo Insoluto", "Objeto Imp DR"
    ]
    sheet = get_or_create_sheet(wb, "Pagos", headers)

    for pago in data['Pagos']:
        for docto in pago['DoctosRelacionados']:
            sheet.append([
                data['TimbreFiscal']['UUID'], data['TimbreFiscal']['FechaTimbrado'],
                data['Emisor'].get('Rfc'), data['Emisor'].get('Nombre'), data['Emisor'].get('RegimenFiscal'),
                data['Receptor'].get('Rfc'), data['Receptor'].get('Nombre'),
                pago.get('FechaPago'), pago.get('FormaDePagoP'), pago.get('MonedaP'),
                pago.get('TipoCambioP'), pago.get('Monto'),
                docto.get('IdDocumento'), docto.get('Serie'), docto.get('Folio'),
                docto.get('MonedaDR'), docto.get('EquivalenciaDR'),
                docto.get('NumParcialidad'), docto.get('ImpSaldoAnt'),
                docto.get('ImpPagado'), docto.get('ImpSaldoInsoluto'), docto.get('ObjetoImpDR')
            ])

    wb.save(output_file)


# -------------------------
# Parsing and saving Ingreso/Egreso CFDI
# -------------------------
def parse_IE(xml_file):
    """
    Parses a CFDI of type 'Ingreso' or 'Egreso' and extracts relevant data.
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    comprobante = {k: root.attrib.get(k) for k in [
        'Version', 'Serie', 'Folio', 'Fecha', 'SubTotal', 'Total', 'FormaPago', 'TipoDeComprobante', 'Moneda'
    ]}

    emisor = root.find('cfdi:Emisor', NAMESPACES).attrib
    receptor = root.find('cfdi:Receptor', NAMESPACES).attrib

    conceptos = []
    for concepto in root.findall('cfdi:Conceptos/cfdi:Concepto', NAMESPACES):
        concepto_data = {
            'Descripcion': concepto.attrib.get('Descripcion'),
            'Cantidad': concepto.attrib.get('Cantidad'),
            'ValorUnitario': concepto.attrib.get('ValorUnitario'),
            'Importe': concepto.attrib.get('Importe'),
        }
        traslado = concepto.find('cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado', NAMESPACES)
        if traslado is not None:
            concepto_data['Traslado_Base'] = traslado.attrib.get('Base')
            concepto_data['Traslado_Importe'] = traslado.attrib.get('Importe')
        conceptos.append(concepto_data)

    complemento = root.find('cfdi:Complemento/tfd:TimbreFiscalDigital', NAMESPACES)
    timbre_fiscal = {
        'UUID': complemento.attrib.get('UUID'),
        'FechaTimbrado': complemento.attrib.get('FechaTimbrado')
    }

    return {
        'Comprobante': comprobante,
        'Emisor': emisor,
        'Receptor': receptor,
        'Conceptos': conceptos,
        'TimbreFiscal': timbre_fiscal
    }


def saveIE_to_excel(data, output_file):
    """
    Writes Ingreso/Egreso CFDI data to Excel.
    """
    try:
        wb = load_workbook(output_file)
    except FileNotFoundError:
        wb = Workbook()

    sheet_name = "Ingresos" if data['Comprobante']['TipoDeComprobante'] == 'I' else "Egresos"
    headers = [
        "UUID", "Fecha", "Serie", "Folio", "Tipo",
        "RFC Emisor", "Nombre Emisor", "Régimen Fiscal",
        "Cantidad", "Valor Unitario", "Importe", "Traslado Importe",
        "Subtotal", "Total", "Forma de Pago", "Descripción",
        "Moneda", "Uso CFDI",
        "RFC Receptor", "Nombre Receptor", "Domicilio", "Régimen Fiscal",
        "Traslado Base", "Fecha Timbrado", "Versión"
    ]
    sheet = get_or_create_sheet(wb, sheet_name, headers)

    for concepto in data['Conceptos']:
        sheet.append([
            data['TimbreFiscal']['UUID'], data['Comprobante']['Fecha'],
            data['Comprobante']['Serie'], data['Comprobante']['Folio'],
            data['Comprobante']['TipoDeComprobante'],
            data['Emisor'].get('Rfc'), data['Emisor'].get('Nombre'), data['Emisor'].get('RegimenFiscal'),
            concepto.get('Cantidad'), concepto.get('ValorUnitario'), concepto.get('Importe'), concepto.get('Traslado_Importe'),
            data['Comprobante']['SubTotal'], data['Comprobante']['Total'], data['Comprobante']['FormaPago'],
            concepto.get('Descripcion'), data['Comprobante']['Moneda'],
            data['Receptor'].get('UsoCFDI'),
            data['Receptor'].get('Rfc'), data['Receptor'].get('Nombre'), data['Receptor'].get('DomicilioFiscalReceptor'),
            data['Receptor'].get('RegimenFiscalReceptor'), concepto.get('Traslado_Base'),
            data['TimbreFiscal']['FechaTimbrado'], data['Comprobante']['Version']
        ])

    wb.save(output_file)


# -------------------------
# Parsing and saving Nómina CFDI
# -------------------------
def parse_N(xml_file):
    """
    Parses a CFDI of type 'Nómina' and extracts relevant information.
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    comprobante = {
        'Serie': root.attrib.get('Serie', 'N/A'),
        'Folio': root.attrib.get('Folio', 'N/A'),
        'Fecha': root.attrib.get('Fecha', 'N/A'),
        'Moneda': root.attrib.get('Moneda', 'N/A'),
        'SubTotal': root.attrib.get('SubTotal', '0.00'),
        'Descuento': root.attrib.get('Descuento', '0.00'),
        'Total': root.attrib.get('Total', '0.00'),
    }

    complemento = root.find('cfdi:Complemento/tfd:TimbreFiscalDigital', NAMESPACES)
    timbre_fiscal = {
        'UUID': complemento.attrib.get('UUID', 'N/A'),
        'FechaTimbrado': complemento.attrib.get('FechaTimbrado', 'N/A')
    }

    emisor = root.find('cfdi:Emisor', NAMESPACES).attrib
    receptor = root.find('cfdi:Receptor', NAMESPACES).attrib

    conceptos = []
    for concepto in root.findall('cfdi:Conceptos/cfdi:Concepto', NAMESPACES):
        conceptos.append({
            'Descripcion': concepto.attrib.get('Descripcion', 'N/A'),
            'Cantidad': concepto.attrib.get('Cantidad', '0'),
            'ValorUnitario': concepto.attrib.get('ValorUnitario', '0.00'),
            'Importe': concepto.attrib.get('Importe', '0.00'),
        })

    complemento_nomina = root.find('cfdi:Complemento/nomina12:Nomina', NAMESPACES)
    nomina = {
        'Version': complemento_nomina.attrib.get('Version', 'N/A') if complemento_nomina is not None else 'N/A',
        'TipoNomina': complemento_nomina.attrib.get('TipoNomina', 'N/A') if complemento_nomina is not None else 'N/A',
        'TotalPercepciones': complemento_nomina.attrib.get('TotalPercepciones', '0.00') if complemento_nomina is not None else '0.00',
        'TotalDeducciones': complemento_nomina.attrib.get('TotalDeducciones', '0.00') if complemento_nomina is not None else '0.00',
        'TotalOtrosPagos': complemento_nomina.attrib.get('TotalOtrosPagos', '0.00') if complemento_nomina is not None else '0.00'
    }

    percepciones, deducciones, otros_pagos = [], [], []
    if complemento_nomina is not None:
        percepciones_elem = complemento_nomina.find('nomina12:Percepciones', NAMESPACES)
        if percepciones_elem is not None:
            for percepcion in percepciones_elem.findall('nomina12:Percepcion', NAMESPACES):
                percepciones.append({
                    'Clave': percepcion.attrib.get('Clave', 'N/A'),
                    'Concepto': percepcion.attrib.get('Concepto', 'N/A'),
                    'ImporteGravado': percepcion.attrib.get('ImporteGravado', '0.00'),
                    'ImporteExento': percepcion.attrib.get('ImporteExento', '0.00')
                })

        deducciones_elem = complemento_nomina.find('nomina12:Deducciones', NAMESPACES)
        if deducciones_elem is not None:
            for deduccion in deducciones_elem.findall('nomina12:Deduccion', NAMESPACES):
                deducciones.append({
                    'Clave': deduccion.attrib.get('Clave', 'N/A'),
                    'Concepto': deduccion.attrib.get('Concepto', 'N/A'),
                    'Importe': deduccion.attrib.get('Importe', '0.00')
                })

        otros_pagos_elem = complemento_nomina.find('nomina12:OtrosPagos', NAMESPACES)
        if otros_pagos_elem is not None:
            for otro_pago in otros_pagos_elem.findall('nomina12:OtroPago', NAMESPACES):
                otros_pagos.append({
                    'Clave': otro_pago.attrib.get('Clave', 'N/A'),
                    'Concepto': otro_pago.attrib.get('Concepto', 'N/A'),
                    'Importe': otro_pago.attrib.get('Importe', '0.00')
                })

    return {
        'TimbreFiscal': timbre_fiscal,
        'Comprobante': comprobante,
        'Emisor': emisor,
        'Receptor': receptor,
        'Conceptos': conceptos,
        'Nomina': nomina,
        'Percepciones': percepciones,
        'Deducciones': deducciones,
        'OtrosPagos': otros_pagos
    }


def saveN_to_excel(data, output_file):
    """
    Writes Nómina CFDI data to Excel.
    """
    try:
        wb = load_workbook(output_file)
    except FileNotFoundError:
        wb = Workbook()

    headers = [
        "UUID", "Fecha Timbrado",
        "Serie", "Folio", "Fecha", "Moneda", "SubTotal", "Descuento", "Total",
        "RFC Emisor", "Nombre Emisor",
        "RFC Receptor", "Nombre Receptor",
        "Descripcion", "Cantidad", "Valor Unitario", "Importe",
        "Version Nómina", "Tipo Nómina", "Total Percepciones", "Total Deducciones", "Total Otros Pagos"
    ]
    sheet = get_or_create_sheet(wb, "Nómina", headers)

    for concepto in data['Conceptos']:
        sheet.append([
            data['TimbreFiscal']['UUID'], data['TimbreFiscal']['FechaTimbrado'],
            data['Comprobante']['Serie'], data['Comprobante']['Folio'], data['Comprobante']['Fecha'],
            data['Comprobante']['Moneda'], data['Comprobante']['SubTotal'], data['Comprobante']['Descuento'], data['Comprobante']['Total'],
            data['Emisor'].get('Rfc', 'N/A'), data['Emisor'].get('Nombre', 'N/A'),
            data['Receptor'].get('Rfc', 'N/A'), data['Receptor'].get('Nombre', 'N/A'),
            concepto.get('Descripcion', 'N/A'), concepto.get('Cantidad', '0'), concepto.get('ValorUnitario', '0.00'), concepto.get('Importe', '0.00'),
            data['Nomina']['Version'], data['Nomina']['TipoNomina'], data['Nomina']['TotalPercepciones'], data['Nomina']['TotalDeducciones'], data['Nomina']['TotalOtrosPagos']
        ])

    wb.save(output_file)
