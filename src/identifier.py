import xml.etree.ElementTree as ET

def determine_xml_type(xml_file):
    """
    Determines the CFDI type (TipoDeComprobante) from an XML file.

    Args:
        xml_file (str): Path to the XML file.

    Returns:
        str: One of:
             - "I" for Ingreso
             - "E" for Egreso
             - "P" for Pago
             - "N" for Nómina
             - "Unknown" if the attribute is missing
             - Error message if parsing fails
    """
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        tipo_comprobante = root.attrib.get('TipoDeComprobante')
        if tipo_comprobante:
            return tipo_comprobante.strip()
        return "Unknown"

    except ET.ParseError as e:
        return f"Error parsing XML: {e}"

    except FileNotFoundError:
        return "File not found"

    except Exception as e:
        return f"Unexpected error: {e}"
