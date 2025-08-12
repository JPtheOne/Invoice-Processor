import xml.etree.ElementTree as ET

def determine_xml_type(xml_file):
    # Parse the XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()
    try:
        tipo_comprobante = root.attrib.get('TipoDeComprobante')      
        if tipo_comprobante:
            return tipo_comprobante.strip()
        else:
            return "Unknown"
    except ET.ParseError as e:
        return f"Error parsing XML: {e}"
    except FileNotFoundError:
        return "File not found"




