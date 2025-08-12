import os
from zipfile import ZipFile
from identifier import determine_xml_type
from extractors import (parse_IE, parse_P, parse_N,saveIE_to_excel, writeP_to_excel, saveN_to_excel)


def process_cfdi(cfdi_filename, output_filename, counters):
    """
    Processes a single CFDI XML file based on its type and updates counters.

    Args:
        cfdi_filename (str): Path to the CFDI XML file.
        output_filename (str): Path to the Excel file where data will be saved.
        counters (dict): Dictionary tracking totals for each CFDI type.

    Returns:
        None
    """
    print(f"Processing CFDI: {cfdi_filename}")

    cfdi_type = determine_xml_type(cfdi_filename)
    print(f"Detected CFDI type: {cfdi_type}")

    # Mapping CFDI types to parsing and saving functions
    type_actions = {
        "I": (parse_IE, saveIE_to_excel, "I/E"),
        "E": (parse_IE, saveIE_to_excel, "I/E"),
        "P": (parse_P, writeP_to_excel, "P"),
        "N": (parse_N, saveN_to_excel, "N")
    }

    if cfdi_type in type_actions:
        parse_func, save_func, counter_key = type_actions[cfdi_type]
        counters[counter_key] += 1
        extracted_data = parse_func(cfdi_filename)
        print(f"Extracted data ({cfdi_type}): {extracted_data}")
        save_func(extracted_data, output_filename)
    else:
        counters["Desconocido"] += 1
        print(f"Unknown CFDI type: {cfdi_type}")


def unzip_folder(origin_zip_filename, destination_folder):
    """
    Extracts all files from a ZIP archive to a destination folder.

    Args:
        origin_zip_filename (str): Path to the .zip file.
        destination_folder (str): Path where files will be extracted.

    Returns:
        list: List of paths to the extracted files.
    """
    with ZipFile(origin_zip_filename, 'r') as zip_ref:
        zip_ref.extractall(destination_folder)

    return [
        os.path.join(destination_folder, f)
        for f in os.listdir(destination_folder)
    ]


def main():
    """
    Standalone execution for processing all ZIPs in the test folder.
    Uses a temporary folder for extraction to avoid leaving artifacts.
    """
    import tempfile

    zips_folder = "./test"
    output_filename = "./Excel_final.xlsx"
    counters = {"Total": 0, "I/E": 0, "P": 0, "N": 0, "Unknown": 0}

    with tempfile.TemporaryDirectory() as unzipped_folder:
        for zip_file in os.listdir(zips_folder):
            zip_path = os.path.join(zips_folder, zip_file)
            if zip_file.lower().endswith(".zip"):
                print(f"Processing zip file: {zip_file}")
                extracted_files = unzip_folder(zip_path, unzipped_folder)

                for cfdi_file in extracted_files:
                    if cfdi_file.lower().endswith(".xml"):
                        counters["Total"] += 1
                        print(f"Processing CFDI file: {cfdi_file}")
                        process_cfdi(cfdi_file, output_filename, counters)

    print("\nProcessing Summary:")
    print(f"Total XML files processed: {counters['Total']}")
    print(f" - I/E (Ingreso/Egreso): {counters['I/E']}")
    print(f" - P (Pago): {counters['P']}")
    print(f" - N (Nómina): {counters['N']}")
    print(f" - Unknown: {counters['Unknown']}")
