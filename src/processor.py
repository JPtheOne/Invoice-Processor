import os
from extractors import *  # Ensure these modules are implemented and imported correctly
from identifier import *  # Ensure these modules are implemented and imported correctly
from zipfile import ZipFile

def process_cfdi(cfdi_filename, output_filename, counters):
    print(f"Procesando CFDI: {cfdi_filename}")  # Debug
    cfdi_type = determine_xml_type(cfdi_filename)
    print(f"Tipo de CFDI detectado: {cfdi_type}")  # Debug
    if cfdi_type == "I" or cfdi_type == "E":
        counters["I/E"] += 1
        extracted_IE = parse_IE(cfdi_filename)
        print(f"Datos extraídos (IE): {extracted_IE}")  # Debug
        saveIE_to_excel(extracted_IE, output_filename)
    elif cfdi_type == "P":
        counters["P"] += 1
        extracted_P = parse_P(cfdi_filename)
        print(f"Datos extraídos (P): {extracted_P}")  # Debug
        writeP_to_excel(extracted_P, output_filename)
    elif cfdi_type == "N":
        counters["N"] += 1
        extracted_N = parse_N(cfdi_filename)
        print(f"Datos extraídos (N): {extracted_N}")  # Debug
        saveN_to_excel(extracted_N, output_filename)
    else:
        counters["Desconocido"] += 1
        print(f"Tipo de CFDI desconocido: {cfdi_type}")


def unzip_folder(origin_zip_filename, destination_folder):
    with ZipFile(origin_zip_filename, 'r') as zip_ref:
        zip_ref.extractall(destination_folder)
    # Return the list of extracted files
    return [os.path.join(destination_folder, f) for f in os.listdir(destination_folder)]


def main():
    zips_folder = "./test"  # Folder containing zip files
    unzipped_folder = "./unzipped"  # Folder to extract files into
    output_filename = "./Excel_final.xlsx"

    # Counters for XML files and types
    counters = {
        "Total": 0,
        "I/E": 0,
        "P": 0,
        "N": 0,
        "Unknown": 0
    }
    
    # Ensure unzipped folder exists
    os.makedirs(unzipped_folder, exist_ok=True)

    # Process each zip file in the folder
    for zip_file in os.listdir(zips_folder):
        zip_path = os.path.join(zips_folder, zip_file)
        if zip_file.endswith(".zip"):
            print(f"Processing zip file: {zip_file}")
            extracted_files = unzip_folder(zip_path, unzipped_folder)

            # Process each extracted file
            for cfdi_file in extracted_files:
                if cfdi_file.endswith(".xml"):  # Only process XML files
                    counters["Total"] += 1
                    print(f"Processing CFDI file: {cfdi_file}")
                    process_cfdi(cfdi_file, output_filename, counters)

    # Summary of processing
    print("\nProcessing Summary:")
    print(f"Total XML files processed: {counters['Total']}")
    print(f" - I/E (Ingreso/Egreso): {counters['I/E']}")
    print(f" - P (Pago): {counters['P']}")
    print(f" - N (Nómina): {counters['N']}")
    print(f" - Desconocido: {counters['Desconocido']}")

if __name__ == "__main__":
    main()
