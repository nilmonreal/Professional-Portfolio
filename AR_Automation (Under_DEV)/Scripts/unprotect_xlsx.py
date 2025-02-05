import os
import zipfile
from lxml import etree
from io import BytesIO

def remove_excel_protection(input_path, output_path):
    # Create a temporary directory
    temp_dir = "temp_excel"
    os.makedirs(temp_dir, exist_ok=True)

    # Unzip the Excel file
    with zipfile.ZipFile(input_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Remove sheet protection from all worksheets
    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.startswith("sheet") and file.endswith(".xml"):
                sheet_path = os.path.join(root, file)
                tree = etree.parse(sheet_path)
                root_xml = tree.getroot()
                
                # Find and remove the <sheetProtection> tag
                for elem in root_xml.getiterator():
                    if etree.QName(elem).localname == 'sheetProtection':
                        root_xml.remove(elem)
                        break
                
                # Save the modified XML
                tree.write(sheet_path, encoding="UTF-8", xml_declaration=True)

    # Repackage the files into a new Excel file
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                zip_out.write(file_path, os.path.relpath(file_path, temp_dir))

    # Clean up temporary files
    import shutil
    shutil.rmtree(temp_dir)

# Usage
input_file = r"C:\Users\nil.monreal\Desktop\7. Cierre de mes AR\5.- EZA1 Accounts Receivable Itercompany January 2025\1. Embraer\1.- EZA1 Accounts Receivable Intercompany JANUARY - General 2025.xlsx"
output_file = r"C:\Users\nil.monreal\Desktop\7. Cierre de mes AR\5.- EZA1 Accounts Receivable Itercompany January 2025\1. Embraer\1.- EZA1 Accounts Receivable Intercompany JANUARY - General 2025.xlsx"
remove_excel_protection(input_file, output_file)