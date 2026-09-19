import json
from pathlib import Path
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

sys.stdout.reconfigure(encoding="utf-8")

def test():
    print("This is a test function.")

def read_electronics_data():
    # Define the path to the JSON file
    json_file_path = Path("data/electronics_data.json")

    # 1. Gehe vom Skript-Ordner (src) eine Ebene nach oben zum Projekt-Hauptordner
    project_root = Path(__file__).parent.parent
    json_file_path = project_root / "Dateien" / "electronics.json"
     
    # Check if the file exists
    if not json_file_path.exists():
        print(f"File not found: {json_file_path}")

    # Read the JSON data from the file
    with open(json_file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # Print the loaded data (for testing purposes)
    #print(json.dumps(data, indent=4, ensure_ascii=False))
    return data

def extract_field_keys(data):
    extracted_keys = []

    # Überprüfen, ob das Array "fields" existiert und ob es mindestens ein Element enthält
    if "fields" in data and isinstance(data["fields"], list):
        # Liste aller Werte für den Schlüssel "key" erstellen
        extracted_keys = [
            field["key"] for field in data["fields"]
            if isinstance(field, dict) and "key" in field
        ]

        print(f"Erfolgreich {len(extracted_keys)} Keys extrahiert:\n")
        #for key in extracted_keys:
        #    print(f"- {key}")

        return extracted_keys
    else:
     print("Das JSON-Dokument enthält kein gültiges 'fields'-Array.")

#def create_excel_project(data, output_file):
    # Create a new Excel workbook and select the active worksheet
#    wb = Workbook()
#    ws = wb.active
#    ws.title = "Electronics Data"

    # Define header style
#    header_font = Font(bold=True, color="FFFFFF")
#    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    # Write headers to the first row
#    if "fields" in data and isinstance(data["fields"], list) and len(data["fields"]) > 0:
#        headers = list(data["fields"][0].keys())
#        for col_num, header in enumerate(headers, start=1):
#            cell = ws.cell(row=1, column=col_num, value=header)
#            cell.font = header_font
#            cell.fill = header_fill

        # Write data rows
#        for row_num, field in enumerate(data["fields"], start=2):
#            for col_num, header in enumerate(headers, start=1):
#                ws.cell(row=row_num, column=col_num, value=field.get(header))

        # Save the workbook to the specified output file
#        wb.save(output_file)
#        print(f"Excel file created successfully: {output_file}")
#    else:
#        print("No valid 'fields' data found to create Excel.")

def create_excel_project(data):
    # Pfade dynamisch ermitteln
    project_root = Path(__file__).parent.parent
    excel_path = project_root / "Dateien" / "datenpunkte.xlsx"
    
    # Excel-Arbeitsmappe initialisieren
    wb = Workbook()
    ws = wb.active
    ws.title = "Datenpunkte"
    
    # Spaltenüberschrift setzen und formatieren
    header = "Datenpunktbezeichnung"
    ws["A1"] = header
    ws["A1"].font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    #extracted_keys = extract_field_keys(data)

    # Keys in die Zeilen eintragen
    for index, key_value in enumerate(data, start=2):
     cell = ws[f"A{index}"]
     cell.value = key_value
     cell.font = Font(name="Arial", size=11)
     
    # Optimale Spaltenbreite berechnen
    max_len = max(len(str(k)) for k in extracted_keys + [header])
    ws.column_dimensions["A"].width = max_len + 4 
    
    # Excel-Datei abspeichern
    wb.save(excel_path)
    print("Projekt erfolgreich ausgeführt! 🎉")
    print(f"Excel-Datei erstellt unter: {excel_path.absolute()}")

data = read_electronics_data()
extracted_keys = extract_field_keys(data)
for key in extracted_keys:
 print(f"- {key}")
create_excel_project(extracted_keys)
test()
