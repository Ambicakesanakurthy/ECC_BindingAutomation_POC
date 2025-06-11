"""
import xml.etree.ElementTree as ET
import sys

def replace_all_binds_for_text(xml_file, text_name, new_bind_name, output_file=None):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    in_group = False
    inside_target_text = False

    for elem in root.iter():
        if elem.tag == "Group":
            in_group = True

        elif elem.tag == "Text" and in_group:
            current_text = elem.attrib.get("Name", "")
            inside_target_text = (text_name in current_text)

        elif elem.tag == "Bind" and in_group and inside_target_text:
            old_bind = elem.attrib.get("Name", "")
            print(f"Replacing Bind '{old_bind}' with '{new_bind_name}' for Text containing '{text_name}'")
            elem.set("Name", new_bind_name)

        elif elem.tag == "Text" and inside_target_text:
            # Exiting this Text element scope
            inside_target_text = False

    # Save output
    output = output_file if output_file else xml_file
    tree.write(output, encoding="utf-8", xml_declaration=True)
    print(f"\n Updated XML saved to: {output}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python replace_all_binds.py <xml_file> <text_name_substring> <new_bind_name>")
    else:
        xml_path = sys.argv[1]
        text_to_find = sys.argv[2]
        new_bind = sys.argv[3]
        replace_all_binds_for_text(xml_path, text_to_find, new_bind)

-----------------------------------------------------------------------------------
import xml.etree.ElementTree as ET
import pandas as pd
import sys
 
def update_binds_from_excel(xml_file, excel_file, output_file=None):
    # Load XML
    tree = ET.parse(xml_file)
    root = tree.getroot()
 
    # Load Excel
    df = pd.read_excel(excel_file)
    
    # Flatten Excel to label -> nomenclature pairs
    label_to_bind = {}
    for _, row in df.iterrows():
        nomenclature = str(row.get("Nomenclature", "")).strip()
        for col in ["First Label", "Second Label", "Third Label"]:
            label = str(row.get(col, "")).strip()
            if label:  # Skip empty labels
                label_to_bind[label] = nomenclature
 
    in_group = False
    current_text = None
    inside_target_text = False
 
    for elem in root.iter():
        if elem.tag == "Group":
            in_group = True
 
        elif elem.tag == "Text" and in_group:
            current_text = elem.attrib.get("Name", "").strip()
            inside_target_text = current_text in label_to_bind
 
        elif elem.tag == "Bind" and in_group and inside_target_text:
            new_bind = label_to_bind.get(current_text)
            old_bind = elem.attrib.get("Name", "")
            if new_bind:
                print(f" Replacing Bind '{old_bind}' with '{new_bind}' for Text '{current_text}'")
                elem.set("Name", new_bind)
 
        elif elem.tag == "Text" and inside_target_text:
            inside_target_text = False
 
    # Save updated XML
    output = output_file if output_file else xml_file
    tree.write(output, encoding="utf-8", xml_declaration=True)
    print(f"\n Updated XML saved to: {output}")
 
 
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python replace_bind_from_excel.py <tgml_file> <excel_file>")
    else:
        tgml_path = sys.argv[1]
        excel_path = sys.argv[2]
        update_binds_from_excel(tgml_path, excel_path)
 
"""
import xml.etree.ElementTree as ET
import pandas as pd
import sys
import os
 
def update_binds_from_excel_sheet(xml_file, excel_file, sheet_name, output_file=None):
    # Load XML
    try:
        tree = ET.parse(xml_file)
    except Exception as e:
        print(f"Error reading XML file '{xml_file}': {e}")
        return
    root = tree.getroot()
 
    # Load Excel sheet
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading Excel file '{excel_file}', sheet '{sheet_name}': {e}")
        return
 
    # Build label -> bind mapping
    label_to_bind = {}
    for _, row in df.iterrows():
        nomenclature = str(row.get("Nomenclature", "")).strip()
        for col in ["First Label", "Second Label", "Third Label"]:
            label = str(row.get(col, "")).strip()
            if label:
                label_to_bind[label] = nomenclature
 
    in_group = False
    current_text = None
    inside_target_text = False
 
    #Checking the label and replacing the bindname
    for elem in root.iter():
        if elem.tag == "Group":
            in_group = True
 
        elif elem.tag == "Text" and in_group:
            current_text = elem.attrib.get("Name", "").strip()
            inside_target_text = current_text in label_to_bind
 
        elif elem.tag == "Bind" and in_group and inside_target_text:
            new_bind = label_to_bind.get(current_text)
            old_bind = elem.attrib.get("Name", "")
            if new_bind:
                print(f"Replacing Bind '{old_bind}' with '{new_bind}' for Text '{current_text}'")
                elem.set("Name", new_bind)
 
        elif elem.tag == "Text" and inside_target_text:
            inside_target_text = False
 
    # Save updated XML
    output = output_file if output_file else xml_file
    tree.write(output, encoding="utf-8", xml_declaration=True)
    print(f"\nUpdated XML saved to: {output}")
 
 
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage:\n python replace_bind_multiple_ddc.py <tgml_file> <excel_file> <sheet_name>")
        print("Example:\n python replace_bind_multiple_ddc.py DDC37.tgml DDC-All.xlsx DDC37")
    else:
        tgml_path = sys.argv[1]
        excel_path = sys.argv[2]
        sheet = sys.argv[3]
        update_binds_from_excel_sheet(tgml_path, excel_path, sheet)