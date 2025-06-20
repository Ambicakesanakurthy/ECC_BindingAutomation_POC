import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import os

st.set_page_config(page_title="TGML Binder Tool", layout="centered")

st.markdown("""
    <style>
    body {
        background-color: #2980b9
    .title {
        font-size: 30px;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 20px;
    }
 
    .subtitle {
        font-size: 20px;
        color: #4b4b4b;
        text-align: center;
        margin-bottom: 40px;
    }
 
    .block-container {
        background-color: #f9f9f9;
        padding: 30px;
        border-radius: 10px;
    }
 
    .css-1aumxhk {  /* Applies to main content box */
        background-color: #ffffff;
    }
 
    .stButton > button {
        background-color: #1f77b4;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 40px;
        width: 200px;
    }
 
    .stButton > button:hover {
        background-color: #0070AD;
    }

    h1 {
    margin-bottom: 10px;
    color: #114488;
    font-size: 24px;
    text-align: center;
}
 
    </style>
""", unsafe_allow_html=True)
 
st.markdown('<div class="title"> TGML Auto-Binder Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your TGML and Excel file below to start binding</div>', unsafe_allow_html=True)

 
 
#st.title("TGML Automatic Binder")
#st.markdown("Upload your TGML file, Excel file, and select the sheet name to bind automatically.")
 
tgml_file = st.file_uploader("Upload TGML (.tgml) file", type="tgml")
excel_file = st.file_uploader("Upload Excel (.xlsx) file", type="xlsx")
sheet_name = st.text_input("Enter Sheet Name (e.g., DDC-1, DDC37)", "")
 
if st.button("Run Binding") and tgml_file and excel_file and sheet_name:
    try:
        # Parse TGML
        tree = ET.parse(tgml_file)
        root = tree.getroot()
 
        # Load Excel
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
 
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
                    elem.set("Name", new_bind)
            elif elem.tag == "Text" and inside_target_text:
                inside_target_text = False
 
        # Save to output file
        output_file_path = "output.tgml"
        tree.write(output_file_path, encoding="utf-8", xml_declaration=True)
 
        with open(output_file_path, "rb") as f:
            st.download_button(
                label="Download Updated TGML File",
                data=f,
file_name="updated_" + tgml_file.name,
                mime="application/xml"
            )
 
        st.success("Binding completed successfully!")
 
    except Exception as e:
        st.error(f"Error: {e}")
