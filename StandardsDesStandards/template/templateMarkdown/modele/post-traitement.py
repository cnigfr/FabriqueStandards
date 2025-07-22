import zipfile
import shutil
import os
import sys
import tempfile
from lxml import etree
import xml.etree.ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', NS['w'])

def remove_before_fiche(xml_data):
    root = ET.fromstring(xml_data)
    body = root.find('w:body', NS)

    new_body = []
    found = False

    for elem in list(body):
        texts = elem.findall(".//w:t", NS)
        full_text = " ".join(t.text for t in texts if t.text).lower()
        if not found and "fiche descriptive" in full_text:
            found = True
            new_body.append(elem)
        elif found:
            new_body.append(elem)

    body.clear()
    for elem in new_body:
        body.append(elem)

    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def enable_autofit_tables(xml_data):
    tree = etree.fromstring(xml_data)

    for tbl in tree.xpath('//w:tbl', namespaces=NS):
        tblW = tbl.find('w:tblPr/w:tblW', namespaces=NS)
        if tblW is not None:
            tbl.find('w:tblPr', namespaces=NS).remove(tblW)

        tblLayout = tbl.find('w:tblPr/w:tblLayout', namespaces=NS)
        if tblLayout is not None:
            tbl.find('w:tblPr', namespaces=NS).remove(tblLayout)

        tblPr = tbl.find('w:tblPr', namespaces=NS)
        layout = etree.Element('{%s}tblLayout' % NS['w'])
        layout.set('{%s}type' % NS['w'], 'autofit')
        tblPr.append(layout)

    return etree.tostring(tree, xml_declaration=True, encoding='utf-8')

def process_docx(input_path):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        temp_path = tmp.name

    with zipfile.ZipFile(input_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    # Étape 1 : nettoyage
                    data = remove_before_fiche(data)
                    # Étape 2 : ajustement tableau
                    data = enable_autofit_tables(data)
                zout.writestr(item, data)

    shutil.move(temp_path, input_path)
    print(f"✅ Post-traitement terminé : {input_path}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Utilisation : python postprocess_docx.py Document.docx")
    else:
        process_docx(sys.argv[1])
