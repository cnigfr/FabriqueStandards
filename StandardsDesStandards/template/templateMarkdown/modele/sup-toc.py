import zipfile
import shutil
import os
import sys
import tempfile
import xml.etree.ElementTree as ET

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', NAMESPACE['w'])

def remove_before_fiche(xml_data):
    root = ET.fromstring(xml_data)
    body = root.find('w:body', NAMESPACE)

    new_body = []
    found = False

    for elem in list(body):
        texts = elem.findall(".//w:t", NAMESPACE)
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

def overwrite_docx(input_docx):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmpfile:
        temp_path = tmpfile.name

    with zipfile.ZipFile(input_docx, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    cleaned_xml = remove_before_fiche(data)
                    zout.writestr(item, cleaned_xml)
                else:
                    zout.writestr(item, data)

    shutil.move(temp_path, input_docx)
    print(f"✅ Contenu nettoyé → {input_docx}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Utilisation : python remove_before_fiche_overwrite.py Document.docx")
    else:
        overwrite_docx(sys.argv[1])
