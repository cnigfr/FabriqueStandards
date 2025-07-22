import zipfile
import shutil
import tempfile
import sys
from lxml import etree

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def enable_autofit_tables(xml_data):
    tree = etree.fromstring(xml_data)

    for tbl in tree.xpath('//w:tbl', namespaces=NS):
        # Supprime la largeur fixe du tableau
        tblW = tbl.find('w:tblPr/w:tblW', namespaces=NS)
        if tblW is not None:
            tbl.find('w:tblPr', namespaces=NS).remove(tblW)

        # Supprime toute disposition fixe
        tblLayout = tbl.find('w:tblPr/w:tblLayout', namespaces=NS)
        if tblLayout is not None:
            tbl.find('w:tblPr', namespaces=NS).remove(tblLayout)

        # (optionnel) Ajoute tblLayout auto (non nécessaire, mais propre)
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
                    data = enable_autofit_tables(data)
                zout.writestr(item, data)

    shutil.move(temp_path, input_path)
    print(f"✅ Ajustement automatique appliqué : {input_path}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Utilisation : python autofit_tables.py Document.docx")
    else:
        process_docx(sys.argv[1])
