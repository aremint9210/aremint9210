import zipfile
import os

docx_path = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\TEMPLATE\00. EXECUTIVE SUMMARY.docx"
extract_dir = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\TEMPLATE\extracted"
output_text = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\GENERATE REPORT\TEMPLATE\summary_text.txt"

if not os.path.exists(extract_dir):
    os.makedirs(extract_dir)

with zipfile.ZipFile(docx_path, 'r') as zip_ref:
    zip_ref.extractall(extract_dir)

doc_xml = os.path.join(extract_dir, 'word', 'document.xml')
with open(doc_xml, 'r', encoding='utf-8') as f:
    content = f.read()

with open(output_text, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"Content written to {output_text}")
