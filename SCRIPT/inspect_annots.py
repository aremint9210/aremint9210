import pymupdf
import os

def log(msg, log_file):
    with open(log_file, 'a') as f:
        f.write(msg + '\n')

def inspect_annots(input_pdf, log_file):
    doc = pymupdf.open(input_pdf)
    for page_num in range(len(doc)):
        page = doc[page_num]
        annots = page.annots()
        if annots:
            log(f"Page {page_num + 1} has {len(list(page.annots()))} annotations:", log_file)
            for i, annot in enumerate(page.annots()):
                log(f"  Annot {i}: type={annot.type}, rect={annot.rect}", log_file)
                # Check for Sound type (usually 16 or 17? No, check annot.type[1])
                log(f"    Subtype: {annot.type}", log_file)
    doc.close()

if __name__ == "__main__":
    work_dir = r"GENERATE REPORT\INSERT SOUND"
    pdf_files = [f for f in os.listdir(work_dir) if f.lower().endswith('.pdf')]
    if pdf_files:
        inspect_annots(os.path.join(work_dir, pdf_files[0]), os.path.join(work_dir, "annot_inspect.txt"))
