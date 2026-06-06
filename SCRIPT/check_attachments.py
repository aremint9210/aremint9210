import PyPDF2
import os

def check_attachments(file_path, output_file):
    info = []
    try:
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            
            # Check for embedded files in the catalog
            try:
                embedded_files = reader.trailer["/Root"]["/Names"]["/EmbeddedFiles"]["/Names"]
                info.append(f"Embedded Files count: {len(embedded_files) // 2}")
            except KeyError:
                info.append("No /EmbeddedFiles found in catalog.")

            # Check for annotations (Sound, Screen, RichMedia) on each page
            total_sound_annots = 0
            for i, page in enumerate(reader.pages):
                if "/Annots" in page:
                    for annot in page["/Annots"]:
                        obj = annot.get_object()
                        subtype = obj.get("/Subtype")
                        if subtype in ["/Sound", "/Screen", "/RichMedia"]:
                            total_sound_annots += 1
            info.append(f"Total Sound/Media Annotations: {total_sound_annots}")

    except Exception as e:
        info.append(f"Error: {str(e)}")

    with open(output_file, 'w') as f:
        f.write("\n".join(info))

if __name__ == "__main__":
    pdf_path = r'GENERATE REPORT\EXAMPLE\17. MINI PPU TERLA 33-11kV (VI).pdf'
    output_path = 'attachments_info.txt'
    check_attachments(pdf_path, output_path)
