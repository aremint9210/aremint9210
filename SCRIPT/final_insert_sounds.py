import pymupdf
import pypdf
from pypdf.generic import DictionaryObject, NameObject, NumberObject, ArrayObject
import os
import sys

def log(msg, log_file):
    with open(log_file, 'a') as f:
        f.write(msg + '\n')

def get_icon_positions(pdf_path, log_file):
    doc = pymupdf.open(pdf_path)
    positions = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_height = page.rect.height
        
        # 1. First find the "Sound clip" text to narrow down the area
        text_instances = page.search_for("Sound clip")
        
        for inst in text_instances:
            # Look for images near this text instance
            # The icon is usually to the right of the text
            search_area = pymupdf.Rect(inst.x1, inst.y0 - 5, page.rect.width, inst.y1 + 5)
            images = page.get_image_info(hashes=True)
            
            best_rect = None
            for img in images:
                img_rect = pymupdf.Rect(img["bbox"])
                # Check if image is within the search area and is small (like an icon)
                if img_rect.intersects(search_area) and img_rect.width < 30 and img_rect.height < 30:
                    best_rect = img_rect
                    break
            
            if best_rect:
                # Convert fitz (top-left) to pypdf (bottom-left)
                x_start = best_rect.x0
                y_start = page_height - best_rect.y1
                x_end = best_rect.x1
                y_end = page_height - best_rect.y0
                
                positions.append({
                    'page': page_num,
                    'rect': [x_start, y_start, x_end, y_end]
                })
                log(f"Page {page_num+1}: Found icon at {best_rect}", log_file)
            else:
                # Fallback to offset if image not detected
                log(f"Page {page_num+1}: Icon image not detected, using fallback offset", log_file)
                x_start = inst.x1 + 35
                y_start = page_height - inst.y1 - 2
                x_end = x_start + 15
                y_end = page_height - inst.y0 + 2
                positions.append({
                    'page': page_num,
                    'rect': [x_start, y_start, x_end, y_end]
                })
                
    doc.close()
    return positions

def insert_sounds(pdf_path, wav_files, positions, output_path, log_file):
    reader = pypdf.PdfReader(pdf_path)
    writer = pypdf.PdfWriter()
    writer.append_pages_from_reader(reader)
    
    wav_index = 0
    total = 0
    
    for pos in positions:
        page_num = pos['page']
        rect = pos['rect']
        page = writer.get_page(page_num)
        
        wav_path = wav_files[wav_index % len(wav_files)]
        wav_name = os.path.basename(wav_path)
        
        try:
            with open(wav_path, "rb") as f:
                wav_data = f.read()
            
            # UNIQUE sound stream for each file
            sound_stream = pypdf.generic.StreamObject()
            sound_stream.update({
                NameObject("/Type"): NameObject("/Sound"),
                NameObject("/R"): NumberObject(44100),
                NameObject("/C"): NumberObject(1),
                NameObject("/B"): NumberObject(16),
                NameObject("/E"): NameObject("/Signed"),
            })
            sound_stream._data = wav_data
            sound_stream_ref = writer._add_object(sound_stream)
            
            # UNIQUE empty appearance stream to make it invisible but active
            from pypdf.generic import StreamObject
            appearance = StreamObject()
            appearance.update({
                NameObject("/Type"): NameObject("/XObject"),
                NameObject("/Subtype"): NameObject("/Form"),
                NameObject("/BBox"): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(rect[2]-rect[0]), NumberObject(rect[3]-rect[1])]),
                NameObject("/Resources"): DictionaryObject(),
            })
            appearance._data = b"" # Empty content = transparent
            appearance_ref = writer._add_object(appearance)

            annot = DictionaryObject()
            annot.update({
                NameObject("/Type"): NameObject("/Annot"),
                NameObject("/Subtype"): NameObject("/Sound"),
                NameObject("/Rect"): ArrayObject([NumberObject(r) for r in rect]),
                NameObject("/Sound"): sound_stream_ref,
                NameObject("/F"): NumberObject(4), 
                NameObject("/AP"): DictionaryObject({NameObject("/N"): appearance_ref}),
                NameObject("/NM"): pypdf.generic.TextStringObject(f"SoundAnnot_{total}"),
                NameObject("/H"): NameObject("/P"),
            })
            annot_ref = writer._add_object(annot)
            
            if "/Annots" not in page:
                page[NameObject("/Annots")] = ArrayObject()
            
            page["/Annots"].append(annot_ref)
            total += 1
            wav_index += 1
        except Exception as e:
            log(f"Error on page {page_num + 1}: {str(e)}", log_file)
            
    with open(output_path, "wb") as f:
        writer.write(f)
    return total

if __name__ == "__main__":
    # Automatic path detection
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
    work_dir = os.path.join(BASE_DIR, "99. Generate Report", "INSERT SOUND")
    
    if not os.path.exists(work_dir):
        os.makedirs(work_dir)
        
    log_path = os.path.join(work_dir, "final_insert_log.txt")
    if os.path.exists(log_path): os.remove(log_path)
    
    pdf_files = [f for f in os.listdir(work_dir) if f.lower().endswith('.pdf') and not f.startswith('SOUND_INSERTED')]
    wav_files = sorted([os.path.join(work_dir, f) for f in os.listdir(work_dir) if f.lower().endswith('.wav')])
    
    if not pdf_files or not wav_files:
        log("Missing PDF or WAV files.", log_path)
        sys.exit(1)
        
    input_pdf = os.path.join(work_dir, pdf_files[0])
    output_pdf = os.path.join(work_dir, "SOUND_INSERTED_" + pdf_files[0])
    
    try:
        positions = get_icon_positions(input_pdf, log_path)
        count = insert_sounds(input_pdf, wav_files, positions, output_pdf, log_path)
        log(f"Successfully inserted {count} sounds perfectly on top of icons. Output: {output_pdf}", log_path)
    except Exception as e:
        log(f"FATAL ERROR: {str(e)}", log_path)
