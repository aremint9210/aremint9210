import pymupdf  # PyMuPDF
import os
import sys

def log(msg, log_file):
    with open(log_file, 'a') as f:
        f.write(msg + '\n')

def insert_sounds(input_dir, log_file):
    log(f"Starting sound insertion in {input_dir}", log_file)
    
    # 1. Find PDF and WAV files
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
    wav_files = sorted([f for f in os.listdir(input_dir) if f.lower().endswith('.wav')])
    
    if not pdf_files:
        log("No PDF file found.", log_file)
        return
    if not wav_files:
        log("No WAV files found.", log_file)
        return
    
    input_pdf = os.path.join(input_dir, pdf_files[0])
    output_pdf = os.path.join(input_dir, "SOUND_INSERTED_" + pdf_files[0])
    
    log(f"Processing PDF: {input_pdf}", log_file)
    log(f"Found {len(wav_files)} WAV files: {wav_files}", log_file)
    
    doc = pymupdf.open(input_pdf)
    wav_index = 0
    total_inserted = 0
    
    # Speaker icon character might be \uf028 or similar depending on the font, 
    # but we can also look for "Sound clip" text.
    target_text = "Sound clip"
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Search for "Sound clip"
        text_instances = page.search_for(target_text)
        
        if text_instances:
            log(f"Page {page_num + 1}: Found '{target_text}'", log_file)
            for inst in text_instances:
                # We want to place the sound icon slightly to the right of "Sound clip"
                # Typically the icon is around x + 60-80 points
                sound_rect = pymupdf.Rect(inst.x1 + 5, inst.y0 - 2, inst.x1 + 25, inst.y1 + 2)
                
                # Use current WAV file
                wav_path = os.path.join(input_dir, wav_files[wav_index % len(wav_files)])
                
                try:
                    # Add sound annotation
                    # Note: add_sound_annot is available in newer PyMuPDF
                    with open(wav_path, "rb") as sound_stream:
                        page.add_sound_annot(sound_rect, sound_stream.read(), filename=wav_files[wav_index % len(wav_files)])
                    
                    log(f"  Inserted {wav_files[wav_index % len(wav_files)]} at {sound_rect}", log_file)
                    total_inserted += 1
                    wav_index += 1
                except Exception as e:
                    log(f"  Error inserting sound: {str(e)}", log_file)
        else:
            # log(f"Page {page_num + 1}: No '{target_text}' found.", log_file)
            pass

    doc.save(output_pdf)
    doc.close()
    log(f"Finished. Total inserted: {total_inserted}. Saved to {output_pdf}", log_file)

if __name__ == "__main__":
    work_dir = r"GENERATE REPORT\INSERT SOUND"
    log_path = os.path.join(work_dir, "script_log.txt")
    if os.path.exists(log_path):
        os.remove(log_path)
    
    try:
        insert_sounds(work_dir, log_path)
    except Exception as e:
        log(f"FATAL ERROR: {str(e)}", log_path)
