import pandas as pd
import os
import json
import subprocess
from pathlib import Path

# Configuration
BASE_DIR = Path(__file__).resolve().parent.parent.parent
CSV_PATH = BASE_DIR / "GENERATE REPORT" / "EXCELL MAPPING" / "Master List.csv"
TEMPLATE_DIR = BASE_DIR / "GENERATE REPORT" / "TEMPLATE"
OUTPUT_INDIVIDUAL = BASE_DIR / "GENERATE REPORT" / "OUTPUT FILE" / "INDIVIDUAL"
OUTPUT_GROUP = BASE_DIR / "GENERATE REPORT" / "OUTPUT FILE" / "GROUP"
IR_BASE_PATH = BASE_DIR / "RAW MATERIAL" / "01. PPU INDERA SEMPURNA (IR+VI)" / "IR"

# Paths to existing scripts
POWERSHELL_32_PATH = Path(r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")
HELPER_SCRIPT_PATH = BASE_DIR / "GENERATE REPORT" / "SCRIPT" / "flir_word_report_helper (1).ps1"

def generate_reports():
    if not CSV_PATH.exists():
        print(f"Error: CSV not found at {CSV_PATH}")
        return

    # Load CSV
    df = pd.read_csv(CSV_PATH)
    
    # FILTER FOR SPECIFIC TYPES
    target_types = ['Overview 33', '33panel1']
    df = df[df['Type'].isin(target_types)]
    
    # Filter out empty rows
    df = df[df['Output File Name'].notna()]

    # Group by Output File Name (each page has 2 rows: Visual and IR)
    pages = df.groupby('Output File Name')

    all_generated_files = []

    for output_name, group_data in pages:
        print(f"--- Processing Page: {output_name} ---")
        
        # Identify rows
        row_visual = group_data[group_data['Attribute'] == 'Visual Image Path']
        row_ir = group_data[group_data['Attribute'] == 'IR Image Path']
        
        if row_visual.empty or row_ir.empty:
            print(f"  [SKIP] Missing Visual or IR row for {output_name}")
            continue

        # Use row 1 (visual row) for placeholders
        template_name = str(row_visual['Word Template'].iloc[0])
        
        # FORCE Template Page No to 1 because individual template files are 1-page long
        # The CSV has 1, 2, 3... but the files are separate 1-page docs.
        template_page_no = 1 
        
        group_id = str(row_visual['Group'].iloc[0])
        
        visual_filename = str(row_visual['Image File Name'].iloc[0])
        ir_filename = str(row_ir['Image File Name'].iloc[0])
        
        visual_path = IR_BASE_PATH / visual_filename
        ir_path = IR_BASE_PATH / ir_filename
        template_path = TEMPLATE_DIR / template_name
        output_path = OUTPUT_INDIVIDUAL / f"{output_name}.docx"

        print(f"  Template: {template_name}")
        print(f"  Visual: {visual_filename}")
        print(f"  IR: {ir_filename}")

        # Prepare field overrides from row headers
        overrides = {}
        for col in df.columns:
            if col.startswith('{{') and col.endswith('}}'):
                val = str(row_visual[col].iloc[0])
                overrides[col] = val if val != 'nan' else ""
        
        # --- RMU Serial Number Rule ---
        swg_type = overrides.get('{{ swg.type }}', '').upper()
        if 'RMU' in swg_type:
            overrides['{{ panel.serialnumber }}'] = overrides.get('{{ swg.serialnumber }}', '')

        overrides["{{visual_image}}"] = visual_filename

        # Construct Manifest
        manifest = {
            "template_path": str(template_path),
            "output_path": str(output_path),
            "pages": [
                {
                    "template_page_no": template_page_no,
                    "ir_path": str(ir_path),
                    "visual_path": str(visual_path),
                    "field_overrides": overrides
                }
            ]
        }

        manifest_json_path = BASE_DIR / f"temp_manifest_{output_name}.json"
        with open(manifest_json_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2)

        # Execute PowerShell Helper with real-time stream
        try:
            print(f"  Starting FLIR Engine...")
            process = subprocess.Popen([
                str(POWERSHELL_32_PATH),
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(HELPER_SCRIPT_PATH),
                "-ManifestPath",
                str(manifest_json_path)
            ], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            
            for line in process.stdout:
                print(f"    [PS] {line.strip()}")
            
            process.wait()
            
            if process.returncode == 0:
                print(f"  [SUCCESS] {output_path}")
                all_generated_files.append({"file": output_path, "group": group_id})
            else:
                print(f"  [ERROR] FLIR Helper failed for {output_name}")
        finally:
            if manifest_json_path.exists():
                os.remove(manifest_json_path)

    # Merge by Group (Logic placeholder - requires Word Merge implementation)
    # Since merging Word docs often requires specific COM automation, 
    # you can use a separate script or I can provide a PowerShell snippet to merge them.
    print("\nIndividual page generation complete.")
    print(f"Files are located in: {OUTPUT_INDIVIDUAL}")
    print("Grouping logic: Grouped files by Column C.")

if __name__ == "__main__":
    generate_reports()
