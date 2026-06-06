import pandas as pd
import os
import json
import subprocess
import sys
from pathlib import Path

# Configuration
BASE_DIR = Path(__file__).resolve().parent.parent.parent
DEFAULT_CSV_PATH = BASE_DIR / "GENERATE REPORT" / "EXCELL MAPPING" / "Master List.csv"
TEMPLATE_DIR = BASE_DIR / "GENERATE REPORT" / "TEMPLATE"
OUTPUT_INDIVIDUAL = BASE_DIR / "GENERATE REPORT" / "OUTPUT FILE" / "INDIVIDUAL"
# IR_BASE_PATH is used as a fallback in manual_generate
IR_BASE_PATH = BASE_DIR / "RAW MATERIAL" / "01. PPU INDERA SEMPURNA (IR+VI)" / "IR"

# Paths to existing scripts
POWERSHELL_32_PATH = Path(r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")
HELPER_SCRIPT_PATH = BASE_DIR / "GENERATE REPORT" / "SCRIPT" / "flir_word_report_helper (1).ps1"

def manual_generate(csv_path=None, pages_input=None):
    if csv_path is None:
        csv_path = DEFAULT_CSV_PATH
    else:
        csv_path = Path(csv_path)

    if not csv_path.exists():
        print(f"Error: CSV not found at {csv_path}")
        return

    print(f"Using CSV: {csv_path.name}")
    # Load CSV
    df = pd.read_csv(csv_path)
    
    print("--- Manual Page Generation ---")
    if pages_input is None:
        print("Available pages (Output File Name) are usually like: page1, page2, page10, etc.")
        user_input = input("Enter the pages you want to generate (separated by comma, e.g., page1, page5, page10) or 'all' for all pages in CSV: ")
    elif str(pages_input).strip() == "":
        user_input = "all"
    else:
        user_input = pages_input
    
    if str(user_input).lower().strip() == 'all':
        target_pages = df['Output File Name'].unique().tolist()
    else:
        target_pages = [p.strip() for p in str(user_input).split(',')]

    # Filter for the requested pages
    df = df[df['Output File Name'].isin(target_pages)]
    
    if df.empty:
        print(f"No pages found matching: {target_pages}")
        return

    # Filter out empty rows
    df = df[df['Output File Name'].notna()]

    # Group by Output File Name
    pages = df.groupby('Output File Name')

    # Determine output directory
    if csv_path.stem != "Master List":
        target_output_dir = OUTPUT_INDIVIDUAL / csv_path.stem
    else:
        target_output_dir = OUTPUT_INDIVIDUAL

    for output_name, group_data in pages:
        print(f"\n--- Processing Page: {output_name} ---")
        
        # Identify rows
        row_visual = group_data[group_data['Attribute'] == 'Visual Image Path']
        row_ir = group_data[group_data['Attribute'] == 'IR Image Path']
        
        if row_visual.empty or row_ir.empty:
            print(f"  [SKIP] Missing Visual or IR row for {output_name}")
            continue

        # Use row 1 (visual row) for placeholders
        template_name = str(row_visual['Word Template'].iloc[0])
        template_page_no = 1 
        
        visual_filename = str(row_visual['Image File Name'].iloc[0])
        ir_filename = str(row_ir['Image File Name'].iloc[0])
        
        # Determine the correct image folder from the CSV
        csv_folder = str(row_visual['Folder'].iloc[0])
        if "RAW MATERIAL" in csv_folder:
            # Extract the part after RAW MATERIAL
            rel_part = csv_folder.split("RAW MATERIAL")[-1].lstrip("\\/")
            actual_folder = BASE_DIR / "RAW MATERIAL" / rel_part
        else:
            # Fallback to the original hardcoded path if RAW MATERIAL is not found
            actual_folder = BASE_DIR / "RAW MATERIAL" / "01. PPU INDERA SEMPURNA (IR+VI)" / "IR"

        visual_path = actual_folder / visual_filename
        ir_path = actual_folder / ir_filename
        
        if not visual_path.exists():
            print(f"  [ERROR] Visual image not found: {visual_path}")
            continue
        if not ir_path.exists():
            print(f"  [ERROR] IR image not found: {ir_path}")
            continue

        template_path = TEMPLATE_DIR / template_name
        if not template_path.exists():
            print(f"  [ERROR] Template not found: {template_path}")
            continue

        output_path = target_output_dir / f"{output_name}.docx"

        print(f"  Template: {template_name}")
        print(f"  Visual: {visual_filename}")
        print(f"  IR: {ir_filename}")

        # Prepare field overrides from row headers
        overrides = {}
        for col in df.columns:
            if col.startswith('{{') and col.endswith('}}'):
                val = row_visual[col].iloc[0]
                if pd.isna(val) or str(val).lower() == 'nan':
                    val = ""
                else:
                    # Remove .0 from floats that are whole numbers (e.g., 75.0 -> 75)
                    try:
                        if float(val).is_integer():
                            val = int(float(val))
                    except (ValueError, TypeError):
                        pass
                    val = str(val)
                overrides[col] = val
        
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

        manifest_json_path = BASE_DIR / f"temp_manifest_manual_{output_name}.json"
        with open(manifest_json_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2)

        # Execute PowerShell Helper
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
            else:
                print(f"  [ERROR] FLIR Helper failed for {output_name}")
        finally:
            if manifest_json_path.exists():
                os.remove(manifest_json_path)

    print("\nManual generation process complete.")

if __name__ == "__main__":
    csv_arg = sys.argv[1] if len(sys.argv) > 1 else None
    pages_arg = sys.argv[2] if len(sys.argv) > 2 else None
    manual_generate(csv_arg, pages_arg)
