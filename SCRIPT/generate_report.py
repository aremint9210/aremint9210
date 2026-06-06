import pandas as pd
import os
import json
import subprocess
import sys
from pathlib import Path

# Automatic Path Detection
SCRIPT_PATH = Path(__file__).resolve()
BASE_DIR = SCRIPT_PATH.parents[2] # C:\Users\aremi\Desktop\REPORT CBM\PO TIANU JB
MAPPING_DIR = BASE_DIR / "99. Generate Report" / "EXCELL MAPPING"
TEMPLATE_DIR = BASE_DIR / "99. Generate Report" / "TEMPLATE"
OUTPUT_DIR = BASE_DIR / "99. Generate Report" / "OUTPUT FILE"
OUTPUT_INDIVIDUAL = OUTPUT_DIR / "INDIVIDUAL"
OUTPUT_GROUP = OUTPUT_DIR / "GROUP"
HELPER_SCRIPT_PATH = SCRIPT_PATH.parent / "flir_word_report_helper.ps1"
MERGE_SCRIPT_PATH = SCRIPT_PATH.parent / "merge_reports.ps1"
POWERSHELL_32_PATH = Path(r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")

# Ensure output directories exist
OUTPUT_INDIVIDUAL.mkdir(parents=True, exist_ok=True)
OUTPUT_GROUP.mkdir(parents=True, exist_ok=True)

from generate_summary import generate_summary

def list_csv_files():
    files = list(MAPPING_DIR.glob("*.csv"))
    return sorted(files)

def merge_reports(csv_path):
    print(f"\n--- Merging Reports for: {csv_path.name} ---")
    csv_prefix = csv_path.stem
    csv_output_dir = OUTPUT_INDIVIDUAL / csv_prefix
    
    try:
        subprocess.run([
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(MERGE_SCRIPT_PATH),
            "-CsvPath",
            str(csv_path),
            "-IndividualDir",
            str(csv_output_dir), # Look for pages in the substation folder
            "-GroupOutputDir",
            str(csv_output_dir)   # Save merged reports in the same substation folder
        ], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error during merging: {e}")

def process_csv(csv_path, target_pages_input=None, no_merge=False):
    print(f"\n--- Loading: {csv_path.name} ---")
    
    # 1st: Generate/Update the Executive Summary for this substation
    print("Generating Executive Summary...")
    generate_summary(csv_path)

    try:
        # Keep N/A as literal string instead of converting to NaN
        df = pd.read_csv(csv_path, na_filter=False)
    except Exception as e:
        print(f"Error reading CSV: {e}")
        return

    # Filter out empty rows
    if 'Output File Name' not in df.columns:
        print(f"Error: 'Output File Name' column not found in {csv_path.name}")
        return

    df = df[df['Output File Name'].notna()]
    df = df[df['Output File Name'] != ""]
    
    available_pages = sorted(df['Output File Name'].unique().tolist())
    
    # --- Check for Executive Summary Template ---
    exec_template = TEMPLATE_DIR / "00. EXECUTIVE SUMMARY.docx"
    has_exec = exec_template.exists()
    if has_exec:
        available_pages.insert(0, "page0")

    # Create a specific output folder for this CSV inside INDIVIDUAL
    csv_prefix = csv_path.stem # Name without extension
    csv_output_dir = OUTPUT_INDIVIDUAL / csv_prefix
    csv_output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Total pages found: {len(available_pages)}")
    print(f"Output folder: {csv_output_dir}")
    
    if target_pages_input:
        user_input = target_pages_input
    else:
        user_input = input("Enter pages to generate (e.g., page0, page1, page5) or 'all': ").strip()
    
    if user_input.lower() == 'all':
        target_pages = available_pages
    else:
        target_pages = [p.strip() for p in user_input.split(',')]
        # Also support space-separated pages if coming from batch
        if len(target_pages) == 1 and ' ' in target_pages[0]:
             target_pages = [p.strip() for p in user_input.split(' ')]
             
        # Validate pages
        target_pages = [p for p in target_pages if p in available_pages]
        if not target_pages:
            print("No valid pages selected.")
            return

    # Filter for selected pages
    df_selected = df[df['Output File Name'].isin(target_pages)]
    
    # --- Executive Summary (page0) logic ---
    if has_exec and "page0" in target_pages:
        print("\nGenerating Executive Summary (page0)...")
        first_row = df.iloc[0]
        output_name = "page0"
        output_path = csv_output_dir / f"{output_name}.docx"
        
        # Prepare field overrides from first row
        overrides = {}
        for col in df.columns:
            if col.startswith('{{') and col.endswith('}}'):
                val = str(first_row[col])
                overrides[col] = val if val.lower() != 'nan' else ""
        
        # --- RMU Serial Number Rule ---
        swg_type = overrides.get('{{ swg.type }}', '').upper()
        if 'RMU' in swg_type:
            overrides['{{ panel.serialnumber }}'] = overrides.get('{{ swg.serialnumber }}', '')

        manifest = {
            "template_path": str(exec_template),
            "output_path": str(output_path),
            "pages": [
                {
                    "template_page_no": 1,
                    "ir_path": "", # No images for executive summary
                    "visual_path": "",
                    "field_overrides": overrides
                }
            ]
        }
        
        manifest_json_path = BASE_DIR / f"temp_{output_name}.json"
        with open(manifest_json_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2)

        try:
            subprocess.run([
                str(POWERSHELL_32_PATH),
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(HELPER_SCRIPT_PATH),
                "-ManifestPath",
                str(manifest_json_path)
            ], check=False)
        finally:
            if manifest_json_path.exists():
                os.remove(manifest_json_path)

    pages_group = df_selected.groupby('Output File Name')

    for output_name, group_data in pages_group:
        if output_name == "page0": continue # Already handled
        
        print(f"\nProcessing {output_name}...")

        # Identify rows (Visual and IR)
        row_visual = group_data[group_data['Attribute'].str.contains('Visual', case=False, na=False)]
        row_ir = group_data[group_data['Attribute'].str.contains('IR', case=False, na=False)]

        if row_visual.empty or row_ir.empty:
            print(f"  [SKIP] Missing Visual or IR row.")
            continue

        template_name = str(row_visual['Word Template'].iloc[0])
        template_path = TEMPLATE_DIR / template_name
        
        if not template_path.exists():
            print(f"  [ERROR] Template not found: {template_name}")
            continue

        # --- Robust Image Path Resolution ---
        visual_filename = str(row_visual['Image File Name'].iloc[0])
        ir_filename = str(row_ir['Image File Name'].iloc[0])
        
        # Helper to find a file in prioritised subfolders
        def find_local_image(base_dir, filename):
            # Prioritise direct hit, then IR, then DIGITAL/DG, then base
            options = [
                base_dir / filename,
                base_dir / "IR" / filename,
                base_dir / "DIGITAL" / filename,
                base_dir / "DG" / filename
            ]
            return next((p for p in options if p.exists()), None)

        # 1. Try path from CSV (as a starting point)
        csv_folder = Path(str(row_visual['Folder'].iloc[0]))
        visual_path = find_local_image(csv_folder, visual_filename)
        ir_path = find_local_image(csv_folder, ir_filename)

        # 2. Fallback: Search in local 'RAW MATERIAL'
        if not visual_path or not ir_path:
            raw_material_root = BASE_DIR / "RAW MATERIAL"
            prefix_match = csv_path.name.split('.')[0].strip()
            possible_folders = [d for d in raw_material_root.iterdir() if d.is_dir() and (d.name == prefix_match or d.name.startswith(prefix_match + "."))]
            
            if possible_folders:
                local_base = possible_folders[0]
                if not visual_path: visual_path = find_local_image(local_base, visual_filename)
                if not ir_path: ir_path = find_local_image(local_base, ir_filename)

        # Final Validation
        if not visual_path or not visual_path.exists():
            print(f"  [ERROR] Visual image not found: {visual_filename}")
            continue
        if not ir_path or not ir_path.exists():
            print(f"  [ERROR] IR image not found: {ir_filename}")
            continue

        output_path = csv_output_dir / f"{output_name}.docx"
        template_page_no = int(row_visual['Template Page No'].iloc[0])

        # Prepare field overrides
        overrides = {}
        for col in df.columns:
            if col.startswith('{{') and col.endswith('}}'):
                val = str(row_visual[col].iloc[0])
                if val.lower() == 'nan':
                    val = ""
                
                # --- Custom Formatting Rules ---
                if 'time' in col.lower() and val:
                    val = val.upper()
                elif 'humidity' in col.lower() and val:
                    if '%' not in val: val = f"{val}%"
                elif 'ambient' in col.lower() and val:
                    if '°' not in val: val = f"{val} °C"
                
                overrides[col] = val

        # --- RMU Serial Number Rule ---
        swg_type = overrides.get('{{ swg.type }}', '').upper()
        if 'RMU' in swg_type:
            overrides['{{ panel.serialnumber }}'] = overrides.get('{{ swg.serialnumber }}', '')

        overrides["{{visual_image}}"] = visual_filename

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

        manifest_json_path = BASE_DIR / f"temp_{output_name}.json"
        with open(manifest_json_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, indent=2)

        try:
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
                print(f"    {line.strip()}")

            process.wait()
        finally:
            if manifest_json_path.exists():
                os.remove(manifest_json_path)

    print("\nGeneration complete.")
    
    if no_merge:
        return
    
    # NEW: Automatically merge reports into the same folder
    merge_reports(csv_path)

def main():
    if len(sys.argv) > 1:
        # Argument mode: python generate_report.py <csv_path> [pages] [--no-merge]
        csv_path_arg = Path(sys.argv[1])
        if not csv_path_arg.is_absolute():
            csv_path_arg = MAPPING_DIR / csv_path_arg.name
        
        if not csv_path_arg.exists():
            print(f"Error: CSV file not found at {csv_path_arg}")
            sys.exit(1)
            
        no_merge = "--no-merge" in sys.argv
        
        # Optional pages argument
        if len(sys.argv) > 2 and sys.argv[2] != "--no-merge":
            target_pages_arg = sys.argv[2]
            process_csv(csv_path_arg, target_pages_arg, no_merge)
        else:
            process_csv(csv_path_arg, no_merge=no_merge)
        return

    # Interactive mode (if run directly)
    csv_files = list_csv_files()
    if not csv_files:
        print(f"No CSV files found in {MAPPING_DIR}")
        return

    print("\n=== Select CSV File ===")
    for i, f in enumerate(csv_files):
        print(f"{i+1}. {f.name}")
    print("0. Exit")

    choice = input("\nEnter choice (number): ").strip()
    if choice != '0':
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(csv_files):
                process_csv(csv_files[idx])
            else:
                print("Invalid choice.")
        except ValueError:
            print("Please enter a number.")

    print("\nProcess finished.")

if __name__ == "__main__":
    main()
