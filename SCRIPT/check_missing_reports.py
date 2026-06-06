import pandas as pd
import os
import re
import sys
from pathlib import Path

# Configuration
SCRIPT_PATH = Path(__file__).resolve()
BASE_DIR = SCRIPT_PATH.parents[2]
OUTPUT_ROOT = BASE_DIR / "99. Generate Report" / "OUTPUT FILE"
MAPPING_DIR = BASE_DIR / "99. Generate Report" / "EXCELL MAPPING"
OUTPUT_INDIVIDUAL = OUTPUT_ROOT / "INDIVIDUAL"

def natural_sort_key(s):
    """Helper to sort strings containing numbers in a natural way (page2 before page10)."""
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split('([0-9]+)', str(s))]

def find_report_file(output_name, csv_prefix):
    """Searches for the report file in the specific INDIVIDUAL subfolder."""
    file_name = f"{output_name}.docx"
    
    # Check ONLY the specific CSV subfolder in INDIVIDUAL
    path_csv_folder = OUTPUT_INDIVIDUAL / csv_prefix / file_name
    if path_csv_folder.exists():
        return path_csv_folder

    return None

def resolve_image_path(csv_folder, filename, csv_path):
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

    # 1. Try path from CSV
    found = find_local_image(csv_folder, filename)
    if found: return found

    # 2. Fallback: Search in local 'RAW MATERIAL'
    raw_material_root = BASE_DIR / "RAW MATERIAL"
    match = re.match(r'^(\d+)', csv_path.name)
    prefix_match = match.group(1) if match else csv_path.name.split('.')[0]
    possible_folders = [d for d in raw_material_root.iterdir() if d.is_dir() and (d.name == prefix_match or d.name.startswith(prefix_match + "."))]
    
    if possible_folders:
        local_base = possible_folders[0]
        return find_local_image(local_base, filename)
    return None

def check_missing(specific_files=None, auto_fix=False):
    print("========================================")
    print("      Detailed Report Audit Tool")
    print("========================================")
    # Match the reference output's path display style (showing where it's checking)
    print(f"Checking output in: {OUTPUT_ROOT}\n")

    # Filter out --auto-fix from specific_files if present
    if specific_files:
        if "--auto-fix" in specific_files:
            auto_fix = True
            specific_files.remove("--auto-fix")

    if specific_files:
        csv_files = []
        for f in specific_files:
            p = MAPPING_DIR / f
            if not p.exists():
                p = MAPPING_DIR / (f + ".csv")
            if p.exists():
                csv_files.append(p)
            else:
                print(f"  [ERROR] CSV not found: {f}")
    else:
        csv_files = list(MAPPING_DIR.glob("*.csv"))
    
    if not csv_files:
        print("No CSV files found.")
        return

    for csv_path in csv_files:
        if csv_path.name in ["image number.csv", "Master List.xlsx", "Master List.csv"]:
            continue
            
        csv_stem = csv_path.stem
        print(f"Analyzing CSV: {csv_path.name}")
        
        try:
            df = pd.read_csv(csv_path, na_filter=False)
        except Exception as e:
            print(f"  [ERROR] Could not read CSV: {e}")
            continue

        if 'Output File Name' not in df.columns:
            print(f"  [SKIP] 'Output File Name' column not found.")
            continue

        df = df[df['Output File Name'].notna()]
        df = df[df['Output File Name'] != ""]
        pages = df.groupby('Output File Name')

        stats = {
            "Total Potential Reports": 0,
            "Successfully Generated": 0,
            "Incomplete Data (Skip)": 0,
            "TRULY MISSING": 0,
            "Log": []
        }

        # Sort page names naturally
        sorted_output_names = sorted(pages.groups.keys(), key=natural_sort_key)

        for output_name in sorted_output_names:
            if output_name == "Output File Name": continue
            
            group_data = pages.get_group(output_name)
            stats["Total Potential Reports"] += 1
            
            # Identify report type and details for the log
            report_type = group_data['Type'].iloc[0] if 'Type' in group_data.columns else "Unknown"
            group_no = group_data['Group'].iloc[0] if 'Group' in group_data.columns else "N/A"
            area = group_data['{{ swg.area }}'].iloc[0] if '{{ swg.area }}' in group_data.columns else "N/A"
            panel_no = group_data['{{ panel.linknumber }}'].iloc[0] if '{{ panel.linknumber }}' in group_data.columns else "N/A"
            panel_name = group_data['{{ panel.name }}'].iloc[0] if '{{ panel.name }}' in group_data.columns else "N/A"
            
            details = f"Grp:{group_no} | {area}"
            if panel_no != "N/A" and panel_no != "":
                details += f" | Panel:{panel_no}"
            if panel_name != "N/A" and panel_name != "":
                details += f" ({panel_name})"

            # --- Check Data Integrity (Images existence) ---
            image_errors = []
            
            # Look for Visual and IR rows
            row_visual = group_data[group_data['Attribute'].str.contains('Visual', case=False, na=False)]
            row_ir = group_data[group_data['Attribute'].str.contains('IR', case=False, na=False)]
            
            if row_visual.empty:
                image_errors.append("Visual row missing in CSV")
            else:
                v_folder = Path(str(row_visual['Folder'].iloc[0]))
                v_file = str(row_visual['Image File Name'].iloc[0])
                if not resolve_image_path(v_folder, v_file, csv_path):
                    image_errors.append(f"Visual image not on disk: {v_file}")
                    
            if row_ir.empty:
                image_errors.append("IR row missing in CSV")
            else:
                ir_folder = Path(str(row_ir['Folder'].iloc[0]))
                ir_file = str(row_ir['Image File Name'].iloc[0])
                if not resolve_image_path(ir_folder, ir_file, csv_path):
                    image_errors.append(f"IR image not on disk: {ir_file}")

            # --- Check Report existence ---
            found_report = find_report_file(output_name, csv_stem)

            if image_errors:
                stats["Incomplete Data (Skip)"] += 1
                stats["Log"].append({
                    "type": "incomplete",
                    "msg": f"  [!] {output_name} ({report_type}): {', '.join(image_errors)}",
                    "details": details
                })
            elif not found_report:
                stats["TRULY MISSING"] += 1
                stats["Log"].append({
                    "type": "missing",
                    "msg": f"  [X] {output_name} ({report_type}): All data present but file missing!",
                    "details": details
                })
            else:
                stats["Successfully Generated"] += 1

        # Print Statistics
        print(f"  - Total Potential Reports: {stats['Total Potential Reports']}")
        print(f"  - Successfully Generated:  {stats['Successfully Generated']}")
        print(f"  - Incomplete Data (Skip): {stats['Incomplete Data (Skip)']}")
        print(f"  - TRULY MISSING:          {stats['TRULY MISSING']}")
        
        if stats["Log"]:
            missing = [l for l in stats["Log"] if l['type'] == 'missing']
            incomplete = [l for l in stats["Log"] if l['type'] == 'incomplete']

            if missing:
                print("\n  >>> TRULY MISSING REPORTS (Action Required) <<<")
                print("  These reports have images but the .docx was NOT found in INDIVIDUAL folder.")
                missing_pages = []
                for item in missing:
                    print(f"{item['msg']}")
                    print(f"      Context: {item['details']}")
                    # Extract page name from the message (it's at the start of the msg)
                    page_name = item['msg'].split('(')[0].replace('[X]', '').strip()
                    missing_pages.append(page_name)
                
                print(f"\n  [COPY-PASTE LIST]: {', '.join(missing_pages)}")
                
                # Option to re-generate
                if auto_fix:
                    print(f"\n[AUTO-FIX] Re-generating {len(missing_pages)} missing reports...")
                    do_fix = True
                else:
                    print(f"\nFound {len(missing_pages)} missing reports.")
                    choice = input(f"Do you want to re-generate these {len(missing_pages)} reports now? (y/n): ").strip().lower()
                    do_fix = (choice == 'y')

                if do_fix:
                    gen_script = SCRIPT_PATH.parent / "generate_report.py"
                    pages_str = ",".join(missing_pages)
                    # If we are in auto_fix mode, we should pass --no-merge to prevent prompts
                    cmd = [sys.executable, str(gen_script), str(csv_path), pages_str]
                    if auto_fix:
                        cmd.append("--no-merge")
                    
                    print(f"Running: {' '.join(cmd)}")
                    import subprocess
                    subprocess.run(cmd)
            
            if incomplete:
                print("\n  >>> INCOMPLETE DATA (Skipped by Generator) <<<")
                print("  These reports were skipped because images are missing from the 'Raw Material' folder.")
                for item in incomplete:
                    print(f"{item['msg']}")
                    print(f"      Context: {item['details']}")
        
        print("-" * 50)

    print("\nAudit Complete.")

if __name__ == "__main__":
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    check_missing(args)
