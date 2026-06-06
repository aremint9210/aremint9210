import os
from pathlib import Path
try:
    from PIL import Image, ImageOps
except ImportError:
    print("Error: The 'Pillow' library is required.")
    print("Please run 'pip install Pillow' in your terminal and try again.")
    input("Press Enter to exit...")
    exit()

# Target size (Width 7.37cm x Height 3.5cm)
# At 300 DPI (for higher quality printing), we calculate:
# 7.37cm ≈ 2.9in -> 870px
# 3.5cm ≈ 1.38in -> 414px
WIDTH_PX = 870
HEIGHT_PX = 414

def resize_images(folder_path):
    folder = Path(folder_path)
    output_folder = folder / "RESIZED_HIGH_QUALITY"
    output_folder.mkdir(exist_ok=True)
    
    print(f"Resizing images in: {folder}")
    
    for filename in os.listdir(folder):
        if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
            file_path = folder / filename
            try:
                with Image.open(file_path) as img:
                    # Using ImageOps.fit to crop and maintain aspect ratio
                    resized_img = ImageOps.fit(img, (WIDTH_PX, HEIGHT_PX), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
                    
                    # Save with explicit DPI metadata to help Word interpret the physical size correctly
                    resized_img.save(output_folder / filename, quality=100, subsampling=0, dpi=(300, 300))
                    print(f"Processed: {filename}")
            except Exception as e:
                print(f"Failed to process {filename}: {e}")

if __name__ == "__main__":
    target_dir = r"C:\Users\aremi\Desktop\REPORT CBM\PO 42289160 - 33kV THUNDER CYCLE2 2026\UE SPECTRUM\05.01.2026\001"
    resize_images(target_dir)
    print("\nDone! High-quality resized images are in the 'RESIZED_HIGH_QUALITY' folder.")
    input("Press Enter to exit...")
