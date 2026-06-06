import PyPDF2
import sys

def count_pages(file_path):
    try:
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            return len(reader.pages)
    except Exception as e:
        return str(e)

if __name__ == "__main__":
    path = r'GENERATE REPORT\EXAMPLE\17. MINI PPU TERLA 33-11kV (VI).pdf'
    print(count_pages(path))
