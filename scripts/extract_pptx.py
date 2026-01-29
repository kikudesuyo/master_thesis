
import zipfile
import re
import sys
import os

def extract_text_from_pptx(pptx_path):
    if not os.path.exists(pptx_path):
        print(f"File not found: {pptx_path}")
        return

    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            # Find all slide files
            slides = [f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            # Sort slides (slide1, slide2, ..., slide10, etc.)
            slides.sort(key=lambda x: int(re.search(r'slide(\d+)', x).group(1)))

            for slide in slides:
                content = z.read(slide).decode('utf-8')
                # Simple regex to extract text within <a:t>...</a:t>
                # Note: This might miss some text or split words if formatting changes mid-word, but good enough for reading.
                text_content = re.findall(r'<a:t>(.*?)</a:t>', content)
                
                if text_content:
                    print(f"--- {slide} ---")
                    print("".join(text_content))
                    print("\n")
    except Exception as e:
        print(f"Error reading pptx: {e}")

if __name__ == "__main__":
    pptx_file = "/Users/kiku/directory/University/master_thesis/ref/修論発表2026_菊池裕夢.pptx"
    extract_text_from_pptx(pptx_file)
