import os
import xml.etree.ElementTree as ET
import zipfile

pptx_path = '/Users/kiku/directory/University/master_thesis/ref/修論発表2026_菊池裕夢.pptx'

def get_slide_texts(pptx_path):
    with zipfile.ZipFile(pptx_path) as z:
        xml_files = [f for f in z.namelist() if f.startswith('ppt/slides/slide')]
        xml_files.sort(key=lambda x: int(x.replace('ppt/slides/slide', '').replace('.xml', '')))
        
        for xml_file in xml_files:
            slide_content = ""
            with z.open(xml_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                # Namespaces
                ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                
                for t in root.findall('.//a:t', ns):
                    if t.text:
                        slide_content += t.text + "\n"
            
            print(f"--- Slide {xml_file} ---")
            print(slide_content)

get_slide_texts(pptx_path)
