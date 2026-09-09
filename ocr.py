#!/usr/bin/env python3
import sys, re, pytesseract
from PIL import Image

def extrair(img_path):
    img  = Image.open(img_path)
    w, h = img.size
    topo = img.crop((0, 0, w, int(h * 0.13)))
    text = pytesseract.image_to_string(topo, lang='por')
    match = re.search(
        r'(\d{3,6})\s+([A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ][A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ\s]{5,})',
        text
    )
    if match:
        mat  = match.group(1).strip()
        nome = ' '.join(match.group(2).strip().split())
        print(f"{mat}||{nome}")
    else:
        print("NAO_IDENTIFICADO")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("NAO_IDENTIFICADO")
        sys.exit(1)
    extrair(sys.argv[1])
