#!/usr/bin/env python3
import sys, re, os, pytesseract
from PIL import Image

LETRA = r'A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ'
# número que não faz parte de uma data (ex.: 19/10/2026)
_D = os.environ.get('MAT_DIGITOS', '')   # tamanho da matrícula da empresa (vem do servidor)
NUM = r'(?<![\d/])(\d{' + (_D if _D.isdigit() and int(_D) > 0 else '3,6') + r'})(?![\d/])'


def ocr(img, lang):
    try:
        return pytesseract.image_to_string(img, lang=lang)
    except Exception:
        return pytesseract.image_to_string(img, lang='eng')


def limpa_nome(txt):
    return ' '.join(txt.split())


def layout_contracheque(img, lang):
    """Contracheque: matrícula e nome na faixa do topo (13% da página)."""
    w, h = img.size
    text = ocr(img.crop((0, 0, w, int(h * 0.13))), lang)
    m = re.search(NUM + r'\s+([' + LETRA + r'][' + LETRA + r'\s]{5,})', text)
    if m:
        return m.group(1), limpa_nome(m.group(2))
    return None


def layout_ficha(img, lang):
    """Fichas (uniforme/EPI): quadro MATRÍCULA | NOME | ADMISSÃO | DEMISSÃO abaixo do título."""
    w, h = img.size
    text = ocr(img.crop((0, int(h * 0.10), w, int(h * 0.30))), lang)
    # pega só o que vem depois do rótulo MATRÍCULA, se ele foi lido
    i = re.search(r'MATR[IÍ]CULA', text, re.I)
    corpo = text[i.end():] if i else text
    m = re.search(NUM + r'\s+([' + LETRA + r'][' + LETRA + r'\s]{3,}?)(?=\s*(?:\d{2}/\d{2}/\d{4}|ADMISS|UNIDADE|$))',
                  corpo, re.S)
    if m:
        return m.group(1), limpa_nome(m.group(2))
    # nome não lido: devolve só a matrícula (o servidor usa o nome da lista)
    m = re.search(NUM, corpo)
    if m:
        return m.group(1), ''
    return None


def extrair(img_path, lang='por'):
    img = Image.open(img_path)
    achou = layout_contracheque(img, lang) or layout_ficha(img, lang)
    if achou:
        print(f"{achou[0].strip()}||{achou[1]}")
    else:
        print("NAO_IDENTIFICADO")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("NAO_IDENTIFICADO")
        sys.exit(1)
    if '--texto' in sys.argv:
        # texto da página inteira (a identificação é feita pelo servidor)
        print(ocr(Image.open(sys.argv[1]), 'por'))
    else:
        extrair(sys.argv[1])
