# -*- coding: utf-8 -*-
"""Funções utilitárias de normalização/conversão.

Portadas sem alteração de lógica a partir de `gerar_base_matrizes.py`.
"""
import hashlib
import re
import unicodedata


def sa(s):
    """normaliza: sem acento, maiuscula, espacos colapsados"""
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().upper()


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return re.sub(r'\s+', ' ', str(v)).strip()


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    s = str(v).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def inteiro(v):
    n = num(v)
    return int(n) if n is not None and n == int(n) else n


def md5(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for bloco in iter(lambda: f.read(1 << 20), b''):
            h.update(bloco)
    return h.hexdigest()


def md5_texto(*partes):
    """hash estável de uma tupla de valores de texto — usado para o fingerprint
    por disciplina (detecta mudança de conteúdo mesmo quando o arquivo inteiro
    não muda de md5, ex.: célula editada e depois revertida em outra linha)."""
    h = hashlib.md5()
    for p in partes:
        h.update((txt(p) if not isinstance(p, (int, float)) else str(p)).encode("utf-8"))
        h.update(b"\x1f")
    return h.hexdigest()
