# -*- coding: utf-8 -*-
"""Mapeamento manual arquivo -> código de curso (`_DE_PARA_CURSOS.xlsx`).

Portado de `gerar_base_matrizes.py` sem alteração de lógica. Única mudança:
as funções recebem o caminho do arquivo como parâmetro em vez de um global
fixo `DEPARA`, para poderem ser chamadas com o caminho definido pelo
`orquestrador` (o de-para continua fora de `_historico/`, é o único arquivo
editado à mão — ver `snapshot.copiar_depara_usado` para a cópia de auditoria
gravada dentro de cada snapshot).
"""
import os
import re

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from pipeline.util import txt


def carregar_depara(caminho_depara):
    """{(marca, nome_do_arquivo): codigo_confirmado}

    Chave é (marca, arquivo), não só arquivo: o mesmo nome de arquivo/
    template pode existir em instituições diferentes (ex.: um template
    genérico reaproveitado) — sem a marca, uma confirmação feita pra um
    arquivo de uma instituição vazaria pro arquivo homônimo de outra.
    """
    if not os.path.exists(caminho_depara):
        return {}
    wb = openpyxl.load_workbook(caminho_depara, data_only=True)
    ws = wb.active
    d = {}
    # dados começam na linha 3 (linha 1 = instrução, linha 2 = cabeçalho —
    # ver gravar_depara). O script original (gerar_base_matrizes.py) começa
    # em range(2, ...), um off-by-one que lê o cabeçalho como se fosse dado;
    # inofensivo na prática (gera uma entrada fantasma "ARQUIVO" que nunca
    # bate com um nome de arquivo real), mas corrigido aqui no port.
    for i in range(3, ws.max_row + 1):
        marca = txt(ws.cell(i, 1).value)
        arq = txt(ws.cell(i, 2).value)
        conf = txt(ws.cell(i, 5).value)
        if arq and conf:
            d[(marca, arq)] = conf
    wb.close()
    return d


def gravar_depara(caminho_depara, pendentes, ja):
    """pendentes = [(marca, arquivo, curso_cabecalho, sugestao)]"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DE_PARA"
    ws["A1"] = ("PREENCHA A COLUNA E. Esta é a única planilha editada à mão. O gerador a lê a cada execução. "
                "Se um arquivo novo chegar sem SIGLA INTERNA preenchida, ele aparece aqui na próxima execução — "
                "basta completar a coluna CÓDIGO CONFIRMADO. Não altere as colunas A a D: elas são reescritas.")
    ws["A1"].font = Font(italic=True, size=9, color="595959")
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 46
    H = ["MARCA", "ARQUIVO", "CURSO (do cabeçalho da matriz)", "CÓDIGO SUGERIDO", "CÓDIGO CONFIRMADO"]
    for j, h in enumerate(H, 1):
        c = ws.cell(2, j, h)
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30
    r = 3
    for marca, arq, curso, sug in sorted(pendentes):
        ws.cell(r, 1, marca)
        ws.cell(r, 2, arq)
        ws.cell(r, 3, curso)
        ws.cell(r, 4, sug)
        c = ws.cell(r, 5, ja.get((marca, arq), ""))
        c.fill = PatternFill("solid", fgColor="FFF2CC")
        r += 1
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = "A2:E%d" % (r - 1)
    for col, w in zip("ABCDE", [13, 58, 46, 20, 22]):
        ws.column_dimensions[col].width = w
    wb.save(caminho_depara)
    return r - 3


def sugerir_codigo(curso, nome_arquivo):
    """tenta achar um codigo no campo Curso: 'ADG - X' ou 'X (CTB)'"""
    m = re.match(r'^\s*([A-Z][A-Z0-9_]{1,14})\s*[-–]\s*\S', curso or '')
    if m:
        return m.group(1)
    m = re.search(r'\(([A-Z][A-Z0-9_]{1,14})\)\s*$', curso or '')
    if m:
        return m.group(1)
    return ""
