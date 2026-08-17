# -*- coding: utf-8 -*-
"""Gravação dos artefatos de saída: BASE_MATRIZES_<MARCA>.xlsx, CSV e JSON.

`escreve_aba` e as listas `COLS_*` são portadas sem alteração de
`gerar_base_matrizes.py`. Novidade: aba `QUALIDADE` com os achados
estruturados das 4 camadas (severidade/camada/regra), substituindo o antigo
texto livre `AVISOS` como principal fonte de verdade — `CONFERENCIA` continua
existindo por compatibilidade, mas passa a ser alimentada a partir dos
mesmos achados.
"""
import csv
import json
import os
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

AZUL = PatternFill("solid", fgColor="1F3864")
BEGE = PatternFill("solid", fgColor="FFF2CC")
ROSA = PatternFill("solid", fgColor="FFC7CE")
VERDE = PatternFill("solid", fgColor="E2EFDA")
CINZA = PatternFill("solid", fgColor="D9D9D9")
AZUL_CLARO = PatternFill("solid", fgColor="DDEBF7")
NUMF = '#,##0.##'

COLS_DISC = ["MARCA", "MODALIDADE", "AREA_CINE", "ARQUIVO", "COD_CURSO", "CURSO", "CURRICULO",
             "IES", "GRAU", "REGIME", "ANO", "SEMESTRE", "MODULO", "SERIE", "ORDEM_NO_MODULO",
             "COD_DISCIPLINA", "NOME_DISCIPLINA", "TIPO", "SUBTIPO", "PERFIL",
             "DOMINIO_FORMATIVO", "FIXA", "OFERTA_ESPECIAL", "OBSERVACAO",
             "CH_EAD", "CH_PRES_TEORICA", "CH_PRES_PRATICA", "CH_PROVA_PRESENCIAL",
             "CH_SINCRONA_MEDIADA", "CH_TOTAL", "COD_EQUIVALENTE", "NOME_EQUIVALENTE",
             "LINHA_ORIGEM"]
COLS_CURSO = ["MARCA", "MODALIDADE", "AREA_CINE", "COD_CURSO", "ORIGEM_COD", "CURSO", "CURRICULO",
              "IES", "GRAU", "MODALIDADE_DECL", "REGIME", "INTEGRALIZACAO_MIN", "INTEGRALIZACAO_MAX",
              "ANO_MATRIZ", "EDITAL_OFERTA", "ATIVIDADE_COMPLEMENTAR",
              "N_DISCIPLINAS", "N_DISCIPLINAS_SEM_CH", "N_SEMESTRES", "N_MODULOS", "MODULO_INICIAL",
              "CH_EAD", "CH_PRES_TEORICA", "CH_PRES_PRATICA", "CH_PROVA_PRESENCIAL",
              "CH_SINCRONA_MEDIADA", "CH_TOTAL", "CH_TOTAL_DECLARADA",
              "DIF_CH_SOMADA_x_DECLARADA", "ARQUIVO", "OBSERVACOES"]
COLS_CONF = ["MARCA", "MODALIDADE", "ARQUIVO", "COD_CURSO", "ORIGEM_COD", "N_DISCIPLINAS",
             "CH_TOTAL", "CH_TOTAL_DECLARADA", "DIF_CH_SOMADA_x_DECLARADA",
             "LINHAS_MODELO_IGNORADAS", "LINHAS_TOTAL_IGNORADAS",
             "COD_DUPLICADO_COM", "AVISOS"]
COLS_MUD = ["SITUACAO", "MARCA", "ARQUIVO", "COD_CURSO", "N_DISCIPLINAS", "CH_TOTAL",
            "CH_TOTAL_ANTERIOR", "VARIACAO_CH", "DETALHE"]
COLS_QUALIDADE = ["SEVERIDADE", "CAMADA", "REGRA", "MARCA", "ARQUIVO", "COD_CURSO",
                   "COD_DISCIPLINA", "MENSAGEM"]
LARG = {"ARQUIVO": 50, "CURSO": 44, "NOME_DISCIPLINA": 46, "AREA_CINE": 34, "IES": 30,
        "CURRICULO": 22, "NOME_EQUIVALENTE": 40, "COD_EQUIVALENTE": 30, "OBSERVACAO": 26,
        "OBSERVACOES": 30, "AVISOS": 60, "DETALHE": 46, "COD_DUPLICADO_COM": 40,
        "MENSAGEM": 70}


def escreve_aba(wb, nome, colunas, linhas, larguras=None, destaques=None):
    ws = wb.create_sheet(nome)
    ws.sheet_view.showGridLines = False
    for j, h in enumerate(colunas, 1):
        c = ws.cell(1, j, h)
        c.fill = AZUL
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    for i, reg in enumerate(linhas, 2):
        for j, h in enumerate(colunas, 1):
            v = reg.get(h)
            c = ws.cell(i, j, v)
            c.font = Font(size=9)
            if isinstance(v, float):
                c.number_format = NUMF
        if destaques:
            f = destaques(reg)
            if f:
                for j in range(1, len(colunas) + 1):
                    ws.cell(i, j).fill = f
    ws.freeze_panes = "A2"
    if linhas:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(colunas)), len(linhas) + 1)
    for j, h in enumerate(colunas, 1):
        w = (larguras or {}).get(h, min(max(11, len(h) + 2), 34))
        ws.column_dimensions[get_column_letter(j)].width = w
    return ws


def _cor_severidade(reg):
    return {"ERRO": ROSA, "ALERTA": BEGE, "INFO": AZUL_CLARO}.get(reg.get("SEVERIDADE"))


def gravar_workbook(caminho, marca, cursos, disciplinas, confs, mudancas, achados_qualidade):
    """Grava BASE_MATRIZES_<MARCA>.xlsx com as 5 abas."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    escreve_aba(wb, "DISCIPLINAS", COLS_DISC, disciplinas, LARG)
    escreve_aba(wb, "CURSOS", COLS_CURSO, cursos, LARG,
                lambda r: ROSA if r.get("DIF_CH_SOMADA_x_DECLARADA") not in (None, 0) else None)
    escreve_aba(wb, "CONFERENCIA", COLS_CONF, confs, LARG,
                lambda r: ROSA if (r.get("AVISOS") or r.get("COD_DUPLICADO_COM")) else VERDE)
    ordem_situacao = {"FALHA_LEITURA": 0, "NOVO": 1, "SUBSTITUÍDO": 2, "REMOVIDO": 3, "inalterado": 4}
    mud_ordenado = sorted(mudancas, key=lambda x: (ordem_situacao.get(x["SITUACAO"], 9), x["ARQUIVO"]))
    escreve_aba(wb, "MUDANCAS", COLS_MUD, mud_ordenado, LARG,
                lambda r: {"NOVO": VERDE, "SUBSTITUÍDO": BEGE, "REMOVIDO": ROSA,
                           "FALHA_LEITURA": ROSA}.get(r["SITUACAO"]))
    ordem_sev = {"ERRO": 0, "ALERTA": 1, "INFO": 2}
    qual_ordenado = sorted((a.como_linha() for a in achados_qualidade),
                           key=lambda r: (ordem_sev.get(r["SEVERIDADE"], 9), r["CAMADA"], r["REGRA"]))
    escreve_aba(wb, "QUALIDADE", COLS_QUALIDADE, qual_ordenado, LARG, _cor_severidade)
    wb.save(caminho)


def gravar_csv_disciplinas(caminho, disciplinas):
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLS_DISC, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(disciplinas)


def _chave_arquivo(r):
    # ARQUIVO sozinho não é único (mesmo nome de template pode se repetir em
    # áreas CINE diferentes) — inclui MODALIDADE/AREA_CINE pra não misturar
    # disciplinas de dois cursos distintos que só coincidem de nome.
    return (r["MODALIDADE"], r["AREA_CINE"], r["ARQUIVO"])


def gravar_json_matrizes(caminho, marca, cursos, disciplinas):
    porarq = defaultdict(list)
    for d in disciplinas:
        porarq[_chave_arquivo(d)].append({k: d[k] for k in COLS_DISC
                                          if k not in ("MARCA", "MODALIDADE", "AREA_CINE", "ARQUIVO",
                                                       "COD_CURSO", "CURSO", "IES", "GRAU", "REGIME")})
    payload = {"marca": marca, "gerado_em": datetime.now().isoformat(timespec="seconds"),
               "cursos": [dict(c, DISCIPLINAS=porarq.get(_chave_arquivo(c), [])) for c in cursos]}
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
