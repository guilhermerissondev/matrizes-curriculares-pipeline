# -*- coding: utf-8 -*-
"""Leitura de uma matriz curricular (.xlsx) para (curso, disciplinas, avisos).

Portado de `gerar_base_matrizes.py` sem alteração de lógica — é o parser
tolerante a diferentes layouts de planilha (mapeia colunas por NOME, nunca por
posição) e não deve ser reescrito. A única mudança em relação ao script
original é onde `full`/`rel` vêm de: antes vinham de `os.walk` sobre uma
pasta local já extraída manualmente; agora vêm da árvore que o
`drive_client` baixa, mas a estrutura relativa (MARCA/MODALIDADE/AREA/arquivo)
é a mesma, então a assinatura da função não mudou.
"""
import os
import re
from collections import Counter

import openpyxl

from pipeline.util import sa, txt, num

LIN_CAB_MAX = 60          # ate onde procurar o cabecalho da matriz
VAZIAS_PARA_PARAR = 25    # linhas seguidas em branco encerram a leitura
LIN_MAX_LEITURA = 400     # a matriz mais profunda termina por volta da linha 130

# O template traz linhas de exemplo ("Disciplina 1".."Disciplina 8") para o
# preenchimento de semestres futuros, e uma linha de TOTAL sem nome no fim.
# Nenhuma das duas e disciplina.
RE_PLACEHOLDER = re.compile(r'^\s*disciplina\s*\d+\s*$', re.I)


def _chave_modulo(m):
    """Chave de ordenação pra MODULO_INICIAL: números primeiro (por valor
    numérico, não lexicográfico — senão "10" vem antes de "2"), rótulos
    não-numéricos depois (em ordem alfabética)."""
    return (0, int(m)) if m.isdigit() else (1, m)


# ----------------------------------------------- mapeamento das colunas por nome
# (campo canonico, teste sobre o nome normalizado da coluna)
REGRAS = [
    ("ANO",                 lambda h: h == "ANO"),
    ("SEMESTRE",            lambda h: h == "SEMESTRE"),
    ("MODULO",              lambda h: h == "MODULO"),
    ("SERIE",               lambda h: h == "SERIE"),
    ("ORDEM_NO_MODULO",     lambda h: h.startswith("ORDEM NO MODULO")),
    ("SIGLA_CURSO",         lambda h: "SIGLA INTERNA" in h),
    ("CURRICULO",           lambda h: h == "CURRICULO"),
    ("DOMINIO_FORMATIVO",   lambda h: "DOMINIO FORMATIVO" in h),
    ("COD_DISCIPLINA",      lambda h: "CODIGO DA DISCIPLINA" in h and "EQUIVALENTE" not in h),
    ("NOME_DISCIPLINA",     lambda h: "NOME DA DISCIPLINA" in h and "EQUIVALENTE" not in h and "ORIGEM" not in h),
    ("TIPO",                lambda h: h == "TIPO" or h.startswith("TIPO/")),
    ("SUBTIPO",             lambda h: "SUB TIPO" in h or h == "SUBTIPO"),
    ("PERFIL",              lambda h: h == "PERFIL"),
    ("FIXA",                lambda h: h.startswith("FIXA")),
    ("OFERTA_ESPECIAL",     lambda h: "OFERTA ESPECIAL" in h),
    ("OBSERVACAO",          lambda h: h == "OBSERVACAO"),
    ("QUESTIONARIO_BONUS",  lambda h: "QUESTIONARIO BONUS" in h),
    ("CH_EAD",              lambda h: "ATIVIDADE EAD" in h),
    ("CH_PRES_TEORICA",     lambda h: "PRESENCIAL TEORICA" in h),
    ("CH_PRES_PRATICA",     lambda h: "PRESENCIAL PRATICA" in h),
    ("CH_PROVA_PRESENCIAL", lambda h: "PROVA PRESENCIAL" in h),
    ("CH_SINCRONA_MEDIADA", lambda h: "SINCRONA MEDIADA" in h),
    ("CH_TOTAL",            lambda h: "CH TOTAL" in h),
    ("COD_EQUIVALENTE",     lambda h: "CODIGO DA DISCIPLINA EQUIVALENTE" in h),
    ("NOME_EQUIVALENTE",    lambda h: "NOME DA DISCIPLINA EQUIVALENTE" in h),
]
CAMPOS_CH = ["CH_EAD", "CH_PRES_TEORICA", "CH_PRES_PRATICA",
             "CH_PROVA_PRESENCIAL", "CH_SINCRONA_MEDIADA", "CH_TOTAL"]

# rotulos do bloco de cabecalho (linhas 1..25) -> campo canonico
ROTULOS = {
    "IES": "IES",
    "CURSO": "CURSO_CABECALHO",
    "GRADUACAO": "GRAU",
    "MODALIDADE": "MODALIDADE_DECL",
    "REGIME": "REGIME",
    "INTEGRALIZACAO MINIMA": "INTEGRALIZACAO_MIN",
    "INTEGRALIZACAO MAXIMA": "INTEGRALIZACAO_MAX",
    "ANO DE APROVACAO DA MATRIZ": "ANO_MATRIZ",
    "EDITAL DE INICIO DE OFERTA DA MATRIZ": "EDITAL_OFERTA",
    "ATIVIDADE COMPLEMENTAR": "ATIVIDADE_COMPLEMENTAR",
    "OBSERVACOES": "OBSERVACOES",
}


def mapear_colunas(cab):
    """cab = lista de valores da linha de cabecalho. Devolve {campo: indice 0-based}."""
    m = {}
    for j, v in enumerate(cab):
        h = sa(v)
        if not h:
            continue
        for campo, teste in REGRAS:
            if campo in m:
                continue          # primeira ocorrencia vence
            try:
                if teste(h):
                    m[campo] = j
                    break
            except Exception:
                pass
    return m


def ler_cabecalho(grade):
    """extrai rotulo -> valor do bloco 1..25 (valor = 1a celula nao vazia a direita)"""
    out = {}
    for i in range(min(25, len(grade))):
        row = grade[i]
        for j, v in enumerate(row):
            r = sa(v)
            if r in ROTULOS and ROTULOS[r] not in out:
                for k in range(j + 1, min(j + 9, len(row))):
                    val = txt(row[k])
                    if val:
                        out[ROTULOS[r]] = val
                        break
    # CH total declarada: rotulo com numeros a esquerda
    for i in range(min(25, len(grade))):
        row = grade[i]
        for j, v in enumerate(row):
            if "CARGA HORARIA TOTAL DO CURSO" in sa(v):
                for k in range(j - 1, max(j - 4, -1), -1):
                    n = num(row[k]) if k < len(row) else None
                    if n is not None and n > 1:
                        out["CH_TOTAL_DECLARADA"] = n
                        break
                break
    return out


# ---------------------------------------------------------------- leitura
def ler_matriz(full, rel):
    """devolve (curso_dict, [disciplinas], [avisos])

    full = caminho absoluto do .xlsx
    rel  = caminho relativo dentro da árvore MARCA/MODALIDADE/AREA_CINE/arquivo.xlsx
    """
    avisos = []
    partes = rel.split(os.sep)
    marca = partes[0]
    modalidade = partes[1] if len(partes) > 1 else ""
    area = partes[2] if len(partes) > 3 else ""
    arquivo = partes[-1]

    wb = openpyxl.load_workbook(full, data_only=True, read_only=True)
    ws = None
    for w in wb.worksheets:
        if w.title.strip().lower().startswith("template"):
            ws = w
            break
    if ws is None:
        wb.close()
        return None, [], ["sem aba 'Template'"]

    grade = [list(r) for r in ws.iter_rows(max_row=LIN_MAX_LEITURA, values_only=True)]
    wb.close()

    # cabecalho da matriz
    lin_cab = None
    for i in range(min(LIN_CAB_MAX, len(grade))):
        if any("CODIGO DA DISCIPLINA" in sa(v) for v in grade[i]):
            lin_cab = i
            break
    if lin_cab is None:
        return None, [], ["cabeçalho da matriz não encontrado"]
    if lin_cab != 25:
        avisos.append("cabeçalho na linha %d (esperado 26)" % (lin_cab + 1))

    col = mapear_colunas(grade[lin_cab])
    faltando = [c for c in ("COD_DISCIPLINA", "NOME_DISCIPLINA", "CH_TOTAL") if c not in col]
    if faltando:
        avisos.append("colunas ausentes: " + ", ".join(faltando))

    cab = ler_cabecalho(grade)

    # ---- disciplinas
    disc = []
    vazias = n_modelo = n_total = 0
    ultimo = {"ANO": "", "SEMESTRE": "", "MODULO": "", "SERIE": ""}
    for i in range(lin_cab + 1, len(grade)):
        row = grade[i]

        def g(campo):
            j = col.get(campo)
            return row[j] if (j is not None and j < len(row)) else None

        nome = txt(g("NOME_DISCIPLINA"))
        cod = txt(g("COD_DISCIPLINA"))
        cht = num(g("CH_TOTAL"))
        if not nome and not cod:
            # linha de TOTAL: sem nome e sem codigo, mas com carga horaria
            if cht:
                n_total += 1
            vazias += 1
            if vazias >= VAZIAS_PARA_PARAR:
                break
            continue
        vazias = 0
        if RE_PLACEHOLDER.match(nome):
            n_modelo += 1
            continue
        if not nome:
            # tem código mas não tem nome (célula vazia/mesclada por engano)
            # — NÃO descarta silenciosamente: a linha ainda tem CH e conta
            # pro total do curso. Mantém em `disc` (com nome vazio) pra
            # camada1_schema sinalizar NOME_DISCIPLINA_VAZIO, em vez de a
            # carga horária simplesmente desaparecer da soma sem rastro.
            avisos.append("linha %d: código '%s' sem nome de disciplina — mantida com nome vazio" % (i + 1, cod))
        # ANO/SEMESTRE/MODULO/SERIE vem preenchido so na 1a linha do bloco: propaga
        for k in ("ANO", "SEMESTRE", "MODULO", "SERIE"):
            v = txt(g(k))
            if v:
                ultimo[k] = v
        d = {
            "MARCA": marca, "MODALIDADE": modalidade, "AREA_CINE": area, "ARQUIVO": arquivo,
            "ANO": ultimo["ANO"], "SEMESTRE": ultimo["SEMESTRE"],
            "MODULO": ultimo["MODULO"], "SERIE": ultimo["SERIE"],
            "ORDEM_NO_MODULO": txt(g("ORDEM_NO_MODULO")),
            "SIGLA_CURSO": txt(g("SIGLA_CURSO")), "CURRICULO": txt(g("CURRICULO")),
            "DOMINIO_FORMATIVO": txt(g("DOMINIO_FORMATIVO")),
            "COD_DISCIPLINA": cod, "NOME_DISCIPLINA": nome,
            "TIPO": txt(g("TIPO")), "SUBTIPO": txt(g("SUBTIPO")), "PERFIL": txt(g("PERFIL")),
            "FIXA": txt(g("FIXA")), "OFERTA_ESPECIAL": txt(g("OFERTA_ESPECIAL")),
            "OBSERVACAO": txt(g("OBSERVACAO")),
            "COD_EQUIVALENTE": txt(g("COD_EQUIVALENTE")),
            "NOME_EQUIVALENTE": txt(g("NOME_EQUIVALENTE")),
            "LINHA_ORIGEM": i + 1,
        }
        for c in CAMPOS_CH:
            d[c] = num(g(c))
        d["TEM_CH"] = "SIM" if (d["CH_TOTAL"] or 0) > 0 else "NÃO"
        disc.append(d)
    ignoradas = {"LINHAS_MODELO_IGNORADAS": n_modelo, "LINHAS_TOTAL_IGNORADAS": n_total}

    # ---- identidade do curso
    # sigla puramente numerica e codigo de disciplina que vazou para a coluna
    # do curso — nao serve como codigo de curso.
    siglas = [d["SIGLA_CURSO"] for d in disc
              if d["SIGLA_CURSO"] and not re.fullmatch(r'\d+', d["SIGLA_CURSO"])]
    numericas = {d["SIGLA_CURSO"] for d in disc
                 if d["SIGLA_CURSO"] and re.fullmatch(r'\d+', d["SIGLA_CURSO"])}
    if numericas and not siglas:
        avisos.append("SIGLA INTERNA numérica (%s) — ignorada, parece código de disciplina"
                      % ", ".join(sorted(numericas)[:3]))
    cod_curso = Counter(siglas).most_common(1)[0][0] if siglas else ""
    if len(set(siglas)) > 1:
        avisos.append("mais de uma SIGLA INTERNA no arquivo: " + ", ".join(sorted(set(siglas))[:4]))
    curric = [d["CURRICULO"] for d in disc if d["CURRICULO"]]
    curriculo = Counter(curric).most_common(1)[0][0] if curric else ""

    soma = {c: round(sum(d[c] or 0 for d in disc), 2) for c in CAMPOS_CH}
    curso = {
        "MARCA": marca, "MODALIDADE": modalidade, "AREA_CINE": area, "ARQUIVO": arquivo,
        "COD_CURSO": cod_curso, "ORIGEM_COD": "matriz" if cod_curso else "",
        "CURSO": cab.get("CURSO_CABECALHO", ""), "CURRICULO": curriculo,
        "IES": cab.get("IES", ""), "GRAU": cab.get("GRAU", ""),
        "MODALIDADE_DECL": cab.get("MODALIDADE_DECL", ""), "REGIME": cab.get("REGIME", ""),
        "INTEGRALIZACAO_MIN": cab.get("INTEGRALIZACAO_MIN", ""),
        "INTEGRALIZACAO_MAX": cab.get("INTEGRALIZACAO_MAX", ""),
        "ANO_MATRIZ": cab.get("ANO_MATRIZ", ""), "EDITAL_OFERTA": cab.get("EDITAL_OFERTA", ""),
        "ATIVIDADE_COMPLEMENTAR": cab.get("ATIVIDADE_COMPLEMENTAR", ""),
        "OBSERVACOES": cab.get("OBSERVACOES", ""),
        "N_DISCIPLINAS": sum(1 for d in disc if d["TEM_CH"] == "SIM"),
        "N_DISCIPLINAS_SEM_CH": sum(1 for d in disc if d["TEM_CH"] == "NÃO"),
        "N_SEMESTRES": len({d["SEMESTRE"] for d in disc if d["SEMESTRE"]}),
        "N_MODULOS": len({d["MODULO"] for d in disc if d["MODULO"]}),
        "MODULO_INICIAL": min([d["MODULO"] for d in disc if d["MODULO"]], key=_chave_modulo, default=""),
        "CH_TOTAL_DECLARADA": cab.get("CH_TOTAL_DECLARADA", None),
    }
    curso.update(soma)
    if curso["CH_TOTAL_DECLARADA"] and soma["CH_TOTAL"]:
        curso["DIF_CH_SOMADA_x_DECLARADA"] = round(soma["CH_TOTAL"] - curso["CH_TOTAL_DECLARADA"], 2)
    else:
        curso["DIF_CH_SOMADA_x_DECLARADA"] = None
    curso.update(ignoradas)
    if not disc:
        avisos.append("nenhuma disciplina lida")
    if disc and curso["N_DISCIPLINAS"] == 0:
        avisos.append("nenhuma disciplina com carga horária")
    return curso, disc, avisos
