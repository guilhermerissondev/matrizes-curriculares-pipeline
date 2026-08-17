# -*- coding: utf-8 -*-
"""Camada 2 — regras de negócio acadêmicas/regulatórias, por curso.

Os limiares NÃO ficam hardcoded aqui: vêm de `config/regras_negocio.yaml`
(carregado pelo orquestrador e passado como `regras`) porque são parâmetros
institucionais/regulatórios (tetos de EAD, piso de carga horária por grau)
que mudam com a legislação (MEC/CNE) e precisam ser revisados por alguém da
área acadêmica — os valores que vêm por padrão no yaml são conservadores e
estão marcados como "AJUSTAR" até essa validação acontecer.
"""
import re

from pipeline.util import sa
from pipeline.qualidade.base import Achado, ERRO, ALERTA

_RE_NUM_INICIAL = re.compile(r'(\d+)')


def _extrair_numero(texto):
    m = _RE_NUM_INICIAL.search(texto or "")
    return int(m.group(1)) if m else None


def avaliar(curso, disciplinas, marca, regras=None):
    regras = regras or {}
    achados = []
    cod = curso.get("COD_CURSO", "")
    arq = curso.get("ARQUIVO", "")

    # ---- CH somada x declarada, com limiar (o script original só calculava
    # a diferença; aqui ela vira ERRO/ALERTA conforme a magnitude).
    dif = curso.get("DIF_CH_SOMADA_x_DECLARADA")
    if dif is not None:
        tol = regras.get("tolerancia_dif_ch_horas", {})
        alerta_h = tol.get("alerta", 10)
        erro_h = tol.get("erro", 40)
        if abs(dif) >= erro_h:
            achados.append(Achado(ERRO, 2, "CH_DIVERGENTE", marca,
                                   "CH somada difere %.1fh da declarada (limite ERRO: %sh)" % (dif, erro_h),
                                   arq, cod))
        elif abs(dif) >= alerta_h:
            achados.append(Achado(ALERTA, 2, "CH_DIVERGENTE", marca,
                                   "CH somada difere %.1fh da declarada (limite ALERTA: %sh)" % (dif, alerta_h),
                                   arq, cod))

    # ---- % EAD sobre o total, vs. teto regulatório por modalidade
    ch_total = curso.get("CH_TOTAL") or 0
    ch_ead = curso.get("CH_EAD") or 0
    tetos = regras.get("teto_pct_ead_por_modalidade", {})
    modalidade = sa(curso.get("MODALIDADE", ""))
    teto = tetos.get(modalidade)
    if ch_total and teto is not None:
        pct_ead = 100.0 * ch_ead / ch_total
        if pct_ead > teto:
            achados.append(Achado(ERRO, 2, "PCT_EAD_ACIMA_DO_TETO", marca,
                                   "%.1f%% da CH é EAD, acima do teto de %s%% para modalidade %s"
                                   % (pct_ead, teto, curso.get("MODALIDADE", "")),
                                   arq, cod))

    # ---- CH total abaixo do piso regulatório por grau
    pisos = regras.get("piso_ch_por_grau", {})
    grau = sa(curso.get("GRAU", ""))
    piso = next((v for k, v in pisos.items() if sa(k) in grau or grau in sa(k)), None)
    if ch_total and piso is not None and ch_total < piso:
        achados.append(Achado(ERRO, 2, "CH_ABAIXO_DO_PISO", marca,
                               "CH total %.0fh abaixo do piso de %sh para o grau '%s'"
                               % (ch_total, piso, curso.get("GRAU", "")),
                               arq, cod))

    # ---- integralização coerente com nº de semestres/módulos (heurística —
    # os campos de integralização são texto livre, então isso é ALERTA, nunca ERRO)
    n_min = _extrair_numero(curso.get("INTEGRALIZACAO_MIN"))
    n_periodos = curso.get("N_SEMESTRES") or curso.get("N_MODULOS") or 0
    if n_min and n_periodos and n_min > n_periodos * 12:
        # heurística grosseira: integralização mínima em meses não deveria
        # superar de longe semestres*12 nem módulos*mês-equivalente
        achados.append(Achado(ALERTA, 2, "INTEGRALIZACAO_INCOERENTE", marca,
                               "integralização mínima declarada (%s) parece incoerente com %d período(s) na matriz"
                               % (curso.get("INTEGRALIZACAO_MIN"), n_periodos),
                               arq, cod))

    # ---- curso sem nenhuma disciplina com carga horária
    if curso.get("N_DISCIPLINAS", 0) == 0:
        achados.append(Achado(ERRO, 2, "CURSO_SEM_CH", marca,
                               "curso não tem nenhuma disciplina com carga horária", arq, cod))

    return achados
