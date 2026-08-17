# -*- coding: utf-8 -*-
"""Camada 3 — consistência entre a execução atual e a execução anterior.

Não bloqueia nada sozinha (tudo aqui é ALERTA, no máximo) — o objetivo é
levantar bandeira para revisão humana quando algo muda de um jeito que
"parece errado" mesmo sendo tecnicamente uma leitura válida da planilha nova.
`historico` é montado pelo orquestrador a partir do snapshot publicado
anterior (ver `pipeline.snapshot`).
"""
from collections import Counter

from pipeline.qualidade.base import Achado, ALERTA

PADRAO_LIMIARES = {
    "variacao_ch_pct": 20,          # % de variação de CH que dispara alerta
    "queda_disciplinas_pct": 30,     # % de queda no nº de disciplinas do curso
}


def avaliar(curso, disciplinas, marca, historico, limiares=None):
    limiares = {**PADRAO_LIMIARES, **(limiares or {})}
    achados = []
    cod = curso.get("COD_CURSO", "")
    arq = curso.get("ARQUIVO", "")
    if not cod:
        return achados  # sem código de curso resolvido, não há o que comparar

    cursos_ant = historico.get("cursos_anteriores", {})
    disc_ant = historico.get("disciplinas_anteriores", {})
    depara_confirmado = historico.get("depara_confirmado", {})

    anterior = cursos_ant.get((marca, cod))
    if anterior:
        ch_a, ch_n = anterior.get("CH_TOTAL") or 0, curso.get("CH_TOTAL") or 0
        if ch_a > 0:
            var_pct = 100.0 * abs(ch_n - ch_a) / ch_a
            if var_pct >= limiares["variacao_ch_pct"]:
                achados.append(Achado(ALERTA, 3, "CH_VARIACAO_ABRUPTA", marca,
                                       "CH total variou %.0f%% desde a execução anterior (%.0fh -> %.0fh)"
                                       % (var_pct, ch_a, ch_n), arq, cod))

        n_a = anterior.get("N_DISCIPLINAS") or 0
        n_n = curso.get("N_DISCIPLINAS") or 0
        if n_a > 0 and n_n < n_a * (1 - limiares["queda_disciplinas_pct"] / 100.0):
            achados.append(Achado(ALERTA, 3, "QUEDA_DISCIPLINAS", marca,
                                   "nº de disciplinas caiu de %d para %d — matriz pode estar incompleta"
                                   % (n_a, n_n), arq, cod))

        for campo in ("IES", "GRAU", "REGIME"):
            va, vn = anterior.get(campo), curso.get(campo)
            if va and vn and va != vn:
                achados.append(Achado(ALERTA, 3, "CAMPO_CABECALHO_MUDOU", marca,
                                       "%s mudou de '%s' para '%s' sem o curso ter sido marcado REMOVIDO+NOVO"
                                       % (campo, va, vn), arq, cod))

    # ---- disciplinas: mesmo código, nome ou CH mudando bruscamente
    # Alguns cursos usam códigos-placeholder genéricos ("NOVA", "Novo código"
    # etc.) repetidos em várias disciplinas distintas do mesmo curso enquanto
    # aguardam um código definitivo — nesses casos COD_DISCIPLINA não
    # identifica unicamente a disciplina, então comparar por essa chave
    # geraria falsos positivos (cada execução "parece" ter trocado nome/CH).
    # Ignoramos códigos duplicados dentro do próprio curso nesta comparação.
    contagem_cod = Counter(d.get("COD_DISCIPLINA") for d in disciplinas if d.get("COD_DISCIPLINA"))
    for d in disciplinas:
        cd = d.get("COD_DISCIPLINA")
        if not cd or contagem_cod[cd] > 1:
            continue
        da = disc_ant.get((marca, cod, cd))
        if not da:
            continue
        if da.get("NOME_DISCIPLINA") and d.get("NOME_DISCIPLINA") and da["NOME_DISCIPLINA"] != d["NOME_DISCIPLINA"]:
            achados.append(Achado(ALERTA, 3, "DISCIPLINA_NOME_MUDOU", marca,
                                   "disciplina %s mudou de nome: '%s' -> '%s'"
                                   % (cd, da["NOME_DISCIPLINA"], d["NOME_DISCIPLINA"]), arq, cod, cd))
        cha, chn = da.get("CH_TOTAL") or 0, d.get("CH_TOTAL") or 0
        if cha > 0 and abs(chn - cha) / cha * 100 >= limiares["variacao_ch_pct"]:
            achados.append(Achado(ALERTA, 3, "DISCIPLINA_CH_MUDOU", marca,
                                   "disciplina %s variou CH de %.0fh para %.0fh" % (cd, cha, chn),
                                   arq, cod, cd))

    # ---- conflito com de-para já confirmado
    # chave "marca::arquivo" — ver o comentário em orquestrador.py sobre por
    # que depara_confirmado é achatado assim ao gravar o manifesto.
    confirmado = depara_confirmado.get("%s::%s" % (marca, arq))
    if confirmado and curso.get("ORIGEM_COD") == "matriz" and cod != confirmado:
        achados.append(Achado(ALERTA, 3, "CONFLITO_DE_PARA", marca,
                               "matriz traz SIGLA INTERNA '%s' mas o de-para já tinha '%s' confirmado para este arquivo"
                               % (cod, confirmado), arq, cod))

    return achados
