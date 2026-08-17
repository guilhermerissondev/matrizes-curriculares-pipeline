# -*- coding: utf-8 -*-
"""Camada 1 — schema/formato de um arquivo isolado.

Verifica que a matriz foi lida corretamente e que os campos estruturais
mínimos estão presentes, antes de qualquer regra de negócio. Roda por
arquivo, logo após `pipeline.parsing.ler_matriz`.
"""
import re
from collections import Counter

from pipeline.qualidade.base import Achado, ERRO, ALERTA

CAMPOS_CABECALHO_OBRIGATORIOS = ("IES", "GRAU", "MODALIDADE_DECL", "REGIME",
                                  "INTEGRALIZACAO_MIN", "INTEGRALIZACAO_MAX")


def avaliar(curso, disciplinas, avisos_parsing, marca, arquivo, padrao_cod_disciplina=None):
    """`padrao_cod_disciplina`: regex opcional (string) validado contra
    COD_DISCIPLINA de cada linha — vem de `config.yaml` por marca; se None,
    a checagem de padrão é pulada (formato ainda não conhecido/parametrizado).
    """
    achados = []

    if curso is None:
        # ler_matriz não conseguiu nem achar a aba/cabeçalho da matriz.
        achados.append(Achado(ERRO, 1, "ARQUIVO_ILEGIVEL", marca,
                               "; ".join(avisos_parsing) or "arquivo não pôde ser lido", arquivo))
        return achados

    for rotulo in CAMPOS_CABECALHO_OBRIGATORIOS:
        if not curso.get(rotulo):
            achados.append(Achado(ALERTA, 1, "CABECALHO_INCOMPLETO", marca,
                                   "campo obrigatório do cabeçalho vazio: %s" % rotulo,
                                   arquivo, curso.get("COD_CURSO", "")))

    regex = re.compile(padrao_cod_disciplina) if padrao_cod_disciplina else None
    for d in disciplinas:
        cod, nome = d.get("COD_DISCIPLINA"), d.get("NOME_DISCIPLINA")
        if not cod:
            # ALERTA, não ERRO: confirmado em 2026-08-14 que esse padrão já
            # existia nas bases antigas (BASE_MATRIZES_* geradas pelo script
            # anterior) nas mesmas proporções (~28-45% das disciplinas) — não
            # é uma regressão nem algo que muda de execução pra execução,
            # então tratar como ERRO bloqueante só gera ruído toda semana.
            # Se depois for confirmado que é uma lacuna real a cobrar de quem
            # alimenta as matrizes, voltar pra ERRO aqui.
            achados.append(Achado(ALERTA, 1, "COD_DISCIPLINA_VAZIO", marca,
                                   "disciplina sem código: %r" % nome, arquivo,
                                   curso.get("COD_CURSO", "")))
        elif regex and not regex.fullmatch(cod):
            achados.append(Achado(ALERTA, 1, "COD_DISCIPLINA_FORA_DO_PADRAO", marca,
                                   "código '%s' não bate com o padrão esperado (%s)"
                                   % (cod, padrao_cod_disciplina),
                                   arquivo, curso.get("COD_CURSO", ""), cod))
        if not nome:
            achados.append(Achado(ERRO, 1, "NOME_DISCIPLINA_VAZIO", marca,
                                   "disciplina sem nome (código %s)" % cod, arquivo,
                                   curso.get("COD_CURSO", ""), cod))

        for campo in ("CH_EAD", "CH_PRES_TEORICA", "CH_PRES_PRATICA",
                      "CH_PROVA_PRESENCIAL", "CH_SINCRONA_MEDIADA", "CH_TOTAL"):
            v = d.get(campo)
            if v is not None and v < 0:
                achados.append(Achado(ERRO, 1, "CH_NEGATIVA", marca,
                                       "%s negativa (%s) em '%s'" % (campo, v, nome),
                                       arquivo, curso.get("COD_CURSO", ""), cod))

    if not disciplinas:
        achados.append(Achado(ERRO, 1, "SEM_DISCIPLINAS", marca,
                               "nenhuma disciplina lida no arquivo", arquivo,
                               curso.get("COD_CURSO", "")))

    # códigos de disciplina repetidos dentro do MESMO curso (ex.: várias
    # disciplinas usando "NOVA" como placeholder) — não impede a leitura,
    # mas quebra a identificação única da disciplina (ver camada 3, que
    # deliberadamente ignora esses códigos na comparação histórica).
    contagem = Counter(d.get("COD_DISCIPLINA") for d in disciplinas if d.get("COD_DISCIPLINA"))
    for cod, n in contagem.items():
        if n > 1:
            achados.append(Achado(ALERTA, 1, "COD_DISCIPLINA_DUPLICADO_NO_CURSO", marca,
                                   "código '%s' usado em %d disciplinas diferentes do mesmo curso "
                                   "(provável placeholder) — comparação histórica por código é pulada para ele"
                                   % (cod, n), arquivo, curso.get("COD_CURSO", ""), cod))

    return achados
