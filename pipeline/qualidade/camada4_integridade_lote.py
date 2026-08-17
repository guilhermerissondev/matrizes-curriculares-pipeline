# -*- coding: utf-8 -*-
"""Camada 4 — integridade do lote inteiro (visão consolidada de uma marca).

Roda uma vez por marca, depois que todos os arquivos dela já foram
processados pelas camadas 1-3. É aqui que mora o julgamento de "essa
execução deve ser publicada como válida ou fica marcada para revisão".
"""
from collections import defaultdict

from pipeline.qualidade.base import Achado, ERRO, ALERTA

PADRAO_LIMIARES = {
    "queda_total_pct": 15,           # % de queda no total geral de cursos/disciplinas
    "execucoes_pendente_alerta": 3,  # nº de execuções seguidas PENDENTE antes de alertar
}


def avaliar(marca, cursos, resultado_drive, historico_agregado, limiares=None):
    """
    cursos: lista de curso_dict já resolvidos (após de-para) desta marca
    resultado_drive: retorno de `drive_client.sincronizar_marca` para esta marca
                      (None se a execução rodou em modo local, sem Drive)
    historico_agregado: {
        "total_cursos_anterior": int, "total_disciplinas_anterior": int,
        "pendentes_consecutivos": {arquivo: n_execucoes_seguidas_pendente},
    }
    """
    limiares = {**PADRAO_LIMIARES, **(limiares or {})}
    achados = []

    # ---- COD_CURSO duplicado entre arquivos diferentes
    porcod = defaultdict(list)
    for c in cursos:
        if c.get("COD_CURSO"):
            porcod[c["COD_CURSO"]].append(c["ARQUIVO"])
    for cod, arqs in porcod.items():
        if len(arqs) > 1:
            achados.append(Achado(ERRO, 4, "COD_CURSO_DUPLICADO", marca,
                                   "código '%s' aparece em %d arquivos: %s"
                                   % (cod, len(arqs), ", ".join(sorted(arqs))), cod_curso=cod))

    # ---- download parcial (arquivos listados no Drive mas não baixados/processados)
    if resultado_drive is not None:
        falhas = resultado_drive.get("falhas", [])
        for f in falhas:
            achados.append(Achado(ERRO, 4, "FALHA_DOWNLOAD", marca,
                                   "arquivo listado no Drive mas falhou ao baixar: %s" % f.get("erro", ""),
                                   arquivo=f.get("caminho_relativo", "")))
        if falhas:
            achados.append(Achado(ALERTA, 4, "EXECUCAO_INCOMPLETA", marca,
                                   "%d arquivo(s) não puderam ser baixados nesta execução — "
                                   "NÃO tratar como removidos, ver FALHA_DOWNLOAD" % len(falhas)))

    # ---- queda abrupta no total geral
    total_cursos_ant = historico_agregado.get("total_cursos_anterior")
    if total_cursos_ant:
        queda_pct = 100.0 * (total_cursos_ant - len(cursos)) / total_cursos_ant
        if queda_pct >= limiares["queda_total_pct"]:
            achados.append(Achado(ERRO, 4, "QUEDA_TOTAL_CURSOS", marca,
                                   "total de cursos caiu %.0f%% desde a execução anterior (%d -> %d) — "
                                   "possível falha sistêmica (pasta errada, permissão revogada, Drive fora do ar); "
                                   "revisar antes de considerar este snapshot válido"
                                   % (queda_pct, total_cursos_ant, len(cursos))))

    # ---- cursos PENDENTE há várias execuções seguidas
    pendentes = historico_agregado.get("pendentes_consecutivos", {})
    for arq, n in pendentes.items():
        if n >= limiares["execucoes_pendente_alerta"]:
            achados.append(Achado(ALERTA, 4, "PENDENTE_PERSISTENTE", marca,
                                   "código de curso pendente há %d execuções seguidas — preencher _DE_PARA_CURSOS.xlsx"
                                   % n, arquivo=arq))

    return achados
