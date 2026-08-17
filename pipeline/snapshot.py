# -*- coding: utf-8 -*-
"""Histórico versionado por execução: uma pasta datada e imutável por rodada.

Evolui o mecanismo de `manifesto.json`/aba MUDANCAS do script original: além
do hash por arquivo, guarda um "fingerprint" leve por curso e por disciplina
(sem precisar reabrir os .xlsx da execução anterior) para alimentar a camada
3 de qualidade (consistência histórica) e o changelog legível.
"""
import json
import os
import shutil
from collections import Counter, defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from pipeline.util import md5
from pipeline.consolidacao import AZUL, VERDE, BEGE, ROSA

NOME_MANIFESTO = "manifesto.json"
NOME_CHANGELOG = "CHANGELOG.md"
NOME_PONTEIRO_ULTIMO = "ULTIMO.txt"
FORMATO_TS = "%Y-%m-%d_%H%M%S"


# ------------------------------------------------------------- ciclo de vida
def limpar_incompletos(raiz_historico):
    """Remove pastas de execuções anteriores que começaram mas não
    terminaram (sem `manifesto.json` — ex.: processo caiu/foi encerrado no
    meio de um download). Chamar no INÍCIO de uma nova execução, antes de
    `nova_pasta`, como faxina defensiva — o caminho normal já limpa isso na
    hora via try/except em `orquestrador.py`; isso cobre o caso de queda
    abrupta (energia, processo morto) que pula até o `except`."""
    if not os.path.isdir(raiz_historico):
        return []
    todas = [n for n in os.listdir(raiz_historico)
             if os.path.isdir(os.path.join(raiz_historico, n)) and _eh_timestamp(n)]
    incompletas = [n for n in todas if not os.path.exists(os.path.join(raiz_historico, n, NOME_MANIFESTO))]
    for n in incompletas:
        shutil.rmtree(os.path.join(raiz_historico, n), ignore_errors=True)
    return incompletas


def nova_pasta(raiz_historico, agora=None):
    agora = agora or datetime.now()
    ts = agora.strftime(FORMATO_TS)
    pasta = os.path.join(raiz_historico, ts)
    os.makedirs(pasta, exist_ok=False)  # nunca sobrescreve um snapshot já publicado
    return pasta, ts


def listar_snapshots(raiz_historico, apenas_completos=True):
    """`apenas_completos=True` (padrão) ignora pastas de execuções que
    começaram mas não terminaram (ex.: processo caiu no meio de um download
    do Drive) — identificadas pela ausência de `manifesto.json`, gravado só
    no fim da execução. Isso é essencial: uma pasta vazia/incompleta NUNCA
    pode virar "snapshot anterior" da próxima rodada (a comparação
    histórica trataria tudo como NOVO/removido e a camada 4 de integridade
    perderia a base de comparação real)."""
    if not os.path.isdir(raiz_historico):
        return []
    nomes = [n for n in os.listdir(raiz_historico)
             if os.path.isdir(os.path.join(raiz_historico, n)) and _eh_timestamp(n)]
    if apenas_completos:
        nomes = [n for n in nomes if os.path.exists(os.path.join(raiz_historico, n, NOME_MANIFESTO))]
    return sorted(nomes)  # o formato AAAA-MM-DD_HHMMSS ordena lexicograficamente = cronológico


def _eh_timestamp(nome):
    try:
        datetime.strptime(nome, FORMATO_TS)
        return True
    except ValueError:
        return False


def snapshot_anterior(raiz_historico, pasta_atual=None):
    """Devolve o caminho do snapshot publicado mais recente ANTES do atual (ou None)."""
    nomes = listar_snapshots(raiz_historico)
    atual_nome = os.path.basename(pasta_atual) if pasta_atual else None
    anteriores = [n for n in nomes if n != atual_nome]
    if not anteriores:
        return None
    return os.path.join(raiz_historico, anteriores[-1])


def atualizar_ponteiro_ultimo(raiz_historico, pasta_snapshot):
    """Grava um arquivo texto simples apontando pro snapshot mais recente —
    evita symlink (exige privilégio elevado no Windows por padrão)."""
    with open(os.path.join(raiz_historico, NOME_PONTEIRO_ULTIMO), "w", encoding="utf-8") as f:
        f.write(os.path.basename(pasta_snapshot))


# ------------------------------------------------------------------ manifesto
def carregar_manifesto(pasta_snapshot):
    if not pasta_snapshot:
        return {}
    caminho = os.path.join(pasta_snapshot, NOME_MANIFESTO)
    if not os.path.exists(caminho):
        return {}
    try:
        return json.load(open(caminho, encoding="utf-8"))
    except Exception:
        return {}


def gravar_manifesto(pasta_snapshot, manifesto):
    json.dump(manifesto, open(os.path.join(pasta_snapshot, NOME_MANIFESTO), "w", encoding="utf-8"),
               ensure_ascii=False, indent=1)


def montar_manifesto(arquivos_processados, cursos, disciplinas, depara_confirmado, pendentes_consecutivos):
    """
    arquivos_processados: {rel_path: {"md5_local", "md5_drive", "tamanho", "modificado",
                                       "cod_curso", "n_disciplinas", "ch_total"}}
    """
    # totais explícitos por marca — NÃO derivar de len(m["cursos"]) por marca:
    # cursos com COD_CURSO vazio (PENDENTE) não entram no dict "cursos"
    # (colidiriam todos na mesma chave "MARCA::"), então contar pelo dict
    # subcontaria o total real de cursos processados nesta execução.
    total_cursos_por_marca = Counter(c["MARCA"] for c in cursos)
    total_disc_por_marca = Counter(d["MARCA"] for d in disciplinas)
    m = {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "arquivos": arquivos_processados,
        "total_cursos_por_marca": dict(total_cursos_por_marca),
        "total_disciplinas_por_marca": dict(total_disc_por_marca),
        "cursos": {}, "disciplinas": {},
        "depara_confirmado": depara_confirmado,
        "pendentes_consecutivos": pendentes_consecutivos,
    }
    for c in cursos:
        if not c.get("COD_CURSO"):
            continue  # PENDENTE: sem código não dá pra rastrear entre execuções por chave
        chave = "%s::%s" % (c["MARCA"], c["COD_CURSO"])
        m["cursos"][chave] = {"CH_TOTAL": c.get("CH_TOTAL"), "N_DISCIPLINAS": c.get("N_DISCIPLINAS"),
                               "IES": c.get("IES"), "GRAU": c.get("GRAU"), "REGIME": c.get("REGIME"),
                               "ARQUIVO": c.get("ARQUIVO")}
    # Códigos-placeholder ("NOVA" etc.) repetidos várias vezes dentro do
    # mesmo curso não identificam uma disciplina de forma única — não
    # gravamos fingerprint pra eles (ver mesmo cuidado em
    # qualidade/camada3_consistencia_historica.py), senão a comparação da
    # próxima execução gera falsos positivos de "nome/CH mudou".
    contagem_por_curso = defaultdict(Counter)
    for d in disciplinas:
        if d.get("COD_CURSO") and d.get("COD_DISCIPLINA"):
            contagem_por_curso[(d["MARCA"], d["COD_CURSO"])][d["COD_DISCIPLINA"]] += 1
    for d in disciplinas:
        if not d.get("COD_CURSO") or not d.get("COD_DISCIPLINA"):
            continue
        if contagem_por_curso[(d["MARCA"], d["COD_CURSO"])][d["COD_DISCIPLINA"]] > 1:
            continue
        chave = "%s::%s::%s" % (d["MARCA"], d["COD_CURSO"], d["COD_DISCIPLINA"])
        m["disciplinas"][chave] = {"NOME_DISCIPLINA": d.get("NOME_DISCIPLINA"), "CH_TOTAL": d.get("CH_TOTAL")}
    return m


def construir_contexto_historico(manifesto_anterior, marca):
    """Monta o `historico` consumido pela camada 3/4 de qualidade a partir do
    manifesto do snapshot anterior — sem reabrir nenhum .xlsx antigo."""
    cursos_ant = {}
    for chave, v in manifesto_anterior.get("cursos", {}).items():
        m, cod = chave.split("::", 1)
        if m == marca:
            cursos_ant[(marca, cod)] = v
    disc_ant = {}
    for chave, v in manifesto_anterior.get("disciplinas", {}).items():
        partes = chave.split("::", 2)
        if len(partes) == 3 and partes[0] == marca:
            disc_ant[(marca, partes[1], partes[2])] = v
    # pendentes_consecutivos é uma chave global (ver _chave_arquivo: já
    # inclui a marca como primeiro segmento "::") — filtra só as desta marca.
    pendentes = {k: n for k, n in manifesto_anterior.get("pendentes_consecutivos", {}).items()
                 if k.startswith(marca + "::")}
    return {
        "cursos_anteriores": cursos_ant,
        "disciplinas_anteriores": disc_ant,
        "depara_confirmado": manifesto_anterior.get("depara_confirmado", {}),
        "total_cursos_anterior": manifesto_anterior.get("total_cursos_por_marca", {}).get(marca),
        "total_disciplinas_anterior": manifesto_anterior.get("total_disciplinas_por_marca", {}).get(marca),
        "pendentes_consecutivos": pendentes,
    }


def _chave_arquivo(c):
    """Chave única de arquivo: MARCA/ARQUIVO sozinho NÃO é único (o mesmo
    nome de template pode se repetir em áreas CINE diferentes da mesma
    marca) — usa marca+modalidade+área+arquivo, no mesmo estilo "::" das
    outras chaves compostas deste módulo."""
    return "%s::%s::%s::%s" % (c.get("MARCA", ""), c.get("MODALIDADE", ""),
                                c.get("AREA_CINE", ""), c.get("ARQUIVO", ""))


def calcular_pendentes_consecutivos(cursos, manifesto_anterior):
    """Incrementa a contagem de execuções seguidas em que um arquivo segue
    com código de curso PENDENTE; reseta pra 0 assim que resolve."""
    anteriores = manifesto_anterior.get("pendentes_consecutivos", {})
    novo = {}
    for c in cursos:
        if c.get("ORIGEM_COD") == "PENDENTE":
            chave = _chave_arquivo(c)
            novo[chave] = anteriores.get(chave, 0) + 1
    return novo


# ------------------------------------------------------------------ mudanças
def calcular_mudancas(marca, arquivos_processados, manifesto_anterior, falhas_rel=None):
    """Compara arquivos desta execução com o manifesto anterior — mesma
    lógica de classificação NOVO/SUBSTITUÍDO/REMOVIDO/inalterado do script
    original, agora baseada em `md5_local` do conteúdo do arquivo.

    `falhas_rel`: caminhos relativos de arquivos que EXISTEM (no Drive ou na
    pasta local) mas não puderam ser baixados/lidos nesta execução — NUNCA
    devem cair no ramo REMOVIDO (isso afirmaria que o arquivo sumiu, o que é
    falso e pode levar a tratar um curso como descontinuado por engano).
    Viram a situação própria FALHA_LEITURA."""
    falhas_rel = falhas_rel or set()
    arq_ant = {k: v for k, v in manifesto_anterior.get("arquivos", {}).items() if k.startswith(marca + os.sep)}
    mud = []
    for rel, m in sorted(arquivos_processados.items()):
        if not rel.startswith(marca + os.sep):
            continue
        arq = os.path.basename(rel)
        a = arq_ant.get(rel)
        if a is None:
            mud.append({"SITUACAO": "NOVO", "MARCA": marca, "ARQUIVO": arq, "COD_CURSO": m.get("cod_curso"),
                        "N_DISCIPLINAS": m.get("n_disciplinas"), "CH_TOTAL": m.get("ch_total"),
                        "CH_TOTAL_ANTERIOR": None, "VARIACAO_CH": None,
                        "DETALHE": "arquivo não existia na execução anterior"})
        elif a.get("md5_local") != m.get("md5_local"):
            va = (m.get("ch_total") or 0) - (a.get("ch_total") or 0)
            mud.append({"SITUACAO": "SUBSTITUÍDO", "MARCA": marca, "ARQUIVO": arq, "COD_CURSO": m.get("cod_curso"),
                        "N_DISCIPLINAS": m.get("n_disciplinas"), "CH_TOTAL": m.get("ch_total"),
                        "CH_TOTAL_ANTERIOR": a.get("ch_total"), "VARIACAO_CH": round(va, 2),
                        "DETALHE": "disciplinas: %s → %s" % (a.get("n_disciplinas"), m.get("n_disciplinas"))})
        else:
            mud.append({"SITUACAO": "inalterado", "MARCA": marca, "ARQUIVO": arq, "COD_CURSO": m.get("cod_curso"),
                        "N_DISCIPLINAS": m.get("n_disciplinas"), "CH_TOTAL": m.get("ch_total"),
                        "CH_TOTAL_ANTERIOR": a.get("ch_total"), "VARIACAO_CH": 0, "DETALHE": ""})
    processados_rel = set(k for k in arquivos_processados if k.startswith(marca + os.sep))
    ausentes = set(arq_ant) - processados_rel
    for rel in sorted(ausentes & falhas_rel):
        a = arq_ant[rel]
        mud.append({"SITUACAO": "FALHA_LEITURA", "MARCA": marca, "ARQUIVO": os.path.basename(rel),
                    "COD_CURSO": a.get("cod_curso"), "N_DISCIPLINAS": None, "CH_TOTAL": None,
                    "CH_TOTAL_ANTERIOR": a.get("ch_total"), "VARIACAO_CH": None,
                    "DETALHE": "arquivo existe mas não pôde ser baixado/lido nesta execução — "
                               "NÃO é remoção, ver log/aba QUALIDADE"})
    for rel in sorted(ausentes - falhas_rel):
        a = arq_ant[rel]
        mud.append({"SITUACAO": "REMOVIDO", "MARCA": marca, "ARQUIVO": os.path.basename(rel),
                    "COD_CURSO": a.get("cod_curso"), "N_DISCIPLINAS": None, "CH_TOTAL": None,
                    "CH_TOTAL_ANTERIOR": a.get("ch_total"), "VARIACAO_CH": None,
                    "DETALHE": "estava na execução anterior e não está mais na pasta"})
    return mud


# ------------------------------------------------------------------ changelog
def gerar_changelog(pasta_snapshot, ts, resultado_por_marca, tipo="manual", motivo=""):
    """`resultado_por_marca`: {marca: {"mudancas": [...], "achados": [Achado, ...],
                                        "n_cursos": int, "n_disciplinas": int}}"""
    linhas = ["# Changelog — execução %s" % ts, ""]
    linhas.append("**Tipo:** %s" % tipo.upper())
    if motivo:
        linhas.append("**Motivo:** %s" % motivo)
    linhas.append("")
    for marca, r in sorted(resultado_por_marca.items()):
        mud = r["mudancas"]
        cont = {"NOVO": 0, "SUBSTITUÍDO": 0, "REMOVIDO": 0, "inalterado": 0, "FALHA_LEITURA": 0}
        for m in mud:
            cont[m["SITUACAO"]] = cont.get(m["SITUACAO"], 0) + 1
        achados = r["achados"]
        n_erro = sum(1 for a in achados if a.severidade == "ERRO")
        n_alerta = sum(1 for a in achados if a.severidade == "ALERTA")
        linhas.append("## %s" % marca)
        linhas.append("- cursos: %d | disciplinas: %d" % (r["n_cursos"], r["n_disciplinas"]))
        linhas.append("- arquivos: %d novos, %d substituídos, %d removidos, %d inalterados%s"
                       % (cont["NOVO"], cont["SUBSTITUÍDO"], cont["REMOVIDO"], cont["inalterado"],
                          (", %d com FALHA DE LEITURA/DOWNLOAD (não são removidos!)" % cont["FALHA_LEITURA"])
                          if cont["FALHA_LEITURA"] else ""))
        linhas.append("- qualidade: %d ERRO, %d ALERTA" % (n_erro, n_alerta))
        if n_erro:
            linhas.append("")
            linhas.append("### Erros que pedem revisão antes de considerar este snapshot confiável")
            erros = [a for a in achados if a.severidade == "ERRO"]
            por_regra = Counter(a.regra for a in erros)
            linhas.append("Por regra: " + ", ".join("%s (%d)" % (r, n) for r, n in por_regra.most_common()))
            linhas.append("")
            LIMITE = 25
            for a in erros[:LIMITE]:
                linhas.append("- `%s` (%s / %s): %s" % (a.regra, a.arquivo or a.marca, a.cod_curso, a.mensagem))
            if len(erros) > LIMITE:
                linhas.append("- ... +%d erro(s) — ver aba QUALIDADE de BASE_MATRIZES_%s.xlsx para a lista completa"
                               % (len(erros) - LIMITE, marca))
        substituidos_ou_novos = [m for m in mud if m["SITUACAO"] in ("SUBSTITUÍDO", "NOVO", "REMOVIDO")]
        if substituidos_ou_novos:
            linhas.append("")
            linhas.append("### Arquivos alterados")
            for m in substituidos_ou_novos:
                linhas.append("- **%s** — %s (%s)" % (m["SITUACAO"], m["ARQUIVO"], m.get("DETALHE", "")))
        linhas.append("")
    texto = "\n".join(linhas)
    with open(os.path.join(pasta_snapshot, NOME_CHANGELOG), "w", encoding="utf-8") as f:
        f.write(texto)
    return texto


# ------------------------------------------------------------- de-para/retenção
def copiar_depara_usado(pasta_snapshot, caminho_depara):
    if os.path.exists(caminho_depara):
        shutil.copy2(caminho_depara, os.path.join(pasta_snapshot, "de_para_usado.xlsx"))


def arquivar_raw(pasta_snapshot, arquivos):
    """Copia (não move) cada matriz-fonte usada nesta execução pra dentro de
    `pasta_snapshot/RAW/<caminho_relativo>` — rastreabilidade real: pra
    qualquer snapshot passado dá pra abrir exatamente os arquivos brutos que
    geraram aquela base, mesmo que a versão no Drive/`_staging` já tenha
    mudado depois. `arquivos` = [(caminho_absoluto, caminho_relativo), ...].
    Falha em copiar UM arquivo não deve derrubar a execução inteira — só
    fica registrado e o snapshot segue sem aquele arquivo em RAW/.
    """
    falhas = []
    for full, rel in arquivos:
        destino = os.path.join(pasta_snapshot, "RAW", rel)
        try:
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            shutil.copy2(full, destino)
        except OSError as e:
            falhas.append((rel, str(e)))
    return falhas


def aplicar_retencao(raiz_historico, manter_raw_ultimos_n=8):
    """Remove a subpasta RAW/ (matrizes brutas baixadas) dos snapshots mais
    antigos que os últimos N, mantendo os artefatos consolidados (xlsx/csv/
    json/manifesto/changelog) por prazo indefinido — são leves; RAW é que
    cresce rápido."""
    nomes = listar_snapshots(raiz_historico)
    antigos = nomes[:-manter_raw_ultimos_n] if manter_raw_ultimos_n > 0 else nomes
    removidos = []
    for nome in antigos:
        raw = os.path.join(raiz_historico, nome, "RAW")
        if os.path.isdir(raw):
            shutil.rmtree(raw, ignore_errors=True)
            removidos.append(nome)
    return removidos


# --------------------------------------------------------------------- índice
NOME_INDICE = "INDICE.xlsx"
COLS_INDICE = ["TIMESTAMP", "TIPO", "MOTIVO", "MARCA", "SITUACAO", "CURSOS", "DISCIPLINAS",
               "ERROS", "ALERTAS", "ARQ_NOVOS", "ARQ_SUBSTITUIDOS", "ARQ_REMOVIDOS", "DURACAO_S"]
_COR_SITUACAO = {"FALHA_SISTEMICA": ROSA, "COM_ACHADOS": BEGE, "OK": VERDE}


def atualizar_indice(raiz_historico, ts, tipo, motivo, resultado_por_marca, duracao_s):
    """Acrescenta uma linha por marca em `_historico/INDICE.xlsx` — visão
    consolidada de TODAS as execuções (uma ao lado da outra), pra identificar
    rápido qual rodada teve problema, quando, e se foi agendada ou manual/
    urgente, sem precisar abrir cada snapshot. Nunca reescreve linhas
    antigas, só acrescenta (o índice em si não é um snapshot versionado —
    é o painel de navegação por cima deles)."""
    caminho = os.path.join(raiz_historico, NOME_INDICE)
    if os.path.exists(caminho):
        wb = openpyxl.load_workbook(caminho)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "INDICE"
        for j, h in enumerate(COLS_INDICE, 1):
            c = ws.cell(1, j, h)
            c.fill = AZUL
            c.font = Font(bold=True, color="FFFFFF", size=9)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 20

    for marca, r in sorted(resultado_por_marca.items()):
        achados = r.get("achados", [])
        n_erro = sum(1 for a in achados if a.severidade == "ERRO")
        n_alerta = sum(1 for a in achados if a.severidade == "ALERTA")
        n_sistemico = sum(1 for a in achados if a.severidade == "ERRO" and a.camada == 4)
        situacao = "FALHA_SISTEMICA" if n_sistemico else ("COM_ACHADOS" if (n_erro or n_alerta) else "OK")
        cont = Counter(m["SITUACAO"] for m in r.get("mudancas", []))
        linha = ws.max_row + 1
        valores = [ts, tipo, motivo or "", marca, situacao, r.get("n_cursos", 0), r.get("n_disciplinas", 0),
                   n_erro, n_alerta, cont.get("NOVO", 0), cont.get("SUBSTITUÍDO", 0), cont.get("REMOVIDO", 0),
                   round(duracao_s, 1)]
        fill = _COR_SITUACAO[situacao]
        for j, v in enumerate(valores, 1):
            c = ws.cell(linha, j, v)
            c.font = Font(size=9)
            c.fill = fill

    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(COLS_INDICE)), ws.max_row)
    larguras = {"TIMESTAMP": 20, "TIPO": 12, "MOTIVO": 34, "MARCA": 14, "SITUACAO": 16}
    for j, h in enumerate(COLS_INDICE, 1):
        ws.column_dimensions[get_column_letter(j)].width = larguras.get(h, 13)
    wb.save(caminho)
