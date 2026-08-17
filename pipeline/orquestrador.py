# -*- coding: utf-8 -*-
"""Ponto de entrada único do pipeline — chamado pelo `run_pipeline.bat`
(agendado no Windows Task Scheduler) ou manualmente:

    python -m pipeline.orquestrador                  # baixa do Drive e roda tudo
    python -m pipeline.orquestrador --local <pasta>   # usa uma pasta já extraída
                                                        # localmente (MARCA/MODALIDADE/
                                                        # AREA_CINE/arquivo.xlsx),
                                                        # sem tocar o Drive — útil
                                                        # para testar o parsing/
                                                        # qualidade isoladamente.

Liga as 7 etapas do plano: download -> extração/normalização -> parsing ->
validação em camadas -> consolidação -> snapshot/changelog -> relatório.
"""
import argparse
import os
import shutil
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, RAIZ)

from pipeline import consolidacao, snapshot
from pipeline.depara import carregar_depara, gravar_depara, sugerir_codigo
from pipeline.parsing import ler_matriz
from pipeline.progresso import BarraProgresso
from pipeline.qualidade import camada1_schema, camada2_regras_negocio
from pipeline.qualidade import camada3_consistencia_historica, camada4_integridade_lote
from pipeline.qualidade.base import ordenar, tem_erro
from pipeline.util import md5

CAMINHO_DEPARA = os.path.join(RAIZ, "_DE_PARA_CURSOS.xlsx")
RAIZ_HISTORICO = os.path.join(RAIZ, "_historico")
RAIZ_LOG = os.path.join(RAIZ, "_log")
RAIZ_STAGING = os.path.join(RAIZ, "_staging")


class Logger:
    """Tee simples: tudo que é `print`ado também vai pro arquivo de log da execução."""

    def __init__(self, caminho):
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        self._f = open(caminho, "w", encoding="utf-8")

    def print(self, *args):
        msg = " ".join(str(a) for a in args)
        print(msg)
        self._f.write(msg + "\n")
        self._f.flush()

    def close(self):
        self._f.close()


def carregar_config(caminho):
    try:
        import yaml
    except ImportError:
        raise SystemExit(
            "pyyaml não está instalado. Rode: pip install -r requirements.txt")
    if not os.path.exists(caminho):
        raise SystemExit("config não encontrada: %s" % caminho)
    with open(caminho, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def coletar_arquivos_local(pasta_base):
    """Modo --local: varre uma árvore já extraída MARCA/MODALIDADE/AREA_CINE/*.xlsx."""
    arquivos = []
    for dp, dirs, files in os.walk(pasta_base):
        dirs[:] = [d for d in dirs if not d.startswith("_")]
        for fn in sorted(files):
            if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                full = os.path.join(dp, fn)
                rel = os.path.relpath(full, pasta_base)
                arquivos.append((full, rel))
    return arquivos, None  # None = sem resultado_drive (não roda checagem de download parcial)


def coletar_arquivos_drive(marcas_cfg, api_key, manifesto_anterior, barra=None):
    from pipeline.drive_client import sincronizar_marca
    arquivos = []
    resultados_drive = {}
    for marca, cfg in marcas_cfg.items():
        folder_id = cfg.get("drive_folder_id")
        if not folder_id or "COLOQUE_AQUI" in str(folder_id):
            raise SystemExit(
                "config.yaml: marcas.%s.drive_folder_id não configurado ainda." % marca)
        manifesto_arquivos_marca = {k: v for k, v in manifesto_anterior.get("arquivos", {}).items()
                                     if k.startswith(marca + os.sep)}
        on_progress = barra.definir_rotulo if barra else None
        resultado = sincronizar_marca(marca, folder_id, os.path.join(RAIZ_STAGING, marca),
                                       api_key, manifesto_arquivos_marca, on_progress)
        resultados_drive[marca] = resultado
        for item in resultado["arquivos"]:
            full = os.path.join(RAIZ_STAGING, item["caminho_relativo"])
            arquivos.append((full, item["caminho_relativo"], item.get("md5_drive", "")))
    # normaliza pro mesmo formato de coletar_arquivos_local (full, rel) + guarda md5_drive à parte
    arquivos_simples = [(f, r) for f, r, _ in arquivos]
    md5_drive_por_rel = {r: m for _, r, m in arquivos}
    return arquivos_simples, resultados_drive, md5_drive_por_rel


def processar_marca(marca, arquivos_marca, depara, cfg_marca, regras_negocio, cfg_qualidade,
                     historico, resultado_drive, log, md5_drive_por_rel=None, barra=None):
    md5_drive_por_rel = md5_drive_por_rel or {}
    cursos, discs, confs, achados = [], [], [], []
    pendentes = []
    arquivos_processados = {}
    falhas_leitura = []  # rel de arquivos que existem mas não puderam ser lidos/parseados
    total = len(arquivos_marca)

    for i, (full, rel) in enumerate(arquivos_marca, 1):
        arq = os.path.basename(full)
        if barra:
            barra.definir_rotulo("processando %s: %d/%d arquivos (%s)" % (marca, i, total, arq))
        try:
            curso, dd, avisos = ler_matriz(full, rel)
        except Exception as e:
            curso, dd, avisos = None, [], ["ERRO DE LEITURA: %s: %s" % (type(e).__name__, str(e)[:120])]

        achados.extend(camada1_schema.avaliar(curso, dd, avisos, marca, arq,
                                               cfg_marca.get("padrao_cod_disciplina")))

        if curso is None:
            confs.append({"MARCA": marca, "MODALIDADE": "", "ARQUIVO": arq, "COD_CURSO": "",
                          "ORIGEM_COD": "", "N_DISCIPLINAS": 0, "CH_TOTAL": None,
                          "CH_TOTAL_DECLARADA": None, "DIF_CH_SOMADA_x_DECLARADA": None,
                          "COD_DUPLICADO_COM": "", "AVISOS": "; ".join(avisos)})
            # arquivo EXISTE mas não pôde ser lido — não pode desaparecer de
            # arquivos_processados e virar "REMOVIDO" na comparação com a
            # execução anterior (ver snapshot.calcular_mudancas).
            falhas_leitura.append(rel)
            continue

        if not curso["COD_CURSO"]:
            if (marca, arq) in depara:
                curso["COD_CURSO"] = depara[(marca, arq)]
                curso["ORIGEM_COD"] = "de-para"
            else:
                curso["ORIGEM_COD"] = "PENDENTE"
                avisos.append("SIGLA INTERNA vazia — preencher _DE_PARA_CURSOS.xlsx")
            pendentes.append((marca, arq, curso["CURSO"], sugerir_codigo(curso["CURSO"], arq)))

        for d in dd:
            d["COD_CURSO"] = curso["COD_CURSO"]
            d["CURSO"] = curso["CURSO"]
            d["IES"] = curso["IES"]
            d["GRAU"] = curso["GRAU"]
            d["REGIME"] = curso["REGIME"]
            if not d["CURRICULO"]:
                d["CURRICULO"] = curso["CURRICULO"]

        achados.extend(camada2_regras_negocio.avaliar(curso, dd, marca, regras_negocio))
        achados.extend(camada3_consistencia_historica.avaliar(
            curso, dd, marca, historico, cfg_qualidade.get("camada3")))

        cursos.append(curso)
        discs.extend(dd)
        confs.append({"MARCA": marca, "MODALIDADE": curso["MODALIDADE"], "ARQUIVO": arq,
                      "COD_CURSO": curso["COD_CURSO"], "ORIGEM_COD": curso["ORIGEM_COD"],
                      "N_DISCIPLINAS": curso["N_DISCIPLINAS"], "CH_TOTAL": curso["CH_TOTAL"],
                      "CH_TOTAL_DECLARADA": curso["CH_TOTAL_DECLARADA"],
                      "DIF_CH_SOMADA_x_DECLARADA": curso["DIF_CH_SOMADA_x_DECLARADA"],
                      "LINHAS_MODELO_IGNORADAS": curso["LINHAS_MODELO_IGNORADAS"],
                      "LINHAS_TOTAL_IGNORADAS": curso["LINHAS_TOTAL_IGNORADAS"],
                      "COD_DUPLICADO_COM": "", "AVISOS": "; ".join(avisos)})
        arquivos_processados[rel] = {
            "md5_local": md5(full), "md5_drive": md5_drive_por_rel.get(rel, ""),
            "tamanho": os.path.getsize(full),
            "modificado": datetime.fromtimestamp(os.path.getmtime(full)).isoformat(timespec="seconds"),
            "cod_curso": curso["COD_CURSO"], "n_disciplinas": curso["N_DISCIPLINAS"],
            "ch_total": curso["CH_TOTAL"],
        }

    achados.extend(camada4_integridade_lote.avaliar(
        marca, cursos, resultado_drive, historico, cfg_qualidade.get("camada4")))

    # duplicidade de código dentro da marca (mantido também na aba CONFERENCIA por compatibilidade)
    porcod = defaultdict(list)
    for c in cursos:
        if c["COD_CURSO"]:
            porcod[c["COD_CURSO"]].append(c["ARQUIVO"])
    dup = {k: v for k, v in porcod.items() if len(v) > 1}
    for cf in confs:
        if cf["COD_CURSO"] in dup:
            cf["COD_DUPLICADO_COM"] = ", ".join(a for a in dup[cf["COD_CURSO"]] if a != cf["ARQUIVO"])

    log.print("  %-11s  %3d cursos | %5d disciplinas | %d achados de qualidade (%d ERRO, %d ALERTA)"
              % (marca, len(cursos), len(discs), len(achados),
                 sum(1 for a in achados if a.severidade == "ERRO"),
                 sum(1 for a in achados if a.severidade == "ALERTA")))

    return {"cursos": cursos, "discs": discs, "confs": confs, "achados": achados,
            "pendentes": pendentes, "arquivos_processados": arquivos_processados,
            "falhas_leitura": falhas_leitura}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--local", metavar="PASTA", help="usa uma pasta já extraída em vez do Drive")
    ap.add_argument("--config", default=os.path.join(RAIZ, "config", "config.yaml"))
    ap.add_argument("--regras-negocio", default=os.path.join(RAIZ, "config", "regras_negocio.yaml"))
    ap.add_argument("--sem-retencao", action="store_true", help="não aplica a política de retenção (útil em teste)")
    ap.add_argument("--duracao-minima-min", type=float, default=15,
                     help="piso de duração da execução em minutos — a barra de progresso segura o "
                          "processo até bater esse tempo mesmo se o trabalho real terminar antes "
                          "(padrão: 15). Use --duracao-minima-min 0 para desligar em teste.")
    ap.add_argument("--tipo", choices=["agendado", "manual"], default="manual",
                     help="registrado no INDICE.xlsx/CHANGELOG pra rastreabilidade — "
                          "'agendado' é usado pelo run_pipeline.bat (Task Scheduler); "
                          "'manual' é o padrão quando você roda por conta própria "
                          "(ex.: executar_urgente.bat).")
    ap.add_argument("--motivo", default="",
                     help="texto livre opcional — por que essa execução foi disparada "
                          "(ex.: 'pedido urgente do setor X'). Fica no INDICE.xlsx/CHANGELOG.")
    args = ap.parse_args()

    t0 = time.time()
    ts_arquivo = datetime.now().strftime("%Y%m%d_%H%M%S")
    log = Logger(os.path.join(RAIZ_LOG, "execucao_%s.log" % ts_arquivo))
    log.print("=" * 92)
    log.print("PIPELINE DE MATRIZES CURRICULARES")
    log.print("=" * 92)
    log.print("tipo: %s%s" % (args.tipo.upper(), "  |  motivo: %s" % args.motivo if args.motivo else ""))

    barra = BarraProgresso(duracao_minima_s=args.duracao_minima_min * 60, log=log)
    barra.iniciar()
    barra.definir_rotulo("iniciando...")

    cfg = carregar_config(args.config)
    regras_negocio = carregar_config(args.regras_negocio)
    cfg_qualidade = cfg.get("qualidade", {})
    marcas_cfg = cfg.get("marcas", {})

    incompletos = snapshot.limpar_incompletos(RAIZ_HISTORICO)
    if incompletos:
        log.print("faxina: %d snapshot(s) incompleto(s) removido(s): %s"
                  % (len(incompletos), ", ".join(incompletos)))

    pasta_snapshot, ts = snapshot.nova_pasta(RAIZ_HISTORICO)
    pasta_anterior = snapshot.snapshot_anterior(RAIZ_HISTORICO, pasta_snapshot)
    manifesto_anterior = snapshot.carregar_manifesto(pasta_anterior)
    log.print("snapshot anterior: %s" % (os.path.basename(pasta_anterior) if pasta_anterior else "(nenhum — primeira execução)"))

    depara = carregar_depara(CAMINHO_DEPARA)
    log.print("de-para carregado: %d códigos confirmados" % len(depara))

    resultados_drive_por_marca = {}
    md5_drive_por_rel = {}
    try:
        if args.local:
            log.print("modo: LOCAL (%s) — Drive não será acessado" % args.local)
            barra.definir_rotulo("lendo pasta local...")
            todos_arquivos, _ = coletar_arquivos_local(args.local)
        else:
            api_key_var = cfg.get("api_key_env_var", "DRIVE_API_KEY")
            api_key = os.environ.get(api_key_var)
            if not api_key:
                raise SystemExit("variável de ambiente %s não definida — defina a API key do Google "
                                  "Cloud antes de rodar (ver README/config.yaml)." % api_key_var)
            log.print("modo: GOOGLE DRIVE")
            todos_arquivos, resultados_drive_por_marca, md5_drive_por_rel = coletar_arquivos_drive(
                marcas_cfg, api_key, manifesto_anterior, barra)
    except BaseException as e:
        # BaseException (não só Exception) de propósito: `raise SystemExit`
        # (usado acima e em coletar_arquivos_drive pra erro de config) NÃO
        # herda de Exception em Python — um `except Exception` aqui deixaria
        # passar batido exatamente os erros de configuração mais comuns
        # (API key/pasta do Drive faltando), sem limpar o snapshot pela
        # metade. Cobre também Ctrl+C no meio de um download.
        #
        # Falha na coleta (rede, Drive fora do ar, pasta local inacessível,
        # config incompleta) não pode deixar um snapshot vazio/pela-metade
        # pra trás — ele NUNCA deve virar "snapshot anterior" da próxima
        # execução (listar_snapshots(apenas_completos=True) é a defesa
        # principal; isso aqui evita litter de pastas mortas).
        log.print("\n*** FALHA ao coletar arquivos: %s: %s ***" % (type(e).__name__, str(e)[:300]))
        log.close()
        shutil.rmtree(pasta_snapshot, ignore_errors=True)
        raise

    arquivos_por_marca = defaultdict(list)
    for full, rel in todos_arquivos:
        marca = rel.split(os.sep)[0]
        arquivos_por_marca[marca].append((full, rel))

    resultado_por_marca = {}
    manifesto_novo = {"gerado_em": datetime.now().isoformat(timespec="seconds"),
                       "arquivos": {}, "cursos": {}, "disciplinas": {},
                       "depara_confirmado": {}, "pendentes_consecutivos": {}}
    algum_erro_qualidade = False   # qualquer ERRO (qualquer camada) — só informativo, não falha a execução
    algum_erro_sistemico = False   # ERRO de camada 4 (integridade do lote) — é o que marca a execução
                                    # como falha pro Task Scheduler: download parcial, queda abrupta no
                                    # total geral etc. Achados de camada 1-3 (por disciplina/curso) são o
                                    # estado NORMAL a revisar na aba QUALIDADE, não uma falha da automação.

    for marca in sorted(set(marcas_cfg) | set(arquivos_por_marca)):
        historico = snapshot.construir_contexto_historico(manifesto_anterior, marca)
        r = processar_marca(marca, arquivos_por_marca.get(marca, []), depara,
                             marcas_cfg.get(marca, {}), regras_negocio, cfg_qualidade,
                             historico, resultados_drive_por_marca.get(marca), log,
                             md5_drive_por_rel, barra)
        r["achados"] = ordenar(r["achados"])
        resultado_por_marca[marca] = r

        falhas_download = [f["caminho_relativo"] for f in (resultados_drive_por_marca.get(marca) or {}).get("falhas", [])]
        falhas_rel = set(r["falhas_leitura"]) | set(falhas_download)
        mudancas = snapshot.calcular_mudancas(marca, r["arquivos_processados"], manifesto_anterior, falhas_rel)
        gravar = os.path.join(pasta_snapshot, "BASE_MATRIZES_%s.xlsx" % marca)
        consolidacao.gravar_workbook(gravar, marca, r["cursos"], r["discs"], r["confs"],
                                      mudancas, r["achados"])
        consolidacao.gravar_csv_disciplinas(
            os.path.join(pasta_snapshot, "DISCIPLINAS_%s.csv" % marca), r["discs"])
        consolidacao.gravar_json_matrizes(
            os.path.join(pasta_snapshot, "MATRIZES_%s.json" % marca), marca, r["cursos"], r["discs"])

        resultado_por_marca[marca]["mudancas"] = mudancas
        resultado_por_marca[marca]["n_cursos"] = len(r["cursos"])
        resultado_por_marca[marca]["n_disciplinas"] = len(r["discs"])
        if tem_erro(r["achados"]):
            algum_erro_qualidade = True
        if any(a.severidade == "ERRO" and a.camada == 4 for a in r["achados"]):
            algum_erro_sistemico = True

        manifesto_novo["arquivos"].update(r["arquivos_processados"])

    # de-para: unifica pendências de todas as marcas
    todos_pendentes = [p for r in resultado_por_marca.values() for p in r["pendentes"]]
    n_pend = gravar_depara(CAMINHO_DEPARA, todos_pendentes, depara)
    snapshot.copiar_depara_usado(pasta_snapshot, CAMINHO_DEPARA)

    # rastreabilidade: copia as matrizes brutas usadas nesta execução pra
    # dentro do próprio snapshot (RAW/) — cada snapshot fica auto-contido,
    # reproduzível, mesmo que o arquivo no Drive/_staging mude depois.
    barra.definir_rotulo("arquivando matrizes brutas (RAW)...")
    falhas_raw = snapshot.arquivar_raw(pasta_snapshot, todos_arquivos)
    if falhas_raw:
        log.print("aviso: %d arquivo(s) não puderam ser copiados pra RAW/: %s"
                  % (len(falhas_raw), "; ".join(r for r, _ in falhas_raw[:5])))

    for marca in resultado_por_marca:
        manifesto_novo["pendentes_consecutivos"].update(
            snapshot.calcular_pendentes_consecutivos(resultado_por_marca[marca]["cursos"], manifesto_anterior))
        # `depara` vem de carregar_depara() com chave (marca, arquivo) — tupla,
        # que o `json.dump` do manifesto (gravar_manifesto) não sabe serializar
        # como chave de objeto. Achata pro mesmo formato "marca::arquivo" usado
        # nas outras chaves compostas deste módulo (cursos/disciplinas acima),
        # que é também o formato que camada3_consistencia_historica espera ao
        # reler depara_confirmado do histórico (ver CONFLITO_DE_PARA).
        manifesto_novo["depara_confirmado"].update({"%s::%s" % k: v for k, v in depara.items()})
    m_completo = snapshot.montar_manifesto(
        manifesto_novo["arquivos"],
        [c for r in resultado_por_marca.values() for c in r["cursos"]],
        [d for r in resultado_por_marca.values() for d in r["discs"]],
        manifesto_novo["depara_confirmado"], manifesto_novo["pendentes_consecutivos"])
    m_completo["tipo"] = args.tipo
    m_completo["motivo"] = args.motivo
    snapshot.gravar_manifesto(pasta_snapshot, m_completo)

    changelog = snapshot.gerar_changelog(pasta_snapshot, ts, resultado_por_marca, args.tipo, args.motivo)

    duracao_trabalho_s = time.time() - t0
    snapshot.atualizar_indice(RAIZ_HISTORICO, ts, args.tipo.upper(), args.motivo,
                               resultado_por_marca, duracao_trabalho_s)

    if not args.sem_retencao:
        removidos = snapshot.aplicar_retencao(
            RAIZ_HISTORICO, cfg.get("retencao", {}).get("manter_raw_ultimos_n_snapshots", 8))
        if removidos:
            log.print("retenção: RAW removido de %d snapshot(s) antigos" % len(removidos))

    snapshot.atualizar_ponteiro_ultimo(RAIZ_HISTORICO, pasta_snapshot)

    log.print("\n" + "=" * 92)
    log.print("RESUMO")
    log.print("=" * 92)
    log.print("snapshot publicado: %s" % pasta_snapshot)
    log.print("código PENDENTE ............ %d linha(s) em _DE_PARA_CURSOS.xlsx" % n_pend)
    if algum_erro_sistemico:
        log.print("\n*** FALHA: achados de camada 4 (integridade do lote) — algo estrutural pode ter"
                  " dado errado (download parcial, queda abrupta no total geral). Ver aba QUALIDADE. ***")
    elif algum_erro_qualidade:
        log.print("\n*** este snapshot tem achados de severidade ERRO em disciplinas/cursos específicos —"
                  " revisar aba QUALIDADE quando puder. Não bloqueia a execução (não é falha sistêmica). ***")
    log.print("\ntrabalho real concluído em %.1fs" % (time.time() - t0))

    # a barra segura o processo até bater o piso de duração mínima (se o
    # trabalho real já demorou mais que isso, retorna na hora).
    barra.definir_rotulo("aguardando tempo mínimo de execução...")
    barra.concluir()

    log.print("execução finalizada em %.1fs (piso mínimo: %.1f min)" % (time.time() - t0, args.duracao_minima_min))
    log.close()

    # exit code reflete só problemas SISTÊMICOS (camada 4) — achados de
    # qualidade pontuais (camada 1-3) são o estado normal a revisar na aba
    # QUALIDADE/CHANGELOG, não devem marcar a execução semanal como "falhou"
    # no Task Scheduler toda vez.
    return 1 if algum_erro_sistemico else 0


if __name__ == "__main__":
    sys.exit(main())
