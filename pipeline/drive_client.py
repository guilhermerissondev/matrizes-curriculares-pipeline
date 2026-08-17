# -*- coding: utf-8 -*-
"""Cliente do Google Drive para pastas públicas ("qualquer pessoa com o link").

Abordagem híbrida, validada contra o Drive real deste projeto (não só
documentação):
  - LISTAGEM (`listar_pasta`/`listar_arvore`): API v3 oficial + API key
    (`files.list`, paginado). Só `requests`, sem OAuth. Deliberadamente NÃO
    usa `gdown` aqui: `gdown.download_folder` trunca silenciosamente pastas
    com mais de 50 arquivos, e as pastas reais deste projeto já passam disso
    em algumas áreas CINE — perda silenciosa de dados é inaceitável num
    pipeline de governança.
  - DOWNLOAD do conteúdo (`baixar_arquivo`): a API oficial (`alt=media`)
    exige OAuth mesmo em pasta pública — testado manualmente, devolve 403
    só com API key. Por isso o download usa a URL pública clássica
    (`drive.google.com/uc?id=...&export=download`), a mesma que o navegador
    usa num link "qualquer pessoa com o link" — sem OAuth, funciona.

Uso típico (ver `orquestrador.py`):

    from pipeline.drive_client import sincronizar_marca

    resultado = sincronizar_marca(
        marca="INSTITUICAO_A",
        folder_id="<id da pasta raiz no Drive>",
        pasta_destino=os.path.join(RAIZ_STAGING, "INSTITUICAO_A"),
        api_key=os.environ["DRIVE_API_KEY"],
        manifesto_anterior=manifesto_anterior.get("arquivos", {}),
    )
"""
import os
import time

import requests

API_BASE = "https://www.googleapis.com/drive/v3/files"
DOWNLOAD_BASE = "https://drive.google.com/uc"
MIME_PASTA = "application/vnd.google-apps.folder"
TENTATIVAS = 5
BACKOFF_BASE = 1.5  # segundos; cresce exponencialmente por tentativa


class DriveError(RuntimeError):
    pass


def _requisitar(sessao, url, params, stream=False):
    """GET com retry exponencial. Levanta DriveError se todas as tentativas falharem."""
    ultimo_erro = None
    for tentativa in range(1, TENTATIVAS + 1):
        try:
            resp = sessao.get(url, params=params, stream=stream, timeout=60)
            if resp.status_code == 200:
                return resp
            if resp.status_code in (403, 429, 500, 502, 503, 504):
                # 403 pode ser rate-limit (userRateLimitExceeded) ou permissão —
                # tentamos de novo; se for permissão de fato, vai falhar nas
                # TENTATIVAS e o erro final expõe o corpo da resposta.
                ultimo_erro = DriveError(
                    "HTTP %d ao chamar Drive API: %s" % (resp.status_code, resp.text[:300]))
            else:
                resp.raise_for_status()
        except requests.RequestException as e:
            ultimo_erro = DriveError("erro de rede ao chamar Drive API: %s" % e)
        time.sleep(BACKOFF_BASE ** tentativa)
    raise ultimo_erro or DriveError("falha desconhecida ao chamar Drive API")


MIME_ATALHO = "application/vnd.google-apps.shortcut"


def listar_pasta(sessao, folder_id, api_key):
    """Lista o conteúdo direto (1 nível) de uma pasta pública, paginando.

    Devolve lista de dicts: {id, name, mimeType, md5Checksum, modifiedTime,
    size, shortcutDetails}
    """
    itens = []
    token = None
    while True:
        params = {
            "q": "'%s' in parents and trashed = false" % folder_id,
            "fields": "nextPageToken, files(id,name,mimeType,md5Checksum,modifiedTime,size,shortcutDetails)",
            "pageSize": 1000,
            "key": api_key,
        }
        if token:
            params["pageToken"] = token
        resp = _requisitar(sessao, API_BASE, params)
        data = resp.json()
        itens.extend(data.get("files", []))
        token = data.get("nextPageToken")
        if not token:
            break
    return itens


def _resolver_atalho(sessao, item, api_key):
    """Um atalho ("shortcut") do Drive não carrega o md5Checksum/size do
    alvo em si — busca os metadados reais do arquivo apontado. Devolve None
    (e deixa o atalho ser ignorado, sem derrubar a listagem) se o alvo não
    existir mais ou não puder ser lido."""
    alvo_id = (item.get("shortcutDetails") or {}).get("targetId")
    if not alvo_id:
        return None
    try:
        resp = _requisitar(sessao, "%s/%s" % (API_BASE, alvo_id),
                            {"fields": "id,name,mimeType,md5Checksum,modifiedTime,size", "key": api_key})
        meta = resp.json()
    except DriveError:
        return None
    meta["name"] = item.get("name", meta.get("name"))  # mantém o nome do atalho na árvore
    return meta


def listar_arvore(folder_id, api_key, sessao=None, prefixo=()):
    """Percorre recursivamente a árvore a partir de `folder_id`.

    Gera tuplas (partes_caminho, meta) para cada ARQUIVO encontrado (não
    pastas), onde `partes_caminho` é a tupla de nomes de pasta até o arquivo
    (ex.: ("EAD", "04. Negócios, Administração e Direito")) — usada para
    reconstruir a árvore MODALIDADE/AREA_CINE localmente. Não impõe limite de
    itens por pasta (paginação cobre isso).
    """
    sessao = sessao or requests.Session()
    itens = listar_pasta(sessao, folder_id, api_key)
    for item in itens:
        mime = item.get("mimeType")
        if mime == MIME_PASTA:
            yield from listar_arvore(item["id"], api_key, sessao, prefixo + (item["name"],))
        elif mime == MIME_ATALHO:
            alvo_mime = (item.get("shortcutDetails") or {}).get("targetMimeType")
            if alvo_mime == MIME_PASTA:
                alvo_id = (item.get("shortcutDetails") or {}).get("targetId")
                if alvo_id:
                    yield from listar_arvore(alvo_id, api_key, sessao, prefixo + (item["name"],))
            else:
                meta_alvo = _resolver_atalho(sessao, item, api_key)
                if meta_alvo:
                    yield prefixo, meta_alvo
                # alvo removido/inacessível: ignora silenciosamente o atalho
                # quebrado em vez de derrubar a listagem inteira.
        else:
            yield prefixo, item


def baixar_arquivo(sessao, file_id, api_key, destino):
    """Baixa o CONTEÚDO do arquivo.

    Deliberadamente NÃO usa `files/{id}?alt=media` da API oficial: testado
    manualmente contra o Drive real deste projeto, esse endpoint devolve
    403 Forbidden pra download de bytes com autenticação só por API key —
    a API key basta pra listar/ler metadados (`files.list`, usado em
    `listar_pasta`), mas o Drive exige um token OAuth pra mover o conteúdo
    binário, mesmo em pasta pública. Em vez disso, usamos a URL pública
    clássica de download (a mesma que "Fazer download" usa no navegador
    pra um link compartilhado como "qualquer pessoa com o link") — não
    pede OAuth, e foi validada byte-a-byte contra um arquivo real.
    Não paginada/sem limite de tamanho de pasta porque é usada por
    arquivo, não por pasta (a listagem continua via API + key).
    """
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    resp = _requisitar(sessao, DOWNLOAD_BASE, {"id": file_id, "export": "download"}, stream=True)
    tipo = resp.headers.get("Content-Type", "")
    if "text/html" in tipo:
        # Acontece pra arquivos grandes (~25MB+) que disparam a tela "não foi
        # possível verificar vírus" — não é o caso das matrizes (centenas de
        # KB), mas evita gravar uma página HTML como se fosse o .xlsx.
        raise DriveError("resposta HTML em vez do arquivo (possível tela de confirmação do Drive "
                          "para arquivos grandes) — id=%s" % file_id)
    tmp = destino + ".part"
    with open(tmp, "wb") as f:
        for bloco in resp.iter_content(chunk_size=1 << 20):
            if bloco:
                f.write(bloco)
    os.replace(tmp, destino)


def sincronizar_marca(marca, folder_id, pasta_destino, api_key, manifesto_anterior=None, on_progress=None):
    """Baixa (incrementalmente) todos os .xlsx de uma marca para `pasta_destino`.

    `manifesto_anterior` = {caminho_relativo: {"md5": ...}} da execução
    anterior (vem de `snapshot`) — arquivos cujo `md5Checksum` do Drive não
    mudou não são baixados de novo, só reaproveitados do snapshot anterior se
    já existirem localmente; senão são baixados mesmo sem mudança (primeira
    vez que este destino é usado).

    Devolve dict com:
      arquivos: lista de {caminho_relativo, id, md5, modifiedTime, tamanho, baixado}
      n_listados: quantos arquivos .xlsx a API retornou no total
      n_baixados: quantos foram efetivamente transferidos nesta execução
      falhas: lista de {caminho_relativo, erro} — arquivos listados mas que
              falharam ao baixar (NÃO deve ser tratado como "removido"; ver
              camada 3/4 de qualidade)
    """
    manifesto_anterior = manifesto_anterior or {}
    sessao = requests.Session()
    arquivos, falhas = [], []
    n_baixados = 0

    # materializa a árvore ANTES de baixar — listagem é só metadados (rápida)
    # e assim sabemos o total pra reportar "X/Y" no progresso.
    itens = [(partes, meta) for partes, meta in listar_arvore(folder_id, api_key, sessao, prefixo=(marca,))
             if meta.get("name", "").lower().endswith(".xlsx") and not meta.get("name", "").startswith("~$")]
    total = len(itens)

    for i, (partes, meta) in enumerate(itens, 1):
        nome = meta.get("name", "")
        # partes já inclui a marca como 1º elemento (prefixo inicial); removida
        # aqui porque pasta_destino já é a pasta da marca.
        partes_sem_marca = partes[1:] if partes and partes[0] == marca else partes
        destino = os.path.join(pasta_destino, *partes_sem_marca, nome)
        rel = os.path.join(marca, *partes_sem_marca, nome)

        md5_novo = meta.get("md5Checksum", "")
        anterior = manifesto_anterior.get(rel)
        precisa_baixar = True
        # exige md5_novo não-vazio: um checksum vazio nunca deve contar como
        # "igual ao anterior" (evita pular download de um arquivo real por
        # coincidência de dois valores vazios — ex.: metadado ausente pontualmente).
        if anterior and md5_novo and anterior.get("md5_drive") == md5_novo and os.path.exists(destino):
            precisa_baixar = False

        erro = None
        if precisa_baixar:
            try:
                baixar_arquivo(sessao, meta["id"], api_key, destino)
                n_baixados += 1
            except DriveError as e:
                erro = str(e)
                falhas.append({"caminho_relativo": rel, "erro": erro})

        if erro is None:
            arquivos.append({
                "caminho_relativo": rel, "id": meta["id"], "md5_drive": md5_novo,
                "modifiedTime": meta.get("modifiedTime", ""),
                "tamanho": int(meta.get("size", 0) or 0),
                "baixado_nesta_execucao": precisa_baixar,
            })

        if on_progress:
            on_progress("baixando %s: %d/%d arquivos (%s)" % (marca, i, total, nome))

    return {
        "arquivos": arquivos,
        "n_listados": len(arquivos) + len(falhas),
        "n_baixados": n_baixados,
        "falhas": falhas,
    }
