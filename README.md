# Pipeline de Matrizes Curriculares

Pipeline em Python que consolida matrizes curriculares (grade de disciplinas +
carga horária) de múltiplas instituições de ensino a distância a partir do
Google Drive, valida a qualidade dos dados em **4 camadas independentes** e
mantém um **histórico versionado e auditável** por execução.

> **Sobre este repositório.** É uma cópia anonimizada de um pipeline que
> rodo em produção (execução semanal automática via Task Scheduler + botão de
> execução urgente sob demanda). Nomes de instituições, IDs de pasta do Drive
> e todo o conteúdo curricular foram substituídos por dados **100% sintéticos**
> — nada aqui identifica instituições, cursos, disciplinas ou pessoas reais.
> Os dados de exemplo (`dados_exemplo/`) e a saída já gerada por eles
> (`_historico/`) ficam no repositório justamente para que o pipeline possa
> ser inspecionado — e rodado — sem precisar de acesso a nenhum Drive real.

## Por que este projeto existe

Duas instituições parceiras publicam matrizes curriculares (uma planilha por
curso) em pastas do Google Drive, em **6 layouts diferentes** de planilha que
foram mudando ao longo do tempo. Antes deste pipeline, consolidar isso era um
processo manual e frágil: alguém abria cada arquivo, copiava os dados a mão
para uma base central, e não havia como saber com confiança se uma matriz
"sumiu" porque foi removida de propósito ou porque o download falhou.

O pipeline resolve isso com três decisões de design que guiam o projeto
inteiro:

1. **Mapear colunas por nome, nunca por posição** — o parser (`pipeline/parsing.py`)
   absorve os diferentes layouts de planilha existentes (e tende a absorver
   os que vierem) porque identifica cada campo pelo texto do cabeçalho, não
   pela letra da coluna.
2. **Nunca publicar um snapshot pela metade.** Uma falha de rede, Drive fora
   do ar ou config incompleta no meio de uma execução remove a pasta que
   estava sendo montada e mantém o último snapshot válido como referência —
   nunca deixa um histórico incompleto se fazer passar por um resultado real.
3. **Separar "isso quebrou de verdade" de "isso merece revisão humana".**
   Achados de qualidade pontuais (uma disciplina sem carga horária, uma matriz
   com carga horária divergente) não derrubam a execução semanal — só uma
   queda abrupta no total geral de cursos, ou uma falha de download, marcam a
   execução como falha sistêmica. Ver [Camadas de qualidade](#camadas-de-qualidade-o-diferencial-do-projeto).

## Arquitetura

```
                 ┌─────────────┐
   Google Drive  │  download   │  pastas públicas ("qualquer pessoa com
   (pastas       │  incremental│  o link"); só baixa de novo o que mudou
   públicas) ───▶│  (md5)      │  (md5Checksum comparado ao manifesto
                 └──────┬──────┘  da execução anterior)
                        ▼
                 ┌─────────────┐
                 │  parsing    │  1 planilha .xlsx = 1 curso; colunas
                 │  tolerante  │  mapeadas por nome (6 layouts distintos
                 │  a layout   │  hoje, tende a absorver os que vierem)
                 └──────┬──────┘
                        ▼
                 ┌─────────────┐
                 │ qualidade   │  4 camadas independentes — schema,
                 │ (4 camadas) │  regras de negócio, histórico, lote
                 └──────┬──────┘
                        ▼
                 ┌─────────────┐
                 │ consolidação│  BASE_MATRIZES_<MARCA>.xlsx (5 abas),
                 │ + snapshot  │  CSV, JSON, CHANGELOG.md, manifesto.json,
                 │ versionado  │  cópia bruta (RAW/) — pasta imutável por
                 └─────────────┘  execução em _historico/AAAA-MM-DD_HHMMSS/
```

Ponto de entrada único: [`pipeline/orquestrador.py`](pipeline/orquestrador.py).
Cada módulo tem uma responsabilidade e pode ser lido isoladamente:

| Módulo | Responsabilidade |
|---|---|
| [`drive_client.py`](pipeline/drive_client.py) | Listagem (API v3 + key) e download incremental (URL pública) do Drive |
| [`parsing.py`](pipeline/parsing.py) | Lê 1 planilha `.xlsx` → `(curso, disciplinas, avisos)`, mapeando colunas por nome |
| [`depara.py`](pipeline/depara.py) | Resolve o código interno do curso quando a planilha não traz um — a única peça do pipeline com um humano no loop |
| [`qualidade/`](pipeline/qualidade) | As 4 camadas de validação (uma responsabilidade por arquivo) |
| [`consolidacao.py`](pipeline/consolidacao.py) | Grava os artefatos de saída (xlsx com 5 abas, CSV, JSON) |
| [`snapshot.py`](pipeline/snapshot.py) | Histórico versionado: manifesto, changelog, retenção, índice consolidado |
| [`progresso.py`](pipeline/progresso.py) | Barra de progresso + piso de duração mínima da execução |
| [`orquestrador.py`](pipeline/orquestrador.py) | Liga as etapas acima; único ponto de entrada |

## Camadas de qualidade — o diferencial do projeto

Em vez de um único "isso está errado" binário, a validação é dividida em 4
camadas com escopos e severidades diferentes — cada achado carrega
`severidade` (ERRO/ALERTA/INFO), `camada`, uma `regra` (slug) e a
mensagem, e vira uma linha na aba **QUALIDADE** da planilha final.

| Camada | Escopo | Exemplos de regra |
|---|---|---|
| **1 — Schema** | 1 arquivo isolado | cabeçalho incompleto, disciplina sem nome, carga horária negativa |
| **2 — Regras de negócio** | 1 curso, contra `config/regras_negocio.yaml` | CH somada diverge da declarada, % de EAD acima do teto regulatório da modalidade, CH total abaixo do piso do grau |
| **3 — Consistência histórica** | 1 curso, contra o snapshot anterior | variação abrupta de carga horária, queda no nº de disciplinas, conflito com um de-para já confirmado |
| **4 — Integridade do lote** | 1 instituição inteira, nesta execução | queda abrupta no total geral de cursos, download parcial, código de curso duplicado |

Só a camada 4 pode marcar a execução inteira como falha sistêmica (o exit
code que o Task Scheduler usa para alertar); as camadas 1-3 ficam registradas
para revisão humana sem travar a publicação semanal — ver a seção
[Rodando a demonstração](#rodando-a-demonstração-dados-sintéticos) para ver
isso acontecendo com dados de verdade.

Os limiares de negócio (tolerância de CH, teto de EAD por modalidade, piso de
CH por grau) **não ficam hardcoded** — vêm de
[`config/regras_negocio.yaml`](config/regras_negocio.yaml), justamente porque
são parâmetros institucionais/regulatórios que mudam com a legislação e
precisam ser revisados por quem entende do assunto acadêmico, não por quem
programa.

## Histórico versionado e auditável

Cada execução publica uma pasta imutável em `_historico/AAAA-MM-DD_HHMMSS/`
com:

- `BASE_MATRIZES_<INSTITUICAO>.xlsx` — abas DISCIPLINAS, CURSOS, CONFERENCIA,
  MUDANCAS e QUALIDADE (cores por severidade), `.csv` e `.json` equivalentes
- `CHANGELOG.md` — resumo legível do que mudou desde a execução anterior
- `manifesto.json` — fingerprint por arquivo/curso/disciplina, usado para
  comparar com a próxima execução sem reabrir nenhum `.xlsx` antigo
- `RAW/` — cópia de todas as matrizes brutas usadas nesta execução (qualquer
  snapshot passado é auto-contido e reproduzível, mesmo que o arquivo de
  origem mude depois — podado após N execuções por política de retenção)

`_historico/INDICE.xlsx` acumula uma linha por instituição a cada execução
(tipo, situação, contagens, duração) — visão do histórico inteiro sem abrir
cada snapshot.

## Rodando a demonstração (dados sintéticos)

O repositório já traz duas "safras" de dados fictícios em `dados_exemplo/` —
`v1` e `v2` simulam duas execuções semanais consecutivas da mesma instituição,
com mudanças propositais entre elas (um curso removido, uma divergência de
carga horária corrigida, uma queda abrupta de carga horária, um curso novo)
para que o histórico e as 4 camadas de qualidade tenham algo real para
mostrar — não é um "melhor caso" artificial.

```bash
pip install -r requirements.txt

# execução 1 (carga inicial)
python pipeline/orquestrador.py --local dados_exemplo/v1 --duracao-minima-min 0

# execução 2 (uma semana depois — mesma instituição, dados mudaram)
python pipeline/orquestrador.py --local dados_exemplo/v2 --duracao-minima-min 0
```

`--local` roda o pipeline inteiro sem tocar o Drive; `--duracao-minima-min 0`
desliga o piso de duração mínima da execução (existe só para a versão
agendada em produção, ver [Setup com Google Drive](#setup-com-google-drive)).
Os `_historico/2026-08-17_.../` já commitados neste repositório **são o
resultado real dessas duas execuções** — inclusive o `CHANGELOG.md` da
segunda mostra a camada 4 marcando a execução como falha sistêmica de
propósito (queda de 25% no total de cursos de uma instituição, por causa do
curso removido):

```
### Erros que pedem revisão antes de considerar este snapshot confiável
Por regra: QUEDA_TOTAL_CURSOS (1)

- `QUEDA_TOTAL_CURSOS` (INSTITUICAO_A / ): total de cursos caiu 25% desde a
  execução anterior (4 -> 3) — possível falha sistêmica (pasta errada,
  permissão revogada, Drive fora do ar); revisar antes de considerar este
  snapshot válido
```

Para regenerar (ou alterar) os dados sintéticos do zero:
[`scripts/gerar_dados_exemplo.py`](scripts/gerar_dados_exemplo.py).

## Setup com Google Drive

Para apontar o pipeline para pastas reais do Drive em vez do modo `--local`:

1. **API key do Google Cloud** (só para *listar* pastas/arquivos — ver nota
   abaixo sobre download):
   - Console do Google Cloud → criar projeto → "APIs e serviços" →
     "Biblioteca" → ativar **Google Drive API**.
   - "Credenciais" → "Criar credenciais" → **Chave de API** (não "ID do
     cliente OAuth" nem "Conta de serviço").
   - (Recomendado) restringir a key à Google Drive API.
   - Definir a variável de ambiente `DRIVE_API_KEY` (no Windows:
     `setx DRIVE_API_KEY "sua-key-aqui"`, feche/reabra o terminal depois).
2. **IDs das pastas do Drive**: em `config/config.yaml`, preencher
   `marcas.<NOME>.drive_folder_id` para cada instituição — o trecho da URL
   depois de `/folders/`.
3. **Agendamento** (opcional): `run_pipeline.bat` é um wrapper fino pensado
   para o Windows Task Scheduler (`python pipeline/orquestrador.py --tipo agendado`).

**Nota sobre download**: a API key só é usada para *listar* pastas/arquivos
(`files.list`) — a API oficial de download de conteúdo (`alt=media`) devolve
403 mesmo em pasta pública sem OAuth (testado manualmente). O download de
fato usa a URL pública clássica (`drive.google.com/uc?id=...&export=download`),
a mesma que o navegador usa num link "qualquer pessoa com o link". Ver
comentários em [`pipeline/drive_client.py`](pipeline/drive_client.py).

## Estrutura do repositório

```
config/                  parâmetros de execução (config.yaml) e regras de
                          negócio acadêmicas/regulatórias (regras_negocio.yaml)
pipeline/                todo o código do pipeline (ver tabela de módulos acima)
dados_exemplo/            dados sintéticos usados na demonstração (v1, v2)
scripts/                 script que gera os dados sintéticos
_historico/               saída publicada — já contém as 2 execuções de exemplo
_DE_PARA_CURSOS.xlsx      único arquivo editado à mão: resolve manualmente o
                          código de um curso quando a planilha não traz um
                          (ver pipeline/depara.py)
```

## Stack

Python (biblioteca padrão + `openpyxl` para ler/gravar `.xlsx`, `requests`
para a API do Drive, `pyyaml` para config) — deliberadamente sem
framework de dados pesado: o volume não pede isso, e manter as dependências
mínimas facilita rodar em qualquer máquina Windows sem infraestrutura
adicional (é assim que roda em produção, via Task Scheduler).

## Licença

[MIT](LICENSE).
