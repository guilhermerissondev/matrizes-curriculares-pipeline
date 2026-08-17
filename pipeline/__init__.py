"""Pipeline automatizado de matrizes curriculares (multi-instituição).

Consolida grades de disciplinas + carga horária de múltiplas instituições
parceiras a partir de uma origem externa (Google Drive), valida a qualidade
dos dados em 4 camadas independentes e mantém um histórico versionado por
execução. Ver `pipeline/orquestrador.py` para o ponto de entrada e o README
na raiz do projeto para a visão geral da arquitetura.
"""
