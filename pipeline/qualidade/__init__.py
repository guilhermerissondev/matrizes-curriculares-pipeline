"""Verificação de qualidade de dados em 4 camadas.

Cada camada expõe uma função `avaliar(...)` que devolve uma lista de achados
no formato `pipeline.qualidade.base.Achado`. Camadas:

  1. camada1_schema             — estrutura/formato de um arquivo isolado
  2. camada2_regras_negocio     — regras acadêmicas/regulatórias de um curso
  3. camada3_consistencia_historica — comparação com a execução anterior
  4. camada4_integridade_lote   — visão consolidada de todos os arquivos da marca
"""
