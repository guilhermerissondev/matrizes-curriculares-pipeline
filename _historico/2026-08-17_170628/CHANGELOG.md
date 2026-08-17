# Changelog — execução 2026-08-17_170628

**Tipo:** MANUAL
**Motivo:** atualizacao semanal simulada (dados de exemplo)

## INSTITUICAO_A
- cursos: 3 | disciplinas: 52
- arquivos: 0 novos, 1 substituídos, 1 removidos, 2 inalterados
- qualidade: 1 ERRO, 0 ALERTA

### Erros que pedem revisão antes de considerar este snapshot confiável
Por regra: QUEDA_TOTAL_CURSOS (1)

- `QUEDA_TOTAL_CURSOS` (INSTITUICAO_A / ): total de cursos caiu 25% desde a execução anterior (4 -> 3) — possível falha sistêmica (pasta errada, permissão revogada, Drive fora do ar); revisar antes de considerar este snapshot válido

### Arquivos alterados
- **SUBSTITUÍDO** — engenharia_producao_bacharelado.xlsx (disciplinas: 18 → 18)
- **REMOVIDO** — marketing_digital_tecnologo.xlsx (estava na execução anterior e não está mais na pasta)

## INSTITUICAO_B
- cursos: 3 | disciplinas: 50
- arquivos: 1 novos, 1 substituídos, 0 removidos, 1 inalterados
- qualidade: 2 ERRO, 1 ALERTA

### Erros que pedem revisão antes de considerar este snapshot confiável
Por regra: CH_ABAIXO_DO_PISO (1), PCT_EAD_ACIMA_DO_TETO (1)

- `CH_ABAIXO_DO_PISO` (pedagogia_licenciatura.xlsx / PED-LIC): CH total 2560h abaixo do piso de 3200h para o grau 'Licenciatura'
- `PCT_EAD_ACIMA_DO_TETO` (direito_bacharelado_semi.xlsx / DIR-SEMI): 80.0% da CH é EAD, acima do teto de 70% para modalidade SEMI

### Arquivos alterados
- **SUBSTITUÍDO** — pedagogia_licenciatura.xlsx (disciplinas: 17 → 13)
- **NOVO** — contabilidade_bacharelado.xlsx (arquivo não existia na execução anterior)
