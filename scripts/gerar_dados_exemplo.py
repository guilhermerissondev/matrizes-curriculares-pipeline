# -*- coding: utf-8 -*-
"""Gera os dados sintéticos usados na demonstração do pipeline (`dados_exemplo/`).

Nenhum dado real de nenhuma instituição — cursos, disciplinas e cargas
horárias abaixo são fictícios, criados só para exercitar o parser e as 4
camadas de qualidade de ponta a ponta. Gera duas "safras" (v1 e v2) para
simular duas execuções semanais consecutivas: v2 parte de v1 com alterações
deliberadas (arquivo removido, carga horária corrigida, carga horária caindo
de forma abrupta, curso novo) para que o histórico versionado e as camadas 3/4
de qualidade tenham algo real para acusar — não é um "melhor caso" artificial.

Uso:
    python scripts/gerar_dados_exemplo.py

Regrava `dados_exemplo/v1`, `dados_exemplo/v2` e `_DE_PARA_CURSOS.xlsx` do
zero (idempotente). Depois disso, rode a demo com:
    python pipeline/orquestrador.py --local dados_exemplo/v1 --duracao-minima-min 0 --motivo "carga inicial (dados de exemplo)"
    python pipeline/orquestrador.py --local dados_exemplo/v2 --duracao-minima-min 0 --motivo "atualizacao semanal simulada (dados de exemplo)"
"""
import os
import shutil
import sys

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, RAIZ)

from pipeline.depara import gravar_depara  # noqa: E402

RAIZ_DADOS = os.path.join(RAIZ, "dados_exemplo")
CAMINHO_DEPARA = os.path.join(RAIZ, "_DE_PARA_CURSOS.xlsx")

# cabeçalho da matriz (linhas 1-13) -> valor em B da mesma linha
LINHA = {
    "IES": 1, "CURSO": 2, "GRADUAÇÃO": 3, "MODALIDADE": 4, "REGIME": 5,
    "INTEGRALIZAÇÃO MÍNIMA": 6, "INTEGRALIZAÇÃO MÁXIMA": 7,
    "ANO DE APROVAÇÃO DA MATRIZ": 8, "EDITAL DE INÍCIO DE OFERTA DA MATRIZ": 9,
    "ATIVIDADE COMPLEMENTAR": 10, "OBSERVAÇÕES": 11,
}
LINHA_CH_DECLARADA = 13   # A13 = numero, B13 = rotulo "CARGA HORARIA TOTAL DO CURSO"
LINHA_CABECALHO_COLUNAS = 26  # onde ficam os titulos de coluna das disciplinas (esperado pelo parser)

COLUNAS = [
    "ANO", "SEMESTRE", "MÓDULO", "SÉRIE", "ORDEM NO MÓDULO Nº", "SIGLA INTERNA DO CURSO",
    "CURRÍCULO", "DOMÍNIO FORMATIVO", "CÓDIGO DA DISCIPLINA", "NOME DA DISCIPLINA",
    "TIPO", "SUB TIPO", "PERFIL", "FIXA", "OFERTA ESPECIAL", "OBSERVAÇÃO",
    "QUESTIONÁRIO BÔNUS", "CH ATIVIDADE EAD", "CH PRESENCIAL TEÓRICA",
    "CH PRESENCIAL PRÁTICA", "CH PROVA PRESENCIAL", "CH SÍNCRONA MEDIADA",
    "CH TOTAL DA DISCIPLINA", "CÓDIGO DA DISCIPLINA EQUIVALENTE", "NOME DA DISCIPLINA EQUIVALENTE",
]


def _distribuir_ch(total, n, passo=10):
    """n valores múltiplos de `passo`, com leve variação, somando exatamente `total`."""
    base = (total // n) // passo * passo
    valores = [base] * n
    resto = total - base * n
    i = 0
    while resto >= passo:
        valores[i % n] += passo
        resto -= passo
        i += 1
    valores[0] += resto  # sobra residual (não múltipla de passo) absorvida no 1º item
    return valores


def montar_matriz(caminho, ies, curso, grau, modalidade, regime, disciplinas,
                   ch_total_declarada=None, sigla_curso=None, integralizacao_min="36 meses",
                   integralizacao_max="72 meses", curriculo="2026/1"):
    """disciplinas: [(semestre, cod, nome, ch_ead, ch_pres_teo, ch_pres_pra, ch_total), ...]
    ch_total_declarada: se None, usa a soma real (arquivo "limpo"); passe um valor
    diferente da soma para simular divergência (camada 2 de qualidade)."""
    soma = round(sum(d[6] for d in disciplinas), 2)
    declarada = soma if ch_total_declarada is None else ch_total_declarada

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"

    valores_cab = {
        "IES": ies, "CURSO": curso, "GRADUAÇÃO": grau, "MODALIDADE": modalidade, "REGIME": regime,
        "INTEGRALIZAÇÃO MÍNIMA": integralizacao_min, "INTEGRALIZAÇÃO MÁXIMA": integralizacao_max,
        "ANO DE APROVAÇÃO DA MATRIZ": "2026", "EDITAL DE INÍCIO DE OFERTA DA MATRIZ": "2026/1",
        "ATIVIDADE COMPLEMENTAR": "20h", "OBSERVAÇÕES": "",
    }
    for rotulo, linha in LINHA.items():
        ws.cell(linha, 1, rotulo)
        ws.cell(linha, 2, valores_cab[rotulo])
    ws.cell(LINHA_CH_DECLARADA, 1, declarada)
    ws.cell(LINHA_CH_DECLARADA, 2, "CARGA HORARIA TOTAL DO CURSO")

    for j, h in enumerate(COLUNAS, 1):
        ws.cell(LINHA_CABECALHO_COLUNAS, j, h)

    r = LINHA_CABECALHO_COLUNAS + 1
    semestre_anterior = None
    for semestre, cod, nome, ch_ead, ch_pres_teo, ch_pres_pra, ch_total in disciplinas:
        primeira_do_bloco = semestre != semestre_anterior
        semestre_anterior = semestre
        ws.cell(r, 1, "")                                        # ANO (não usado nesta demo)
        ws.cell(r, 2, semestre if primeira_do_bloco else "")      # SEMESTRE (propaga p/ linhas seguintes)
        ws.cell(r, 6, sigla_curso if (sigla_curso and r == LINHA_CABECALHO_COLUNAS + 1) else "")
        ws.cell(r, 7, curriculo if r == LINHA_CABECALHO_COLUNAS + 1 else "")
        ws.cell(r, 9, cod)
        ws.cell(r, 10, nome)
        ws.cell(r, 11, "OBRIGATORIA")
        ws.cell(r, 18, ch_ead)
        ws.cell(r, 19, ch_pres_teo)
        ws.cell(r, 20, ch_pres_pra)
        ws.cell(r, 23, ch_total)
        r += 1

    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    wb.save(caminho)
    return soma


# --------------------------------------------------------------------- cursos
def _disciplinas(prefixo, nomes, n, ch_total, pct_ead=1.0, n_semestres=None):
    """Gera n disciplinas com nomes de `nomes` (cicla se precisar), espalhadas
    em `n_semestres` semestres (padrão: um a cada ~5 disciplinas), com CH
    somando `ch_total`. `pct_ead`: fração da CH de cada disciplina que é EAD
    (o resto vira presencial teórica) — usado para simular cursos SEMI."""
    n_semestres = n_semestres or max(1, -(-n // 5))
    chs = _distribuir_ch(ch_total, n)
    out = []
    for i in range(n):
        semestre = i // -(-n // n_semestres) + 1
        cod = "%s%03d" % (prefixo, i + 1)
        nome = nomes[i % len(nomes)]
        ch = chs[i]
        ch_ead = round(ch * pct_ead, 2)
        ch_pres = round(ch - ch_ead, 2)
        out.append((semestre, cod, nome, ch_ead, ch_pres, 0.0, ch))
    return out


NOMES_ADMINISTRACAO = [
    "Fundamentos de Administração", "Contabilidade Geral", "Matemática Financeira",
    "Economia Aplicada", "Gestão de Pessoas", "Marketing Institucional",
    "Direito Empresarial", "Comportamento Organizacional", "Logística Empresarial",
    "Gestão Financeira", "Planejamento Estratégico", "Empreendedorismo",
    "Gestão de Processos", "Estatística Aplicada", "Ética Profissional",
    "Sistemas de Informação Gerencial", "Gestão da Qualidade", "Negociação",
    "Projeto Integrador", "Trabalho de Conclusão de Curso",
]
NOMES_GESTAO_COMERCIAL = [
    "Técnicas de Vendas", "Comportamento do Consumidor", "Gestão de Varejo",
    "Negociação Comercial", "Marketing Digital Aplicado", "Gestão de Equipes Comerciais",
    "Matemática Comercial", "Atendimento ao Cliente", "Plano de Negócios",
    "Legislação Comercial", "Indicadores de Desempenho Comercial", "Trade Marketing",
    "E-commerce", "Projeto Integrador",
]
NOMES_MARKETING_DIGITAL = [
    "Fundamentos de Marketing", "Redes Sociais e Conteúdo", "SEO e Performance",
    "Design para Marketing Digital", "Analytics e Métricas", "Branding",
    "Marketing de Conteúdo", "Growth Marketing", "Mídia Paga",
    "E-mail Marketing e Automação", "UX para Marketing", "Empreendedorismo Digital",
    "Projeto Integrador",
]
NOMES_ENGENHARIA_PRODUCAO = [
    "Cálculo Aplicado", "Física para Engenharia", "Desenho Técnico",
    "Gestão da Produção", "Logística Industrial", "Pesquisa Operacional",
    "Controle de Qualidade", "Ergonomia e Segurança do Trabalho", "Gestão de Projetos",
    "Planejamento e Controle da Produção", "Automação Industrial", "Cadeia de Suprimentos",
    "Sustentabilidade Industrial", "Manutenção Industrial", "Engenharia Econômica",
    "Simulação de Processos", "Projeto Integrador", "Trabalho de Conclusão de Curso",
]
NOMES_PEDAGOGIA = [
    "Fundamentos da Educação", "Psicologia da Aprendizagem", "Didática Geral",
    "Educação Infantil", "Alfabetização e Letramento", "Currículo e Avaliação",
    "Libras", "Educação Especial e Inclusiva", "Gestão Escolar",
    "Metodologia do Ensino de Matemática", "Metodologia do Ensino de Ciências",
    "Tecnologias na Educação", "Estágio Supervisionado I", "Estágio Supervisionado II",
    "Políticas Educacionais", "Educação de Jovens e Adultos", "Trabalho de Conclusão de Curso",
]
NOMES_DIREITO = [
    "Introdução ao Estudo do Direito", "Direito Constitucional", "Direito Civil",
    "Direito Penal", "Direito Administrativo", "Direito do Trabalho",
    "Direito Empresarial", "Direito Tributário", "Processo Civil",
    "Processo Penal", "Direito Internacional", "Ética e Deontologia Jurídica",
    "Direito do Consumidor", "Direitos Humanos", "Prática Jurídica I",
    "Prática Jurídica II", "Filosofia do Direito", "Trabalho de Conclusão de Curso",
]


def gerar_v1():
    if os.path.isdir(RAIZ_DADOS):
        shutil.rmtree(RAIZ_DADOS)

    area_negocios = "04 - Negocios Administracao e Direito"
    area_engenharia = "07 - Engenharia Producao e Construcao"
    area_educacao = "01 - Educacao"

    # --- INSTITUICAO_A --------------------------------------------------
    disc_adm = _disciplinas("ADM", NOMES_ADMINISTRACAO, 20, 3320)
    montar_matriz(
        os.path.join(RAIZ_DADOS, "v1", "INSTITUICAO_A", "EAD", area_negocios, "administracao_bacharelado.xlsx"),
        ies="Instituição de Ensino A", curso="Administração", grau="Bacharelado",
        modalidade="EAD", regime="Semestral", disciplinas=disc_adm, sigla_curso="ADM-BACH",
    )

    disc_gescom = _disciplinas("GES", NOMES_GESTAO_COMERCIAL, 14, 1680)
    montar_matriz(
        os.path.join(RAIZ_DADOS, "v1", "INSTITUICAO_A", "EAD", area_negocios, "gestao_comercial_tecnologo.xlsx"),
        ies="Instituição de Ensino A", curso="Gestão Comercial", grau="Tecnólogo",
        modalidade="EAD", regime="Semestral", disciplinas=disc_gescom, sigla_curso=None,  # sem SIGLA -> de-para
    )

    disc_mkt = _disciplinas("MKT", NOMES_MARKETING_DIGITAL, 13, 1620)
    montar_matriz(
        os.path.join(RAIZ_DADOS, "v1", "INSTITUICAO_A", "EAD", area_negocios, "marketing_digital_tecnologo.xlsx"),
        ies="Instituição de Ensino A", curso="Marketing Digital", grau="Tecnólogo",
        modalidade="EAD", regime="Semestral", disciplinas=disc_mkt, sigla_curso=None,  # sem SIGLA e sem de-para -> PENDENTE
    )

    disc_eng = _disciplinas("ENG", NOMES_ENGENHARIA_PRODUCAO, 18, 3260)
    montar_matriz(
        os.path.join(RAIZ_DADOS, "v1", "INSTITUICAO_A", "EAD", area_engenharia, "engenharia_producao_bacharelado.xlsx"),
        ies="Instituição de Ensino A", curso="Engenharia de Produção", grau="Bacharelado",
        modalidade="EAD", regime="Semestral", disciplinas=disc_eng, sigla_curso="ENGPROD-BACH",
        ch_total_declarada=3190,  # diverge 70h da soma real (3260) -> ERRO CH_DIVERGENTE (camada 2)
    )

    # --- INSTITUICAO_B ----------------------------------------------------
    disc_ped = _disciplinas("PED", NOMES_PEDAGOGIA, 17, 3320)
    montar_matriz(
        os.path.join(RAIZ_DADOS, "v1", "INSTITUICAO_B", "EAD", area_educacao, "pedagogia_licenciatura.xlsx"),
        ies="Instituição de Ensino B", curso="Pedagogia", grau="Licenciatura",
        modalidade="EAD", regime="Semestral", disciplinas=disc_ped, sigla_curso="PED-LIC",
    )

    disc_dir = _disciplinas("DIR", NOMES_DIREITO, 18, 3300, pct_ead=0.80)  # 80% EAD -> acima do teto SEMI (70%)
    montar_matriz(
        os.path.join(RAIZ_DADOS, "v1", "INSTITUICAO_B", "SEMI", area_negocios, "direito_bacharelado_semi.xlsx"),
        ies="Instituição de Ensino B", curso="Direito", grau="Bacharelado",
        modalidade="SEMI", regime="Semestral", disciplinas=disc_dir, sigla_curso="DIR-SEMI",
    )

    # de-para: GESCOM já confirmado (mostra a resolução automática); Marketing
    # Digital fica de fora de propósito -> aparece como PENDENTE na 1ª execução.
    gravar_depara(
        CAMINHO_DEPARA,
        pendentes=[("INSTITUICAO_A", "gestao_comercial_tecnologo.xlsx", "Gestão Comercial", "GES-TEC")],
        ja={("INSTITUICAO_A", "gestao_comercial_tecnologo.xlsx"): "GESCOM-TEC"},
    )


def gerar_v2():
    """v2 = v1 com 4 mudanças deliberadas, para exercitar histórico e qualidade:
    1) marketing_digital_tecnologo.xlsx REMOVIDO (estava PENDENTE)
    2) engenharia_producao: CH declarada corrigida (resolve o ERRO da v1)
    3) pedagogia: CH cai ~23% (camada 3: variação abrupta; e passa a violar o
       piso de CH do grau Licenciatura -> camada 2)
    4) contabilidade_bacharelado.xlsx NOVO em INSTITUICAO_B
    direito_bacharelado_semi e administracao_bacharelado e gestao_comercial
    ficam EXATAMENTE iguais -> aparecem como "inalterado".
    """
    v1 = os.path.join(RAIZ_DADOS, "v1")
    v2 = os.path.join(RAIZ_DADOS, "v2")
    if os.path.isdir(v2):
        shutil.rmtree(v2)
    shutil.copytree(v1, v2)

    area_negocios = "04 - Negocios Administracao e Direito"
    area_educacao = "01 - Educacao"

    # 1) remove o curso que estava PENDENTE
    os.remove(os.path.join(v2, "INSTITUICAO_A", "EAD", area_negocios, "marketing_digital_tecnologo.xlsx"))

    # 2) corrige a divergência de CH da Engenharia de Produção
    caminho_eng = os.path.join(v2, "INSTITUICAO_A", "EAD", "07 - Engenharia Producao e Construcao",
                                "engenharia_producao_bacharelado.xlsx")
    wb = openpyxl.load_workbook(caminho_eng)
    wb.active.cell(LINHA_CH_DECLARADA, 1, 3260)  # agora bate com a soma real
    wb.save(caminho_eng)

    # 3) Pedagogia: menos disciplinas / menos CH -> queda abrupta
    disc_ped_v2 = _disciplinas("PED", NOMES_PEDAGOGIA, 13, 2560)
    montar_matriz(
        os.path.join(v2, "INSTITUICAO_B", "EAD", area_educacao, "pedagogia_licenciatura.xlsx"),
        ies="Instituição de Ensino B", curso="Pedagogia", grau="Licenciatura",
        modalidade="EAD", regime="Semestral", disciplinas=disc_ped_v2, sigla_curso="PED-LIC",
    )

    # 4) curso novo em INSTITUICAO_B
    disc_ctb = _disciplinas("CTB", NOMES_ADMINISTRACAO, 19, 3350)  # área afim, reaproveita nomes de negócios
    montar_matriz(
        os.path.join(v2, "INSTITUICAO_B", "EAD", area_negocios, "contabilidade_bacharelado.xlsx"),
        ies="Instituição de Ensino B", curso="Ciências Contábeis", grau="Bacharelado",
        modalidade="EAD", regime="Semestral", disciplinas=disc_ctb, sigla_curso="CTB-BACH",
    )


if __name__ == "__main__":
    gerar_v1()
    gerar_v2()
    print("dados_exemplo/v1 e dados_exemplo/v2 gerados em %s" % RAIZ_DADOS)
    print("_DE_PARA_CURSOS.xlsx gerado em %s" % CAMINHO_DEPARA)
