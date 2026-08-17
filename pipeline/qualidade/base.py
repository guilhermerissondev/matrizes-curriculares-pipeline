# -*- coding: utf-8 -*-
"""Estrutura comum dos achados de qualidade, compartilhada pelas 4 camadas."""
from dataclasses import dataclass, field

ERRO = "ERRO"
ALERTA = "ALERTA"
INFO = "INFO"
_ORDEM_SEVERIDADE = {ERRO: 0, ALERTA: 1, INFO: 2}


@dataclass
class Achado:
    severidade: str          # ERRO | ALERTA | INFO
    camada: int               # 1..4
    regra: str                 # slug curto, ex.: "CH_DIVERGENTE"
    marca: str
    mensagem: str
    arquivo: str = ""
    cod_curso: str = ""
    cod_disciplina: str = ""

    def como_linha(self):
        return {
            "SEVERIDADE": self.severidade, "CAMADA": self.camada, "REGRA": self.regra,
            "MARCA": self.marca, "ARQUIVO": self.arquivo, "COD_CURSO": self.cod_curso,
            "COD_DISCIPLINA": self.cod_disciplina, "MENSAGEM": self.mensagem,
        }


def ordenar(achados):
    return sorted(achados, key=lambda a: (_ORDEM_SEVERIDADE.get(a.severidade, 9), a.camada, a.regra))


def tem_erro(achados):
    return any(a.severidade == ERRO for a in achados)
