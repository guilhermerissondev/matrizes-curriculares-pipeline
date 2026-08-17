# -*- coding: utf-8 -*-
"""Barra de progresso em porcentagem + piso de duração mínima da execução.

Duas coisas deliberadamente combinadas aqui:

1. Uma barra no terminal que se atualiza sozinha (thread em background),
   então aparece se movendo mesmo durante fases sem callback explícito
   (ex.: um download lento na rede) — não depende de "tick" manual em cada
   ponto do código.
2. Um piso de duração: a execução (manual ou pelo Task Scheduler) nunca
   termina antes de `duracao_minima_s`, mesmo que o trabalho real acabe bem
   antes — `concluir()` segura a barra em andamento até bater o piso. Se o
   trabalho real demorar MAIS que o piso, não há espera extra: a barra segue
   o progresso real e termina assim que o trabalho termina.
"""
import sys
import threading
import time

PREENCHIDO = "#"
VAZIO = "-"


class BarraProgresso:
    def __init__(self, duracao_minima_s=900, largura=30, intervalo_redesenho_s=1.0, log=None):
        self.duracao_minima_s = duracao_minima_s
        self.largura = largura
        self.intervalo = intervalo_redesenho_s
        self.log = log  # objeto com .print(...) — opcional, registra marcos de 10% no log persistente

        self._inicio = time.time()
        self._rotulo = ""
        self._trabalho_concluido = False
        self._parar = threading.Event()
        self._lock = threading.RLock()  # reentrante: _desenhar chama _pct() sob o mesmo lock
        self._ultimo_marco_logado = -1
        self._thread = None

    # ------------------------------------------------------------ controle
    def iniciar(self):
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def definir_rotulo(self, texto):
        with self._lock:
            self._rotulo = texto

    def marcar_trabalho_concluido(self):
        """Chamar quando o trabalho REAL (download+processamento) terminou.
        A barra continua desenhando até bater `duracao_minima_s`."""
        with self._lock:
            self._trabalho_concluido = True

    def concluir(self):
        """Bloqueia até a barra realmente terminar (trabalho concluído E
        duração mínima atingida), depois para a thread e finaliza a linha."""
        self.marcar_trabalho_concluido()
        while self._pct() < 100:
            time.sleep(self.intervalo)
        self._parar.set()
        if self._thread:
            self._thread.join(timeout=2)
        self._desenhar(final=True)
        sys.stdout.write("\n")
        sys.stdout.flush()

    # ------------------------------------------------------------- interno
    def _pct(self):
        with self._lock:
            concluido = self._trabalho_concluido
        if self.duracao_minima_s <= 0:
            # piso desligado (--duracao-minima-min 0, usado em teste): não
            # há tempo mínimo a esperar, só reflete se o trabalho terminou.
            return 100 if concluido else 99
        decorrido = time.time() - self._inicio
        pct_tempo = 100.0 * decorrido / self.duracao_minima_s
        if not concluido:
            # nunca mostra 100% enquanto o trabalho real não terminou —
            # mesmo que o tempo mínimo já tenha passado (trabalho mais
            # lento que o piso: a barra some no fim, refletindo o real).
            return min(99, int(pct_tempo))
        # trabalho concluído: falta só bater o piso de tempo, se ainda não bateu
        return min(100, int(pct_tempo))

    def _desenhar(self, final=False):
        # tudo que lê/muda estado compartilhado (_rotulo, _ultimo_marco_logado)
        # fica sob o lock — _desenhar pode ser chamado tanto pela thread de
        # redesenho quanto pelo `concluir()` do thread principal perto do fim
        # (join com timeout não garante a outra chamada já ter retornado).
        with self._lock:
            pct = self._pct()
            rotulo = self._rotulo
            novo_marco = self.log is not None and (pct // 10) != (self._ultimo_marco_logado // 10)
            if novo_marco:
                self._ultimo_marco_logado = pct

        preenchido = int(self.largura * pct / 100)
        barra = PREENCHIDO * preenchido + VAZIO * (self.largura - preenchido)
        linha = "\r[%s] %3d%%  %s" % (barra, pct, rotulo)
        linha = linha.ljust(100)  # limpa sobras de linhas anteriores mais longas
        try:
            # defensivo: sem console anexado (ex.: Task Scheduler sem sessão
            # interativa) ou codepage exótico não devem derrubar o pipeline
            # por causa de um efeito cosmético.
            sys.stdout.write(linha)
            sys.stdout.flush()
        except Exception:
            pass
        if novo_marco:
            self.log.print("progresso: %d%% — %s" % (pct, rotulo))

    def _loop(self):
        while not self._parar.is_set():
            self._desenhar()
            time.sleep(self.intervalo)
