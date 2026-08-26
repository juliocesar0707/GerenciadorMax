"""Fluxos dos botoes Abrir Sistema e Atualizar Sistema."""
import os

import apoio
from apoio import AmbienteFalso, criar_app

import ui_app

# ui_app.sevenzip E o modulo sevenzip: trocar um atributo aqui vaza
# para toda a suite, entao guardamos os originais para restaurar.
_EXTRAIR_ORIGINAL = ui_app.sevenzip.extrair
_POPEN_ORIGINAL = ui_app.subprocess.Popen


class Executados:
    """Captura as chamadas externas em vez de executa-las."""

    def __init__(self):
        self.extraidos = []
        self.abertos = []


def preparar(respostas=None):
    """Monta ambiente, app e substitui as chamadas externas."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    reg = Executados()

    versoes = amb.criar_versoes("2.4.152.101_Max_Manager.RAR",
                                "2.4.152.97_Max_Manager.RAR")
    app._popular_versoes_locais()
    app.update()

    apoio.silenciar_dialogos(ui_app, respostas)
    ui_app.sevenzip.extrair = lambda sete, arq, destino: reg.extraidos.append(arq)
    ui_app.subprocess.Popen = lambda cmd, **kw: reg.abertos.append(cmd)
    # o guarda de "sistema aberto" tem teste proprio; aqui nao interfere
    app._bloqueio_de_extracao = lambda: []
    # extrai de forma sincrona, sem depender de timing de thread
    app._extrair_em_thread = lambda arq, cb: app._thread_extrair(arq, cb)

    return amb, app, reg, sorted(versoes, reverse=True)


def encerrar(app, amb):
    """Desfaz os monkeypatches globais e fecha o ambiente."""
    ui_app.sevenzip.extrair = _EXTRAIR_ORIGINAL
    ui_app.subprocess.Popen = _POPEN_ORIGINAL
    apoio.restaurar_dialogos(ui_app)
    apoio.destruir(app)
    amb.fechar()


def selecionar(app, indice):
    filhos = app.lb_versoes.get_children()
    app.lb_versoes.selection_set(filhos[indice])
    app.update()


def test_abrir_extrai_a_versao_selecionada_e_abre_o_manager():
    amb, app, reg, versoes = preparar()
    try:
        selecionar(app, 1)
        app._lancar_erp()
        app.update()

        assert reg.extraidos == [os.path.join(amb.versoes, versoes[1])], reg.extraidos
        assert reg.abertos == [[amb.cfg.caminho_do_erp_cliente]], reg.abertos
    finally:
        encerrar(app, amb)


def test_trocar_a_selecao_troca_o_arquivo_extraido():
    amb, app, reg, versoes = preparar()
    try:
        selecionar(app, 0)
        app._lancar_erp()
        app.update()

        assert reg.extraidos == [os.path.join(amb.versoes, versoes[0])], reg.extraidos
    finally:
        encerrar(app, amb)


def test_recusar_a_confirmacao_cancela():
    amb, app, reg, _ = preparar(respostas={"askyesno": False})
    try:
        selecionar(app, 0)
        app._lancar_erp()
        app.update()

        assert not reg.extraidos and not reg.abertos, \
            f"extraidos={reg.extraidos} abertos={reg.abertos}"
    finally:
        encerrar(app, amb)


def test_atualizar_abre_o_atualizador():
    amb, app, reg, versoes = preparar()
    try:
        selecionar(app, 1)
        app._lancar_atualizacao()
        app.update()

        assert reg.extraidos == [os.path.join(amb.versoes, versoes[1])], reg.extraidos
        assert reg.abertos == [[amb.cfg.caminho_do_max_atualiza]], reg.abertos
    finally:
        encerrar(app, amb)


def test_sem_selecao_abre_direto_sem_extrair():
    amb, app, reg, _ = preparar()
    try:
        app.lb_versoes.selection_remove(*app.lb_versoes.get_children())
        app.lbl_versao_sql.config(text="---")   # desliga a checagem de versao
        app.update()

        app._lancar_erp()
        app.update()

        assert not reg.extraidos, "nao deveria extrair sem selecao"
        assert reg.abertos == [[amb.cfg.caminho_do_erp_cliente]]
    finally:
        encerrar(app, amb)


def test_falha_na_extracao_nao_abre_o_sistema():
    amb, app, reg, _ = preparar()
    try:
        def explode(sete, arq, destino):
            raise RuntimeError("7-Zip falhou (codigo 2).")
        ui_app.sevenzip.extrair = explode

        selecionar(app, 0)
        app._lancar_erp()
        app.update()

        assert not reg.abertos, "nao deveria abrir apos falha na extracao"
    finally:
        encerrar(app, amb)


def test_sistema_aberto_impede_a_extracao():
    """O guarda deve barrar antes de chamar o 7-Zip."""
    amb, app, reg, _ = preparar()
    try:
        vistos = apoio.silenciar_dialogos(ui_app)
        app._bloqueio_de_extracao = lambda: ["MAX_manager2.exe"]
        # restaura o caminho real, sem atalho sincrono
        del app._extrair_em_thread

        selecionar(app, 0)
        app._lancar_erp()
        app.update()

        assert not reg.extraidos, "nao deveria chamar o 7-Zip com o sistema aberto"
        assert vistos["aviso"], "deveria avisar que o sistema esta aberto"
        assert "MAX_manager2.exe" in str(vistos["aviso"])
    finally:
        encerrar(app, amb)
