"""Configuracao: ofuscacao, ancoragem do arquivo e janela de setup."""
import os
import tempfile

import apoio
from apoio import AmbienteFalso, criar_app

import app_config
from app_config import CONFIG_FILE_PATH, AppConfig, desofuscar, diretorio_base, ofuscar


def test_ofuscacao_ida_e_volta():
    for original in ("", "senha123", "acao#$% com espaco", "acentuacao ~ ao"):
        assert desofuscar(ofuscar(original)) == original, original


def test_config_antigo_em_texto_puro_continua_legivel():
    """Instalacoes anteriores gravavam a senha sem ofuscacao."""
    assert desofuscar("texto_puro_antigo") == "texto_puro_antigo"


def test_valor_ofuscado_invalido_nao_derruba():
    assert desofuscar("b64:isto-nao-e-base64-valido!!") == ""


def test_caminho_do_config_e_absoluto_e_ancorado_no_app():
    assert os.path.isabs(CONFIG_FILE_PATH)
    assert os.path.dirname(CONFIG_FILE_PATH) == diretorio_base()


def test_mudar_de_diretorio_nao_move_o_config():
    """O .exe roda com qualquer cwd; o config nao pode seguir o cwd."""
    antes = app_config.CONFIG_FILE_PATH
    origem = os.getcwd()
    with tempfile.TemporaryDirectory() as tmp:
        try:
            os.chdir(tmp)
            assert app_config.CONFIG_FILE_PATH == antes
            assert os.listdir(tmp) == [], "criou arquivo no cwd errado"
        finally:
            # o Windows nao remove um diretorio que ainda e o cwd
            os.chdir(origem)


def test_janela_de_setup_monta_nos_dois_modos():
    from ui_config_window import ConfigWindow

    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        for primeira in (True, False):
            win = ConfigWindow(app, amb.cfg, is_first_run=primeira)
            app.update()
            try:
                assert win.winfo_exists()
                assert win.var_sistema.get() == amb.cfg.pasta_do_sistema
                assert win.var_ini.get() == amb.cfg.caminho_do_ini
            finally:
                win.destroy()
                app.update()
    finally:
        apoio.destruir(app); amb.fechar()


def test_get_e_set_campo_percorrem_o_mapa():
    cfg = AppConfig()
    cfg.set_campo("CAMINHOS", "PASTA_DO_SISTEMA", r"X:/qualquer")
    assert cfg.get_campo("CAMINHOS", "PASTA_DO_SISTEMA") == r"X:/qualquer"
    assert cfg.pasta_do_sistema == r"X:/qualquer"


def test_ambiente_de_teste_nao_toca_no_config_real():
    """Guarda-corpo: um teste que salve configuracao nao pode escrever no
    gerenciador_config.ini do usuario. Ja aconteceu uma vez."""
    real = app_config.CONFIG_FILE_PATH
    antes = None
    if os.path.exists(real):
        with open(real, "rb") as f:
            antes = f.read()

    amb = AmbienteFalso()
    try:
        assert app_config.CONFIG_FILE_PATH != real, "CONFIG_FILE_PATH nao foi isolado"
        amb.cfg.url_cloud = "https://gravacao-de-teste.invalido"
        amb.cfg.salvar()
        assert os.path.exists(app_config.CONFIG_FILE_PATH), "salvou fora do temporario"
    finally:
        amb.fechar()

    assert app_config.CONFIG_FILE_PATH == real, "nao restaurou o caminho real"
    if antes is not None:
        with open(real, "rb") as f:
            assert f.read() == antes, "o config real foi alterado por um teste"
