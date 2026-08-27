"""Painel da nuvem: estado visivel, gating de carga e troca de credenciais."""
import os
import tempfile

import apoio
from apoio import AmbienteFalso, criar_app, linhas

from app_config import EXTENSOES_VERSAO
from webdav_client import (
    ArquivoRemoto, WebDAVClient, formatar_data, formatar_tamanho,
)


class ClienteQuebrado:
    caminho_atual = "/"
    def listar(self, **kw):
        raise OSError("host inacessivel")


class ClienteOk:
    caminho_atual = "/VERSOES/"
    def listar(self, **kw):
        return (["v152", "v151"],
                [ArquivoRemoto("a.rar", "12,3 MB", "10/08/2026 09:15")])


class ClienteVazio:
    caminho_atual = "/vazia/"
    def listar(self, **kw):
        return ([], [])


def test_reconfigurar_mantem_instancia_e_limpa_cache():
    """Os botoes do painel capturam o cliente em closures.

    Trocar credenciais precisa acontecer na MESMA instancia, senao os
    botoes continuariam usando as credenciais antigas.
    """
    w = WebDAVClient("https://exemplo.com/", "u1", "s1")
    w._cache[("/", ())] = (0, [], [])
    antes = id(w)

    w.reconfigurar("https://novo.com", "u2", "s2")

    assert id(w) == antes, "reconfigurar trocou a instancia"
    assert (w.url, w.usuario, w.senha) == ("https://novo.com", "u2", "s2")
    assert w._cache == {}, "cache nao foi limpo"


def test_falha_de_rede_vira_linha_visivel():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._popular_nuvem(ClienteQuebrado(), app.lb_nuvem_versoes,
                           app.lbl_caminho_versoes, EXTENSOES_VERSAO)
        app.update()

        l = linhas(app.lb_nuvem_versoes)
        assert len(l) == 1 and l[0][0] == "\u26a0", f"esperava linha de erro, veio {l}"
        assert "host inacessivel" in l[0][1]
        assert id(app.lb_nuvem_versoes) not in app._nuvem_carregada, \
            "erro nao deve marcar a aba como carregada"
    finally:
        apoio.destruir(app); amb.fechar()


def test_listagem_popula_e_marca_como_carregada():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._popular_nuvem(ClienteOk(), app.lb_nuvem_versoes,
                           app.lbl_caminho_versoes, EXTENSOES_VERSAO)
        app.update()

        l = linhas(app.lb_nuvem_versoes)
        assert [x[0] for x in l] == ["\U0001f4c1", "\U0001f4c1", "\U0001f4c4"], f"veio {l}"
        assert app.lbl_caminho_versoes.cget("text") == "/VERSOES/"
        assert id(app.lb_nuvem_versoes) in app._nuvem_carregada
    finally:
        apoio.destruir(app); amb.fechar()


def test_tamanho_e_data_aparecem_na_linha_do_arquivo():
    """O PROPFIND ja devolve os dois; jogar fora obrigava a abrir a nuvem
    no navegador so para conferir qual backup e o certo."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._popular_nuvem(ClienteOk(), app.lb_nuvem_versoes,
                           app.lbl_caminho_versoes, EXTENSOES_VERSAO)
        app.update()

        arquivo = [l for l in linhas(app.lb_nuvem_versoes) if l[0] == "\U0001f4c4"][0]
        assert arquivo[1] == "a.rar", arquivo
        assert arquivo[2] == "12,3 MB", arquivo
        assert arquivo[3] == "10/08/2026 09:15", arquivo

        # pastas nao tem tamanho nem data
        pasta = [l for l in linhas(app.lb_nuvem_versoes) if l[0] == "\U0001f4c1"][0]
        assert pasta[2] == "" and pasta[3] == "", pasta
    finally:
        apoio.destruir(app); amb.fechar()


def test_formatacao_de_tamanho_e_data():
    assert formatar_tamanho(0) == "0 B"
    assert formatar_tamanho(512) == "512 B"
    assert formatar_tamanho(1536) == "1,5 KB"
    assert formatar_tamanho(155807877) == "148,6 MB"
    assert formatar_tamanho(None) == ""

    assert formatar_data("Mon, 10 Aug 2026 09:15:00 GMT") == "10/08/2026 09:15"
    assert formatar_data("") == ""
    assert formatar_data("data invalida") == ""


def test_download_interrompido_nao_deixa_arquivo_com_nome_final():
    """Uma queda no meio deixava um .rar truncado com o nome certo, que so
    falhava depois, na extracao, com um erro do 7-Zip que nao explicava nada."""
    import urllib.request

    class RespostaQueCai:
        def __enter__(self): return self
        def __exit__(self, *a): return False
        def getheader(self, nome): return "1000"
        def read(self, n): raise OSError("conexao perdida")

    original = urllib.request.urlopen
    urllib.request.urlopen = lambda *a, **k: RespostaQueCai()
    try:
        with tempfile.TemporaryDirectory() as d:
            w = WebDAVClient("https://exemplo.com", "u", "s")
            try:
                w.download("versao.rar", d)
            except OSError:
                pass
            else:
                raise AssertionError("deveria propagar a falha")

            assert os.listdir(d) == [], f"sobrou lixo: {os.listdir(d)}"
    finally:
        urllib.request.urlopen = original


def test_pasta_vazia_tem_rotulo_explicito():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._popular_nuvem(ClienteVazio(), app.lb_nuvem_backups,
                           app.lbl_caminho_backups, EXTENSOES_VERSAO)
        app.update()

        l = linhas(app.lb_nuvem_backups)
        assert len(l) == 1 and l[0][1] == "(pasta vazia)", f"veio {l}"
    finally:
        apoio.destruir(app); amb.fechar()


def test_salvar_config_reconfigura_os_mesmos_clientes():
    import ui_app
    amb = AmbienteFalso()
    app = criar_app(amb)
    original = ui_app.threading.Thread
    try:
        apoio.silenciar_dialogos(ui_app)
        id_v, id_b = id(app.webdav_versoes), id(app.webdav_backups)

        app.cfg_vars = {}
        app.cfg.url_cloud = "https://outro.example"
        app.cfg.usuario_cloud = "novo_user"
        app.cfg.senha_cloud = "nova_senha"
        # impede o recarregamento SQL disparado ao salvar
        ui_app.threading.Thread = type(
            "ThreadFalsa", (), {"__init__": lambda s, **k: None,
                                "start": lambda s: None})

        app._salvar_config_aba()

        assert id(app.webdav_versoes) == id_v and id(app.webdav_backups) == id_b, \
            "clientes recriados: as closures dos botoes ficariam obsoletas"
        assert app.webdav_versoes.usuario == "novo_user"
        assert app.webdav_backups.senha == "nova_senha"
        assert not app._nuvem_carregada, "deveria forcar nova listagem"
    finally:
        ui_app.threading.Thread = original
        apoio.restaurar_dialogos(ui_app)
        apoio.destruir(app); amb.fechar()
