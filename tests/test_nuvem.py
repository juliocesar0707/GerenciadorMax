"""Painel da nuvem: estado visivel, gating de carga e troca de credenciais."""
import apoio
from apoio import AmbienteFalso, criar_app, linhas

from app_config import EXTENSOES_VERSAO
from webdav_client import WebDAVClient


class ClienteQuebrado:
    caminho_atual = "/"
    def listar(self, **kw):
        raise OSError("host inacessivel")


class ClienteOk:
    caminho_atual = "/VERSOES/"
    def listar(self, **kw):
        return (["v152", "v151"], ["a.rar"])


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
