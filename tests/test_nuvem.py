"""Painel da nuvem: estado visivel, gating de carga e troca de credenciais."""
import os
import tempfile

import apoio
from apoio import AmbienteFalso, criar_app, linhas

from app_config import EXTENSOES_VERSAO
from webdav_client import (
    ArquivoRemoto, WebDAVClient, formatar_data, formatar_tamanho,
)


class ClienteFalso:
    """Base dos dubles: sem nada em cache, o painel sempre vai ao 'servidor'."""
    caminho_atual = "/"

    def consultar_cache(self, **kw):
        return None


class ClienteQuebrado(ClienteFalso):
    def listar(self, **kw):
        raise OSError("host inacessivel")


class ClienteOk(ClienteFalso):
    caminho_atual = "/VERSOES/"
    def listar(self, **kw):
        return (["v152", "v151"],
                [ArquivoRemoto("a.rar", "12,3 MB", "10/08/2026 09:15")])


class ClienteVazio(ClienteFalso):
    caminho_atual = "/vazia/"
    def listar(self, **kw):
        return ([], [])


class ClienteComCache(ClienteFalso):
    """Tem listagem guardada e conta quantas vezes foi ao servidor."""
    caminho_atual = "/VERSOES/"

    def __init__(self, vencido, quebra_ao_atualizar=False):
        self.vencido = vencido
        self.quebra_ao_atualizar = quebra_ao_atualizar
        self.idas_ao_servidor = 0

    def consultar_cache(self, **kw):
        return (["antiga"], [ArquivoRemoto("cache.rar", "1,0 MB", "01/01/2026 00:00")],
                self.vencido)

    def listar(self, **kw):
        self.idas_ao_servidor += 1
        if self.quebra_ao_atualizar:
            raise OSError("host inacessivel")
        return (["nova"], [ArquivoRemoto("fresca.rar", "2,0 MB", "02/02/2026 00:00")])


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


XML_PROPFIND = b"""<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:">
  <d:response><d:href>/remote.php/webdav/</d:href><d:propstat><d:prop>
    <d:resourcetype><d:collection/></d:resourcetype></d:prop></d:propstat></d:response>
  <d:response><d:href>/remote.php/webdav/PASTA/</d:href><d:propstat><d:prop>
    <d:resourcetype><d:collection/></d:resourcetype></d:prop></d:propstat></d:response>
  <d:response><d:href>/remote.php/webdav/versao.rar</d:href><d:propstat><d:prop>
    <d:resourcetype/><d:getcontentlength>1048576</d:getcontentlength>
    <d:getlastmodified>Mon, 10 Aug 2026 09:15:00 GMT</d:getlastmodified>
    </d:prop></d:propstat></d:response>
  <d:response><d:href>/remote.php/webdav/backup.bak</d:href><d:propstat><d:prop>
    <d:resourcetype/><d:getcontentlength>2097152</d:getcontentlength>
    <d:getlastmodified>Tue, 11 Aug 2026 10:00:00 GMT</d:getlastmodified>
    </d:prop></d:propstat></d:response>
</d:multistatus>"""


def test_abas_compartilham_o_cache_e_so_buscam_uma_vez():
    """As duas abas listam a mesma pasta e so mudam o filtro de extensao.

    Com o cache chaveado por (caminho, extensoes), abrir o painel disparava
    dois PROPFIND identicos da mesma pasta.
    """
    chamadas = []
    original = WebDAVClient._propfind
    WebDAVClient._propfind = lambda self, path: (chamadas.append(path) or XML_PROPFIND)
    try:
        compartilhado = {}
        versoes = WebDAVClient("https://ex.com", "u", "s", cache=compartilhado)
        backups = WebDAVClient("https://ex.com", "u", "s", cache=compartilhado)

        pastas_v, arq_v = versoes.listar(path="/", extensoes=(".rar",))
        pastas_b, arq_b = backups.listar(path="/", extensoes=(".bak",))

        assert len(chamadas) == 1, f"foi ao servidor {len(chamadas)} vezes"
        assert pastas_v == pastas_b == ["PASTA"]
        assert [a.nome for a in arq_v] == ["versao.rar"], arq_v
        assert [a.nome for a in arq_b] == ["backup.bak"], arq_b
        # os metadados sobrevivem ao cache
        assert arq_v[0].tamanho == "1,0 MB"
        assert arq_b[0].modificado == "11/08/2026 10:00"
    finally:
        WebDAVClient._propfind = original


def test_conexao_e_reaproveitada_e_se_refaz_quando_cai():
    """O urllib abria socket novo e mandava 'Connection: close' a cada
    chamada, entao todo clique pagava DNS + TCP + TLS de novo.

    Sobe um servidor real em localhost para provar que agora sao N listagens
    em uma conexao so — e que uma conexao derrubada pelo servidor (o que
    acontece o tempo todo, por timeout de ociosidade) e refeita sozinha.
    """
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    conexoes = []
    corpos = []

    class Handler(BaseHTTPRequestHandler):
        protocol_version = 'HTTP/1.1'   # sem isso nao ha keep-alive

        def do_PROPFIND(self):
            # drenar o corpo e obrigatorio para reusar a conexao
            corpos.append(self.rfile.read(int(self.headers.get('Content-Length', 0))))
            self.send_response(207)
            self.send_header('Content-Type', 'text/xml')
            self.send_header('Content-Length', str(len(XML_PROPFIND)))
            self.end_headers()
            self.wfile.write(XML_PROPFIND)

        def log_message(self, *a):
            pass

    class Servidor(ThreadingHTTPServer):
        def process_request(self, request, addr):
            conexoes.append(addr)
            super().process_request(request, addr)

    srv = Servidor(('127.0.0.1', 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        c = WebDAVClient(f'http://127.0.0.1:{srv.server_address[1]}', 'u', 's')

        for i in range(3):
            c.listar(path=f'/pasta{i}/', force_refresh=True)
        assert len(conexoes) == 1, f"abriu {len(conexoes)} conexoes para 3 listagens"

        # o PROPFIND leva corpo: pede so as 3 propriedades, em vez do allprop
        assert corpos[0] and b'getcontentlength' in corpos[0], corpos[0]

        # servidor derruba a conexao ociosa
        c._conn.sock.close()
        _, arquivos = c.listar(path='/depois/', force_refresh=True)

        assert len(conexoes) == 2, "deveria ter reconectado uma unica vez"
        assert [a.nome for a in arquivos] == ["versao.rar"], "retry perdeu o resultado"
    finally:
        srv.shutdown()
        srv.server_close()


def test_consultar_cache_devolve_vencido_sem_ir_ao_servidor():
    chamadas = []
    original = WebDAVClient._propfind
    WebDAVClient._propfind = lambda self, path: (chamadas.append(path) or XML_PROPFIND)
    try:
        c = WebDAVClient("https://ex.com", "u", "s", cache_ttl=0)
        assert c.consultar_cache(path="/") is None, "nao deveria ter nada ainda"

        c.listar(path="/", extensoes=(".rar",))
        assert len(chamadas) == 1

        pastas, arquivos, vencido = c.consultar_cache(path="/", extensoes=(".rar",))
        assert vencido is True, "ttl zero deveria marcar como vencido"
        assert [a.nome for a in arquivos] == ["versao.rar"]
        assert len(chamadas) == 1, "consultar_cache nao pode ir ao servidor"
    finally:
        WebDAVClient._propfind = original


def test_cache_valido_desenha_sem_ir_ao_servidor():
    amb = AmbienteFalso()
    app = criar_app(amb)
    cliente = ClienteComCache(vencido=False)
    try:
        app._popular_nuvem(cliente, app.lb_nuvem_versoes,
                           app.lbl_caminho_versoes, EXTENSOES_VERSAO)
        app.update()

        assert cliente.idas_ao_servidor == 0, "cache valido nao deveria buscar"
        assert [l[1] for l in linhas(app.lb_nuvem_versoes)] == ["antiga", "cache.rar"]
    finally:
        apoio.destruir(app); amb.fechar()


def test_cache_vencido_mostra_o_antigo_e_atualiza_por_baixo():
    amb = AmbienteFalso()
    app = criar_app(amb)
    cliente = ClienteComCache(vencido=True)
    try:
        app._popular_nuvem(cliente, app.lb_nuvem_versoes,
                           app.lbl_caminho_versoes, EXTENSOES_VERSAO)
        app.update()

        assert cliente.idas_ao_servidor == 1, "vencido deveria atualizar"
        # o desenho final e o dado fresco
        assert [l[1] for l in linhas(app.lb_nuvem_versoes)] == ["nova", "fresca.rar"]
    finally:
        apoio.destruir(app); amb.fechar()


def test_falha_ao_atualizar_preserva_a_lista_que_ja_estava_na_tela():
    """Trocar uma lista util por uma linha de erro so piora a vida de quem
    esta atendendo: o cache continua valendo."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    cliente = ClienteComCache(vencido=True, quebra_ao_atualizar=True)
    try:
        app._popular_nuvem(cliente, app.lb_nuvem_versoes,
                           app.lbl_caminho_versoes, EXTENSOES_VERSAO)
        app.update()

        assert [l[1] for l in linhas(app.lb_nuvem_versoes)] == ["antiga", "cache.rar"]
        assert "desatualizada" in app.status.cget("text"), app.status.cget("text")
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
