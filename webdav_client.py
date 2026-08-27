"""Cliente WebDAV reutilizável com cache integrado para Nextcloud.

Três decisões guiam este módulo, todas voltadas à espera que o usuário sente
ao abrir e navegar no painel da nuvem:

1. A conexão é reaproveitada entre listagens. O `urllib` abre um socket novo
   e manda `Connection: close` a cada chamada, então cada clique pagava DNS,
   TCP e handshake TLS de novo. Aqui a `http.client.HTTPSConnection` fica de
   pé, e uma conexão derrubada pelo servidor é refeita de forma transparente.
2. O PROPFIND pede só as três propriedades usadas. Sem corpo, o servidor
   monta e envia todas as propriedades de todos os filhos da pasta.
3. O cache guarda a listagem crua por caminho e pode ser compartilhado entre
   clientes, então as abas Versões e Backups não buscam a mesma pasta duas
   vezes só por filtrarem extensões diferentes.
"""

import urllib.request
import urllib.parse
import urllib.error
import base64
import collections
import email.utils
import http.client
import os
import threading
import time
import xml.etree.ElementTree as ET
import logging

logger = logging.getLogger(__name__)

NAMESPACES = {'d': 'DAV:'}

TIMEOUT = 15

# Só o que a listagem realmente usa. Um PROPFIND sem corpo equivale a
# `allprop`: o servidor monta e devolve tudo, e o parser joga fora.
CORPO_PROPFIND = (
    '<?xml version="1.0" encoding="utf-8"?>'
    '<d:propfind xmlns:d="DAV:">'
    '<d:prop>'
    '<d:resourcetype/>'
    '<d:getcontentlength/>'
    '<d:getlastmodified/>'
    '</d:prop>'
    '</d:propfind>'
).encode('utf-8')

# Um arquivo listado na nuvem. `tamanho` e `modificado` já vêm prontos para
# exibição — o PROPFIND devolve os dois junto com o nome.
ArquivoRemoto = collections.namedtuple('ArquivoRemoto', 'nome tamanho modificado')


class ErroWebDAV(Exception):
    """Falha reportada pelo próprio servidor WebDAV."""


def formatar_tamanho(bytes_):
    """Formata um tamanho em bytes para leitura humana ('1,4 GB')."""
    if bytes_ is None:
        return ''
    valor = float(bytes_)
    for unidade in ('B', 'KB', 'MB', 'GB', 'TB'):
        if valor < 1024 or unidade == 'TB':
            if unidade == 'B':
                return f"{int(valor)} B"
            return f"{valor:.1f} {unidade}".replace('.', ',')
        valor /= 1024


def formatar_data(texto):
    """Converte a data RFC 1123 do WebDAV para 'dd/mm/aaaa hh:mm'."""
    if not texto:
        return ''
    try:
        return email.utils.parsedate_to_datetime(texto).strftime('%d/%m/%Y %H:%M')
    except (TypeError, ValueError):
        return ''


def _texto_da_prop(prop, nome):
    """Texto de uma propriedade do PROPFIND, ou None."""
    if prop is None:
        return None
    elemento = prop.find(f'd:{nome}', NAMESPACES)
    return elemento.text if elemento is not None else None


class WebDAVClient:
    """Cliente WebDAV com cache para navegação e download de arquivos Nextcloud."""

    def __init__(self, url, usuario='', senha='', cache_ttl=600, cache=None):
        """Inicializa o cliente WebDAV.

        Args:
            url: URL base do Nextcloud (ex: https://cloud.maxdata.com.br).
            usuario: Usuário para autenticação Basic.
            senha: Senha para autenticação Basic.
            cache_ttl: Segundos até a listagem ser considerada vencida. Vencida
                ela ainda é exibida — quem chama decide atualizar em segundo
                plano (ver `consultar_cache`).
            cache: Dicionário de cache a compartilhar com outro cliente. Sem
                ele, cada instância mantém o seu.
        """
        self.cache_ttl = cache_ttl
        # {caminho: (timestamp, pastas, arquivos_sem_filtro)}
        self._cache = cache if cache is not None else {}
        self.caminho_atual = '/'
        self._conn = None
        self._lock = threading.Lock()
        self.reconfigurar(url, usuario, senha)

    def reconfigurar(self, url, usuario, senha):
        """Troca as credenciais mantendo a mesma instância.

        A UI guarda referências a este cliente em closures, então substituir a
        instância deixaria os botões apontando para as credenciais antigas.
        """
        self.url = url.rstrip('/') if url else ''
        self.usuario = usuario
        self.senha = senha
        # O host pode ter mudado; uma listagem em curso na conexão antiga
        # falha e o retry a refaz já com os dados novos.
        self._fechar()
        self.limpar_cache()

    # --- Conexão ---------------------------------------------------------
    def _partes(self):
        """Divide a URL configurada em (esquema, host, prefixo de caminho)."""
        url = self.url if self.url.startswith('http') else 'https://' + self.url
        partes = urllib.parse.urlsplit(url)
        return partes.scheme or 'https', partes.netloc, partes.path.rstrip('/')

    def _caminho_webdav(self, path):
        """Caminho absoluto do recurso dentro do host (sem esquema nem host)."""
        _, _, prefixo = self._partes()
        return prefixo + "/remote.php/webdav" + urllib.parse.quote(path)

    def _webdav_url(self, path):
        """URL WebDAV completa — usada pelo download, que vai pelo urllib."""
        esquema, host, _ = self._partes()
        return f"{esquema}://{host}{self._caminho_webdav(path)}"

    def _conexao(self):
        """Conexão persistente, criada sob demanda."""
        if self._conn is None:
            esquema, host, _ = self._partes()
            classe = (http.client.HTTPSConnection if esquema == 'https'
                      else http.client.HTTPConnection)
            self._conn = classe(host, timeout=TIMEOUT)
            logger.debug("Nova conexão WebDAV para %s", host)
        return self._conn

    def _fechar(self):
        """Descarta a conexão persistente, se houver."""
        if self._conn is not None:
            try:
                self._conn.close()
            except OSError:
                pass
            self._conn = None

    def _auth_header(self):
        """Retorna header de autenticação Basic ou None."""
        if self.usuario and self.senha:
            auth = base64.b64encode(
                f"{self.usuario}:{self.senha}".encode('utf-8')
            ).decode('ascii')
            return f"Basic {auth}"
        return None

    def _requisicao(self, path, metodo='GET'):
        """Monta uma Request urllib já autenticada (usada no download)."""
        req = urllib.request.Request(self._webdav_url(path), method=metodo)
        auth = self._auth_header()
        if auth:
            req.add_header("Authorization", auth)
        return req

    def _propfind(self, path):
        """Executa o PROPFIND e devolve o XML cru.

        Tenta duas vezes: servidores fecham conexão ociosa sem avisar, e a
        falha só aparece quando a próxima requisição é enviada nela.
        """
        cabecalhos = {
            "Depth": "1",
            "Content-Type": 'text/xml; charset="utf-8"',
        }
        auth = self._auth_header()
        if auth:
            cabecalhos["Authorization"] = auth

        caminho = self._caminho_webdav(path)

        with self._lock:
            for tentativa in (1, 2):
                try:
                    conexao = self._conexao()
                    conexao.request("PROPFIND", caminho,
                                    body=CORPO_PROPFIND, headers=cabecalhos)
                    resposta = conexao.getresponse()
                    # O corpo precisa ser drenado antes de reusar a conexão
                    dados = resposta.read()
                except (OSError, http.client.HTTPException) as e:
                    self._fechar()
                    if tentativa == 2:
                        raise
                    logger.debug("Conexão reaproveitada caiu (%s); refazendo", e)
                    continue

                if resposta.status == 401:
                    raise ErroWebDAV("Credenciais recusadas (HTTP 401).")
                if resposta.status >= 400:
                    raise ErroWebDAV(f"HTTP {resposta.status} {resposta.reason}")
                return dados

    @staticmethod
    def _interpretar(xml_data):
        """Converte a resposta do PROPFIND em (pastas, arquivos)."""
        root = ET.fromstring(xml_data)
        pastas = []
        arquivos = []

        primeiro = True
        for resp in root.findall('d:response', NAMESPACES):
            href = resp.find('d:href', NAMESPACES).text
            href = urllib.parse.unquote(href)
            nome_item = [p for p in href.split('/') if p][-1]

            # A primeira resposta é o próprio diretório consultado
            if primeiro:
                primeiro = False
                continue

            propstat = resp.find('d:propstat', NAMESPACES)
            if propstat is None:
                continue

            prop = propstat.find('d:prop', NAMESPACES)
            resourcetype = prop.find('d:resourcetype', NAMESPACES) if prop is not None else None
            is_collection = (
                resourcetype is not None
                and resourcetype.find('d:collection', NAMESPACES) is not None
            )

            if is_collection:
                pastas.append(nome_item)
            else:
                arquivos.append(ArquivoRemoto(
                    nome=nome_item,
                    tamanho=formatar_tamanho(_texto_da_prop(prop, 'getcontentlength')),
                    modificado=formatar_data(_texto_da_prop(prop, 'getlastmodified')),
                ))

        pastas.sort()
        arquivos.sort(key=lambda a: a.nome, reverse=True)
        return pastas, arquivos

    # --- Cache -----------------------------------------------------------
    @staticmethod
    def _filtrar(arquivos, extensoes):
        """Aplica o filtro de extensão sobre a listagem guardada."""
        if not extensoes:
            return list(arquivos)
        return [a for a in arquivos if a.nome.lower().endswith(extensoes)]

    def consultar_cache(self, path=None, extensoes=('.rar',)):
        """Listagem já em cache, mesmo vencida.

        Permite à UI desenhar a lista na hora e só então decidir se atualiza
        em segundo plano — mostrar dados de dez minutos atrás é melhor que
        mostrar uma tela vazia por três segundos.

        Returns:
            Tupla (pastas, arquivos, vencido: bool), ou None se este caminho
            nunca foi listado.
        """
        if path is None:
            path = self.caminho_atual
        entrada = self._cache.get(path)
        if entrada is None:
            return None
        momento, pastas, todos = entrada
        vencido = (time.time() - momento) >= self.cache_ttl
        return pastas, self._filtrar(todos, extensoes), vencido

    def listar(self, path=None, force_refresh=False, extensoes=('.rar',)):
        """Lista pastas e arquivos em um caminho WebDAV.

        Args:
            path: Caminho a listar (default: caminho_atual).
            force_refresh: Ignora cache e força nova requisição.
            extensoes: Tupla de extensões de arquivo a incluir (case-insensitive).

        Returns:
            Tupla (pastas: list[str], arquivos: list[ArquivoRemoto]).

        Raises:
            ErroWebDAV, OSError: Se a conexão ou o servidor falharem.
        """
        if path is None:
            path = self.caminho_atual

        if not force_refresh:
            entrada = self._cache.get(path)
            if entrada is not None:
                momento, pastas, todos = entrada
                if time.time() - momento < self.cache_ttl:
                    logger.debug("Cache hit para '%s'", path)
                    return pastas, self._filtrar(todos, extensoes)

        logger.info("WebDAV PROPFIND: %s", path)
        pastas, todos = self._interpretar(self._propfind(path))

        # Guardado sem filtro: as duas abas compartilham este cache e só
        # divergem nas extensões que exibem.
        self._cache[path] = (time.time(), pastas, todos)
        logger.info("WebDAV listou %d pastas, %d arquivos em '%s'",
                    len(pastas), len(todos), path)

        return pastas, self._filtrar(todos, extensoes)

    def download(self, arquivo, destino_dir, on_progress=None):
        """Baixa um arquivo do WebDAV para um diretório local.

        Vai pelo urllib, em conexão própria: o download ocupa o socket por
        minutos, e prendê-lo na conexão das listagens travaria a navegação.
        O ganho do keep-alive aqui seria irrelevante perto do tempo de
        transferência.

        Escreve em um `.part` e só renomeia para o nome final ao terminar: uma
        queda de conexão no meio de um .rar de 150 MB deixava um arquivo
        truncado com o nome certo, que só falhava depois, na extração, com um
        erro do 7-Zip que não apontava para o download.

        Args:
            arquivo: Nome do arquivo a baixar.
            destino_dir: Diretório de destino local.
            on_progress: Callback(percent: int) chamado durante o download.

        Returns:
            str: Caminho completo do arquivo baixado.

        Raises:
            Exception: Se o download falhar.
        """
        path = self.caminho_atual
        if not path.endswith('/'):
            path += '/'

        req = self._requisicao(path + arquivo)

        caminho_local = os.path.join(destino_dir, arquivo)
        parcial = caminho_local + '.part'
        logger.info("Download WebDAV: %s → %s", arquivo, caminho_local)

        baixado = 0
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as response, \
                    open(parcial, 'wb') as out_file:
                tamanho_total = response.getheader('content-length')
                tamanho_total = int(tamanho_total) if tamanho_total else None
                bloco = 1024 * 64  # 64KB

                while True:
                    dados = response.read(bloco)
                    if not dados:
                        break
                    out_file.write(dados)
                    baixado += len(dados)
                    if tamanho_total and on_progress:
                        pct = int((baixado / tamanho_total) * 100)
                        on_progress(pct)

            os.replace(parcial, caminho_local)
        except BaseException:
            # Inclui KeyboardInterrupt e o fechamento da janela: em qualquer
            # caso o .part não pode ficar para trás fingindo ser um download.
            try:
                if os.path.exists(parcial):
                    os.remove(parcial)
            except OSError as e:
                logger.warning("Não foi possível remover o parcial '%s': %s", parcial, e)
            raise

        # O arquivo novo muda a listagem desta pasta
        self._cache.pop(self.caminho_atual, None)
        logger.info("Download concluído: %s (%d bytes)", arquivo, baixado)
        return caminho_local

    def navegar(self, pasta):
        """Entra em uma subpasta."""
        if not self.caminho_atual.endswith('/'):
            self.caminho_atual += '/'
        self.caminho_atual += pasta + '/'
        logger.debug("Navegou para: %s", self.caminho_atual)

    def voltar(self):
        """Volta uma pasta. Retorna True se voltou, False se já estava na raiz."""
        if not self.caminho_atual or self.caminho_atual == '/':
            return False
        partes = [p for p in self.caminho_atual.strip('/').split('/') if p]
        if partes:
            partes.pop()
        self.caminho_atual = '/' + '/'.join(partes) + ('/' if partes else '')
        logger.debug("Voltou para: %s", self.caminho_atual)
        return True

    def limpar_cache(self):
        """Limpa todo o cache."""
        self._cache.clear()
        logger.debug("Cache WebDAV limpo")

    def fechar(self):
        """Encerra a conexão persistente (usado ao fechar o app)."""
        self._fechar()
