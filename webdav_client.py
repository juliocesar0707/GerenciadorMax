"""Cliente WebDAV reutilizável com cache integrado para Nextcloud."""

import urllib.request
import urllib.parse
import urllib.error
import base64
import time
import xml.etree.ElementTree as ET
import logging

logger = logging.getLogger(__name__)


class WebDAVClient:
    """Cliente WebDAV com cache para navegação e download de arquivos Nextcloud."""

    def __init__(self, url, usuario='', senha='', cache_ttl=120):
        """Inicializa o cliente WebDAV.

        Args:
            url: URL base do Nextcloud (ex: https://cloud.maxdata.com.br).
            usuario: Usuário para autenticação Basic.
            senha: Senha para autenticação Basic.
            cache_ttl: Tempo de vida do cache em segundos (default: 120).
        """
        self.cache_ttl = cache_ttl
        self._cache = {}  # {path: (timestamp, pastas, arquivos)}
        self.caminho_atual = '/'
        self.reconfigurar(url, usuario, senha)

    def reconfigurar(self, url, usuario, senha):
        """Troca as credenciais mantendo a mesma instância.

        A UI guarda referências a este cliente em closures, então substituir a
        instância deixaria os botões apontando para as credenciais antigas.
        """
        self.url = url.rstrip('/') if url else ''
        self.usuario = usuario
        self.senha = senha
        self.limpar_cache()

    def _auth_header(self):
        """Retorna header de autenticação Basic ou None."""
        if self.usuario and self.senha:
            auth = base64.b64encode(
                f"{self.usuario}:{self.senha}".encode('utf-8')
            ).decode('ascii')
            return f"Basic {auth}"
        return None

    def _webdav_url(self, path):
        """Monta a URL WebDAV completa para um caminho."""
        url = self.url
        if not url.startswith('http'):
            url = 'https://' + url
        return url + "/remote.php/webdav" + urllib.parse.quote(path)

    def listar(self, path=None, force_refresh=False, extensoes=('.rar',)):
        """Lista pastas e arquivos em um caminho WebDAV.

        Args:
            path: Caminho a listar (default: caminho_atual).
            force_refresh: Ignora cache e força nova requisição.
            extensoes: Tupla de extensões de arquivo a incluir (case-insensitive).

        Returns:
            Tupla (pastas: list[str], arquivos: list[str]).

        Raises:
            Exception: Se a conexão WebDAV falhar.
        """
        if path is None:
            path = self.caminho_atual

        # Verificar cache
        cache_key = (path, extensoes)
        if not force_refresh and cache_key in self._cache:
            cached_time, pastas, arquivos = self._cache[cache_key]
            if time.time() - cached_time < self.cache_ttl:
                logger.debug("Cache hit para '%s'", path)
                return pastas, arquivos

        # Requisição PROPFIND
        logger.info("WebDAV PROPFIND: %s", path)
        webdav_url = self._webdav_url(path)

        req = urllib.request.Request(webdav_url, method='PROPFIND')
        req.add_header("Depth", "1")
        auth = self._auth_header()
        if auth:
            req.add_header("Authorization", auth)

        with urllib.request.urlopen(req, timeout=10) as response:
            xml_data = response.read()

        root = ET.fromstring(xml_data)
        pastas = []
        arquivos = []
        namespaces = {'d': 'DAV:'}

        primeiro = True
        for resp in root.findall('d:response', namespaces):
            href = resp.find('d:href', namespaces).text
            href = urllib.parse.unquote(href)
            nome_item = [p for p in href.split('/') if p][-1]

            if primeiro:
                primeiro = False
                continue

            propstat = resp.find('d:propstat', namespaces)
            if propstat:
                prop = propstat.find('d:prop', namespaces)
                resourcetype = prop.find('d:resourcetype', namespaces)
                is_collection = (
                    resourcetype is not None
                    and resourcetype.find('d:collection', namespaces) is not None
                )
                if is_collection:
                    pastas.append(nome_item)
                elif extensoes and nome_item.lower().endswith(extensoes):
                    arquivos.append(nome_item)

        pastas.sort()
        arquivos.sort(reverse=True)

        # Guardar no cache
        self._cache[cache_key] = (time.time(), pastas, arquivos)
        logger.info("WebDAV listou %d pastas, %d arquivos em '%s'", len(pastas), len(arquivos), path)

        return pastas, arquivos

    def download(self, arquivo, destino_dir, on_progress=None):
        """Baixa um arquivo do WebDAV para um diretório local.

        Args:
            arquivo: Nome do arquivo a baixar.
            destino_dir: Diretório de destino local.
            on_progress: Callback(percent: int) chamado durante o download.

        Returns:
            str: Caminho completo do arquivo baixado.

        Raises:
            Exception: Se o download falhar.
        """
        import os

        path = self.caminho_atual
        if not path.endswith('/'):
            path += '/'

        webdav_url = self._webdav_url(path + arquivo)

        req = urllib.request.Request(webdav_url)
        auth = self._auth_header()
        if auth:
            req.add_header("Authorization", auth)

        caminho_local = os.path.join(destino_dir, arquivo)
        logger.info("Download WebDAV: %s → %s", arquivo, caminho_local)

        with urllib.request.urlopen(req, timeout=15) as response, \
                open(caminho_local, 'wb') as out_file:
            tamanho_total = response.getheader('content-length')
            tamanho_total = int(tamanho_total) if tamanho_total else None
            baixado = 0
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
