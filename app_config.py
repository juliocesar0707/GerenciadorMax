"""Configuração centralizada do GerenciadorMax."""

import base64
import configparser
import dataclasses
import os
import sys
import logging

logger = logging.getLogger(__name__)

CONFIG_FILE_NAME = 'gerenciador_config.ini'

# Prefixo que marca um valor ofuscado no INI.
# ATENÇÃO: base64 é OFUSCAÇÃO, não criptografia — protege apenas contra leitura
# casual do arquivo, e não contra quem tenha acesso à máquina.
OBFUSCATION_PREFIX = 'b64:'
SQL_QUERY_VERSAO = "select cofMaxAtualizaVersao from config"

# Extensões de arquivo aceitas
EXTENSOES_VERSAO = ('.rar',)
EXTENSOES_BACKUP = ('.max', '.bak', '.zip', '.rar')
EXTENSOES_BACKUP_NUVEM = ('.bak', '.zip', '.rar')

# Mapeamento (section, key) → campo do AppConfig
# Usado pela aba de configurações para popular/salvar
CONFIG_FIELD_MAP = {
    ('CAMINHOS', 'PASTA_DO_SISTEMA'): 'pasta_do_sistema',
    ('CAMINHOS', 'PASTA_DAS_VERSOES'): 'pasta_das_versoes',
    ('CAMINHOS', 'PASTA_DE_BACKUP'): 'caminho_base_backup',
    ('CAMINHOS', 'CAMINHO_DO_INI'): 'caminho_do_ini',
    ('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE'): 'caminho_do_7zip',
    ('EXECUTAVEIS', 'NOME_EXE_CLIENTE'): 'nome_exe_cliente',
    ('EXECUTAVEIS', 'NOME_EXE_ATUALIZADOR'): 'nome_exe_atualizador',
    ('SQL_LAUDO', 'SQL_DRIVER_LISTA'): 'sql_driver_lista',
    ('SQL_LAUDO', 'SQL_SERVER_INSTANCE'): 'sql_server_instance',
    ('SQL_RESTORE', 'SERVIDOR'): 'servidor',
    ('SQL_RESTORE', 'USUARIO'): 'usuario',
    ('SQL_RESTORE', 'SENHA'): 'senha',
    ('SQL_RESTORE', 'ODBC_DRIVER_RESTORE'): 'odbc_driver_restore',
    ('CONFIG_INI_MAX', 'INI_SECTION'): 'ini_section',
    ('CONFIG_INI_MAX', 'INI_KEY'): 'ini_key',
    ('CONFIG_INI_MAX', 'INI_SERVER_KEY'): 'ini_server_key',
    ('CLOUD', 'URL_CLOUD'): 'url_cloud',
    ('CLOUD', 'USUARIO_CLOUD'): 'usuario_cloud',
    ('CLOUD', 'SENHA_CLOUD'): 'senha_cloud',
}

# Seções exibidas na aba de configurações (título, chaves, seção INI)
CONFIG_SECTIONS_UI = [
    ("Caminhos", ["PASTA_DO_SISTEMA", "PASTA_DAS_VERSOES", "PASTA_DE_BACKUP", "CAMINHO_DO_INI", "CAMINHO_DO_7ZIP_EXE"], "CAMINHOS"),
    ("Executáveis", ["NOME_EXE_CLIENTE", "NOME_EXE_ATUALIZADOR"], "EXECUTAVEIS"),
    ("SQL Laudo", ["SQL_DRIVER_LISTA", "SQL_SERVER_INSTANCE"], "SQL_LAUDO"),
    ("SQL Restore", ["SERVIDOR", "USUARIO", "SENHA", "ODBC_DRIVER_RESTORE"], "SQL_RESTORE"),
    ("Config INI MAX", ["INI_SECTION", "INI_KEY", "INI_SERVER_KEY"], "CONFIG_INI_MAX"),
    ("Cloud Nuvem", ["URL_CLOUD", "USUARIO_CLOUD", "SENHA_CLOUD"], "CLOUD"),
]


def ofuscar(valor):
    """Ofusca um valor para gravação no INI. Retorna '' para valores vazios."""
    if not valor:
        return ''
    return OBFUSCATION_PREFIX + base64.b64encode(valor.encode('utf-8')).decode('ascii')


def desofuscar(valor):
    """Reverte `ofuscar`. Valores sem o prefixo são devolvidos como estão
    (compatibilidade com configs antigos gravados em texto puro)."""
    if not valor:
        return ''
    if not valor.startswith(OBFUSCATION_PREFIX):
        return valor
    try:
        return base64.b64decode(valor[len(OBFUSCATION_PREFIX):]).decode('utf-8')
    except (ValueError, UnicodeDecodeError) as e:
        logger.warning("Valor ofuscado inválido no INI: %s", e)
        return ''


@dataclasses.dataclass
class AppConfig:
    """Configuração centralizada do aplicativo, substituindo variáveis globais."""

    # Caminhos
    pasta_do_sistema: str = ""
    pasta_das_versoes: str = ""
    caminho_base_backup: str = ""
    caminho_do_ini: str = ""
    caminho_do_7zip: str = ""

    # Executáveis
    nome_exe_cliente: str = "MAX_manager2.exe"
    nome_exe_atualizador: str = "MAX_Atualiza.exe"

    # Config INI MAX
    ini_section: str = "CON"
    ini_key: str = "Initial catalog"
    ini_server_key: str = "Data Source"

    # SQL Laudo
    sql_driver_lista: str = "{ODBC Driver 17 for SQL Server}"
    sql_server_instance: str = "localhost"

    # SQL Restore
    servidor: str = "localhost"
    usuario: str = "sa"
    senha: str = ""
    odbc_driver_restore: str = "{ODBC Driver 17 for SQL Server}"

    # Cloud
    url_cloud: str = "https://cloud.maxdata.com.br"
    usuario_cloud: str = ""
    senha_cloud: str = ""

    # Caminhos derivados (não persistidos)
    caminho_do_erp_cliente: str = dataclasses.field(default="", repr=False)
    caminho_do_max_atualiza: str = dataclasses.field(default="", repr=False)

    # Flags
    primeira_execucao: bool = dataclasses.field(default=False, repr=False)

    # --- Singleton ---
    _instance = None

    @classmethod
    def get(cls):
        """Retorna a instância singleton da configuração."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @classmethod
    def reset(cls):
        """Reseta a instância singleton (útil para testes)."""
        cls._instance = None

    # --- Auto-detect ---
    def auto_detect(self):
        """Detecta automaticamente caminhos padrão do sistema."""
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

        if os.path.exists(os.path.join(base_dir, 'max.ini')):
            default_sistema = base_dir
        elif os.path.exists(r'C:\Max'):
            default_sistema = r'C:\Max'
        elif os.path.exists(r'D:\Max'):
            default_sistema = r'D:\Max'
        else:
            default_sistema = r'C:\Max'

        self.pasta_do_sistema = default_sistema
        self.pasta_das_versoes = os.path.join(default_sistema, 'Versões')
        self.caminho_base_backup = os.path.join(default_sistema, 'backup')
        self.caminho_do_ini = os.path.join(default_sistema, 'max.ini')

        # Auto-detect 7-Zip
        _7z = r'C:\Program Files\7-Zip\7z.exe'
        if not os.path.exists(_7z):
            _7z = r'C:\Program Files (x86)\7-Zip\7z.exe'
        self.caminho_do_7zip = _7z

        logger.info("Auto-detect: pasta_sistema=%s, 7zip=%s", default_sistema, _7z)

    # --- Carregar / Salvar ---
    def carregar(self):
        """Carrega configurações do arquivo INI. Cria defaults na primeira execução."""
        if not os.path.exists(CONFIG_FILE_NAME):
            self.primeira_execucao = True
            self.auto_detect()
            self.salvar()
            logger.info("Primeira execução — configurações padrão criadas")
            return

        try:
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_NAME, encoding='utf-8')

            self.pasta_do_sistema = config.get('CAMINHOS', 'PASTA_DO_SISTEMA', fallback=self.pasta_do_sistema)
            self.pasta_das_versoes = config.get('CAMINHOS', 'PASTA_DAS_VERSOES', fallback=self.pasta_das_versoes)
            self.caminho_base_backup = config.get('CAMINHOS', 'PASTA_DE_BACKUP',
                                                   fallback=os.path.join(self.pasta_do_sistema, 'backup'))
            self.caminho_do_ini = config.get('CAMINHOS', 'CAMINHO_DO_INI', fallback=self.caminho_do_ini)
            self.caminho_do_7zip = config.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', fallback=self.caminho_do_7zip)

            self.nome_exe_cliente = config.get('EXECUTAVEIS', 'NOME_EXE_CLIENTE', fallback=self.nome_exe_cliente)
            self.nome_exe_atualizador = config.get('EXECUTAVEIS', 'NOME_EXE_ATUALIZADOR',
                                                    fallback=self.nome_exe_atualizador)

            self.ini_section = config.get('CONFIG_INI_MAX', 'INI_SECTION', fallback=self.ini_section)
            self.ini_key = config.get('CONFIG_INI_MAX', 'INI_KEY', fallback=self.ini_key)
            self.ini_server_key = config.get('CONFIG_INI_MAX', 'INI_SERVER_KEY', fallback=self.ini_server_key)

            self.sql_driver_lista = config.get('SQL_LAUDO', 'SQL_DRIVER_LISTA', fallback=self.sql_driver_lista)
            self.sql_server_instance = config.get('SQL_LAUDO', 'SQL_SERVER_INSTANCE',
                                                   fallback=self.sql_server_instance)

            self.servidor = config.get('SQL_RESTORE', 'SERVIDOR', fallback=self.servidor)
            self.usuario = config.get('SQL_RESTORE', 'USUARIO', fallback=self.usuario)
            self.senha = desofuscar(config.get('SQL_RESTORE', 'SENHA', fallback=''))
            self.odbc_driver_restore = config.get('SQL_RESTORE', 'ODBC_DRIVER_RESTORE',
                                                   fallback=self.odbc_driver_restore)

            self.url_cloud = config.get('CLOUD', 'URL_CLOUD', fallback=self.url_cloud)
            self.usuario_cloud = config.get('CLOUD', 'USUARIO_CLOUD', fallback='')
            self.senha_cloud = desofuscar(config.get('CLOUD', 'SENHA_CLOUD', fallback=''))

            logger.info("Configurações carregadas de '%s'", CONFIG_FILE_NAME)
        except Exception as e:
            logger.error("Erro ao carregar configurações: %s", e)

    def salvar(self):
        """Salva configurações no arquivo INI com senhas ofuscadas."""
        config = configparser.ConfigParser()

        config['CAMINHOS'] = {
            'PASTA_DO_SISTEMA': self.pasta_do_sistema,
            'PASTA_DAS_VERSOES': self.pasta_das_versoes,
            'PASTA_DE_BACKUP': self.caminho_base_backup,
            'CAMINHO_DO_INI': self.caminho_do_ini,
            'CAMINHO_DO_7ZIP_EXE': self.caminho_do_7zip,
        }
        config['EXECUTAVEIS'] = {
            'NOME_EXE_CLIENTE': self.nome_exe_cliente,
            'NOME_EXE_ATUALIZADOR': self.nome_exe_atualizador,
        }
        config['CONFIG_INI_MAX'] = {
            'INI_SECTION': self.ini_section,
            'INI_KEY': self.ini_key,
            'INI_SERVER_KEY': self.ini_server_key,
        }
        config['SQL_LAUDO'] = {
            'SQL_DRIVER_LISTA': self.sql_driver_lista,
            'SQL_SERVER_INSTANCE': self.sql_server_instance,
        }
        config['SQL_RESTORE'] = {
            'SERVIDOR': self.servidor,
            'USUARIO': self.usuario,
            'SENHA': ofuscar(self.senha),
            'ODBC_DRIVER_RESTORE': self.odbc_driver_restore,
        }
        config['CLOUD'] = {
            'URL_CLOUD': self.url_cloud,
            'USUARIO_CLOUD': self.usuario_cloud,
            'SENHA_CLOUD': ofuscar(self.senha_cloud),
        }

        try:
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f:
                config.write(f)
            logger.info("Configurações salvas em '%s'", CONFIG_FILE_NAME)
        except OSError as e:
            logger.error("Erro ao salvar configurações: %s", e)

    def validar_caminhos(self):
        """Valida caminhos configurados e atualiza caminhos derivados.

        Returns:
            list[str]: Lista de erros encontrados. Vazia se tudo OK.
        """
        self.caminho_do_erp_cliente = os.path.join(self.pasta_do_sistema, self.nome_exe_cliente)
        self.caminho_do_max_atualiza = os.path.join(self.pasta_do_sistema, self.nome_exe_atualizador)

        erros = []
        if not os.path.isdir(self.pasta_do_sistema):
            erros.append(f"Pasta Sistema: {self.pasta_do_sistema}")
        if not os.path.isdir(self.pasta_das_versoes):
            erros.append(f"Pasta Versões: {self.pasta_das_versoes}")
        if not os.path.isdir(self.caminho_base_backup):
            erros.append(f"Pasta Backups: {self.caminho_base_backup}")
        if not os.path.exists(self.caminho_do_ini):
            erros.append(f"Ficheiro INI: {self.caminho_do_ini}")
        if not os.path.exists(self.caminho_do_7zip):
            erros.append(f"7-Zip: {self.caminho_do_7zip}")

        if erros:
            logger.warning("Caminhos inválidos: %s", erros)
        return erros

    def get_campo(self, section, key):
        """Retorna o valor atual de um campo pelo par (section, key)."""
        field_name = CONFIG_FIELD_MAP.get((section, key))
        if field_name:
            return getattr(self, field_name, '')
        return ''

    def set_campo(self, section, key, valor):
        """Define o valor de um campo pelo par (section, key)."""
        field_name = CONFIG_FIELD_MAP.get((section, key))
        if field_name:
            setattr(self, field_name, valor)
