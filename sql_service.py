"""Serviço centralizado de operações SQL Server."""

import contextlib
import datetime
import logging
import os
import re
import shutil
import threading
import time

import pyodbc

import sevenzip

from app_config import SQL_QUERY_VERSAO

logger = logging.getLogger(__name__)

# Palavras que aparecem no nome dos arquivos de backup e não fazem parte do
# nome do cliente. Usadas por `sugerir_nome_banco`.
_RUIDO_NO_NOME = {
    'max', 'manager', 'maxmanager', 'backup', 'bkp', 'bak', 'db', 'base',
    'dados', 'full', 'completo',
}

# Traduções dos erros de conexão mais comuns. O texto do pyodbc é longo e em
# inglês; na barra de status só cabe o essencial.
_ERROS_CONHECIDOS = (
    ('login failed', 'Login recusado: usuário ou senha do SQL incorretos.'),
    ('login timeout', 'Tempo esgotado: o servidor SQL não respondeu.'),
    ('server is not found', 'Servidor SQL não encontrado ou inacessível.'),
    ('could not open a connection', 'Servidor SQL não encontrado ou inacessível.'),
    ('data source name not found', 'Driver ODBC não instalado ou nome errado.'),
    ('permission was denied', 'Permissão negada para esta operação.'),
)


def descrever_erro(e):
    """Traduz um erro do pyodbc para uma frase curta e legível.

    O pyodbc empacota a mensagem como
    ('08001', '[Microsoft][ODBC Driver 17...]Login timeout expired'), que não
    cabe na barra de status nem ajuda quem está no atendimento.
    """
    texto = str(e)
    baixo = texto.lower()
    for marca, traducao in _ERROS_CONHECIDOS:
        if marca in baixo:
            return traducao

    # Sem tradução conhecida: fica com o trecho após o último "[driver]"
    partes = re.findall(r'\]([^\[\]]+)', texto)
    msg = partes[-1] if partes else texto
    return msg.strip(" '\")\n") or texto


def escapar_literal(valor):
    """Escapa um valor para uso dentro de aspas simples em SQL.

    Caminhos e nomes lógicos entram no RESTORE/BACKUP como literais; um
    apóstrofo no caminho quebraria o comando.
    """
    return str(valor).replace("'", "''")


def sugerir_nome_banco(nome_arquivo):
    """Deriva um nome de banco a partir do nome do arquivo de backup.

    'MAX-Manager_FORTUP_10082026.MAX' -> 'FORTUP'

    Descarta a extensão, os blocos de data (6 dígitos ou mais) e as palavras
    genéricas que todo backup do MAX carrega. Não sobrando nada, devolve o
    nome do arquivo sem extensão, que ainda é melhor que um campo vazio.
    """
    if not nome_arquivo:
        return ''
    base = os.path.splitext(nome_arquivo)[0]
    partes = [p for p in re.split(r'[\s_\-.]+', base) if p]

    uteis = [
        p for p in partes
        if p.lower() not in _RUIDO_NO_NOME
        and not (p.isdigit() and len(p) >= 6)
    ]
    return '_'.join(uteis) if uteis else base


class SqlService:
    """Encapsula todas as operações SQL Server do GerenciadorMax."""

    def __init__(self, config):
        """Inicializa o serviço SQL.

        Args:
            config: Instância de AppConfig com credenciais e drivers.
        """
        self.config = config

    @staticmethod
    def sanitize_db_name(name):
        """Sanitiza nome de banco para uso seguro em SQL DDL.

        Dobra colchetes de fechamento para evitar SQL injection em identificadores.
        """
        return name.replace(']', ']]')

    def _conn_str_trusted(self, database='master'):
        """String de conexão com autenticação Windows (Trusted)."""
        return (
            f'DRIVER={self.config.sql_driver_lista};'
            f'SERVER={self.config.sql_server_instance};'
            f'DATABASE={database};'
            f'Trusted_Connection=yes;'
        )

    def _conn_str_auth(self, database='master'):
        """String de conexão com autenticação SQL (usuário/senha).

        A senha vai entre chaves: sem isso, um ';' na senha encerraria o campo
        e o driver leria o resto como outro parâmetro.
        """
        senha = str(self.config.senha).replace('}', '}}')
        return (
            f'DRIVER={self.config.odbc_driver_restore};'
            f'SERVER={self.config.servidor};'
            f'UID={self.config.usuario};'
            f'PWD={{{senha}}};'
            f'DATABASE={database}'
        )

    def listar_bancos(self):
        """Lista bancos de dados SQL Server (excluindo os de sistema).

        Returns:
            list[str]: Lista de nomes de bancos ordenada.

        Raises:
            pyodbc.Error: Se a conexão falhar. Devolver lista vazia aqui
                tornaria "SQL fora do ar" indistinguível de "nenhum banco",
                e quem está atendendo não saberia o que investigar.
        """
        with pyodbc.connect(self._conn_str_trusted(), timeout=5) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name FROM sys.databases "
                "WHERE name NOT IN ('master','tempdb','model','msdb') "
                "ORDER BY name"
            )
            return [row.name for row in cursor.fetchall()]

    def testar_conexao(self):
        """Abre uma conexão com as credenciais de RESTORE (usuário/senha).

        Returns:
            str: Primeira linha do @@VERSION, para confirmar em qual servidor
                a conexão caiu.

        Raises:
            pyodbc.Error: Se a conexão ou a consulta falharem.
        """
        with pyodbc.connect(self._conn_str_auth(), timeout=5) as conn:
            linha = conn.cursor().execute("SELECT @@VERSION").fetchone()
        return str(linha[0]).splitlines()[0].strip() if linha else 'SQL Server'

    def get_versao(self, db):
        """Obtém a versão do sistema Maxdata armazenada no banco.

        Args:
            db: Nome do banco de dados.

        Returns:
            str: Versão encontrada ou '---' em caso de erro.
        """
        if not db or "ERRO" in db or "Nenhum" in db:
            return "---"
        try:
            with pyodbc.connect(self._conn_str_trusted(db), timeout=1) as conn:
                cursor = conn.cursor()
                cursor.execute(SQL_QUERY_VERSAO)
                row = cursor.fetchone()
                return str(row[0]) if row else "N/A"
        except pyodbc.Error as e:
            logger.debug("Erro ao obter versão do banco '%s': %s", db, e)
            return "---"

    def listar_instancias(self):
        """Lista instâncias SQL Server instaladas via registro do Windows.

        Returns:
            list[str]: Instâncias encontradas, sempre com localhost à frente.
        """
        instancias = ["127.0.0.1", "localhost"]
        try:
            import winreg
            registry_key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
            )
            for i in range(1024):
                try:
                    name, value, _ = winreg.EnumValue(registry_key, i)
                    if name != "MSSQLSERVER":
                        inst_name = f"localhost\\{name}"
                        if inst_name not in instancias:
                            instancias.append(inst_name)
                except OSError:
                    break
        except Exception as e:
            logger.debug("Erro ao ler instâncias do registro: %s", e)
        return instancias

    def drop_databases(self, bancos):
        """Elimina múltiplos bancos de dados permanentemente.

        Args:
            bancos: Lista de nomes de bancos a eliminar.

        Raises:
            pyodbc.Error: Se algum DROP falhar.
        """
        logger.warning("DROP DATABASE solicitado para: %s", bancos)
        with pyodbc.connect(self._conn_str_auth(), autocommit=True) as conn:
            cursor = conn.cursor()
            for db in bancos:
                safe_name = self.sanitize_db_name(db)
                cursor.execute(f"ALTER DATABASE [{safe_name}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE")
                cursor.execute(f"DROP DATABASE [{safe_name}]")
                logger.info("Banco eliminado: %s", db)

    # =========================================================================
    # ACOMPANHAMENTO DE PROGRESSO (BACKUP / RESTORE)
    # =========================================================================
    def _poll_progresso(self, spid, parar, on_progress):
        """Consulta `percent_complete` da sessão, por uma conexão separada.

        BACKUP e RESTORE publicam o andamento em sys.dm_exec_requests, mas só
        para quem tem VIEW SERVER STATE. Sem essa permissão a consulta falha e
        o acompanhamento simplesmente não acontece — a barra fica indeterminada,
        que é o comportamento antigo.
        """
        try:
            with pyodbc.connect(self._conn_str_auth(), timeout=5) as conn:
                cursor = conn.cursor()
                while not parar.wait(2):
                    cursor.execute(
                        "SELECT percent_complete FROM sys.dm_exec_requests "
                        "WHERE session_id = ?", spid
                    )
                    linha = cursor.fetchone()
                    if linha and linha[0] is not None:
                        on_progress(float(linha[0]))
        except pyodbc.Error as e:
            logger.debug("Sem acompanhamento de progresso (%s)", e)

    @contextlib.contextmanager
    def _acompanhando(self, conn, on_progress):
        """Reporta o progresso da operação que roda em `conn` durante o bloco."""
        if on_progress is None:
            yield
            return

        spid = conn.cursor().execute("SELECT @@SPID").fetchone()[0]
        parar = threading.Event()
        vigia = threading.Thread(
            target=self._poll_progresso, args=(spid, parar, on_progress), daemon=True
        )
        vigia.start()
        try:
            yield
        finally:
            parar.set()
            vigia.join(timeout=3)

    # =========================================================================
    # BACKUP
    # =========================================================================
    def backup_database(self, dbname, destino_dir, on_progress=None):
        """Gera um .bak do banco na pasta indicada.

        O arquivo leva data e hora no nome, então backups sucessivos do mesmo
        banco não se sobrescrevem.

        Args:
            dbname: Banco a copiar.
            destino_dir: Pasta de destino (a mesma que alimenta o Restaurador).
            on_progress: Callback(percentual: float).

        Returns:
            str: Caminho do arquivo gerado.
        """
        os.makedirs(destino_dir, exist_ok=True)
        nome = f"{dbname}_{datetime.datetime.now():%Y%m%d_%H%M%S}.bak"
        caminho = os.path.join(destino_dir, nome)

        # Sem WITH COMPRESSION: o SQL Server Express não suporta e aborta a
        # operação inteira, e Express é comum nas instalações do MAX.
        sql = (
            f"BACKUP DATABASE [{self.sanitize_db_name(dbname)}] "
            f"TO DISK = '{escapar_literal(caminho)}' "
            f"WITH INIT, STATS = 5"
        )
        logger.info("Executando BACKUP DATABASE [%s] para %s", dbname, caminho)

        with pyodbc.connect(self._conn_str_auth(), autocommit=True) as conn:
            with self._acompanhando(conn, on_progress):
                cursor = conn.cursor()
                cursor.execute(sql)
                while cursor.nextset():
                    pass

        logger.info("BACKUP DATABASE [%s] concluído: %s", dbname, caminho)
        return caminho

    # =========================================================================
    # RESTORE
    # =========================================================================
    def banco_existe(self, dbname):
        """Indica se já existe um banco com esse nome (case-insensitive)."""
        try:
            return dbname.lower() in {b.lower() for b in self.listar_bancos()}
        except pyodbc.Error as e:
            logger.debug("Não foi possível verificar se '%s' existe: %s", dbname, e)
            return False

    def restore_filelistonly(self, backup_path):
        """Obtém nomes lógicos de um arquivo de backup.

        Args:
            backup_path: Caminho completo do arquivo .BAK/.MAX.

        Returns:
            Tupla (logical_data, logical_log).
        """
        with pyodbc.connect(self._conn_str_auth(), autocommit=True) as conn:
            cursor = conn.cursor()
            cursor.execute("RESTORE FILELISTONLY FROM DISK = ?", backup_path)
            ld, ll = None, None
            for r in cursor.fetchall():
                if r.Type == 'D':
                    ld = r.LogicalName
                if r.Type == 'L':
                    ll = r.LogicalName
            return ld, ll

    def restore_database(self, dbname, backup_path, mdf_path, ldf_path,
                         logical_data, logical_log, on_progress=None):
        """Executa RESTORE DATABASE.

        Args:
            dbname: Nome do novo banco.
            backup_path: Caminho do arquivo .BAK/.MAX.
            mdf_path: Caminho destino do arquivo .mdf.
            ldf_path: Caminho destino do arquivo .ldf.
            logical_data: Nome lógico do arquivo de dados.
            logical_log: Nome lógico do arquivo de log.
            on_progress: Callback(percentual: float).
        """
        safe_name = self.sanitize_db_name(dbname)
        sql = (
            f"RESTORE DATABASE [{safe_name}] FROM DISK='{escapar_literal(backup_path)}' "
            f"WITH MOVE '{escapar_literal(logical_data)}' TO '{escapar_literal(mdf_path)}', "
            f"MOVE '{escapar_literal(logical_log)}' TO '{escapar_literal(ldf_path)}', "
            f"REPLACE, STATS = 5"
        )
        logger.info("Executando RESTORE DATABASE [%s]", dbname)
        with pyodbc.connect(self._conn_str_auth(), autocommit=True) as conn:
            with self._acompanhando(conn, on_progress):
                cursor = conn.cursor()
                cursor.execute(sql)
                while cursor.nextset():
                    pass
        logger.info("RESTORE DATABASE [%s] concluído com sucesso", dbname)

    def executar_restore_completo(self, fname, dbname, caminho_backup, pasta_sistema,
                                  caminho_7zip, on_message=None, on_progress=None):
        """Orquestra todo o fluxo de restore (extração + SQL).

        Args:
            fname: Nome do arquivo de backup selecionado.
            dbname: Nome do novo banco de dados.
            caminho_backup: Pasta onde estão os backups.
            pasta_sistema: Pasta do sistema (para criar dados{N}).
            caminho_7zip: Caminho do executável 7-Zip.
            on_message: Callback(msg: str) para progresso.
            on_progress: Callback(percentual: float) da etapa de RESTORE.

        Raises:
            Exception: Se qualquer etapa falhar.
        """
        def msg(texto):
            if on_message:
                on_message(texto)
            logger.info(texto)

        tmp_dir = None
        try:
            msg(f"--- Iniciando Restore: {dbname} ---")
            origem = os.path.join(caminho_backup, fname)
            final = origem

            # Extração se necessário
            if fname.lower().endswith(('.zip', '.rar')):
                msg("A extrair ficheiro...")
                tmp_dir = os.path.join(caminho_backup, f"_tmp_{int(time.time())}")
                os.makedirs(tmp_dir, exist_ok=True)
                sevenzip.extrair(caminho_7zip, origem, tmp_dir)

                encontrado = next(
                    (os.path.join(root, f)
                     for root, _, files in os.walk(tmp_dir)
                     for f in files if f.upper().endswith(('.MAX', '.BAK'))),
                    None
                )
                if not encontrado:
                    raise FileNotFoundError("Ficheiro .MAX/.BAK não encontrado no arquivo.")
                final = encontrado
                msg(f"Encontrado: {os.path.basename(final)}")

            # Criar pasta de dados
            i = 1
            while True:
                d = os.path.join(pasta_sistema, f"dados{i}")
                if not os.path.exists(d):
                    os.makedirs(d)
                    break
                i += 1
            mdf = os.path.join(d, f"{dbname}.mdf")
            ldf = os.path.join(d, f"{dbname}_log.ldf")

            # Obter nomes lógicos
            ld, ll = self.restore_filelistonly(final)

            # Restaurar
            msg("A restaurar SQL...")
            self.restore_database(dbname, final, mdf, ldf, ld, ll,
                                  on_progress=on_progress)

        finally:
            if tmp_dir and os.path.exists(tmp_dir):
                try:
                    shutil.rmtree(tmp_dir)
                except OSError as e:
                    logger.warning("Erro ao limpar temp: %s", e)
