"""Serviço centralizado de operações SQL Server."""

import logging
import os
import shutil
import time

import pyodbc

import sevenzip

from app_config import SQL_QUERY_VERSAO

logger = logging.getLogger(__name__)


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
        """String de conexão com autenticação SQL (usuário/senha)."""
        return (
            f'DRIVER={self.config.odbc_driver_restore};'
            f'SERVER={self.config.servidor};'
            f'UID={self.config.usuario};'
            f'PWD={self.config.senha};'
            f'DATABASE={database}'
        )

    def listar_bancos(self):
        """Lista bancos de dados SQL Server (excluindo os de sistema).

        Returns:
            list[str]: Lista de nomes de bancos ordenada.
        """
        try:
            with pyodbc.connect(self._conn_str_trusted(), timeout=2) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT name FROM sys.databases "
                    "WHERE name NOT IN ('master','tempdb','model','msdb') "
                    "ORDER BY name"
                )
                return [row.name for row in cursor.fetchall()]
        except pyodbc.Error as e:
            logger.warning("Erro ao listar bancos: %s", e)
            return []

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
            list[str]: Lista de instâncias (ex: ['127.0.0.1', 'localhost', 'localhost\\SQLEXPRESS']).
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

    def restore_database(self, dbname, backup_path, mdf_path, ldf_path, logical_data, logical_log):
        """Executa RESTORE DATABASE.

        Args:
            dbname: Nome do novo banco.
            backup_path: Caminho do arquivo .BAK/.MAX.
            mdf_path: Caminho destino do arquivo .mdf.
            ldf_path: Caminho destino do arquivo .ldf.
            logical_data: Nome lógico do arquivo de dados.
            logical_log: Nome lógico do arquivo de log.
        """
        safe_name = self.sanitize_db_name(dbname)
        sql = (
            f"RESTORE DATABASE [{safe_name}] FROM DISK='{backup_path}' "
            f"WITH MOVE '{logical_data}' TO '{mdf_path}', "
            f"MOVE '{logical_log}' TO '{ldf_path}', REPLACE"
        )
        logger.info("Executando RESTORE DATABASE [%s]", dbname)
        with pyodbc.connect(self._conn_str_auth(), autocommit=True) as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            while cursor.nextset():
                pass
        logger.info("RESTORE DATABASE [%s] concluído com sucesso", dbname)

    def executar_restore_completo(self, fname, dbname, caminho_backup, pasta_sistema,
                                  caminho_7zip, on_message=None):
        """Orquestra todo o fluxo de restore (extração + SQL).

        Args:
            fname: Nome do arquivo de backup selecionado.
            dbname: Nome do novo banco de dados.
            caminho_backup: Pasta onde estão os backups.
            pasta_sistema: Pasta do sistema (para criar dados{N}).
            caminho_7zip: Caminho do executável 7-Zip.
            on_message: Callback(msg: str) para progresso.

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
            self.restore_database(dbname, final, mdf, ldf, ld, ll)

        finally:
            if tmp_dir and os.path.exists(tmp_dir):
                try:
                    shutil.rmtree(tmp_dir)
                except OSError as e:
                    logger.warning("Erro ao limpar temp: %s", e)
