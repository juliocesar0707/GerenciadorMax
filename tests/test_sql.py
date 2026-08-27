"""Regras puras do servico SQL: nomes, escapes e traducao de erro.

Nada aqui abre conexao — sao as decisoes que o servico toma antes de falar
com o servidor, e que ja causaram comando quebrado em campo.
"""
import apoio  # noqa: F401  (ajusta o sys.path e silencia o logging)

from app_config import AppConfig
from sql_service import (
    SqlService, descrever_erro, escapar_literal, sugerir_nome_banco,
)


def test_sugere_o_nome_do_cliente_a_partir_do_arquivo():
    casos = {
        "MAX-Manager_FORTUP_10082026.MAX": "FORTUP",
        "Max_FortupFiscal_20260810.bak": "FortupFiscal",
        "MAX-Manager_CLINICA-SAO-JOSE_01092026.MAX": "CLINICA_SAO_JOSE",
        "backup_CLIENTE.zip": "CLIENTE",
    }
    for arquivo, esperado in casos.items():
        assert sugerir_nome_banco(arquivo) == esperado, arquivo


def test_sugestao_nunca_devolve_vazio():
    """Sobrando so ruido, o nome do arquivo ainda e melhor que campo em branco."""
    assert sugerir_nome_banco("MAX_Manager_backup.bak") == "MAX_Manager_backup"
    assert sugerir_nome_banco("20260810.bak") == "20260810"
    assert sugerir_nome_banco("") == ""


def test_apostrofo_no_caminho_nao_quebra_o_comando():
    """RESTORE ... FROM DISK='...' interpola o caminho como literal."""
    assert escapar_literal(r"C:\Backups\D'Angelo\base.bak") == \
        r"C:\Backups\D''Angelo\base.bak"
    assert escapar_literal(r"C:\Backups\normal.bak") == r"C:\Backups\normal.bak"


def test_senha_com_ponto_e_virgula_fica_contida():
    """Sem as chaves, o ';' encerraria o campo PWD e o resto da senha
    seria lido pelo driver como outro parametro da connection string."""
    cfg = AppConfig()
    cfg.usuario = "sa"
    cfg.senha = "abc;def=1"
    conn = SqlService(cfg)._conn_str_auth()

    assert "PWD={abc;def=1};" in conn, conn


def test_erros_comuns_viram_frase_em_portugues():
    login = "('28000', \"[28000][Microsoft][ODBC Driver 17]Login failed for user 'sa'.\")"
    assert "usuário ou senha" in descrever_erro(Exception(login))

    timeout = "('HYT00', '[HYT00][Microsoft][ODBC Driver 17]Login timeout expired')"
    assert "não respondeu" in descrever_erro(Exception(timeout))


def test_erro_desconhecido_perde_o_prefixo_do_driver():
    bruto = "('42000', '[42000][Microsoft][ODBC Driver 17]Coisa estranha aconteceu')"
    assert descrever_erro(Exception(bruto)) == "Coisa estranha aconteceu"


def test_erro_sem_colchetes_passa_inteiro():
    assert descrever_erro(Exception("host inacessivel")) == "host inacessivel"


def test_nome_de_banco_com_colchete_e_neutralizado():
    assert SqlService.sanitize_db_name("Max_A]B") == "Max_A]]B"
