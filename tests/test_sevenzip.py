"""Diagnostico de falhas do 7-Zip e deteccao de executavel em uso."""
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import sevenzip


class Resultado:
    """Imita subprocess.CompletedProcess."""
    def __init__(self, returncode, stdout="", stderr=""):
        self.returncode, self.stdout, self.stderr = returncode, stdout, stderr


def test_mensagem_extrai_o_motivo_real():
    """A saida do 7-Zip deve virar mensagem, nao so o codigo numerico."""
    saida = (
        "7-Zip 24.09 (x64)\n"
        "Scanning the drive for archives:\n"
        "1 file, 155807877 bytes\n"
        "\n"
        "ERROR: Cannot open output file MAX_Manager2.exe\n"
        "The process cannot access the file because it is being used by another process.\n"
    )
    msg = sevenzip.mensagem_de_erro(Resultado(2, stdout=saida))
    assert "codigo 2" in msg.replace("ó", "o")
    assert "MAX_Manager2.exe" in msg, msg
    assert "used by another process" in msg, msg
    # o ruido do cabecalho nao entra
    assert "Scanning the drive" not in msg, msg


def test_mensagem_sem_marcador_usa_o_final_da_saida():
    msg = sevenzip.mensagem_de_erro(Resultado(255, stdout="linha a\nlinha b\n"))
    assert "linha b" in msg


def test_mensagem_remove_repeticoes():
    saida = "ERROR: igual\nERROR: igual\nERROR: outro\n"
    msg = sevenzip.mensagem_de_erro(Resultado(2, stdout=saida))
    assert msg.count("ERROR: igual") == 1, msg


def test_arquivo_inexistente_nao_esta_bloqueado():
    assert sevenzip.arquivo_bloqueado(r"C:\nao\existe\nada.exe") is False


def test_arquivo_livre_nao_esta_bloqueado():
    with tempfile.TemporaryDirectory() as d:
        alvo = os.path.join(d, "livre.exe")
        with open(alvo, "wb") as f:
            f.write(b"x")
        assert sevenzip.arquivo_bloqueado(alvo) is False


def test_arquivo_com_escrita_travada_e_detectado():
    """Simula o .exe em execucao.

    O Windows abre o executavel em uso negando partilha de escrita, e e
    isso que faz o 7-Zip abortar. msvcrt.locking nao serve aqui: ele trava
    faixas de bytes, mas ainda deixa o arquivo ser aberto para escrita.
    """
    import win32con
    import win32file

    with tempfile.TemporaryDirectory() as d:
        alvo = os.path.join(d, "emuso.exe")
        with open(alvo, "wb") as f:
            f.write(b"x" * 16)

        handle = win32file.CreateFile(
            alvo,
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ,   # nao partilha escrita
            None,
            win32con.OPEN_EXISTING,
            0,
            None,
        )
        try:
            travado = sevenzip.arquivo_bloqueado(alvo)
        finally:
            handle.Close()

        assert travado is True, "deveria detectar o arquivo em uso"


def test_executaveis_bloqueados_lista_so_os_nomes():
    with tempfile.TemporaryDirectory() as d:
        livre = os.path.join(d, "livre.exe")
        with open(livre, "wb") as f:
            f.write(b"x")
        nomes = sevenzip.executaveis_bloqueados([livre, r"C:\nao\existe.exe", ""])
        assert nomes == [], nomes


def test_executar_levanta_com_o_motivo():
    """Roda o 7-Zip de verdade contra um arquivo invalido."""
    sete = r"C:\Program Files\7-Zip\7z.exe"
    if not os.path.isfile(sete):
        print("  (pulado: 7-Zip nao instalado)")
        return
    with tempfile.TemporaryDirectory() as d:
        falso = os.path.join(d, "quebrado.rar")
        with open(falso, "wb") as f:
            f.write(b"isto nao e um rar")
        try:
            sevenzip.extrair(sete, falso, d)
        except RuntimeError as e:
            texto = str(e)
            assert "7-Zip falhou" in texto, texto
            assert len(texto.splitlines()) > 2, f"sem detalhe: {texto!r}"
            return
        raise AssertionError("deveria ter levantado RuntimeError")
