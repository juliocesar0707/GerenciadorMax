"""Execução do 7-Zip com diagnóstico de erro legível.

O 7-Zip explica a falha em texto na saída padrão, mas devolve apenas um
código numérico ao processo. Rodando com `check=True` e sem capturar a
saída, o usuário recebia "returned non-zero exit status 2" — que não diz
nada. Este módulo captura a saída e devolve o motivo real.
"""

import os
import subprocess
import logging

logger = logging.getLogger(__name__)

# Termos que marcam a linha útil na saída do 7-Zip (ele é verboso).
_MARCADORES_DE_ERRO = ('error', 'cannot', 'denied', 'access', 'sharing', 'warning')


def arquivo_bloqueado(caminho):
    """Indica se o arquivo existe e está travado por outro processo.

    O Windows mantém o .exe em execução aberto sem compartilhar escrita,
    então abri-lo para escrita falha. É exatamente a condição que faz o
    7-Zip abortar ao tentar sobrescrever o sistema em uso.
    """
    if not os.path.isfile(caminho):
        return False
    try:
        with open(caminho, 'a+b'):
            return False
    except OSError:
        return True


def executaveis_bloqueados(caminhos):
    """Nomes dos executáveis travados, na ordem recebida."""
    return [os.path.basename(c) for c in caminhos if c and arquivo_bloqueado(c)]


def mensagem_de_erro(resultado):
    """Monta a mensagem de falha a partir da saída capturada do 7-Zip."""
    saida = f"{resultado.stdout or ''}\n{resultado.stderr or ''}"
    linhas = [ln.strip() for ln in saida.splitlines() if ln.strip()]

    relevantes = [ln for ln in linhas
                  if any(m in ln.lower() for m in _MARCADORES_DE_ERRO)]
    # dict.fromkeys preserva a ordem e remove repetições
    unicas = list(dict.fromkeys(relevantes))[:6]
    detalhe = "\n".join(unicas)
    if not detalhe:
        detalhe = "\n".join(linhas[-6:])

    return f"7-Zip falhou (código {resultado.returncode}).\n\n{detalhe}".strip()


def executar(cmd):
    """Roda o 7-Zip capturando a saída.

    Raises:
        RuntimeError: com o motivo relatado pelo próprio 7-Zip.
    """
    resultado = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding='oem',        # 7-Zip escreve na codepage do console
        errors='replace',
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if resultado.returncode != 0:
        msg = mensagem_de_erro(resultado)
        logger.error("7-Zip falhou: %s", msg.replace("\n", " | "))
        raise RuntimeError(msg)
    return resultado


def extrair(caminho_7zip, arquivo, destino):
    """Extrai `arquivo` em `destino`, sobrescrevendo."""
    return executar([caminho_7zip, 'x', arquivo, f'-o{destino}', '-y'])
