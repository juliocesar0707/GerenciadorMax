"""Executa a suite de testes sem depender de pytest.

Cada arquivo test_*.py deste diretorio e importado e todas as funcoes
test_* sao chamadas. Uso:

    python tests/run_all.py
"""
import importlib.util
import os
import sys
import traceback

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, RAIZ)


def carregar(caminho):
    nome = os.path.splitext(os.path.basename(caminho))[0]
    spec = importlib.util.spec_from_file_location(nome, caminho)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def main():
    arquivos = sorted(
        os.path.join(AQUI, f) for f in os.listdir(AQUI)
        if f.startswith("test_") and f.endswith(".py")
    )
    if not arquivos:
        print("nenhum teste encontrado")
        return 1

    passou = falhou = 0
    for caminho in arquivos:
        print(f"\n== {os.path.basename(caminho)}")
        try:
            modulo = carregar(caminho)
        except Exception:
            falhou += 1
            print("  ERRO ao importar")
            traceback.print_exc()
            continue

        for nome in sorted(vars(modulo)):
            if not nome.startswith("test_"):
                continue
            funcao = getattr(modulo, nome)
            if not callable(funcao):
                continue
            try:
                funcao()
                passou += 1
                print(f"  ok    {nome}")
            except Exception as e:
                falhou += 1
                print(f"  FALHA {nome}: {type(e).__name__}: {e}")
                traceback.print_exc()

    print(f"\n{passou} passaram, {falhou} falharam")
    return 1 if falhou else 0


if __name__ == "__main__":
    sys.exit(main())
