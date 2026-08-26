"""Apoio comum aos testes: ambiente sintetico, sem tocar na maquina real.

Os testes nao devem depender do gerenciador_config.ini nem das pastas do
usuario. Aqui montamos uma arvore temporaria equivalente e devolvemos um
AppConfig apontando para ela.
"""
import logging
import os
import sys
import tempfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

logging.getLogger().setLevel(logging.CRITICAL)

import app_config
from app_config import AppConfig


class AmbienteFalso:
    """Arvore temporaria com a estrutura que o app espera encontrar."""

    def __init__(self):
        self._tmp = tempfile.TemporaryDirectory()
        base = self._tmp.name

        self.sistema = os.path.join(base, "Max")
        self.versoes = os.path.join(self.sistema, "Versoes")
        self.backup = os.path.join(self.sistema, "backup")
        for d in (self.sistema, self.versoes, self.backup):
            os.makedirs(d, exist_ok=True)

        self.ini = os.path.join(self.sistema, "max.ini")
        with open(self.ini, "w", encoding="windows-1252") as f:
            f.write("[CON]\nInitial catalog=Max_Teste\nData Source=localhost\n")

        self.sete_zip = os.path.join(base, "7z.exe")
        with open(self.sete_zip, "wb") as f:
            f.write(b"nao e um 7zip de verdade")

        # AppConfig.salvar() grava no app_config.CONFIG_FILE_PATH do modulo.
        # Sem redirecionar, qualquer teste que salve configuracao sobrescreve
        # o gerenciador_config.ini REAL do usuario.
        self._config_real = app_config.CONFIG_FILE_PATH
        app_config.CONFIG_FILE_PATH = os.path.join(base, "gerenciador_config.ini")

        self.cfg = AppConfig()
        self.cfg.pasta_do_sistema = self.sistema
        self.cfg.pasta_das_versoes = self.versoes
        self.cfg.caminho_base_backup = self.backup
        self.cfg.caminho_do_ini = self.ini
        self.cfg.caminho_do_7zip = self.sete_zip
        assert self.cfg.validar_caminhos() == [], "ambiente falso incompleto"

    def criar_versoes(self, *nomes):
        """Cria arquivos de versao vazios e devolve os nomes."""
        for n in nomes:
            with open(os.path.join(self.versoes, n), "wb") as f:
                f.write(b"")
        return list(nomes)

    def criar_backups(self, *nomes):
        for n in nomes:
            with open(os.path.join(self.backup, n), "wb") as f:
                f.write(b"")
        return list(nomes)

    def fechar(self):
        app_config.CONFIG_FILE_PATH = self._config_real
        self._tmp.cleanup()


def criar_app(ambiente):
    """Instancia a janela principal com o layout montado, sem carregar dados."""
    from ui_app import GerenciadorMaxApp
    app = GerenciadorMaxApp(ambiente.cfg)
    # Callbacks `after` que sobram do teste disparam apos o destroy e o Tk
    # despeja "application has been destroyed" no stderr. Nao e falha do
    # app; silenciamos para a saida da suite ficar legivel.
    app.report_callback_exception = lambda *a: None
    app.create_layout()
    app.update()
    return app


_DIALOGOS = ("showinfo", "showwarning", "showerror", "askyesno")
_ORIGINAIS = {}


def silenciar_dialogos(modulo, respostas=None):
    """Substitui os messagebox do modulo por no-ops.

    `modulo.messagebox` e o tkinter.messagebox compartilhado, entao a troca
    vaza para os outros testes se nao for desfeita: use restaurar_dialogos.

    Args:
        modulo: modulo que importou `messagebox`.
        respostas: dict opcional com a resposta de askyesno.

    Returns:
        dict com listas 'info', 'aviso', 'erro' do que seria exibido.
    """
    _ORIGINAIS.setdefault(modulo.__name__, {
        nome: getattr(modulo.messagebox, nome) for nome in _DIALOGOS
    })
    vistos = {"info": [], "aviso": [], "erro": []}
    respostas = respostas if respostas is not None else {"askyesno": True}

    modulo.messagebox.showinfo = lambda *a, **k: vistos["info"].append(a)
    modulo.messagebox.showwarning = lambda *a, **k: vistos["aviso"].append(a)
    modulo.messagebox.showerror = lambda *a, **k: vistos["erro"].append(a)
    modulo.messagebox.askyesno = lambda *a, **k: respostas["askyesno"]
    return vistos


def linhas(tree):
    """Valores das linhas de um Treeview."""
    return [tree.item(i, "values") for i in tree.get_children()]


def destruir(app):
    """Fecha a janela cancelando callbacks `after` ainda pendentes.

    Sem isso, um after agendado (barra de status, fila do restore) dispara
    apos o destroy e o Tk imprime "application has been destroyed".
    """
    try:
        app.update()          # escoa afters ja agendados
    except Exception:
        pass
    try:
        pendentes = app.tk.call("after", "info")
        if isinstance(pendentes, str):
            pendentes = pendentes.split()
        for ident in pendentes:
            try:
                app.after_cancel(ident)
            except Exception:
                pass
    except Exception:
        pass
    app.destroy()


def restaurar_dialogos(modulo):
    """Desfaz silenciar_dialogos."""
    for nome, funcao in _ORIGINAIS.pop(modulo.__name__, {}).items():
        setattr(modulo.messagebox, nome, funcao)
