"""Janela de configuração de caminhos (setup inicial e edição)."""

import tkinter as tk
from tkinter import BOTH, YES, W, X, LEFT, RIGHT
from tkinter import ttk, messagebox, filedialog
import logging

import ui_theme
from ui_widgets import RoundedButton

logger = logging.getLogger(__name__)


class ConfigWindow(tk.Toplevel):
    """Diálogo para configuração de caminhos do sistema."""

    def __init__(self, parent, config, is_first_run=False):
        """Inicializa a janela de configuração.

        Args:
            parent: Janela pai.
            config: Instância de AppConfig.
            is_first_run: True se for a primeira execução.
        """
        super().__init__(parent)
        # NÃO usar self.config: sobrescreveria tk.Misc.config() da janela.
        self.cfg = config
        self.salvo = False

        self.title("Bem-vindo! Setup Inicial" if is_first_run else "Configuração de Caminhos")
        self.geometry("750x380")
        self.resizable(False, False)
        if parent.state() != 'withdrawn':
            self.transient(parent)
        self.grab_set()

        self.var_sistema = tk.StringVar(value=config.pasta_do_sistema)
        self.var_versoes = tk.StringVar(value=config.pasta_das_versoes)
        self.var_backup = tk.StringVar(value=config.caminho_base_backup)
        self.var_ini = tk.StringVar(value=config.caminho_do_ini)
        self.var_7zip = tk.StringVar(value=config.caminho_do_7zip)

        main_frame = ttk.Frame(self, style="Card.TFrame", padding=24)
        main_frame.pack(fill=BOTH, expand=YES)

        titulo = ("Primeira execução! Confirme os diretórios de trabalho:"
                  if is_first_run else "Corrija os caminhos:")
        ttk.Label(
            main_frame, text=titulo, style="CardTitle.TLabel",
            font=(ui_theme.FONT_FAMILY, 12, "bold"),
        ).pack(anchor=W, pady=(0, 18))

        self._criar_campo(main_frame, "Pasta do Sistema", self.var_sistema, True)
        self._criar_campo(main_frame, "Pasta de Versões", self.var_versoes, True)
        self._criar_campo(main_frame, "Pasta de Backups", self.var_backup, True)
        self._criar_campo(main_frame, "Arquivo .ini", self.var_ini, False)
        self._criar_campo(main_frame, "Executável 7-Zip", self.var_7zip, False)

        btn_frame = ttk.Frame(main_frame, style="Card.TFrame")
        btn_frame.pack(fill=X, pady=(22, 0))
        RoundedButton(btn_frame, text="Salvar e Continuar",
                      command=self._salvar, variant="primary").pack(side=RIGHT, padx=(6, 0))
        if not is_first_run:
            RoundedButton(btn_frame, text="Cancelar",
                          command=self.destroy, variant="outline").pack(side=RIGHT)

    def _criar_campo(self, parent, label, var, is_dir):
        """Cria uma linha de campo com label, entry e botão de seleção."""
        f = ttk.Frame(parent, style="Card.TFrame")
        f.pack(fill=X, pady=5)
        ttk.Label(f, text=label, width=20, style="CardMuted.TLabel").pack(side=LEFT)
        ttk.Entry(f, textvariable=var, style="Campo.TEntry").pack(
            side=LEFT, fill=X, expand=YES, padx=6)
        RoundedButton(f, text="...", variant="outline", padx=12,
                      command=lambda: self._selecionar(var, is_dir)).pack(side=LEFT)

    @staticmethod
    def _selecionar(var, is_dir):
        """Abre diálogo de seleção de pasta ou arquivo."""
        p = filedialog.askdirectory() if is_dir else filedialog.askopenfilename()
        if p:
            var.set(p)

    def _salvar(self):
        """Salva os caminhos editados na configuração."""
        try:
            self.cfg.pasta_do_sistema = self.var_sistema.get()
            self.cfg.pasta_das_versoes = self.var_versoes.get()
            self.cfg.caminho_base_backup = self.var_backup.get()
            self.cfg.caminho_do_ini = self.var_ini.get()
            self.cfg.caminho_do_7zip = self.var_7zip.get()
            self.cfg.salvar()
            self.salvo = True
            logger.info("Configuração de caminhos salva via ConfigWindow")
            self.destroy()
        except Exception as e:
            logger.error("Erro ao salvar configuração: %s", e)
            messagebox.showerror("Erro", f"{e}", parent=self)
