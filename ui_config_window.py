"""Janela de configuração de caminhos (setup inicial e edição)."""

import tkinter as tk
from tkinter import BOTH, YES, X, LEFT, RIGHT
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as bstrap
import logging

from app_config import CONFIG_FILE_NAME

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
        self.config = config
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

        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)

        if is_first_run:
            bstrap.Label(main_frame, text="Primeira execução! Confirme os diretórios de trabalho:",
                         font=("bold", 12), bootstyle="success").pack(pady=(0, 15))
        else:
            bstrap.Label(main_frame, text="Corrija os caminhos:",
                         font=("bold", 12), bootstyle="danger").pack(pady=(0, 15))

        self._criar_campo(main_frame, "Pasta do Sistema", self.var_sistema, True)
        self._criar_campo(main_frame, "Pasta de Versões", self.var_versoes, True)
        self._criar_campo(main_frame, "Pasta de Backups", self.var_backup, True)
        self._criar_campo(main_frame, "Arquivo .ini", self.var_ini, False)
        self._criar_campo(main_frame, "Executável 7-Zip", self.var_7zip, False)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, pady=(20, 0))
        bstrap.Button(btn_frame, text="Salvar e Continuar",
                      command=self._salvar, bootstyle="success").pack(side=RIGHT, padx=5)
        if not is_first_run:
            bstrap.Button(btn_frame, text="Cancelar",
                          command=self.destroy, bootstyle="secondary").pack(side=RIGHT)

    def _criar_campo(self, parent, label, var, is_dir):
        """Cria uma linha de campo com label, entry e botão de seleção."""
        f = ttk.Frame(parent)
        f.pack(fill=X, pady=5)
        bstrap.Label(f, text=label, width=20, bootstyle="inverse-secondary").pack(side=LEFT)
        bstrap.Entry(f, textvariable=var, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)
        cmd = lambda: self._selecionar(var, is_dir)
        bstrap.Button(f, text="...", command=cmd, bootstyle="info-outline", width=4).pack(side=LEFT)

    @staticmethod
    def _selecionar(var, is_dir):
        """Abre diálogo de seleção de pasta ou arquivo."""
        p = filedialog.askdirectory() if is_dir else filedialog.askopenfilename()
        if p:
            var.set(p)

    def _salvar(self):
        """Salva os caminhos editados na configuração."""
        try:
            self.config.pasta_do_sistema = self.var_sistema.get()
            self.config.pasta_das_versoes = self.var_versoes.get()
            self.config.caminho_base_backup = self.var_backup.get()
            self.config.caminho_do_ini = self.var_ini.get()
            self.config.caminho_do_7zip = self.var_7zip.get()
            self.config.salvar()
            self.salvo = True
            logger.info("Configuração de caminhos salva via ConfigWindow")
            self.destroy()
        except Exception as e:
            logger.error("Erro ao salvar configuração: %s", e)
            messagebox.showerror("Erro", f"{e}", parent=self)
