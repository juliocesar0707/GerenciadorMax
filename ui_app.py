"""Aplicação principal do GerenciadorMax — UI e orquestração."""

import tkinter as tk
from tkinter import W, END, BOTH, YES, X, Y, LEFT, RIGHT, FLAT, BOTTOM
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import ttkbootstrap as bstrap
import os
import threading
import queue
import subprocess
import logging

from app_config import (
    AppConfig, CONFIG_FIELD_MAP, CONFIG_SECTIONS_UI,
    EXTENSOES_VERSAO, EXTENSOES_BACKUP, EXTENSOES_BACKUP_NUVEM,
)
from ini_service import IniService
from sql_service import SqlService
from webdav_client import WebDAVClient

logger = logging.getLogger(__name__)


class GerenciadorMaxApp(bstrap.Window):
    """Janela principal do GerenciadorMax."""

    def __init__(self, config):
        """Inicializa a aplicação.

        Args:
            config: Instância de AppConfig carregada.
        """
        super().__init__(themename="cyborg")
        self.title("Gerenciador Max (Black & Red)")
        self.geometry("900x780")
        self.withdraw()

        self.config = config
        self.sql = SqlService(config)
        self.msg_queue = queue.Queue()

        # WebDAV clients (um para versões, outro para backups)
        self.webdav_versoes = WebDAVClient(
            config.url_cloud, config.usuario_cloud, config.senha_cloud
        )
        self.webdav_backups = WebDAVClient(
            config.url_cloud, config.usuario_cloud, config.senha_cloud
        )

    def iniciar_interface(self):
        """Cria o layout e inicia o carregamento assíncrono."""
        self.create_layout()
        self.process_queue()
        self.deiconify()
        threading.Thread(target=self._carregamento_assincrono, daemon=True).start()

    def _carregamento_assincrono(self):
        """Carrega dados em background (SQL + arquivos locais)."""
        self.after(0, lambda: self.status.config(text="A carregar base de dados e ficheiros locais..."))
        self._popular_versoes_locais()
        self._popular_nuvem(self.webdav_versoes, self.lb_nuvem, self.lbl_caminho_nuvem, EXTENSOES_VERSAO)
        self._load_backups()
        self._carregar_banco_atual_sql()
        self.after(0, lambda: self.status.config(text="Pronto."))

    # =========================================================================
    # LAYOUT PRINCIPAL
    # =========================================================================
    def create_layout(self):
        """Cria toda a estrutura de abas e widgets."""
        main = bstrap.Frame(self, padding=10)
        main.pack(fill=BOTH, expand=YES)

        # Seletor de instância SQL
        top_frame = bstrap.Frame(main, padding=5)
        top_frame.pack(fill=X, pady=(0, 10))
        bstrap.Label(top_frame, text="Instância SQL Ativa:", font=("bold", 12),
                     bootstyle="secondary").pack(side=LEFT)
        self.combo_instancia = bstrap.Combobox(top_frame, state="readonly",
                                               bootstyle="danger", width=30)
        self.combo_instancia.pack(side=LEFT, padx=10)
        self.combo_instancia.bind("<<ComboboxSelected>>", self._on_instancia_changed)

        # Notebook principal
        self.notebook = bstrap.Notebook(main, bootstyle="danger")
        self.notebook.pack(fill=BOTH, expand=YES)

        tab_launch = bstrap.Frame(self.notebook, padding=15)
        tab_restore = bstrap.Frame(self.notebook, padding=15)
        tab_tools = bstrap.Frame(self.notebook, padding=15)
        tab_config = bstrap.Frame(self.notebook, padding=15)

        self.notebook.add(tab_launch, text="🚀 Launcher")
        self.notebook.add(tab_restore, text="📦 Restore")
        self.notebook.add(tab_tools, text="🛠️ Ferramentas")
        self.notebook.add(tab_config, text="⚙️ Configurações")

        self._setup_launcher(tab_launch)
        self._setup_restore(tab_restore)
        self._setup_tools(tab_tools)
        self._setup_config(tab_config)

        # Barra de status
        self.status = bstrap.Label(self, text="A inicializar...", relief=FLAT,
                                   anchor=W, padding=8, bootstyle="secondary")
        self.status.pack(side=BOTTOM, fill=X)

    # =========================================================================
    # ABA 1: LAUNCHER
    # =========================================================================
    def _setup_launcher(self, parent):
        """Configura a aba Launcher."""
        bstrap.Label(parent, text="Launcher Maxdata", font=("bold", 20),
                     bootstyle="danger").pack(pady=5)

        # --- Base de Dados ---
        db_frame = bstrap.Labelframe(parent, text=" Base de Dados ", bootstyle="danger", padding=15)
        db_frame.pack(fill=X, pady=10)

        f1 = bstrap.Frame(db_frame)
        f1.pack(fill=X, pady=5)
        bstrap.Label(f1, text="Banco Atual (INI):", bootstyle="secondary").pack(side=LEFT)
        self.lbl_db_atual = bstrap.Label(f1, text="A carregar...", font=("bold"),
                                         bootstyle="danger")
        self.lbl_db_atual.pack(side=LEFT, padx=5)
        bstrap.Label(f1, text="| Versão:", bootstyle="secondary").pack(side=LEFT, padx=(15, 5))
        self.lbl_versao_sql = bstrap.Label(f1, text="...", font=("bold"), bootstyle="warning")
        self.lbl_versao_sql.pack(side=LEFT)

        f2 = bstrap.Frame(db_frame)
        f2.pack(fill=X, pady=5)
        bstrap.Label(f2, text="Trocar para:", bootstyle="secondary").pack(side=LEFT)
        self.combo_db = bstrap.Combobox(f2, state="readonly", bootstyle="danger")
        self.combo_db.pack(side=LEFT, fill=X, expand=YES, padx=5)
        self.combo_db.bind("<<ComboboxSelected>>", self._preview_version)
        bstrap.Button(f2, text="Guardar", command=self._mudar_banco,
                      bootstyle="danger-outline").pack(side=LEFT)

        # --- Gestão de Versões ---
        v_frame = bstrap.Labelframe(parent, text=" Gestão de Versões ", bootstyle="danger", padding=15)
        v_frame.pack(fill=BOTH, expand=YES, pady=10)

        self.nb_versoes = bstrap.Notebook(v_frame, bootstyle="danger")
        self.nb_versoes.pack(fill=BOTH, expand=YES)

        tab_locais = bstrap.Frame(self.nb_versoes, padding=10)
        tab_nuvem = bstrap.Frame(self.nb_versoes, padding=10)
        self.nb_versoes.add(tab_locais, text="💻 Versões Locais")
        self.nb_versoes.add(tab_nuvem, text="☁️ Nuvem Maxdata")

        # Sub-tab: Locais
        f_busca = bstrap.Frame(tab_locais)
        f_busca.pack(fill=X, pady=(0, 5))
        bstrap.Label(f_busca, text="🔍 Buscar:").pack(side=LEFT)
        self.var_busca_versao = tk.StringVar()
        self.var_busca_versao.trace("w", self._filtrar_versoes)
        bstrap.Entry(f_busca, textvariable=self.var_busca_versao,
                     bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)

        self.lb_versoes = ttk.Treeview(tab_locais, height=6, bootstyle="danger",
                                       columns=("versao"), show="headings")
        self.lb_versoes.heading("versao", text="Arquivos (.rar) no PC")
        self.lb_versoes.column("versao", width=300, anchor=W)
        sb = bstrap.Scrollbar(tab_locais, command=self.lb_versoes.yview, bootstyle="danger-round")
        self.lb_versoes.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_versoes.pack(side=LEFT, fill=BOTH, expand=YES)

        # Sub-tab: Nuvem
        f_busca_n = bstrap.Frame(tab_nuvem)
        f_busca_n.pack(fill=X, pady=(0, 5))
        bstrap.Button(
            f_busca_n, text="🔄 Conectar/Recarregar",
            command=lambda: threading.Thread(
                target=lambda: self._popular_nuvem(
                    self.webdav_versoes, self.lb_nuvem, self.lbl_caminho_nuvem,
                    EXTENSOES_VERSAO, force_refresh=True
                ), daemon=True
            ).start(),
            bootstyle="danger-outline"
        ).pack(side=LEFT)
        bstrap.Button(
            f_busca_n, text="⬅️ Voltar Pasta",
            command=lambda: self._voltar_nuvem(
                self.webdav_versoes, self.lb_nuvem, self.lbl_caminho_nuvem, EXTENSOES_VERSAO
            ),
            bootstyle="secondary"
        ).pack(side=LEFT, padx=5)
        self.lbl_caminho_nuvem = bstrap.Label(f_busca_n, text="/", bootstyle="secondary")
        self.lbl_caminho_nuvem.pack(side=LEFT, padx=5)

        self.lb_nuvem = ttk.Treeview(tab_nuvem, height=6, bootstyle="danger",
                                      columns=("tipo", "nome"), show="headings")
        self.lb_nuvem.heading("tipo", text="Tipo")
        self.lb_nuvem.column("tipo", width=50, stretch=False, anchor=W)
        self.lb_nuvem.heading("nome", text="Nome (Duplo-clique para abrir pasta)")
        self.lb_nuvem.column("nome", width=250, anchor=W)
        self.lb_nuvem.bind("<Double-1>", lambda e: self._double_click_nuvem(
            e, self.webdav_versoes, self.lb_nuvem, self.lbl_caminho_nuvem, EXTENSOES_VERSAO
        ))
        sb_n = bstrap.Scrollbar(tab_nuvem, command=self.lb_nuvem.yview, bootstyle="danger-round")
        self.lb_nuvem.config(yscrollcommand=sb_n.set)
        sb_n.pack(side=RIGHT, fill=Y)
        self.lb_nuvem.pack(side=LEFT, fill=BOTH, expand=YES)

        bstrap.Button(
            tab_nuvem, text="⬇️ BAIXAR ARQUIVO SELECIONADO",
            command=lambda: self._baixar_nuvem(
                self.webdav_versoes, self.lb_nuvem, self.config.pasta_das_versoes,
                self._popular_versoes_locais
            ),
            bootstyle="danger"
        ).pack(side=BOTTOM, fill=X, pady=5)

        # Botões inferiores
        bts = bstrap.Frame(parent)
        bts.pack(fill=X)
        bstrap.Button(
            bts, text="🔄 Recarregar Locais",
            command=lambda: threading.Thread(target=self._popular_versoes_locais, daemon=True).start(),
            bootstyle="secondary-outline"
        ).pack(side=LEFT)
        bstrap.Button(bts, text="▶️ EXECUTAR SISTEMA", command=self._lancar_erp,
                      bootstyle="success-outline").pack(side=RIGHT, padx=5)
        bstrap.Button(bts, text="⚡ ATUALIZAR VERSÃO (Extrair)", command=self._lancar_atualizacao,
                      bootstyle="danger").pack(side=RIGHT)

    # =========================================================================
    # ABA 2: RESTORE
    # =========================================================================
    def _setup_restore(self, parent):
        """Configura a aba Restore."""
        f1 = bstrap.Labelframe(parent, text=" 1. Selecione o Backup ", bootstyle="danger", padding=10)
        f1.pack(fill=BOTH, expand=YES)

        self.nb_backups = bstrap.Notebook(f1, bootstyle="danger")
        self.nb_backups.pack(fill=BOTH, expand=YES)

        tab_locais_bkp = bstrap.Frame(self.nb_backups, padding=10)
        tab_nuvem_bkp = bstrap.Frame(self.nb_backups, padding=10)
        self.nb_backups.add(tab_locais_bkp, text="💻 Backups Locais")
        self.nb_backups.add(tab_nuvem_bkp, text="☁️ Nuvem Maxdata")

        # Sub-tab: Locais
        f_busca = bstrap.Frame(tab_locais_bkp)
        f_busca.pack(fill=X, pady=(0, 5))
        bstrap.Label(f_busca, text="🔍 Buscar:").pack(side=LEFT)
        self.var_busca_backup = tk.StringVar()
        self.var_busca_backup.trace("w", self._filtrar_backups)
        bstrap.Entry(f_busca, textvariable=self.var_busca_backup,
                     bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)

        self.lb_backups = ttk.Treeview(tab_locais_bkp, height=6, bootstyle="danger",
                                       columns=("backup"), show="headings")
        self.lb_backups.heading("backup", text="Backups Disponíveis")
        self.lb_backups.column("backup", width=300, anchor=tk.W)
        sb = bstrap.Scrollbar(tab_locais_bkp, command=self.lb_backups.yview, bootstyle="danger-round")
        self.lb_backups.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_backups.pack(side=LEFT, fill=BOTH, expand=YES)
        bstrap.Button(
            tab_locais_bkp, text="🔄 Atualizar Locais",
            command=lambda: threading.Thread(target=self._load_backups, daemon=True).start(),
            bootstyle="secondary-outline"
        ).pack(fill=X, pady=5)

        # Sub-tab: Nuvem Backups
        f_busca_nb = bstrap.Frame(tab_nuvem_bkp)
        f_busca_nb.pack(fill=X, pady=(0, 5))
        bstrap.Button(
            f_busca_nb, text="🔄 Conectar/Recarregar",
            command=lambda: threading.Thread(
                target=lambda: self._popular_nuvem(
                    self.webdav_backups, self.lb_nuvem_bkp, self.lbl_caminho_nuvem_bkp,
                    EXTENSOES_BACKUP_NUVEM, force_refresh=True
                ), daemon=True
            ).start(),
            bootstyle="danger-outline"
        ).pack(side=LEFT)
        bstrap.Button(
            f_busca_nb, text="⬅️ Voltar Pasta",
            command=lambda: self._voltar_nuvem(
                self.webdav_backups, self.lb_nuvem_bkp, self.lbl_caminho_nuvem_bkp, EXTENSOES_BACKUP_NUVEM
            ),
            bootstyle="secondary"
        ).pack(side=LEFT, padx=5)
        self.lbl_caminho_nuvem_bkp = bstrap.Label(f_busca_nb, text="/", bootstyle="secondary")
        self.lbl_caminho_nuvem_bkp.pack(side=LEFT, padx=5)

        self.lb_nuvem_bkp = ttk.Treeview(tab_nuvem_bkp, height=6, bootstyle="danger",
                                          columns=("tipo", "nome"), show="headings")
        self.lb_nuvem_bkp.heading("tipo", text="Tipo")
        self.lb_nuvem_bkp.column("tipo", width=50, stretch=False, anchor=tk.W)
        self.lb_nuvem_bkp.heading("nome", text="Nome (Duplo-clique para abrir pasta)")
        self.lb_nuvem_bkp.column("nome", width=250, anchor=tk.W)
        self.lb_nuvem_bkp.bind("<Double-1>", lambda e: self._double_click_nuvem(
            e, self.webdav_backups, self.lb_nuvem_bkp, self.lbl_caminho_nuvem_bkp, EXTENSOES_BACKUP_NUVEM
        ))
        sb_nb = bstrap.Scrollbar(tab_nuvem_bkp, command=self.lb_nuvem_bkp.yview, bootstyle="danger-round")
        self.lb_nuvem_bkp.config(yscrollcommand=sb_nb.set)
        sb_nb.pack(side=RIGHT, fill=Y)
        self.lb_nuvem_bkp.pack(side=LEFT, fill=BOTH, expand=YES)

        bstrap.Button(
            tab_nuvem_bkp, text="⬇️ BAIXAR ARQUIVO SELECIONADO",
            command=lambda: self._baixar_nuvem(
                self.webdav_backups, self.lb_nuvem_bkp, self.config.caminho_base_backup,
                self._load_backups
            ),
            bootstyle="danger"
        ).pack(side=BOTTOM, fill=X, pady=5)

        # Nome do novo banco
        f2 = bstrap.Labelframe(parent, text=" 2. Nome do Novo Banco ", bootstyle="danger", padding=10)
        f2.pack(fill=X, pady=10)
        self.entry_new_db = bstrap.Entry(f2, bootstyle="danger")
        self.entry_new_db.pack(fill=X)

        # Botão de restore + progress + log
        f3 = bstrap.Frame(parent)
        f3.pack(fill=BOTH, expand=YES)
        self.btn_restore = bstrap.Button(f3, text="▶️ INICIAR RESTAURAÇÃO",
                                         command=self._iniciar_restore, bootstyle="danger")
        self.btn_restore.pack(fill=X, pady=5)
        self.progress = bstrap.Progressbar(f3, mode='indeterminate', bootstyle="danger-striped")
        self.progress.pack(fill=X, pady=5)
        self.log_txt = scrolledtext.ScrolledText(f3, height=8, state='disabled',
                                                  font=("Consolas", 9), bg="#111", fg="#0f0")
        self.log_txt.pack(fill=BOTH, expand=YES)

    # =========================================================================
    # ABA 3: FERRAMENTAS
    # =========================================================================
    def _setup_tools(self, parent):
        """Configura a aba Ferramentas."""
        bstrap.Label(parent, text="Gestão de Bases de Dados", font=("bold", 14),
                     bootstyle="danger").pack(pady=10)

        flist = bstrap.Labelframe(parent,
                                  text=" Bases SQL (Use Shift/Ctrl para selecionar várias) ",
                                  bootstyle="danger", padding=10)
        flist.pack(fill=BOTH, expand=YES)

        f_busca = bstrap.Frame(flist)
        f_busca.pack(fill=X, pady=(0, 5))
        bstrap.Label(f_busca, text="🔍 Buscar:").pack(side=LEFT)
        self.var_busca_banco = tk.StringVar()
        self.var_busca_banco.trace("w", self._filtrar_bancos)
        bstrap.Entry(f_busca, textvariable=self.var_busca_banco,
                     bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)

        self.lb_tools = ttk.Treeview(flist, height=10, bootstyle="danger",
                                     columns=("banco"), show="headings", selectmode="extended")
        self.lb_tools.heading("banco", text="Bases de Dados")
        self.lb_tools.column("banco", width=300, anchor=tk.W)
        sb = bstrap.Scrollbar(flist, command=self.lb_tools.yview, bootstyle="danger-round")
        self.lb_tools.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_tools.pack(side=LEFT, fill=BOTH, expand=YES)

        fbtn = bstrap.Frame(flist)
        fbtn.pack(side=RIGHT, fill=Y, padx=10)
        bstrap.Button(
            fbtn, text="🔄 Atualizar",
            command=lambda: threading.Thread(target=self._carregar_banco_atual_sql, daemon=True).start(),
            bootstyle="secondary-outline"
        ).pack(fill=X, pady=5)
        bstrap.Button(fbtn, text="🗑️ ELIMINAR (DROP)", command=self._drop_database,
                      bootstyle="danger").pack(fill=X, pady=20)

    # =========================================================================
    # ABA 4: CONFIGURAÇÕES
    # =========================================================================
    def _setup_config(self, parent):
        """Configura a aba de Configurações."""
        scroll_canvas = bstrap.Canvas(parent, highlightthickness=0)
        scrollbar = bstrap.Scrollbar(parent, orient="vertical", command=scroll_canvas.yview)
        scroll_frame = bstrap.Frame(scroll_canvas)
        scroll_frame.bind("<Configure>",
                          lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))
        scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_canvas.configure(yscrollcommand=scrollbar.set)
        scroll_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

        bstrap.Label(scroll_frame, text="Configurações do Sistema", font=("bold", 14),
                     bootstyle="danger").pack(pady=10)

        self.cfg_vars = {}

        for title, keys, ini_section in CONFIG_SECTIONS_UI:
            f = bstrap.Labelframe(scroll_frame, text=f" {title} ", bootstyle="danger", padding=10)
            f.pack(fill=X, pady=5, padx=5)
            for key in keys:
                row = bstrap.Frame(f)
                row.pack(fill=X, pady=2)
                bstrap.Label(row, text=key, width=25, bootstyle="secondary").pack(side=LEFT)
                valor = self.config.get_campo(ini_section, key)
                var = tk.StringVar(value=valor)
                self.cfg_vars[(ini_section, key)] = var
                bstrap.Entry(row, textvariable=var, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES)

        bstrap.Button(scroll_frame, text="💾 Guardar Configurações",
                      command=self._salvar_config_aba, bootstyle="success").pack(pady=20)

    def _salvar_config_aba(self):
        """Salva todas as configurações editadas na aba."""
        try:
            for (section, key), var in self.cfg_vars.items():
                self.config.set_campo(section, key, var.get())
            self.config.salvar()
            self.config.carregar()
            self.config.validar_caminhos()

            # Recriar serviços com novas configurações
            self.sql = SqlService(self.config)
            self.webdav_versoes = WebDAVClient(
                self.config.url_cloud, self.config.usuario_cloud, self.config.senha_cloud
            )
            self.webdav_backups = WebDAVClient(
                self.config.url_cloud, self.config.usuario_cloud, self.config.senha_cloud
            )

            messagebox.showinfo("Sucesso", "Configurações guardadas!")
            logger.info("Configurações da aba salvas com sucesso")
            threading.Thread(target=self._carregamento_assincrono, daemon=True).start()
        except Exception as e:
            logger.error("Erro ao guardar configurações: %s", e)
            messagebox.showerror("Erro", f"Erro ao guardar: {e}")

    # =========================================================================
    # LÓGICA: FILTROS DE BUSCA
    # =========================================================================
    def _filtrar_versoes(self, *args):
        """Filtra a lista de versões locais pelo termo de busca."""
        busca = self.var_busca_versao.get().lower()
        for i in self.lb_versoes.get_children():
            self.lb_versoes.delete(i)
        for v in getattr(self, '_all_versoes', []):
            if busca in v.lower():
                self.lb_versoes.insert("", END, values=(v,))

    def _filtrar_backups(self, *args):
        """Filtra a lista de backups locais pelo termo de busca."""
        busca = self.var_busca_backup.get().lower()
        for i in self.lb_backups.get_children():
            self.lb_backups.delete(i)
        for b in getattr(self, '_all_backups', []):
            if busca in b.lower():
                self.lb_backups.insert("", tk.END, values=(b,))

    def _filtrar_bancos(self, *args):
        """Filtra a lista de bancos SQL pelo termo de busca."""
        busca = self.var_busca_banco.get().lower()
        for i in self.lb_tools.get_children():
            self.lb_tools.delete(i)
        for d in getattr(self, '_all_dbs', []):
            if busca in d.lower():
                self.lb_tools.insert("", tk.END, values=(d,))

    # =========================================================================
    # LÓGICA: SQL + INI
    # =========================================================================
    def _atualizar_ui_sql(self, bancos, db_atual, versao):
        """Atualiza os widgets de SQL com dados novos."""
        self.lbl_db_atual.config(text=db_atual)
        self.lbl_versao_sql.config(text=versao)
        self.combo_db['values'] = bancos
        if db_atual in bancos:
            self.combo_db.set(db_atual)
        self._all_dbs = bancos
        self._filtrar_bancos()

    def _carregar_banco_atual_sql(self):
        """Lê o banco atual do max.ini e atualiza a UI."""
        atual, server_atual = IniService.ler_banco_e_servidor(
            self.config.caminho_do_ini,
            self.config.ini_section,
            self.config.ini_key,
            self.config.ini_server_key
        )

        if server_atual:
            self.config.sql_server_instance = server_atual
            self.config.servidor = server_atual
            if hasattr(self, 'combo_instancia'):
                vals = (list(self.combo_instancia.cget('values'))
                        if self.combo_instancia.cget('values')
                        else self.sql.listar_instancias())
                if server_atual not in vals:
                    vals.append(server_atual)
                self.after(0, lambda v=vals: self.combo_instancia.config(values=v))
                self.after(0, lambda s=server_atual: self.combo_instancia.set(s))
        else:
            if hasattr(self, 'combo_instancia') and not self.combo_instancia.get():
                vals = self.sql.listar_instancias()
                self.after(0, lambda v=vals: self.combo_instancia.config(values=v))
                if vals:
                    self.after(0, lambda: self.combo_instancia.set(vals[0]))

        versao = self.sql.get_versao(atual)
        bancos = self.sql.listar_bancos()
        self.after(0, lambda: self._atualizar_ui_sql(bancos, atual, versao))

    def _preview_version(self, event):
        """Mostra a versão do banco selecionado no combo."""
        db = self.combo_db.get()
        threading.Thread(
            target=lambda: self.after(
                0, lambda v=self.sql.get_versao(db): self.status.config(
                    text=f"Banco selecionado: {db} (Versão: {v})"
                )
            ), daemon=True
        ).start()

    def _on_instancia_changed(self, event=None):
        """Handler para mudança de instância SQL."""
        instancia = self.combo_instancia.get()
        self.config.sql_server_instance = instancia
        self.config.servidor = instancia

        try:
            c = IniService.ler_arquivo(self.config.caminho_do_ini)
            IniService.set_value(c, self.config.ini_section or 'CON', 
                                 self.config.ini_server_key or 'Data Source', instancia)
            IniService.salvar(c, self.config.caminho_do_ini)
        except Exception as e:
            logger.warning("Erro ao salvar instância no INI: %s", e)

        self.after(0, lambda: self.status.config(
            text=f"Instância alterada para: {instancia}. A recarregar bancos..."))
        threading.Thread(target=self._carregar_banco_atual_sql, daemon=True).start()

    def _mudar_banco(self):
        """Troca o banco de dados no max.ini."""
        novo = self.combo_db.get()
        if not novo:
            return
        try:
            c = IniService.ler_arquivo(self.config.caminho_do_ini)
            IniService.set_value(c, self.config.ini_section or 'CON', 
                                 self.config.ini_key or 'Initial catalog', novo)
            if hasattr(self, 'combo_instancia'):
                inst = self.combo_instancia.get()
                if inst:
                    IniService.set_value(c, self.config.ini_section or 'CON', 
                                         self.config.ini_server_key or 'Data Source', inst)
            IniService.salvar(c, self.config.caminho_do_ini)
            messagebox.showinfo("Sucesso", f"Alterado para: {novo}")
            threading.Thread(target=self._carregar_banco_atual_sql, daemon=True).start()
        except Exception as e:
            logger.error("Erro ao mudar banco: %s", e)
            messagebox.showerror("Erro", f"{e}")

    # =========================================================================
    # LÓGICA: VERSÕES LOCAIS
    # =========================================================================
    def _popular_versoes_locais(self):
        """Carrega a lista de versões locais (.rar)."""
        try:
            if os.path.exists(self.config.pasta_das_versoes):
                versoes = [
                    e.name for e in os.scandir(self.config.pasta_das_versoes)
                    if e.is_file() and e.name.lower().endswith(EXTENSOES_VERSAO)
                ]
                self._all_versoes = sorted(versoes, reverse=True)
                self.after(0, self._filtrar_versoes)
        except OSError as e:
            logger.warning("Erro ao listar versões locais: %s", e)

    # =========================================================================
    # LÓGICA: WEBDAV GENÉRICO (reutilizado para versões e backups)
    # =========================================================================
    def _popular_nuvem(self, client, treeview, lbl_caminho, extensoes, force_refresh=False):
        """Popula um treeview com conteúdo WebDAV.

        Args:
            client: Instância de WebDAVClient.
            treeview: Widget Treeview a popular.
            lbl_caminho: Label que mostra o caminho atual.
            extensoes: Tupla de extensões de arquivo a filtrar.
            force_refresh: Ignora cache.
        """
        self.after(0, lambda: lbl_caminho.config(text=client.caminho_atual))

        try:
            pastas, arquivos = client.listar(force_refresh=force_refresh, extensoes=extensoes)
            is_cache = not force_refresh and len(pastas) + len(arquivos) > 0

            def att_ui():
                for i in treeview.get_children():
                    treeview.delete(i)
                for p in pastas:
                    treeview.insert("", END, values=("📁", p))
                for a in arquivos:
                    treeview.insert("", END, values=("📄", a))
                self.status.config(
                    text=f"{'(Cache) ' if is_cache else 'Pronto. '}"
                         f"{len(pastas)} pastas, {len(arquivos)} arquivos."
                )

            self.after(0, att_ui)
        except Exception as e:
            logger.warning("Erro WebDAV: %s", e)
            self.after(0, lambda e=e: self.status.config(text=f"Erro WebDAV: {str(e)[:50]}"))

    def _voltar_nuvem(self, client, treeview, lbl_caminho, extensoes):
        """Volta uma pasta no navegador WebDAV."""
        if client.voltar():
            threading.Thread(
                target=lambda: self._popular_nuvem(client, treeview, lbl_caminho, extensoes),
                daemon=True
            ).start()

    def _double_click_nuvem(self, event, client, treeview, lbl_caminho, extensoes):
        """Handler de duplo-clique: navega para subpasta."""
        sel = treeview.selection()
        if not sel:
            return
        item = treeview.item(sel[0], "values")
        tipo, nome = item[0], item[1]

        if tipo == "📁":
            client.navegar(nome)
            threading.Thread(
                target=lambda: self._popular_nuvem(client, treeview, lbl_caminho, extensoes),
                daemon=True
            ).start()

    def _baixar_nuvem(self, client, treeview, destino_dir, on_complete_callback):
        """Baixa arquivo selecionado do WebDAV.

        Args:
            client: Instância de WebDAVClient.
            treeview: Treeview onde o arquivo está selecionado.
            destino_dir: Diretório de destino local.
            on_complete_callback: Função chamada após download bem-sucedido.
        """
        sel = treeview.selection()
        if not sel:
            return
        item = treeview.item(sel[0], "values")
        tipo, nome = item[0], item[1]

        if tipo == "📁":
            messagebox.showinfo("Aviso", "Selecione um ARQUIVO (📄) para baixar, e não uma pasta.")
            return

        if messagebox.askyesno("Baixar da Nuvem", f"Deseja baixar {nome} para o computador?"):
            self.status.config(text=f"A iniciar download de {nome}...")
            threading.Thread(
                target=self._thread_download,
                args=(client, nome, destino_dir, on_complete_callback),
                daemon=True
            ).start()

    def _thread_download(self, client, nome, destino_dir, on_complete_callback):
        """Thread de download WebDAV."""
        try:
            def on_progress(pct):
                self.after(0, lambda p=pct: self.status.config(text=f"A baixar {nome}... {p}%"))

            client.download(nome, destino_dir, on_progress=on_progress)

            self.after(0, lambda: self.status.config(text=f"Download de {nome} concluído!"))
            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Download concluído!\n{nome}"))
            if on_complete_callback:
                self.after(0, on_complete_callback)
        except Exception as e:
            logger.error("Erro no download: %s", e)
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Download", f"Falha:\n{m}"))
            self.after(0, lambda: self.status.config(text="Erro no download."))

    # =========================================================================
    # LÓGICA: BACKUPS LOCAIS
    # =========================================================================
    def _load_backups(self):
        """Carrega lista de backups locais."""
        try:
            backups = []
            if os.path.exists(self.config.caminho_base_backup):
                for e in os.scandir(self.config.caminho_base_backup):
                    if e.is_file() and e.name.upper().endswith(
                            tuple(ext.upper() for ext in EXTENSOES_BACKUP)):
                        backups.append((e.name, e.stat().st_mtime))
                backups.sort(key=lambda x: x[1], reverse=True)
                self._all_backups = [b[0] for b in backups]
                self.after(0, self._filtrar_backups)
        except OSError as e:
            logger.warning("Erro ao listar backups: %s", e)

    # =========================================================================
    # LÓGICA: LAUNCHER (EXE + ATUALIZAÇÃO)
    # =========================================================================
    def _get_exe_versao(self, caminho_exe):
        """Obtém a versão de um executável Windows."""
        try:
            import win32api
            info = win32api.GetFileVersionInfo(caminho_exe, "\\")
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            return (f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}"
                    f".{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}")
        except Exception as e:
            logger.debug("Não foi possível obter versão do EXE: %s", e)
            return None

    def _lancar_erp(self):
        """Lança o sistema ERP, verificando compatibilidade de versão."""
        try:
            db_atual = self.lbl_db_atual.cget("text")
            db_versao = self.lbl_versao_sql.cget("text")
            exe_versao = self._get_exe_versao(self.config.caminho_do_erp_cliente)

            if (db_versao and exe_versao
                    and db_versao not in ("---", "N/A")
                    and exe_versao != db_versao):
                arquivo_rar = next(
                    (v for v in getattr(self, '_all_versoes', []) if db_versao in v), None
                )
                if arquivo_rar:
                    if messagebox.askyesno(
                        "Atualização Necessária",
                        f"Versão BD: {db_versao} | Versão EXE: {exe_versao}.\n\n"
                        f"Atualizar para '{arquivo_rar}'?"
                    ):
                        threading.Thread(target=self._thread_extrair, args=(arquivo_rar,),
                                         daemon=True).start()
                        return
                else:
                    messagebox.showwarning(
                        "Aviso de Versão",
                        "Versões diferentes e ficheiro de extração não encontrado.\nA ignorar..."
                    )

            subprocess.Popen([self.config.caminho_do_erp_cliente], cwd=self.config.pasta_do_sistema)
        except Exception as e:
            logger.error("Erro ao lançar ERP: %s", e)
            messagebox.showerror("Erro", f"Erro: {e}")

    def _lancar_atualizacao(self):
        """Lança extração da versão selecionada."""
        sel = self.lb_versoes.selection()
        if not sel:
            return
        arq = self.lb_versoes.item(sel[0], "values")[0]
        if messagebox.askyesno("Confirmar", f"Atualizar para {arq}?"):
            threading.Thread(target=self._thread_extrair, args=(arq,), daemon=True).start()

    def _thread_extrair(self, arq):
        """Thread de extração de versão com 7-Zip."""
        try:
            self.after(0, lambda: self.status.config(text="A extrair..."))
            cmd = [
                self.config.caminho_do_7zip, 'x',
                os.path.join(self.config.pasta_das_versoes, arq),
                f'-o{self.config.pasta_do_sistema}', '-y'
            ]
            subprocess.run(cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.after(0, self._lancar_atualizador_callback)
        except Exception as e:
            logger.error("Erro na extração: %s", e)
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Extração", m))

    def _lancar_atualizador_callback(self):
        """Lança o MAX_Atualiza.exe após extração."""
        try:
            subprocess.Popen([self.config.caminho_do_max_atualiza], cwd=self.config.pasta_do_sistema)
            self.status.config(text="Atualizador Aberto.")
        except Exception as e:
            logger.error("Erro ao abrir atualizador: %s", e)
            messagebox.showerror("Erro", f"{e}")

    # =========================================================================
    # LÓGICA: RESTORE
    # =========================================================================
    def _iniciar_restore(self):
        """Inicia o processo de restore em thread separada."""
        try:
            sel = self.lb_backups.selection()[0]
            fname = self.lb_backups.item(sel, "values")[0]
        except (IndexError, KeyError):
            return
        new_db = self.entry_new_db.get().strip()
        if not new_db:
            return

        self.btn_restore.config(state='disabled')
        self.progress.start()
        self.log_txt.config(state='normal')
        self.log_txt.delete('1.0', END)
        self.log_txt.config(state='disabled')
        threading.Thread(target=self._restore_logic, args=(fname, new_db), daemon=True).start()

    def _restore_logic(self, fname, dbname):
        """Orquestra o restore usando SqlService."""
        try:
            self.sql.executar_restore_completo(
                fname=fname,
                dbname=dbname,
                caminho_backup=self.config.caminho_base_backup,
                pasta_sistema=self.config.pasta_do_sistema,
                caminho_7zip=self.config.caminho_do_7zip,
                on_message=lambda msg: self.msg_queue.put(msg)
            )
            self.msg_queue.put("__DONE__")
        except Exception as e:
            logger.error("Erro no restore: %s", e)
            self.msg_queue.put(f"ERRO: {e}")
            self.msg_queue.put("__ERROR__")

    def process_queue(self):
        """Processa mensagens da fila de threads (polling a cada 100ms)."""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg == "__DONE__":
                    self.progress.stop()
                    self.btn_restore.config(state='normal')
                    messagebox.showinfo("Sucesso", "Restore Concluído!")
                    threading.Thread(target=self._carregar_banco_atual_sql, daemon=True).start()
                elif msg == "__ERROR__":
                    self.progress.stop()
                    self.btn_restore.config(state='normal')
                    messagebox.showerror("Erro", "Falhou. Veja o log.")
                else:
                    self.log_txt.config(state='normal')
                    self.log_txt.insert(END, msg + "\n")
                    self.log_txt.see(END)
                    self.log_txt.config(state='disabled')
        except queue.Empty:
            pass
        self.after(100, self.process_queue)

    # =========================================================================
    # LÓGICA: DROP DATABASE
    # =========================================================================
    def _drop_database(self):
        """Elimina bancos selecionados (com confirmação por escrito)."""
        sel = self.lb_tools.selection()
        if not sel:
            return

        bancos_alvo = [self.lb_tools.item(i, "values")[0] for i in sel]
        db_atual = self.lbl_db_atual.cget("text")

        if db_atual in bancos_alvo:
            bancos_alvo.remove(db_atual)
            messagebox.showwarning(
                "Proteção",
                f"O banco atual ({db_atual}) não pode ser eliminado. "
                f"Foi removido da sua lista de exclusão."
            )

        if not bancos_alvo:
            return

        if len(bancos_alvo) == 1:
            msg = f"ELIMINAR '{bancos_alvo[0]}' permanentemente?\n\nEscreva EXCLUIR para confirmar:"
        else:
            lista_preview = ", ".join(bancos_alvo[:5]) + ("..." if len(bancos_alvo) > 5 else "")
            msg = (f"ELIMINAR {len(bancos_alvo)} bases de dados permanentemente?\n\n"
                   f"({lista_preview})\n\nEscreva EXCLUIR para confirmar:")

        confirmacao = simpledialog.askstring("PERIGO", msg)

        if confirmacao == "EXCLUIR":
            threading.Thread(target=self._thread_drop, args=(bancos_alvo,), daemon=True).start()
        elif confirmacao is not None:
            messagebox.showwarning("Cancelado", "Palavra de confirmação inválida.")

    def _thread_drop(self, bancos):
        """Thread de exclusão de bancos."""
        try:
            self.after(0, lambda: self.status.config(text=f"A eliminar {len(bancos)} banco(s)..."))
            self.sql.drop_databases(bancos)
            self.after(0, lambda: messagebox.showinfo("Sucesso", "Bases eliminadas com sucesso!"))
            self._carregar_banco_atual_sql()
            self.after(0, lambda: self.status.config(text="Pronto."))
        except Exception as e:
            logger.error("Erro ao eliminar bancos: %s", e)
            self.after(0, lambda err=str(e): messagebox.showerror("Erro", err))
