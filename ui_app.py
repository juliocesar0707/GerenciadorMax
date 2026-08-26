"""Aplicação principal do GerenciadorMax — UI e orquestração."""

import tkinter as tk
from tkinter import W, E, END, BOTH, YES, NO, X, Y, LEFT, RIGHT, FLAT, BOTTOM
from tkinter import ttk, messagebox, simpledialog
import ttkbootstrap as bstrap
import os
import threading
import queue
import subprocess
import logging

from app_config import (
    CONFIG_SECTIONS_UI,
    EXTENSOES_VERSAO, EXTENSOES_BACKUP, EXTENSOES_BACKUP_NUVEM,
)
from ini_service import IniService
from sql_service import SqlService
from webdav_client import WebDAVClient
import sevenzip
import ui_theme
from ui_widgets import RoundedButton

logger = logging.getLogger(__name__)

CLOUD_PANEL_WIDTH = 330


class GerenciadorMaxApp(bstrap.Window):
    """Janela principal do GerenciadorMax."""

    def __init__(self, config):
        """Inicializa a aplicação.

        Args:
            config: Instância de AppConfig carregada.
        """
        ui_theme.registrar_tema()
        super().__init__(themename=ui_theme.THEME_NAME)
        ui_theme.aplicar_estilos(self.style)

        self.title("Gerenciador Max")
        self.geometry("1180x800")
        self.minsize(1020, 700)
        self.withdraw()

        # NÃO usar self.config: sobrescreveria tk.Misc.config() da janela.
        self.cfg = config
        self.sql = SqlService(config)
        self.msg_queue = queue.Queue()

        self._all_versoes = []
        self._all_backups = []
        self._all_dbs = []
        self._cloud_visivel = False
        self._nuvem_carregada = set()   # id(treeview) das abas já listadas

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
        self._set_status("A carregar base de dados e ficheiros locais...")
        self._popular_versoes_locais()
        self._load_backups()
        self._carregar_banco_atual_sql()
        self._set_status("Pronto.")

    def _set_status(self, texto):
        """Atualiza a barra de status a partir de qualquer thread."""
        self.after(0, lambda t=texto: self.status.config(text=t))

    # =========================================================================
    # LAYOUT PRINCIPAL
    # =========================================================================
    def create_layout(self):
        """Monta a janela: rodapé, painel da nuvem retrátil e as 3 colunas."""
        self._criar_rodape()
        self._criar_aba_nuvem()
        self._criar_painel_nuvem()

        self.main_container = bstrap.Frame(self, padding=(20, 16, 12, 10))
        self.main_container.pack(side=LEFT, fill=BOTH, expand=YES)

        colunas = bstrap.Frame(self.main_container)
        colunas.pack(fill=BOTH, expand=YES)

        # grid com `uniform` mantém as 3 colunas com a mesma largura,
        # independente do tamanho dos nomes de arquivo listados.
        colunas.rowconfigure(0, weight=1)
        for i in range(3):
            colunas.columnconfigure(i, weight=1, uniform="col")

        self.col_left = bstrap.Frame(colunas)
        self.col_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.col_center = bstrap.Frame(colunas)
        self.col_center.grid(row=0, column=1, sticky="nsew", padx=8)

        self.col_right = bstrap.Frame(colunas)
        self.col_right.grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        self._setup_col_left()
        self._setup_col_center()
        self._setup_col_right()

        # Precisa rodar depois: o ttkbootstrap só constrói os estilos de
        # scrollbar/combobox quando o primeiro widget que os usa é criado.
        ui_theme.ajustar_estilos_derivados(self.style)

    # --- Helpers de construção -------------------------------------------
    def _titulo_coluna(self, parent, texto):
        """Título grande e centralizado no topo de uma coluna."""
        ttk.Label(parent, text=texto, style="ColTitle.TLabel",
                  anchor="center").pack(fill=X, pady=(0, 14))

    def _card(self, parent, **kwargs):
        """Cartão branco padrão."""
        kwargs.setdefault("padding", 14)
        return ttk.Frame(parent, style="Card.TFrame", **kwargs)

    @staticmethod
    def _botao(parent, texto, comando, variant="primary", **pack_kwargs):
        """Botão de cantos arredondados, no estilo do login do MaxManager."""
        btn = RoundedButton(parent, text=texto, command=comando, variant=variant)
        if pack_kwargs:
            btn.pack(**pack_kwargs)
        return btn

    def _rotulo(self, parent, texto, estilo="CardMuted.TLabel", **pack_kwargs):
        """Rótulo secundário (cinza) usado acima dos campos."""
        lbl = ttk.Label(parent, text=texto, style=estilo)
        pack_kwargs.setdefault("anchor", W)
        lbl.pack(**pack_kwargs)
        return lbl

    def _entrada_clara(self, parent, textvariable=None):
        """Campo de texto branco com borda suave, como no login do MaxManager."""
        return ttk.Entry(parent, textvariable=textvariable, style="Campo.TEntry")

    @staticmethod
    def _scrollbar(parent, widget):
        """Scrollbar discreta acoplada a um treeview ou text."""
        sb = bstrap.Scrollbar(parent, command=widget.yview,
                              bootstyle="secondary-round")
        widget.config(yscrollcommand=sb.set)
        return sb

    def _tabela(self, parent, coluna_id, cabecalho, height=12, selectmode="browse"):
        """Cria um Treeview escuro de coluna única com scrollbar acoplada.

        Returns:
            ttk.Treeview já empacotado dentro de `parent`.
        """
        wrapper = ttk.Frame(parent, style="Card.TFrame")
        wrapper.pack(fill=BOTH, expand=YES)

        tree = ttk.Treeview(
            wrapper, height=height, columns=(coluna_id,), show="headings",
            style="Claro.Treeview", selectmode=selectmode,
        )
        tree.heading(coluna_id, text=cabecalho, anchor=W)
        tree.column(coluna_id, anchor=W, width=120, stretch=True)

        self._scrollbar(wrapper, tree).pack(side=RIGHT, fill=Y)
        tree.pack(side=LEFT, fill=BOTH, expand=YES)
        return tree

    # --- Coluna 1: Manager -----------------------------------------------
    def _setup_col_left(self):
        """Coluna 1: versões locais + ações de abrir/atualizar o sistema."""
        self._titulo_coluna(self.col_left, "Manager")

        f_btn = bstrap.Frame(self.col_left)
        f_btn.pack(side=BOTTOM, fill=X, pady=(12, 0))
        self._botao(f_btn, "▶  Abrir Sistema", self._lancar_erp,
                    "primary", fill=X, pady=(0, 6))
        self._botao(f_btn, "⚡  Atualizar Sistema", self._lancar_atualizacao,
                    "outline", fill=X, pady=(0, 6))
        self._botao(
            f_btn, "🔄  Recarregar Lista",
            lambda: threading.Thread(
                target=self._popular_versoes_locais, daemon=True).start(),
            "outline", fill=X,
        )

        card = self._card(self.col_left)
        card.pack(fill=BOTH, expand=YES)

        self.var_busca_versao = tk.StringVar()
        self.var_busca_versao.trace_add("write", self._filtrar_versoes)
        self._entrada_clara(card, self.var_busca_versao).pack(fill=X, pady=(0, 10))

        self.lb_versoes = self._tabela(card, "versao", "Versões Locais", height=16)
        self.lb_versoes.bind("<Double-1>", lambda e: self._lancar_atualizacao())

    # --- Coluna 2: Info Base ---------------------------------------------
    def _setup_col_center(self):
        """Coluna 2: banco ativo, instância SQL e ferramentas."""
        self._titulo_coluna(self.col_center, "Info Base")

        card = self._card(self.col_center, padding=16)
        card.pack(fill=X)

        self._rotulo(card, "Banco Atual:")
        self.lbl_db_atual = ttk.Label(card, text="A carregar...", style="CardValue.TLabel")
        self.lbl_db_atual.pack(anchor=W, pady=(0, 12))

        self._rotulo(card, "Servidor:")
        self.combo_instancia = bstrap.Combobox(card, state="readonly", bootstyle="dark")
        self.combo_instancia.pack(fill=X, pady=(2, 12))
        self.combo_instancia.bind("<<ComboboxSelected>>", self._on_instancia_changed)

        self._rotulo(card, "Versão SQL:")
        self.lbl_versao_sql = ttk.Label(card, text="...", style="CardValue.TLabel")
        self.lbl_versao_sql.pack(anchor=W, pady=(0, 16))

        self._rotulo(card, "Trocar Banco:")
        self.combo_db = bstrap.Combobox(card, state="readonly", bootstyle="dark")
        self.combo_db.pack(fill=X, pady=(2, 10))
        self.combo_db.bind("<<ComboboxSelected>>", self._preview_version)
        self._botao(card, "Selecionar Banco", self._mudar_banco,
                    "primary", fill=X)

        # Ferramentas / drop de bancos
        f_tools = self._card(self.col_center, padding=16)
        f_tools.pack(fill=BOTH, expand=YES, pady=(14, 0))
        self._rotulo(f_tools, "Ferramentas", estilo="CardTitle.TLabel", pady=(0, 10))

        # Empacotado antes da lista para reservar sua altura: com a lista
        # em expand=YES, o que vem depois seria espremido a zero.
        self._botao(f_tools, "Excluir Banco(s)", self._drop_database,
                    "outline-danger", side=BOTTOM, fill=X, pady=(10, 0))

        self.var_busca_banco = tk.StringVar()
        self.var_busca_banco.trace_add("write", self._filtrar_bancos)
        self._entrada_clara(f_tools, self.var_busca_banco).pack(fill=X, pady=(0, 8))

        self.lb_tools = self._tabela(
            f_tools, "banco", "Bases de Dados", height=7, selectmode="extended"
        )

    # --- Coluna 3: Restaurador -------------------------------------------
    def _setup_col_right(self):
        """Coluna 3: seleção de backup e restore."""
        self._titulo_coluna(self.col_right, "Restaurador")

        # A área de ações é empacotada primeiro, no rodapé da coluna, para que
        # o log reserve sua altura antes de a lista tomar o espaço restante.
        f_acoes = bstrap.Frame(self.col_right)
        f_acoes.pack(side=BOTTOM, fill=X, pady=(12, 0))

        self._rotulo(f_acoes, "Nome do Banco (Sem Max_):",
                     estilo="Muted.TLabel", pady=(0, 4))
        self.entry_new_db = self._entrada_clara(f_acoes)
        self.entry_new_db.pack(fill=X, pady=(0, 8))

        self.btn_restore = self._botao(
            f_acoes, "Restaurar Banco", self._iniciar_restore, "primary", fill=X
        )

        # Só aparece durante o restore (ver _mostrar_progresso)
        self.progress = ttk.Progressbar(
            f_acoes, mode='indeterminate', style="Restore.Horizontal.TProgressbar"
        )

        self.log_wrap = log_wrap = ttk.Frame(f_acoes, style="Card.TFrame")
        log_wrap.pack(fill=X, pady=(8, 0))
        self.log_txt = tk.Text(
            log_wrap, height=5, state='disabled', font=("Consolas", 8),
            bg=ui_theme.SURFACE, fg=ui_theme.FG_BODY,
            relief=FLAT, borderwidth=0, highlightthickness=0, wrap="none",
            insertbackground=ui_theme.FG,
        )
        # ttkbootstrap redefine o padrão dos widgets tk clássicos, então o
        # highlight precisa ser zerado após a construção, não no construtor.
        self.log_txt.configure(
            highlightthickness=0,
            highlightbackground=ui_theme.SURFACE,
            highlightcolor=ui_theme.SURFACE,
        )
        self._scrollbar(log_wrap, self.log_txt).pack(side=RIGHT, fill=Y)
        self.log_txt.pack(side=LEFT, fill=BOTH, expand=YES)

        card = self._card(self.col_right)
        card.pack(fill=BOTH, expand=YES)

        self.var_busca_backup = tk.StringVar()
        self.var_busca_backup.trace_add("write", self._filtrar_backups)
        self._entrada_clara(card, self.var_busca_backup).pack(fill=X, pady=(0, 10))

        self.lb_backups = self._tabela(
            card, "backup", "Arquivo de Backup (.BAK / .MAX)", height=14
        )

    # --- Rodapé -----------------------------------------------------------
    def _criar_rodape(self):
        """Barra inferior: status + botão de configurações."""
        bts = bstrap.Frame(self, padding=(20, 8, 20, 12))
        bts.pack(side=BOTTOM, fill=X)

        self.status = ttk.Label(
            bts, text="A inicializar...", anchor=W, style="Muted.TLabel"
        )
        self.status.pack(side=LEFT, fill=X, expand=YES)

        self._botao(bts, "⚙  Configurações", self._abrir_configuracoes,
                    "outline", side=RIGHT)

    # =========================================================================
    # PAINEL DA NUVEM (retrátil)
    # =========================================================================
    @staticmethod
    def _label_aba_nuvem(aberto):
        """Texto vertical da aba lateral, com a seta apontando para a ação."""
        return "\n".join(("›" if aberto else "‹") + "☁NUVEM")

    def _criar_aba_nuvem(self):
        """Aba vertical na borda direita que abre/fecha o painel da nuvem."""
        strip = bstrap.Frame(self)
        strip.pack(side=RIGHT, fill=Y)

        self.btn_toggle_cloud = RoundedButton(
            strip, text=self._label_aba_nuvem(False), variant="primary",
            command=self._toggle_cloud, padx=8, pady=16,
        )
        self.btn_toggle_cloud.pack(expand=YES, padx=(2, 6))

    def _criar_painel_nuvem(self):
        """Painel lateral com os navegadores WebDAV (versões e backups)."""
        self.cloud_panel = ttk.Frame(self, style="Card.TFrame", padding=12,
                                     width=CLOUD_PANEL_WIDTH)
        self.cloud_panel.pack_propagate(False)

        ttk.Label(
            self.cloud_panel, text="☁  Cloud Maxdata", style="CardTitle.TLabel",
            font=(ui_theme.FONT_FAMILY, 13, "bold"),
        ).pack(fill=X, pady=(0, 10))

        nb = bstrap.Notebook(self.cloud_panel)
        nb.pack(fill=BOTH, expand=YES)

        tab_versoes = ttk.Frame(nb, style="Card.TFrame", padding=8)
        tab_backups = ttk.Frame(nb, style="Card.TFrame", padding=8)
        nb.add(tab_versoes, text="Versões")
        nb.add(tab_backups, text="Backups")

        self.lbl_caminho_versoes, self.lb_nuvem_versoes = self._montar_aba_nuvem(
            tab_versoes,
            client=self.webdav_versoes,
            extensoes=EXTENSOES_VERSAO,
            destino=lambda: self.cfg.pasta_das_versoes,
            on_complete=self._popular_versoes_locais,
        )
        self.lbl_caminho_backups, self.lb_nuvem_backups = self._montar_aba_nuvem(
            tab_backups,
            client=self.webdav_backups,
            extensoes=EXTENSOES_BACKUP_NUVEM,
            destino=lambda: self.cfg.caminho_base_backup,
            on_complete=self._load_backups,
        )

    def _montar_aba_nuvem(self, parent, client, extensoes, destino, on_complete):
        """Monta uma aba de navegação WebDAV.

        Args:
            parent: Frame da aba.
            client: WebDAVClient a usar.
            extensoes: Extensões de arquivo a listar.
            destino: Callable que devolve o diretório local de destino.
            on_complete: Callable a rodar após um download concluído.

        Returns:
            Tupla (label de caminho, treeview).
        """
        lbl_caminho = ttk.Label(
            parent, text="/", style="CardMuted.TLabel",
            font=(ui_theme.FONT_FAMILY, 8),
        )
        lbl_caminho.pack(fill=X, pady=(0, 6))

        # Barra de ações antes da lista, para não ser espremida por ela
        f_btn = ttk.Frame(parent, style="Card.TFrame")
        f_btn.pack(side=BOTTOM, fill=X, pady=(8, 0))

        wrapper = ttk.Frame(parent, style="Card.TFrame")
        wrapper.pack(fill=BOTH, expand=YES)

        tree = ttk.Treeview(
            wrapper, columns=("tipo", "nome"), show="headings",
            style="Claro.Treeview", height=18,
        )
        tree.heading("tipo", text="")
        tree.heading("nome", text="Nome", anchor=W)
        tree.column("tipo", width=30, stretch=NO, anchor="center")
        tree.column("nome", anchor=W, stretch=True)

        self._scrollbar(wrapper, tree).pack(side=RIGHT, fill=Y)
        tree.pack(side=LEFT, fill=BOTH, expand=YES)

        tree.bind(
            "<Double-1>",
            lambda e: self._double_click_nuvem(e, client, tree, lbl_caminho, extensoes),
        )

        self._botao(
            f_btn, "⬆", lambda: self._voltar_nuvem(client, tree, lbl_caminho, extensoes),
            "outline", side=LEFT, padx=(0, 4),
        )
        self._botao(
            f_btn, "🔄",
            lambda: threading.Thread(
                target=self._popular_nuvem,
                args=(client, tree, lbl_caminho, extensoes, True),
                daemon=True).start(),
            "outline", side=LEFT, padx=(0, 4),
        )
        self._botao(
            f_btn, "⬇  Baixar",
            lambda: self._baixar_nuvem(client, tree, destino(), on_complete),
            "primary", side=LEFT, fill=X, expand=YES,
        )

        return lbl_caminho, tree

    def _toggle_cloud(self):
        """Abre ou fecha o painel da nuvem."""
        if self._cloud_visivel:
            self.cloud_panel.pack_forget()
            self.btn_toggle_cloud.config(text=self._label_aba_nuvem(False))
            self._cloud_visivel = False
            return

        # `before` é essencial: main_container tem expand=YES e consumiria
        # todo o espaço restante se o painel fosse alocado depois dele.
        self.cloud_panel.pack(side=RIGHT, fill=Y, before=self.main_container)
        self.btn_toggle_cloud.config(text=self._label_aba_nuvem(True))
        self._cloud_visivel = True

        # Carrega sob demanda, só na primeira abertura de cada aba
        for client, tree, lbl, ext in (
            (self.webdav_versoes, self.lb_nuvem_versoes,
             self.lbl_caminho_versoes, EXTENSOES_VERSAO),
            (self.webdav_backups, self.lb_nuvem_backups,
             self.lbl_caminho_backups, EXTENSOES_BACKUP_NUVEM),
        ):
            if id(tree) not in self._nuvem_carregada:
                threading.Thread(
                    target=self._popular_nuvem, args=(client, tree, lbl, ext),
                    daemon=True,
                ).start()

    # =========================================================================
    # CONFIGURAÇÕES
    # =========================================================================
    def _abrir_configuracoes(self):
        """Abre janela pop-up de configurações."""
        # bstrap.Toplevel recebe `title` como 1º posicional — o pai vai em `master`.
        cfg_win = bstrap.Toplevel(
            master=self, title="Configurações do Sistema", size=(560, 620)
        )
        cfg_win.transient(self)
        cfg_win.grab_set()

        scroll_canvas = bstrap.Canvas(cfg_win, highlightthickness=0,
                                      background=ui_theme.BG)
        scrollbar = self._scrollbar(cfg_win, scroll_canvas)
        scroll_frame = bstrap.Frame(scroll_canvas, padding=16)
        scroll_frame.configure(style="TFrame")
        scroll_frame.bind(
            "<Configure>",
            lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")),
        )
        scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_canvas.configure(yscrollcommand=scrollbar.set)
        scroll_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.cfg_vars = {}
        for title, keys, ini_section in CONFIG_SECTIONS_UI:
            f = self._card(scroll_frame, padding=14)
            f.pack(fill=X, pady=5)
            ttk.Label(
                f, text=title, style="CardTitle.TLabel",
                font=(ui_theme.FONT_FAMILY, 11, "bold"),
            ).pack(anchor=W, pady=(0, 10))

            for key in keys:
                row = ttk.Frame(f, style="Card.TFrame")
                row.pack(fill=X, pady=3)
                ttk.Label(
                    row, text=key, width=22, style="CardMuted.TLabel",
                    font=(ui_theme.FONT_FAMILY, 8),
                ).pack(side=LEFT)

                var = tk.StringVar(value=self.cfg.get_campo(ini_section, key))
                self.cfg_vars[(ini_section, key)] = var
                entry = self._entrada_clara(row, var)
                if "SENHA" in key.upper():
                    entry.config(show="•")
                entry.pack(side=LEFT, fill=X, expand=YES)

        self._botao(scroll_frame, "💾  Guardar",
                    lambda: self._salvar_config_aba(cfg_win),
                    "primary", pady=18, anchor=E)

    def _salvar_config_aba(self, janela=None):
        """Salva as configurações editadas e recria os serviços dependentes."""
        try:
            for (section, key), var in self.cfg_vars.items():
                self.cfg.set_campo(section, key, var.get())
            self.cfg.salvar()
            self.cfg.carregar()
            self.cfg.validar_caminhos()

            # Recriar serviços com novas configurações
            self.sql = SqlService(self.cfg)
            # Reconfigurar em vez de recriar: o painel da nuvem guarda estes
            # clientes em closures dos botões.
            for cliente in (self.webdav_versoes, self.webdav_backups):
                cliente.reconfigurar(
                    self.cfg.url_cloud, self.cfg.usuario_cloud, self.cfg.senha_cloud
                )
            # Credenciais podem ter mudado: força nova listagem da nuvem.
            self._nuvem_carregada.clear()

            if janela is not None:
                janela.destroy()
            messagebox.showinfo("Sucesso", "Configurações guardadas!")
            logger.info("Configurações salvas com sucesso")
            threading.Thread(target=self._carregamento_assincrono, daemon=True).start()
        except Exception as e:
            logger.error("Erro ao guardar configurações: %s", e)
            messagebox.showerror("Erro", f"Erro ao guardar: {e}")

    # =========================================================================
    # LÓGICA: FILTROS DE BUSCA
    # =========================================================================
    @staticmethod
    def _repovoar(tree, itens, busca):
        """Repopula um treeview de coluna única aplicando o filtro de busca."""
        for i in tree.get_children():
            tree.delete(i)
        termo = busca.lower()
        for item in itens:
            if termo in item.lower():
                tree.insert("", END, values=(item,))

    def _filtrar_versoes(self, *args):
        """Filtra a lista de versões locais pelo termo de busca."""
        self._repovoar(self.lb_versoes, self._all_versoes, self.var_busca_versao.get())

    def _filtrar_backups(self, *args):
        """Filtra a lista de backups locais pelo termo de busca."""
        self._repovoar(self.lb_backups, self._all_backups, self.var_busca_backup.get())

    def _filtrar_bancos(self, *args):
        """Filtra a lista de bancos SQL pelo termo de busca."""
        self._repovoar(self.lb_tools, self._all_dbs, self.var_busca_banco.get())

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
            self.cfg.caminho_do_ini,
            self.cfg.ini_section,
            self.cfg.ini_key,
            self.cfg.ini_server_key
        )

        if server_atual:
            self.cfg.sql_server_instance = server_atual
            self.cfg.servidor = server_atual
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
                    self.after(0, lambda v=vals: self.combo_instancia.set(v[0]))

        versao = self.sql.get_versao(atual)
        bancos = self.sql.listar_bancos()
        self.after(0, lambda: self._atualizar_ui_sql(bancos, atual, versao))

    def _preview_version(self, event):
        """Mostra a versão do banco selecionado no combo."""
        db = self.combo_db.get()
        threading.Thread(
            target=lambda: self._set_status(
                f"Banco selecionado: {db} (Versão: {self.sql.get_versao(db)})"
            ), daemon=True
        ).start()

    def _on_instancia_changed(self, event=None):
        """Handler para mudança de instância SQL."""
        instancia = self.combo_instancia.get()
        self.cfg.sql_server_instance = instancia
        self.cfg.servidor = instancia

        try:
            c = IniService.ler_arquivo(self.cfg.caminho_do_ini)
            IniService.set_value(c, self.cfg.ini_section or 'CON',
                                 self.cfg.ini_server_key or 'Data Source', instancia)
            IniService.salvar(c, self.cfg.caminho_do_ini)
        except Exception as e:
            logger.warning("Erro ao salvar instância no INI: %s", e)

        self._set_status(f"Instância alterada para: {instancia}. A recarregar bancos...")
        threading.Thread(target=self._carregar_banco_atual_sql, daemon=True).start()

    def _mudar_banco(self):
        """Troca o banco de dados no max.ini."""
        novo = self.combo_db.get()
        if not novo:
            return
        try:
            c = IniService.ler_arquivo(self.cfg.caminho_do_ini)
            IniService.set_value(c, self.cfg.ini_section or 'CON',
                                 self.cfg.ini_key or 'Initial catalog', novo)
            if hasattr(self, 'combo_instancia'):
                inst = self.combo_instancia.get()
                if inst:
                    IniService.set_value(c, self.cfg.ini_section or 'CON',
                                         self.cfg.ini_server_key or 'Data Source', inst)
            IniService.salvar(c, self.cfg.caminho_do_ini)
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
            if os.path.exists(self.cfg.pasta_das_versoes):
                versoes = [
                    e.name for e in os.scandir(self.cfg.pasta_das_versoes)
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
        # O PROPFIND do Nextcloud costuma levar vários segundos; sem esta
        # marcação a lista fica vazia nesse intervalo e parece defeito.
        def mostrar_carregando():
            lbl_caminho.config(text=client.caminho_atual)
            for i in treeview.get_children():
                treeview.delete(i)
            treeview.insert("", END, values=("⏳", "A carregar..."))

        self.after(0, mostrar_carregando)
        self._set_status(f"Nuvem: a listar {client.caminho_atual}...")

        try:
            pastas, arquivos = client.listar(force_refresh=force_refresh, extensoes=extensoes)

            def att_ui():
                for i in treeview.get_children():
                    treeview.delete(i)
                for p in pastas:
                    treeview.insert("", END, values=("📁", p))
                for a in arquivos:
                    treeview.insert("", END, values=("📄", a))
                if not pastas and not arquivos:
                    treeview.insert("", END, values=("—", "(pasta vazia)"))
                self.status.config(
                    text=f"Nuvem: {len(pastas)} pastas, {len(arquivos)} arquivos."
                )

            self.after(0, att_ui)
            self._nuvem_carregada.add(id(treeview))
        except Exception as e:
            logger.warning("Erro WebDAV: %s", e)
            erro = str(e)[:60]

            def att_erro():
                for i in treeview.get_children():
                    treeview.delete(i)
                treeview.insert("", END, values=("⚠", f"Erro: {erro}"))

            self.after(0, att_erro)
            self._set_status(f"Erro WebDAV: {erro}")

    def _voltar_nuvem(self, client, treeview, lbl_caminho, extensoes):
        """Volta uma pasta no navegador WebDAV."""
        if client.voltar():
            threading.Thread(
                target=self._popular_nuvem,
                args=(client, treeview, lbl_caminho, extensoes),
                daemon=True
            ).start()

    def _double_click_nuvem(self, event, client, treeview, lbl_caminho, extensoes):
        """Handler de duplo-clique: navega para subpasta."""
        sel = treeview.selection()
        if not sel:
            return
        tipo, nome = treeview.item(sel[0], "values")

        if tipo == "📁":
            client.navegar(nome)
            threading.Thread(
                target=self._popular_nuvem,
                args=(client, treeview, lbl_caminho, extensoes),
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
            messagebox.showinfo("Aviso", "Selecione um arquivo na lista da nuvem.")
            return
        tipo, nome = treeview.item(sel[0], "values")

        # Só linhas de arquivo: exclui pastas e as linhas de estado
        # (a carregar / erro / pasta vazia).
        if tipo != "📄":
            messagebox.showinfo("Aviso", "Selecione um ARQUIVO (📄) para baixar.")
            return

        if not os.path.isdir(destino_dir):
            messagebox.showerror("Erro", f"Pasta de destino inexistente:\n{destino_dir}")
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
            client.download(
                nome, destino_dir,
                on_progress=lambda pct: self._set_status(f"A baixar {nome}... {pct}%"),
            )

            self._set_status(f"Download de {nome} concluído!")
            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Download concluído!\n{nome}"))
            if on_complete_callback:
                threading.Thread(target=on_complete_callback, daemon=True).start()
        except Exception as e:
            logger.error("Erro no download: %s", e)
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Download", f"Falha:\n{m}"))
            self._set_status("Erro no download.")

    # =========================================================================
    # LÓGICA: BACKUPS LOCAIS
    # =========================================================================
    def _load_backups(self):
        """Carrega lista de backups locais, mais recentes primeiro."""
        try:
            if not os.path.exists(self.cfg.caminho_base_backup):
                return
            sufixos = tuple(ext.upper() for ext in EXTENSOES_BACKUP)
            backups = [
                (e.name, e.stat().st_mtime)
                for e in os.scandir(self.cfg.caminho_base_backup)
                if e.is_file() and e.name.upper().endswith(sufixos)
            ]
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

    def _versao_selecionada(self):
        """Nome do arquivo de versão selecionado na lista, ou None."""
        sel = self.lb_versoes.selection()
        if not sel:
            return None
        return self.lb_versoes.item(sel[0], "values")[0]

    def _lancar_erp(self):
        """Abre o sistema.

        Com uma versão selecionada na lista, extrai-a primeiro e só então abre
        o Manager. Sem seleção, mantém a checagem de compatibilidade com a
        versão gravada no banco.
        """
        arq = self._versao_selecionada()
        if arq:
            if messagebox.askyesno(
                "Abrir Sistema",
                f"Extrair '{arq}' sobre a pasta do sistema e abrir o Manager?"
            ):
                self._extrair_em_thread(arq, self._abrir_manager)
            return

        self._abrir_com_checagem_de_versao()

    def _abrir_com_checagem_de_versao(self):
        """Abre o Manager comparando a versão do EXE com a do banco."""
        try:
            db_versao = self.lbl_versao_sql.cget("text")
            exe_versao = self._get_exe_versao(self.cfg.caminho_do_erp_cliente)

            if (db_versao and exe_versao
                    and db_versao not in ("---", "N/A")
                    and exe_versao != db_versao):
                arquivo_rar = next(
                    (v for v in self._all_versoes if db_versao in v), None
                )
                if arquivo_rar:
                    if messagebox.askyesno(
                        "Atualização Necessária",
                        f"Versão BD: {db_versao} | Versão EXE: {exe_versao}.\n\n"
                        f"Atualizar para '{arquivo_rar}'?"
                    ):
                        self._extrair_em_thread(arquivo_rar, self._abrir_atualizador)
                        return
                else:
                    messagebox.showwarning(
                        "Aviso de Versão",
                        "Versões diferentes e ficheiro de extração não encontrado.\nA ignorar..."
                    )

            self._abrir_manager()
        except Exception as e:
            logger.error("Erro ao lançar ERP: %s", e)
            messagebox.showerror("Erro", f"Erro: {e}")

    def _lancar_atualizacao(self):
        """Extrai a versão selecionada e abre o atualizador."""
        arq = self._versao_selecionada()
        if not arq:
            messagebox.showinfo("Aviso", "Selecione uma versão na lista.")
            return
        if messagebox.askyesno("Confirmar", f"Atualizar para {arq}?"):
            self._extrair_em_thread(arq, self._abrir_atualizador)

    def _bloqueio_de_extracao(self):
        """Executáveis do sistema que estão abertos e impediriam a extração."""
        return sevenzip.executaveis_bloqueados([
            self.cfg.caminho_do_erp_cliente,
            self.cfg.caminho_do_max_atualiza,
        ])

    def _extrair_em_thread(self, arq, ao_terminar):
        """Dispara a extração em background.

        Args:
            arq: Nome do arquivo de versão em `pasta_das_versoes`.
            ao_terminar: Callback executado na thread da UI após a extração.
        """
        abertos = self._bloqueio_de_extracao()
        if abertos:
            messagebox.showwarning(
                "Sistema aberto",
                "Feche antes de extrair:\n\n  " + "\n  ".join(abertos) +
                "\n\nO 7-Zip não consegue sobrescrever um executável em uso."
            )
            self._set_status("Extração cancelada: sistema aberto.")
            return

        threading.Thread(
            target=self._thread_extrair, args=(arq, ao_terminar), daemon=True
        ).start()

    def _thread_extrair(self, arq, ao_terminar):
        """Thread de extração de versão com 7-Zip."""
        try:
            self._set_status(f"A extrair {arq}...")
            sevenzip.extrair(
                self.cfg.caminho_do_7zip,
                os.path.join(self.cfg.pasta_das_versoes, arq),
                self.cfg.pasta_do_sistema,
            )
            logger.info("Versão extraída: %s", arq)
            self.after(0, ao_terminar)
        except Exception as e:
            logger.error("Erro na extração: %s", e)
            self._set_status("Erro na extração.")
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Extração", m))

    def _abrir_manager(self):
        """Abre o MAX_Manager2.exe."""
        try:
            subprocess.Popen([self.cfg.caminho_do_erp_cliente],
                             cwd=self.cfg.pasta_do_sistema)
            self.status.config(text="Sistema aberto.")
        except Exception as e:
            logger.error("Erro ao abrir o sistema: %s", e)
            messagebox.showerror("Erro", f"{e}")

    def _abrir_atualizador(self):
        """Abre o MAX_Atualiza.exe após extração."""
        try:
            subprocess.Popen([self.cfg.caminho_do_max_atualiza],
                             cwd=self.cfg.pasta_do_sistema)
            self.status.config(text="Atualizador aberto.")
        except Exception as e:
            logger.error("Erro ao abrir atualizador: %s", e)
            messagebox.showerror("Erro", f"{e}")

    # =========================================================================
    # LÓGICA: RESTORE
    # =========================================================================
    def _iniciar_restore(self):
        """Inicia o processo de restore em thread separada."""
        sel = self.lb_backups.selection()
        if not sel:
            messagebox.showinfo("Aviso", "Selecione um arquivo de backup na lista.")
            return
        fname = self.lb_backups.item(sel[0], "values")[0]

        new_db = self.entry_new_db.get().strip()
        if not new_db:
            messagebox.showinfo("Aviso", "Informe o nome do novo banco.")
            self.entry_new_db.focus_set()
            return

        self._mostrar_progresso(True)
        self.log_txt.config(state='normal')
        self.log_txt.delete('1.0', END)
        self.log_txt.config(state='disabled')
        threading.Thread(target=self._restore_logic, args=(fname, new_db), daemon=True).start()

    def _mostrar_progresso(self, ativo):
        """Mostra/esconde a barra de progresso e trava o botão de restore."""
        if ativo:
            self.progress.pack(fill=X, pady=(8, 0), before=self.log_wrap)
            self.progress.start()
            self.btn_restore.config(state='disabled')
        else:
            self.progress.stop()
            self.progress.pack_forget()
            self.btn_restore.config(state='normal')

    def _restore_logic(self, fname, dbname):
        """Orquestra o restore usando SqlService."""
        try:
            self.sql.executar_restore_completo(
                fname=fname,
                dbname=dbname,
                caminho_backup=self.cfg.caminho_base_backup,
                pasta_sistema=self.cfg.pasta_do_sistema,
                caminho_7zip=self.cfg.caminho_do_7zip,
                on_message=self.msg_queue.put
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
                    self._mostrar_progresso(False)
                    messagebox.showinfo("Sucesso", "Restore Concluído!")
                    threading.Thread(target=self._carregar_banco_atual_sql, daemon=True).start()
                elif msg == "__ERROR__":
                    self._mostrar_progresso(False)
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
            messagebox.showinfo("Aviso", "Selecione ao menos um banco na lista.")
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

        confirmacao = simpledialog.askstring("PERIGO", msg, parent=self)

        if confirmacao == "EXCLUIR":
            threading.Thread(target=self._thread_drop, args=(bancos_alvo,), daemon=True).start()
        elif confirmacao is not None:
            messagebox.showwarning("Cancelado", "Palavra de confirmação inválida.")

    def _thread_drop(self, bancos):
        """Thread de exclusão de bancos."""
        try:
            self._set_status(f"A eliminar {len(bancos)} banco(s)...")
            self.sql.drop_databases(bancos)
            self.after(0, lambda: messagebox.showinfo("Sucesso", "Bases eliminadas com sucesso!"))
            self._carregar_banco_atual_sql()
            self._set_status("Pronto.")
        except Exception as e:
            logger.error("Erro ao eliminar bancos: %s", e)
            self.after(0, lambda err=str(e): messagebox.showerror("Erro", err))
