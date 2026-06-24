import tkinter as tk
from tkinter import W, END, BOTH, YES, X, Y, LEFT, RIGHT, FLAT, BOTTOM
from tkinter import ttk, scrolledtext, messagebox, simpledialog
from tkinter import filedialog
import ttkbootstrap as bstrap
import os
import pyodbc
import threading
import queue
import sys
import subprocess
import configparser
import time
import shutil
import urllib.request
import urllib.error
import base64
import re
import urllib.parse
import xml.etree.ElementTree as ET

try:
    from pywinauto.application import Application
except ImportError:
    pass

# =============================================================================
# CONFIGURAÇÕES E VARIÁVEIS GLOBAIS
# =============================================================================
CONFIG_FILE_NAME = 'gerenciador_config.ini'
SQL_QUERY_VERSAO = "select cofMaxAtualizaVersao from config" 

PASTA_DO_SISTEMA = ""
NOME_EXE_CLIENTE = ""
NOME_EXE_ATUALIZADOR = ""
PASTA_DAS_VERSOES = ""
CAMINHO_BASE_MAX_BACKUP = ""
CAMINHO_DO_INI = ""
CAMINHO_DO_7ZIP_EXE = ""
INI_SECTION = ""
INI_KEY = ""
INI_SERVER_KEY = ""
SERVIDOR = ""
USUARIO = ""
SENHA = ""
ODBC_DRIVER_RESTORE = ""
SQL_DRIVER_LISTA = ""
SQL_SERVER_INSTANCE = ""
URL_CLOUD = ""
USUARIO_CLOUD = ""
SENHA_CLOUD = ""
CAMINHO_NUVEM_ATUAL = ""
CAMINHO_NUVEM_BACKUP = ""

CAMINHO_DO_ERP_CLIENTE = ""
CAMINHO_DO_MAX_ATUALIZA = ""

PRIMEIRA_EXECUCAO = False

# =============================================================================
# FUNÇÕES DE CONFIGURAÇÃO (AUTO-DETECT INTELIGENTE)
# =============================================================================
def carregar_ou_criar_configuracoes():
    global PASTA_DO_SISTEMA, NOME_EXE_CLIENTE, NOME_EXE_ATUALIZADOR, PASTA_DAS_VERSOES
    global CAMINHO_BASE_MAX_BACKUP, CAMINHO_DO_INI, CAMINHO_DO_7ZIP_EXE, INI_SECTION, INI_KEY
    global SERVIDOR, USUARIO, SENHA, ODBC_DRIVER_RESTORE, SQL_DRIVER_LISTA, SQL_SERVER_INSTANCE
    global PRIMEIRA_EXECUCAO, INI_SERVER_KEY
    global URL_CLOUD, USUARIO_CLOUD, SENHA_CLOUD

    # Onde o executável está rodando agora?
    BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

    # AUTO-DETECT MULTI-DISCO (C:\ ou D:\)
    if os.path.exists(os.path.join(BASE_DIR, 'max.ini')):
        default_sistema = BASE_DIR
    elif os.path.exists(r'C:\Max'):
        default_sistema = r'C:\Max'
    elif os.path.exists(r'D:\Max'):
        default_sistema = r'D:\Max'
    else:
        default_sistema = r'C:\Max' # Se não achar nada, sugere C:\ como padrão

    config = configparser.ConfigParser()
    
    # Auto-detect do 7-Zip
    _7z_path = r'C:\Program Files\7-Zip\7z.exe'
    if not os.path.exists(_7z_path):
        _7z_path = r'C:\Program Files (x86)\7-Zip\7z.exe'

    # SETUP DA PRIMEIRA EXECUÇÃO
    if not os.path.exists(CONFIG_FILE_NAME):
        PRIMEIRA_EXECUCAO = True
        config['CAMINHOS'] = {
            'PASTA_DO_SISTEMA': default_sistema,
            'PASTA_DAS_VERSOES': os.path.join(default_sistema, 'Versões'),
            'PASTA_DE_BACKUP': os.path.join(default_sistema, 'backup'),
            'CAMINHO_DO_INI': os.path.join(default_sistema, 'max.ini'),
            'CAMINHO_DO_7ZIP_EXE': _7z_path
        }
        config['EXECUTAVEIS'] = {
            'NOME_EXE_CLIENTE': 'MAX_manager2.exe',
            'NOME_EXE_ATUALIZADOR': 'MAX_Atualiza.exe'
        }
        config['CONFIG_INI_MAX'] = {'INI_SECTION': 'CON', 'INI_KEY': 'Initial catalog', 'INI_SERVER_KEY': 'Data Source'}
        config['SQL_LAUDO'] = {'SQL_DRIVER_LISTA': '{ODBC Driver 17 for SQL Server}', 'SQL_SERVER_INSTANCE': 'localhost'}
        config['SQL_RESTORE'] = {'SERVIDOR': 'localhost', 'USUARIO': 'sa', 'SENHA': 'macro01', 'ODBC_DRIVER_RESTORE': '{ODBC Driver 17 for SQL Server}'}
        config['CLOUD'] = {'URL_CLOUD': 'https://cloud.maxdata.com.br', 'USUARIO_CLOUD': '', 'SENHA_CLOUD': ''}
        
        try:
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f: config.write(f)
        except: pass

    # LEITURA DAS VARIÁVEIS (Usa o que está salvo no INI)
    try:
        config.read(CONFIG_FILE_NAME, encoding='utf-8')
        PASTA_DO_SISTEMA = config.get('CAMINHOS', 'PASTA_DO_SISTEMA')
        PASTA_DAS_VERSOES = config.get('CAMINHOS', 'PASTA_DAS_VERSOES')
        CAMINHO_BASE_MAX_BACKUP = config.get('CAMINHOS', 'PASTA_DE_BACKUP', fallback=os.path.join(PASTA_DO_SISTEMA, 'backup'))
        CAMINHO_DO_INI = config.get('CAMINHOS', 'CAMINHO_DO_INI')
        CAMINHO_DO_7ZIP_EXE = config.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE')
        NOME_EXE_CLIENTE = config.get('EXECUTAVEIS', 'NOME_EXE_CLIENTE')
        NOME_EXE_ATUALIZADOR = config.get('EXECUTAVEIS', 'NOME_EXE_ATUALIZADOR')
        INI_SECTION = config.get('CONFIG_INI_MAX', 'INI_SECTION')
        INI_KEY = config.get('CONFIG_INI_MAX', 'INI_KEY')
        INI_SERVER_KEY = config.get('CONFIG_INI_MAX', 'INI_SERVER_KEY', fallback='Data Source')
        SQL_DRIVER_LISTA = config.get('SQL_LAUDO', 'SQL_DRIVER_LISTA')
        SQL_SERVER_INSTANCE = config.get('SQL_LAUDO', 'SQL_SERVER_INSTANCE')
        SERVIDOR = config.get('SQL_RESTORE', 'SERVIDOR')
        USUARIO = config.get('SQL_RESTORE', 'USUARIO')
        SENHA = config.get('SQL_RESTORE', 'SENHA')
        ODBC_DRIVER_RESTORE = config.get('SQL_RESTORE', 'ODBC_DRIVER_RESTORE')
        URL_CLOUD = config.get('CLOUD', 'URL_CLOUD', fallback='https://cloud.maxdata.com.br')
        USUARIO_CLOUD = config.get('CLOUD', 'USUARIO_CLOUD', fallback='')
        SENHA_CLOUD = config.get('CLOUD', 'SENHA_CLOUD', fallback='')
    except Exception as e:
        print(f"Erro Config: {e}")

class ConfigWindow(tk.Toplevel):
    def __init__(self, parent, is_first_run=False):
        super().__init__(parent)
        self.title("Bem-vindo! Setup Inicial" if is_first_run else "Configuração de Caminhos")
        self.geometry("750x380")
        self.resizable(False, False)
        self.salvo = False 
        self.transient(parent); self.grab_set()

        self.var_sistema = tk.StringVar()
        self.var_versoes = tk.StringVar()
        self.var_backup = tk.StringVar()
        self.var_ini = tk.StringVar()
        self.var_7zip = tk.StringVar()

        self.ler_ini_atual()
        
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)
        
        if is_first_run:
            bstrap.Label(main_frame, text="Primeira execução! Confirme os diretórios de trabalho:", font=("bold", 12), bootstyle="success").pack(pady=(0, 15))
        else:
            bstrap.Label(main_frame, text="Corrija os caminhos:", font=("bold", 12), bootstyle="danger").pack(pady=(0, 15))

        self.criar_campo(main_frame, "Pasta do Sistema", self.var_sistema, True)
        self.criar_campo(main_frame, "Pasta de Versões", self.var_versoes, True)
        self.criar_campo(main_frame, "Pasta de Backups", self.var_backup, True)
        self.criar_campo(main_frame, "Arquivo .ini", self.var_ini, False)
        self.criar_campo(main_frame, "Executável 7-Zip", self.var_7zip, False)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, pady=(20, 0))
        bstrap.Button(btn_frame, text="Salvar e Continuar", command=self.salvar, bootstyle="success").pack(side=RIGHT, padx=5)
        if not is_first_run:
            bstrap.Button(btn_frame, text="Cancelar", command=self.destroy, bootstyle="secondary").pack(side=RIGHT)

    def criar_campo(self, parent, label, var, is_dir):
        f = ttk.Frame(parent)
        f.pack(fill=X, pady=5)
        bstrap.Label(f, text=label, width=20, bootstyle="inverse-secondary").pack(side=LEFT)
        bstrap.Entry(f, textvariable=var, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)
        cmd = lambda: self.selecionar(var, is_dir)
        bstrap.Button(f, text="...", command=cmd, bootstyle="info-outline", width=4).pack(side=LEFT)

    def selecionar(self, var, is_dir):
        p = filedialog.askdirectory() if is_dir else filedialog.askopenfilename()
        if p: var.set(p)

    def ler_ini_atual(self):
        try:
            c = configparser.ConfigParser()
            c.read(CONFIG_FILE_NAME, encoding='utf-8')
            self.var_sistema.set(c.get('CAMINHOS', 'PASTA_DO_SISTEMA', fallback=''))
            self.var_versoes.set(c.get('CAMINHOS', 'PASTA_DAS_VERSOES', fallback=''))
            self.var_backup.set(c.get('CAMINHOS', 'PASTA_DE_BACKUP', fallback=''))
            self.var_ini.set(c.get('CAMINHOS', 'CAMINHO_DO_INI', fallback=''))
            self.var_7zip.set(c.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', fallback=''))
        except: pass

    def salvar(self):
        try:
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_NAME, encoding='utf-8')
            if not config.has_section('CAMINHOS'): config.add_section('CAMINHOS')
            config.set('CAMINHOS', 'PASTA_DO_SISTEMA', self.var_sistema.get())
            config.set('CAMINHOS', 'PASTA_DAS_VERSOES', self.var_versoes.get())
            config.set('CAMINHOS', 'PASTA_DE_BACKUP', self.var_backup.get())
            config.set('CAMINHOS', 'CAMINHO_DO_INI', self.var_ini.get())
            config.set('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', self.var_7zip.get())
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f: config.write(f)
            self.salvo = True; self.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"{e}", parent=self)

# =============================================================================
# APLICAÇÃO PRINCIPAL
# =============================================================================
class GerenciadorMaxApp(bstrap.Window):
    def __init__(self):
        super().__init__(themename="cyborg")
        self.title("Gerenciador Max (Black & Red)")
        self.geometry("900x780")
        self.withdraw() 
        self.msg_queue = queue.Queue()

    def iniciar_interface(self):
        self.caminho_backup = CAMINHO_BASE_MAX_BACKUP 
        self.create_layout()
        self.process_queue()
        self.deiconify() 
        
        # Inicia carregamento em Background
        threading.Thread(target=self.carregamento_assincrono, daemon=True).start()

    def carregamento_assincrono(self):
        self.after(0, lambda: self.status.config(text="A carregar base de dados e ficheiros locais..."))
        self.popular_versoes()
        self.popular_versoes_nuvem()
        self.load_backups()
        self.carregar_banco_atual_sql()
        self.after(0, lambda: self.status.config(text="Pronto."))

    def create_layout(self):
        main = bstrap.Frame(self, padding=10)
        main.pack(fill=BOTH, expand=YES)

        top_frame = bstrap.Frame(main, padding=5)
        top_frame.pack(fill=X, pady=(0, 10))
        bstrap.Label(top_frame, text="Instância SQL Ativa:", font=("bold", 12), bootstyle="secondary").pack(side=LEFT)
        self.combo_instancia = bstrap.Combobox(top_frame, state="readonly", bootstyle="danger", width=30)
        self.combo_instancia.pack(side=LEFT, padx=10)
        self.combo_instancia.bind("<<ComboboxSelected>>", self.on_instancia_changed)
        
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

        self.setup_launcher(tab_launch)
        self.setup_restore(tab_restore)
        self.setup_tools(tab_tools)
        self.setup_config(tab_config)

        self.status = bstrap.Label(self, text="A inicializar...", relief=FLAT, anchor=W, padding=8, bootstyle="secondary")
        self.status.pack(side=BOTTOM, fill=X)

    # --- ABA 1: LAUNCHER ---
    def setup_launcher(self, parent):
        bstrap.Label(parent, text="Launcher Maxdata", font=("bold", 20), bootstyle="danger").pack(pady=5)
        db_frame = bstrap.Labelframe(parent, text=" Base de Dados ", bootstyle="danger", padding=15)
        db_frame.pack(fill=X, pady=10)

        f1 = bstrap.Frame(db_frame)
        f1.pack(fill=X, pady=5)
        bstrap.Label(f1, text="Banco Atual (INI):", bootstyle="secondary").pack(side=LEFT)
        self.lbl_db_atual = bstrap.Label(f1, text="A carregar...", font=("bold"), bootstyle="danger")
        self.lbl_db_atual.pack(side=LEFT, padx=5)
        
        bstrap.Label(f1, text="| Versão:", bootstyle="secondary").pack(side=LEFT, padx=(15, 5))
        self.lbl_versao_sql = bstrap.Label(f1, text="...", font=("bold"), bootstyle="warning") 
        self.lbl_versao_sql.pack(side=LEFT)

        f2 = bstrap.Frame(db_frame)
        f2.pack(fill=X, pady=5)
        bstrap.Label(f2, text="Trocar para:", bootstyle="secondary").pack(side=LEFT)
        self.combo_db = bstrap.Combobox(f2, state="readonly", bootstyle="danger")
        self.combo_db.pack(side=LEFT, fill=X, expand=YES, padx=5)
        self.combo_db.bind("<<ComboboxSelected>>", self.preview_version)
        bstrap.Button(f2, text="Guardar", command=self.mudar_banco, bootstyle="danger-outline").pack(side=LEFT)

        v_frame = bstrap.Labelframe(parent, text=" Gestão de Versões ", bootstyle="danger", padding=15)
        v_frame.pack(fill=BOTH, expand=YES, pady=10)

        self.nb_versoes = bstrap.Notebook(v_frame, bootstyle="danger")
        self.nb_versoes.pack(fill=BOTH, expand=YES)

        tab_locais = bstrap.Frame(self.nb_versoes, padding=10)
        tab_nuvem = bstrap.Frame(self.nb_versoes, padding=10)
        self.nb_versoes.add(tab_locais, text="💻 Versões Locais")
        self.nb_versoes.add(tab_nuvem, text="☁️ Nuvem Maxdata")

        # --- TAB LOCAIS ---
        f_busca = bstrap.Frame(tab_locais)
        f_busca.pack(fill=X, pady=(0, 5))
        bstrap.Label(f_busca, text="🔍 Buscar:").pack(side=LEFT)
        self.var_busca_versao = tk.StringVar()
        self.var_busca_versao.trace("w", self.filtrar_versoes)
        bstrap.Entry(f_busca, textvariable=self.var_busca_versao, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)

        self.lb_versoes = ttk.Treeview(tab_locais, height=6, bootstyle="danger", columns=("versao"), show="headings")
        self.lb_versoes.heading("versao", text="Arquivos (.rar) no PC")
        self.lb_versoes.column("versao", width=300, anchor=W)
        sb = bstrap.Scrollbar(tab_locais, command=self.lb_versoes.yview, bootstyle="danger-round")
        self.lb_versoes.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_versoes.pack(side=LEFT, fill=BOTH, expand=YES)
        
        # --- TAB NUVEM ---
        f_busca_n = bstrap.Frame(tab_nuvem)
        f_busca_n.pack(fill=X, pady=(0, 5))
        bstrap.Button(f_busca_n, text="🔄 Conectar/Recarregar", command=lambda: threading.Thread(target=self.popular_versoes_nuvem, daemon=True).start(), bootstyle="danger-outline").pack(side=LEFT)
        bstrap.Button(f_busca_n, text="⬅️ Voltar Pasta", command=self.voltar_pasta_nuvem, bootstyle="secondary").pack(side=LEFT, padx=5)
        
        self.lbl_caminho_nuvem = bstrap.Label(f_busca_n, text="/", bootstyle="secondary")
        self.lbl_caminho_nuvem.pack(side=LEFT, padx=5)

        self.lb_nuvem = ttk.Treeview(tab_nuvem, height=6, bootstyle="danger", columns=("tipo", "nome"), show="headings")
        self.lb_nuvem.heading("tipo", text="Tipo")
        self.lb_nuvem.column("tipo", width=50, stretch=False, anchor=W)
        self.lb_nuvem.heading("nome", text="Nome (Duplo-clique para abrir pasta)")
        self.lb_nuvem.column("nome", width=250, anchor=W)
        self.lb_nuvem.bind("<Double-1>", self.on_nuvem_double_click)
        
        sb_n = bstrap.Scrollbar(tab_nuvem, command=self.lb_nuvem.yview, bootstyle="danger-round")
        self.lb_nuvem.config(yscrollcommand=sb_n.set)
        sb_n.pack(side=RIGHT, fill=Y)
        self.lb_nuvem.pack(side=LEFT, fill=BOTH, expand=YES)

        bstrap.Button(tab_nuvem, text="⬇️ BAIXAR ARQUIVO SELECIONADO", command=self.baixar_da_nuvem, bootstyle="danger").pack(side=BOTTOM, fill=X, pady=5)

        bts = bstrap.Frame(parent)
        bts.pack(fill=X)
        bstrap.Button(bts, text="🔄 Recarregar Locais", command=lambda: threading.Thread(target=self.popular_versoes, daemon=True).start(), bootstyle="secondary-outline").pack(side=LEFT)
        bstrap.Button(bts, text="▶️ EXECUTAR SISTEMA", command=self.lancar_erp, bootstyle="success-outline").pack(side=RIGHT, padx=5)
        bstrap.Button(bts, text="⚡ ATUALIZAR VERSÃO (Extrair)", command=self.lancar_atualizacao, bootstyle="danger").pack(side=RIGHT)

    # --- ABA 2: RESTORE ---
    def setup_restore(self, parent):
        f1 = bstrap.Labelframe(parent, text=" 1. Selecione o Backup ", bootstyle="danger", padding=10)
        f1.pack(fill=BOTH, expand=YES)

        self.nb_backups = bstrap.Notebook(f1, bootstyle="danger")
        self.nb_backups.pack(fill=BOTH, expand=YES)

        tab_locais_bkp = bstrap.Frame(self.nb_backups, padding=10)
        tab_nuvem_bkp = bstrap.Frame(self.nb_backups, padding=10)
        self.nb_backups.add(tab_locais_bkp, text="💻 Backups Locais")
        self.nb_backups.add(tab_nuvem_bkp, text="☁️ Nuvem Maxdata")

        # --- TAB LOCAIS BKP ---
        f_busca = bstrap.Frame(tab_locais_bkp)
        f_busca.pack(fill=X, pady=(0, 5))
        bstrap.Label(f_busca, text="🔍 Buscar:").pack(side=LEFT)
        self.var_busca_backup = tk.StringVar()
        self.var_busca_backup.trace("w", self.filtrar_backups)
        bstrap.Entry(f_busca, textvariable=self.var_busca_backup, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)

        self.lb_backups = ttk.Treeview(tab_locais_bkp, height=6, bootstyle="danger", columns=("backup"), show="headings")
        self.lb_backups.heading("backup", text="Backups Disponíveis")
        self.lb_backups.column("backup", width=300, anchor=tk.W)
        sb = bstrap.Scrollbar(tab_locais_bkp, command=self.lb_backups.yview, bootstyle="danger-round")
        self.lb_backups.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_backups.pack(side=LEFT, fill=BOTH, expand=YES)
        bstrap.Button(tab_locais_bkp, text="🔄 Atualizar Locais", command=lambda: threading.Thread(target=self.load_backups, daemon=True).start(), bootstyle="secondary-outline").pack(fill=X, pady=5)

        # --- TAB NUVEM BKP ---
        f_busca_nb = bstrap.Frame(tab_nuvem_bkp)
        f_busca_nb.pack(fill=X, pady=(0, 5))
        bstrap.Button(f_busca_nb, text="🔄 Conectar/Recarregar", command=lambda: threading.Thread(target=self.popular_nuvem_backups, daemon=True).start(), bootstyle="danger-outline").pack(side=LEFT)
        bstrap.Button(f_busca_nb, text="⬅️ Voltar Pasta", command=self.voltar_pasta_nuvem_bkp, bootstyle="secondary").pack(side=LEFT, padx=5)
        
        self.lbl_caminho_nuvem_bkp = bstrap.Label(f_busca_nb, text="/", bootstyle="secondary")
        self.lbl_caminho_nuvem_bkp.pack(side=LEFT, padx=5)

        self.lb_nuvem_bkp = ttk.Treeview(tab_nuvem_bkp, height=6, bootstyle="danger", columns=("tipo", "nome"), show="headings")
        self.lb_nuvem_bkp.heading("tipo", text="Tipo")
        self.lb_nuvem_bkp.column("tipo", width=50, stretch=False, anchor=tk.W)
        self.lb_nuvem_bkp.heading("nome", text="Nome (Duplo-clique para abrir pasta)")
        self.lb_nuvem_bkp.column("nome", width=250, anchor=tk.W)
        self.lb_nuvem_bkp.bind("<Double-1>", self.on_nuvem_bkp_double_click)
        
        sb_nb = bstrap.Scrollbar(tab_nuvem_bkp, command=self.lb_nuvem_bkp.yview, bootstyle="danger-round")
        self.lb_nuvem_bkp.config(yscrollcommand=sb_nb.set)
        sb_nb.pack(side=RIGHT, fill=Y)
        self.lb_nuvem_bkp.pack(side=LEFT, fill=BOTH, expand=YES)

        bstrap.Button(tab_nuvem_bkp, text="⬇️ BAIXAR ARQUIVO SELECIONADO", command=self.baixar_da_nuvem_bkp, bootstyle="danger").pack(side=BOTTOM, fill=X, pady=5)

        f2 = bstrap.Labelframe(parent, text=" 2. Nome do Novo Banco ", bootstyle="danger", padding=10)
        f2.pack(fill=X, pady=10)
        self.entry_new_db = bstrap.Entry(f2, bootstyle="danger")
        self.entry_new_db.pack(fill=X)

        f3 = bstrap.Frame(parent)
        f3.pack(fill=BOTH, expand=YES)
        
        self.btn_restore = bstrap.Button(f3, text="▶️ INICIAR RESTAURAÇÃO", command=self.iniciar_restore_thread, bootstyle="danger")
        self.btn_restore.pack(fill=X, pady=5)
        
        self.progress = bstrap.Progressbar(f3, mode='indeterminate', bootstyle="danger-striped")
        self.progress.pack(fill=X, pady=5)
        
        self.log_txt = scrolledtext.ScrolledText(f3, height=8, state='disabled', font=("Consolas", 9), bg="#111", fg="#0f0")
        self.log_txt.pack(fill=BOTH, expand=YES)

    # --- ABA 3: TOOLS (COM MULTI-SELEÇÃO) ---
    def setup_tools(self, parent):
        bstrap.Label(parent, text="Gestão de Bases de Dados", font=("bold", 14), bootstyle="danger").pack(pady=10)
        flist = bstrap.Labelframe(parent, text=" Bases SQL (Use Shift/Ctrl para selecionar várias) ", bootstyle="danger", padding=10)
        flist.pack(fill=BOTH, expand=YES)
        
        f_busca = bstrap.Frame(flist)
        f_busca.pack(fill=X, pady=(0, 5))
        bstrap.Label(f_busca, text="🔍 Buscar:").pack(side=LEFT)
        self.var_busca_banco = tk.StringVar()
        self.var_busca_banco.trace("w", self.filtrar_bancos)
        bstrap.Entry(f_busca, textvariable=self.var_busca_banco, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES, padx=5)

        self.lb_tools = ttk.Treeview(flist, height=10, bootstyle="danger", columns=("banco"), show="headings", selectmode="extended")
        self.lb_tools.heading("banco", text="Bases de Dados")
        self.lb_tools.column("banco", width=300, anchor=tk.W)
        sb = bstrap.Scrollbar(flist, command=self.lb_tools.yview, bootstyle="danger-round")
        self.lb_tools.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_tools.pack(side=LEFT, fill=BOTH, expand=YES)
        
        fbtn = bstrap.Frame(flist)
        fbtn.pack(side=RIGHT, fill=Y, padx=10)
        
        bstrap.Button(fbtn, text="🔄 Atualizar", command=lambda: threading.Thread(target=self.carregar_banco_atual_sql, daemon=True).start(), bootstyle="secondary-outline").pack(fill=X, pady=5)
        bstrap.Button(fbtn, text="🗑️ ELIMINAR (DROP)", command=self.drop_database, bootstyle="danger").pack(fill=X, pady=20)

    def listar_instancias_sql(self):
        instancias = ["127.0.0.1", "localhost"]
        try:
            import winreg
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL")
            for i in range(1024):
                try:
                    name, value, _ = winreg.EnumValue(registry_key, i)
                    if name != "MSSQLSERVER":
                        inst_name = f"localhost\\{name}"
                        if inst_name not in instancias: instancias.append(inst_name)
                except OSError:
                    break
        except Exception:
            pass
        return instancias

    def on_instancia_changed(self, event=None):
        instancia = self.combo_instancia.get()
        global SQL_SERVER_INSTANCE, SERVIDOR
        SQL_SERVER_INSTANCE = instancia
        SERVIDOR = instancia
        
        try:
            c = configparser.ConfigParser()
            c.read(CAMINHO_DO_INI)
            if not c.has_section(INI_SECTION): c.add_section(INI_SECTION)
            c.set(INI_SECTION, INI_SERVER_KEY, instancia)
            with open(CAMINHO_DO_INI, 'w') as f: c.write(f)
        except Exception:
            pass

        self.after(0, lambda: self.status.config(text=f"Instância alterada para: {instancia}. A recarregar bancos..."))
        threading.Thread(target=self.carregar_banco_atual_sql, daemon=True).start()

    # --- ABA 4: CONFIGURAÇÕES ---
    def setup_config(self, parent):
        scroll_canvas = bstrap.Canvas(parent, highlightthickness=0)
        scrollbar = bstrap.Scrollbar(parent, orient="vertical", command=scroll_canvas.yview)
        scroll_frame = bstrap.Frame(scroll_canvas)
        scroll_frame.bind("<Configure>", lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))
        scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_canvas.configure(yscrollcommand=scrollbar.set)
        scroll_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)
        bstrap.Label(scroll_frame, text="Configurações do Sistema", font=("bold", 14), bootstyle="danger").pack(pady=10)
        
        self.cfg_vars = {}
        def add_section(title, keys, ini_section):
            f = bstrap.Labelframe(scroll_frame, text=f" {title} ", bootstyle="danger", padding=10)
            f.pack(fill=X, pady=5, padx=5)
            c = configparser.ConfigParser()
            c.read(CONFIG_FILE_NAME, encoding='utf-8')
            for key in keys:
                row = bstrap.Frame(f)
                row.pack(fill=X, pady=2)
                bstrap.Label(row, text=key, width=25, bootstyle="secondary").pack(side=LEFT)
                var = tk.StringVar(value=c.get(ini_section, key, fallback=""))
                self.cfg_vars[(ini_section, key)] = var
                bstrap.Entry(row, textvariable=var, bootstyle="danger").pack(side=LEFT, fill=X, expand=YES)
                
        add_section("Caminhos", ["PASTA_DO_SISTEMA", "PASTA_DAS_VERSOES", "PASTA_DE_BACKUP", "CAMINHO_DO_INI", "CAMINHO_DO_7ZIP_EXE"], "CAMINHOS")
        add_section("Executáveis", ["NOME_EXE_CLIENTE", "NOME_EXE_ATUALIZADOR"], "EXECUTAVEIS")
        add_section("SQL Laudo", ["SQL_DRIVER_LISTA", "SQL_SERVER_INSTANCE"], "SQL_LAUDO")
        add_section("SQL Restore", ["SERVIDOR", "USUARIO", "SENHA", "ODBC_DRIVER_RESTORE"], "SQL_RESTORE")
        add_section("Config INI MAX", ["INI_SECTION", "INI_KEY", "INI_SERVER_KEY"], "CONFIG_INI_MAX")
        add_section("Cloud Nuvem", ["URL_CLOUD", "USUARIO_CLOUD", "SENHA_CLOUD"], "CLOUD")
        bstrap.Button(scroll_frame, text="💾 Guardar Configurações", command=self.salvar_config_aba, bootstyle="success").pack(pady=20)

    def salvar_config_aba(self):
        try:
            c = configparser.ConfigParser()
            c.read(CONFIG_FILE_NAME, encoding='utf-8')
            for (section, key), var in self.cfg_vars.items():
                if not c.has_section(section): c.add_section(section)
                c.set(section, key, var.get())
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f: c.write(f)
            carregar_ou_criar_configuracoes()
            validar_caminhos()
            self.caminho_backup = CAMINHO_BASE_MAX_BACKUP
            messagebox.showinfo("Sucesso", "Configurações guardadas!")
            threading.Thread(target=self.carregamento_assincrono, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar: {e}")

    # =========================================================================
    # LÓGICA E FILTROS DE INTERFACE
    # =========================================================================
    def filtrar_versoes(self, *args):
        busca = self.var_busca_versao.get().lower()
        for i in self.lb_versoes.get_children(): self.lb_versoes.delete(i)
        for v in getattr(self, '_all_versoes', []):
            if busca in v.lower(): self.lb_versoes.insert("", END, values=(v,))

    def filtrar_backups(self, *args):
        busca = self.var_busca_backup.get().lower()
        for i in self.lb_backups.get_children(): self.lb_backups.delete(i)
        for b in getattr(self, '_all_backups', []):
            if busca in b.lower(): self.lb_backups.insert("", tk.END, values=(b,))

    def filtrar_bancos(self, *args):
        busca = self.var_busca_banco.get().lower()
        for i in self.lb_tools.get_children(): self.lb_tools.delete(i)
        for d in getattr(self, '_all_dbs', []):
            if busca in d.lower(): self.lb_tools.insert("", tk.END, values=(d,))

    def atualizar_ui_sql(self, bancos, db_atual, versao):
        self.lbl_db_atual.config(text=db_atual)
        self.lbl_versao_sql.config(text=versao)
        self.combo_db['values'] = bancos
        if db_atual in bancos: self.combo_db.set(db_atual)
        self._all_dbs = bancos
        self.filtrar_bancos()

    def get_versao(self, db):
        if not db or "ERRO" in db or "Nenhum" in db: return "---"
        conn_str = f'DRIVER={SQL_DRIVER_LISTA};SERVER={SQL_SERVER_INSTANCE};DATABASE={db};Trusted_Connection=yes;'
        try:
            with pyodbc.connect(conn_str, timeout=1) as conn:
                cursor = conn.cursor()
                cursor.execute(SQL_QUERY_VERSAO)
                row = cursor.fetchone()
                return str(row[0]) if row else "N/A"
        except: return "---"

    def carregar_banco_atual_sql(self):
        try:
            c = configparser.ConfigParser()
            c.read(CAMINHO_DO_INI)
            atual = c.get(INI_SECTION, INI_KEY)
            server_atual = c.get(INI_SECTION, INI_SERVER_KEY)
        except: 
            atual = "ERRO LER INI"
            server_atual = None
            
        if server_atual:
            global SQL_SERVER_INSTANCE, SERVIDOR
            SQL_SERVER_INSTANCE = server_atual
            SERVIDOR = server_atual
            if hasattr(self, 'combo_instancia'):
                vals = list(self.combo_instancia.cget('values')) if self.combo_instancia.cget('values') else self.listar_instancias_sql()
                if server_atual not in vals: vals.append(server_atual)
                self.after(0, lambda v=vals: self.combo_instancia.config(values=v))
                self.after(0, lambda s=server_atual: self.combo_instancia.set(s))
        else:
            if hasattr(self, 'combo_instancia') and not self.combo_instancia.get():
                vals = self.listar_instancias_sql()
                self.after(0, lambda v=vals: self.combo_instancia.config(values=v))
                if vals: self.after(0, lambda: self.combo_instancia.set(vals[0]))
        
        versao = self.get_versao(atual)
        bancos = self.listar_sql_dbs()
        self.after(0, lambda: self.atualizar_ui_sql(bancos, atual, versao))

    def preview_version(self, e):
        db = self.combo_db.get()
        threading.Thread(target=lambda: self.after(0, lambda v=self.get_versao(db): self.status.config(text=f"Banco selecionado: {db} (Versão: {v})")), daemon=True).start()

    def listar_sql_dbs(self):
        conn_str = f'DRIVER={SQL_DRIVER_LISTA};SERVER={SQL_SERVER_INSTANCE};DATABASE=master;Trusted_Connection=yes;'
        try:
            lst = []
            with pyodbc.connect(conn_str, timeout=2) as conn:
                res = conn.cursor().execute("SELECT name FROM sys.databases WHERE name NOT IN ('master','tempdb','model','msdb')")
                for r in res: lst.append(r.name)
            return sorted(lst)
        except: return []

    def mudar_banco(self):
        novo = self.combo_db.get()
        if not novo: return
        try:
            c = configparser.ConfigParser()
            c.read(CAMINHO_DO_INI)
            if not c.has_section(INI_SECTION): c.add_section(INI_SECTION)
            c.set(INI_SECTION, INI_KEY, novo)
            if hasattr(self, 'combo_instancia'):
                inst = self.combo_instancia.get()
                if inst: c.set(INI_SECTION, INI_SERVER_KEY, inst)
            with open(CAMINHO_DO_INI, 'w') as f: c.write(f)
            messagebox.showinfo("Sucesso", f"Alterado para: {novo}")
            threading.Thread(target=self.carregar_banco_atual_sql, daemon=True).start()
        except Exception as e: messagebox.showerror("Erro", f"{e}")

    def popular_versoes(self):
        try:
            if os.path.exists(PASTA_DAS_VERSOES):
                versoes = [e.name for e in os.scandir(PASTA_DAS_VERSOES) if e.is_file() and e.name.lower().endswith('.rar')]
                self._all_versoes = sorted(versoes, reverse=True)
                self.after(0, self.filtrar_versoes)
        except: pass

    def voltar_pasta_nuvem(self):
        global CAMINHO_NUVEM_ATUAL
        if not CAMINHO_NUVEM_ATUAL or CAMINHO_NUVEM_ATUAL == '/': return
        partes = [p for p in CAMINHO_NUVEM_ATUAL.strip('/').split('/') if p]
        if len(partes) > 0:
            partes.pop()
        CAMINHO_NUVEM_ATUAL = '/' + '/'.join(partes) + ('/' if partes else '')
        threading.Thread(target=self.popular_versoes_nuvem, daemon=True).start()

    def on_nuvem_double_click(self, event):
        sel = self.lb_nuvem.selection()
        if not sel: return
        item = self.lb_nuvem.item(sel[0], "values")
        tipo, nome = item[0], item[1]
        
        if tipo == "📁":
            global CAMINHO_NUVEM_ATUAL
            if not CAMINHO_NUVEM_ATUAL.endswith('/'): CAMINHO_NUVEM_ATUAL += '/'
            CAMINHO_NUVEM_ATUAL += nome + '/'
            threading.Thread(target=self.popular_versoes_nuvem, daemon=True).start()

    def popular_versoes_nuvem(self):
        self.after(0, lambda: self.status.config(text="A conectar à Nuvem Maxdata via WebDAV..."))
        try:
            url = URL_CLOUD.rstrip('/')
            if not url.startswith('http'): url = 'https://' + url
            
            global CAMINHO_NUVEM_ATUAL
            if not CAMINHO_NUVEM_ATUAL: CAMINHO_NUVEM_ATUAL = '/'
            
            self.after(0, lambda: self.lbl_caminho_nuvem.config(text=CAMINHO_NUVEM_ATUAL))

            # Caminho base do WebDAV Nextcloud, ex: https://cloud.maxdata.com.br/remote.php/webdav
            webdav_url = url + "/remote.php/webdav" + urllib.parse.quote(CAMINHO_NUVEM_ATUAL)
            
            req = urllib.request.Request(webdav_url, method='PROPFIND')
            req.add_header("Depth", "1")
            if USUARIO_CLOUD and SENHA_CLOUD:
                auth = base64.b64encode(f"{USUARIO_CLOUD}:{SENHA_CLOUD}".encode('utf-8')).decode('ascii')
                req.add_header("Authorization", f"Basic {auth}")
            
            with urllib.request.urlopen(req, timeout=10) as response:
                xml_data = response.read()

            root = ET.fromstring(xml_data)
            pastas = []
            arquivos = []
            
            # Namespace do WebDAV é DAV:
            namespaces = {'d': 'DAV:'}
            
            # O primeiro response é a própria pasta, pulamos ele comparando o href
            primeiro = True
            for resp in root.findall('d:response', namespaces):
                href = resp.find('d:href', namespaces).text
                href = urllib.parse.unquote(href)
                
                # Pega o nome do item a partir da rota
                nome_item = [p for p in href.split('/') if p][-1]
                
                if primeiro:
                    primeiro = False
                    continue
                
                propstat = resp.find('d:propstat', namespaces)
                if propstat:
                    prop = propstat.find('d:prop', namespaces)
                    resourcetype = prop.find('d:resourcetype', namespaces)
                    if resourcetype is not None and resourcetype.find('d:collection', namespaces) is not None:
                        pastas.append(nome_item)
                    else:
                        if nome_item.lower().endswith('.rar'):
                            arquivos.append(nome_item)
            
            pastas.sort()
            arquivos.sort(reverse=True)
            
            def att_ui():
                for i in self.lb_nuvem.get_children(): self.lb_nuvem.delete(i)
                for p in pastas: self.lb_nuvem.insert("", END, values=("📁", p))
                for a in arquivos: self.lb_nuvem.insert("", END, values=("📄", a))
                self.status.config(text=f"Pronto. {len(pastas)} pastas, {len(arquivos)} arquivos úteis.")
            self.after(0, att_ui)
        except Exception as e:
            self.after(0, lambda e=e: self.status.config(text=f"Erro WebDAV: {str(e)[:50]}"))

    def baixar_da_nuvem(self):
        sel = self.lb_nuvem.selection()
        if not sel: return
        item = self.lb_nuvem.item(sel[0], "values")
        tipo, nome = item[0], item[1]
        
        if tipo == "📁":
            messagebox.showinfo("Aviso", "Você precisa selecionar um ARQUIVO (📄) para baixar, e não uma pasta.")
            return

        if messagebox.askyesno("Baixar da Nuvem", f"Deseja baixar {nome} para o computador?"):
            self.status.config(text=f"A iniciar download de {nome}...")
            threading.Thread(target=self.thread_download, args=(nome,), daemon=True).start()

    def thread_download(self, arq):
        try:
            url = URL_CLOUD.rstrip('/')
            if not url.startswith('http'): url = 'https://' + url
            
            global CAMINHO_NUVEM_ATUAL
            webdav_url = url + "/remote.php/webdav" + urllib.parse.quote(CAMINHO_NUVEM_ATUAL + arq)
            
            req = urllib.request.Request(webdav_url)
            if USUARIO_CLOUD and SENHA_CLOUD:
                auth = base64.b64encode(f"{USUARIO_CLOUD}:{SENHA_CLOUD}".encode('utf-8')).decode('ascii')
                req.add_header("Authorization", f"Basic {auth}")
            
            caminho_local = os.path.join(PASTA_DAS_VERSOES, arq)
            with urllib.request.urlopen(req, timeout=15) as response, open(caminho_local, 'wb') as out_file:
                tamanho_total = response.getheader('content-length')
                tamanho_total = int(tamanho_total) if tamanho_total else None
                baixado = 0
                bloco = 1024 * 64 # 64KB
                while True:
                    dados = response.read(bloco)
                    if not dados: break
                    out_file.write(dados)
                    baixado += len(dados)
                    if tamanho_total:
                        pct = int((baixado / tamanho_total) * 100)
                        self.after(0, lambda p=pct, a=arq: self.status.config(text=f"A baixar {a}... {p}%"))
            
            self.after(0, lambda: self.status.config(text=f"Download de {arq} concluído!"))
            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Download concluído!\n{arq} salvo em Versões Locais."))
            self.after(0, self.popular_versoes) # FIX UI Refresh
        except Exception as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Download", f"Falha ao baixar:\n{m}"))
            self.after(0, lambda: self.status.config(text="Erro no download."))

    def voltar_pasta_nuvem_bkp(self):
        global CAMINHO_NUVEM_BACKUP
        if not CAMINHO_NUVEM_BACKUP or CAMINHO_NUVEM_BACKUP == '/': return
        partes = [p for p in CAMINHO_NUVEM_BACKUP.strip('/').split('/') if p]
        if len(partes) > 0: partes.pop()
        CAMINHO_NUVEM_BACKUP = '/' + '/'.join(partes) + ('/' if partes else '')
        threading.Thread(target=self.popular_nuvem_backups, daemon=True).start()

    def on_nuvem_bkp_double_click(self, event):
        sel = self.lb_nuvem_bkp.selection()
        if not sel: return
        item = self.lb_nuvem_bkp.item(sel[0], "values")
        tipo, nome = item[0], item[1]
        
        if tipo == "📁":
            global CAMINHO_NUVEM_BACKUP
            if not CAMINHO_NUVEM_BACKUP.endswith('/'): CAMINHO_NUVEM_BACKUP += '/'
            CAMINHO_NUVEM_BACKUP += nome + '/'
            threading.Thread(target=self.popular_nuvem_backups, daemon=True).start()

    def popular_nuvem_backups(self):
        self.after(0, lambda: self.status.config(text="A conectar à Nuvem Maxdata via WebDAV (Backups)..."))
        try:
            url = URL_CLOUD.rstrip('/')
            if not url.startswith('http'): url = 'https://' + url
            
            global CAMINHO_NUVEM_BACKUP
            if not CAMINHO_NUVEM_BACKUP: CAMINHO_NUVEM_BACKUP = '/'
            
            self.after(0, lambda: self.lbl_caminho_nuvem_bkp.config(text=CAMINHO_NUVEM_BACKUP))

            webdav_url = url + "/remote.php/webdav" + urllib.parse.quote(CAMINHO_NUVEM_BACKUP)
            
            req = urllib.request.Request(webdav_url, method='PROPFIND')
            req.add_header("Depth", "1")
            if USUARIO_CLOUD and SENHA_CLOUD:
                auth = base64.b64encode(f"{USUARIO_CLOUD}:{SENHA_CLOUD}".encode('utf-8')).decode('ascii')
                req.add_header("Authorization", f"Basic {auth}")
            
            with urllib.request.urlopen(req, timeout=10) as response:
                xml_data = response.read()

            root = ET.fromstring(xml_data)
            pastas = []
            arquivos = []
            namespaces = {'d': 'DAV:'}
            primeiro = True
            for resp in root.findall('d:response', namespaces):
                href = resp.find('d:href', namespaces).text
                href = urllib.parse.unquote(href)
                nome_item = [p for p in href.split('/') if p][-1]
                if primeiro:
                    primeiro = False
                    continue
                propstat = resp.find('d:propstat', namespaces)
                if propstat:
                    prop = propstat.find('d:prop', namespaces)
                    resourcetype = prop.find('d:resourcetype', namespaces)
                    if resourcetype is not None and resourcetype.find('d:collection', namespaces) is not None:
                        pastas.append(nome_item)
                    else:
                        if nome_item.lower().endswith(('.bak', '.zip', '.rar')):
                            arquivos.append(nome_item)
            
            pastas.sort()
            arquivos.sort(reverse=True)
            
            def att_ui():
                for i in self.lb_nuvem_bkp.get_children(): self.lb_nuvem_bkp.delete(i)
                for p in pastas: self.lb_nuvem_bkp.insert("", END, values=("📁", p))
                for a in arquivos: self.lb_nuvem_bkp.insert("", END, values=("📄", a))
                self.status.config(text=f"Pronto. {len(pastas)} pastas, {len(arquivos)} backups.")
            self.after(0, att_ui)
        except Exception as e:
            self.after(0, lambda e=e: self.status.config(text=f"Erro WebDAV: {str(e)[:50]}"))

    def baixar_da_nuvem_bkp(self):
        sel = self.lb_nuvem_bkp.selection()
        if not sel: return
        item = self.lb_nuvem_bkp.item(sel[0], "values")
        tipo, nome = item[0], item[1]
        
        if tipo == "📁":
            messagebox.showinfo("Aviso", "Selecione um arquivo de backup (📄) para baixar.")
            return

        if messagebox.askyesno("Baixar Backup", f"Deseja baixar o backup {nome} para o computador?"):
            self.status.config(text=f"A iniciar download de {nome}...")
            threading.Thread(target=self.thread_download_bkp, args=(nome,), daemon=True).start()

    def thread_download_bkp(self, arq):
        try:
            url = URL_CLOUD.rstrip('/')
            if not url.startswith('http'): url = 'https://' + url
            
            global CAMINHO_NUVEM_BACKUP
            webdav_url = url + "/remote.php/webdav" + urllib.parse.quote(CAMINHO_NUVEM_BACKUP + arq)
            
            req = urllib.request.Request(webdav_url)
            if USUARIO_CLOUD and SENHA_CLOUD:
                auth = base64.b64encode(f"{USUARIO_CLOUD}:{SENHA_CLOUD}".encode('utf-8')).decode('ascii')
                req.add_header("Authorization", f"Basic {auth}")
            
            caminho_local = os.path.join(CAMINHO_BASE_MAX_BACKUP, arq)
            with urllib.request.urlopen(req, timeout=15) as response, open(caminho_local, 'wb') as out_file:
                tamanho_total = response.getheader('content-length')
                tamanho_total = int(tamanho_total) if tamanho_total else None
                baixado = 0
                bloco = 1024 * 64 # 64KB
                while True:
                    dados = response.read(bloco)
                    if not dados: break
                    out_file.write(dados)
                    baixado += len(dados)
                    if tamanho_total:
                        pct = int((baixado / tamanho_total) * 100)
                        self.after(0, lambda p=pct, a=arq: self.status.config(text=f"A baixar {a}... {p}%"))
            
            self.after(0, lambda: self.status.config(text=f"Download de {arq} concluído!"))
            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Backup baixado com sucesso!\n{arq} salvo nos Backups Locais."))
            self.after(0, self.load_backups)
        except Exception as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Download", f"Falha ao baixar backup:\n{m}"))
            self.after(0, lambda: self.status.config(text="Erro no download do backup."))

    def load_backups(self):
        try:
            backups = []
            if os.path.exists(self.caminho_backup):
                for e in os.scandir(self.caminho_backup):
                    if e.is_file() and e.name.upper().endswith((".MAX", ".BAK", ".ZIP", ".RAR")):
                        backups.append((e.name, e.stat().st_mtime))
                backups.sort(key=lambda x: x[1], reverse=True)
                self._all_backups = [b[0] for b in backups]
                self.after(0, self.filtrar_backups)
        except: pass

    def get_exe_versao(self, caminho_exe):
        try:
            import win32api
            info = win32api.GetFileVersionInfo(caminho_exe, "\\")
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            import win32api as wa
            return f"{wa.HIWORD(ms)}.{wa.LOWORD(ms)}.{wa.HIWORD(ls)}.{wa.LOWORD(ls)}"
        except:
            return None

    def lancar_erp(self):
        try:
            db_atual = self.lbl_db_atual.cget("text")
            db_versao = self.lbl_versao_sql.cget("text")
            exe_versao = self.get_exe_versao(CAMINHO_DO_ERP_CLIENTE)

            if db_versao and exe_versao and db_versao != "---" and db_versao != "N/A" and exe_versao != db_versao:
                arquivo_rar = next((v for v in getattr(self, '_all_versoes', []) if db_versao in v), None)
                if arquivo_rar:
                    if messagebox.askyesno("Atualização Necessária", 
                        f"Versão Base de Dados: {db_versao} | Versão Executável: {exe_versao}.\n\nAtualizar para '{arquivo_rar}'?"):
                        threading.Thread(target=self.thread_extrair, args=(arquivo_rar,), daemon=True).start()
                        return
                else:
                    messagebox.showwarning("Aviso de Versão", "As versões da base de dados e do executável são diferentes e o ficheiro de extração não foi encontrado.\nA ignorar...")
            
            subprocess.Popen([CAMINHO_DO_ERP_CLIENTE], cwd=PASTA_DO_SISTEMA)
        except Exception as e: messagebox.showerror("Erro", f"Erro: {e}")

    def lancar_atualizacao(self):
        sel = self.lb_versoes.selection()
        if not sel: return
        arq = self.lb_versoes.item(sel[0], "values")[0]
        if messagebox.askyesno("Confirmar", f"Atualizar para {arq}?"):
            threading.Thread(target=self.thread_extrair, args=(arq,), daemon=True).start()

    def thread_extrair(self, arq):
        try:
            self.after(0, lambda: self.status.config(text="A extrair..."))
            cmd = [CAMINHO_DO_7ZIP_EXE, 'x', os.path.join(PASTA_DAS_VERSOES, arq), f'-o{PASTA_DO_SISTEMA}', '-y']
            subprocess.run(cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.after(0, self.lancar_atualizador_callback)
        except Exception as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Erro Extração", m))

    def lancar_atualizador_callback(self):
        try:
            subprocess.Popen([CAMINHO_DO_MAX_ATUALIZA], cwd=PASTA_DO_SISTEMA)
            self.status.config(text="Atualizador Aberto.")
        except Exception as e: messagebox.showerror("Erro", f"{e}")

    def iniciar_restore_thread(self):
        try: 
            sel = self.lb_backups.selection()[0]
            fname = self.lb_backups.item(sel, "values")[0]
        except: return
        new_db = self.entry_new_db.get().strip()
        if not new_db: return

        self.btn_restore.config(state='disabled')
        self.progress.start()
        self.log_txt.config(state='normal'); self.log_txt.delete('1.0', END); self.log_txt.config(state='disabled')
        threading.Thread(target=self.restore_logic, args=(fname, new_db), daemon=True).start()

    def restore_logic(self, fname, dbname):
        tmp_dir = None
        try:
            self.msg_queue.put(f"--- Iniciando Restore: {dbname} ---")
            origem = os.path.join(self.caminho_backup, fname)
            final = origem

            if fname.lower().endswith(('.zip', '.rar')):
                self.msg_queue.put("A extrair ficheiro...")
                tmp_dir = os.path.join(self.caminho_backup, f"_tmp_{int(time.time())}")
                os.makedirs(tmp_dir, exist_ok=True)
                cmd = [CAMINHO_DO_7ZIP_EXE, 'x', origem, f'-o{tmp_dir}', '-y']
                subprocess.run(cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                encontrado = next((os.path.join(root, f) for root, _, files in os.walk(tmp_dir) for f in files if f.upper().endswith(('.MAX', '.BAK'))), None)
                if not encontrado: raise Exception("Ficheiro .MAX/.BAK não encontrado.")
                final = encontrado
                self.msg_queue.put(f"Encontrado: {os.path.basename(final)}")

            i = 1
            while True:
                d = os.path.join(PASTA_DO_SISTEMA, f"dados{i}")
                if not os.path.exists(d): os.makedirs(d); break
                i += 1
            mdf = os.path.join(d, f"{dbname}.mdf")
            ldf = os.path.join(d, f"{dbname}_log.ldf")

            conn_str = f"DRIVER={ODBC_DRIVER_RESTORE};SERVER={SERVIDOR};UID={USUARIO};PWD={SENHA};DATABASE=master"
            conn = pyodbc.connect(conn_str, autocommit=True)
            cur = conn.cursor()
            cur.execute("RESTORE FILELISTONLY FROM DISK = ?", final)
            ld, ll = None, None
            for r in cur.fetchall():
                if r.Type == 'D': ld = r.LogicalName
                if r.Type == 'L': ll = r.LogicalName
            
            self.msg_queue.put("A restaurar SQL...")
            sql = f"RESTORE DATABASE [{dbname}] FROM DISK='{final}' WITH MOVE '{ld}' TO '{mdf}', MOVE '{ll}' TO '{ldf}', REPLACE"
            cur.execute(sql)
            while cur.nextset(): pass
            conn.close()
            self.msg_queue.put("__DONE__")
        except Exception as e:
            self.msg_queue.put(f"ERRO: {e}")
            self.msg_queue.put("__ERROR__")
        finally:
            if tmp_dir and os.path.exists(tmp_dir):
                try: shutil.rmtree(tmp_dir)
                except: pass

    def process_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg == "__DONE__":
                    self.progress.stop()
                    self.btn_restore.config(state='normal')
                    messagebox.showinfo("Sucesso", "Restore Concluído!")
                    threading.Thread(target=self.carregar_banco_atual_sql, daemon=True).start()
                elif msg == "__ERROR__":
                    self.progress.stop()
                    self.btn_restore.config(state='normal')
                    messagebox.showerror("Erro", "Falhou. Veja o log.")
                else:
                    self.log_txt.config(state='normal')
                    self.log_txt.insert(END, msg + "\n")
                    self.log_txt.see(END)
                    self.log_txt.config(state='disabled')
        except: pass
        self.after(100, self.process_queue)

    def drop_database(self):
        sel = self.lb_tools.selection()
        if not sel: return
        
        bancos_alvo = [self.lb_tools.item(i, "values")[0] for i in sel]
        db_atual = self.lbl_db_atual.cget("text")
        
        if db_atual in bancos_alvo:
            bancos_alvo.remove(db_atual)
            messagebox.showwarning("Proteção", f"O banco atual ({db_atual}) não pode ser eliminado. Foi removido da sua lista de exclusão.")
            
        if not bancos_alvo: return

        if len(bancos_alvo) == 1:
            msg = f"ELIMINAR '{bancos_alvo[0]}' permanentemente?\n\nEscreva EXCLUIR para confirmar:"
        else:
            lista_preview = ", ".join(bancos_alvo[:5]) + ("..." if len(bancos_alvo) > 5 else "")
            msg = f"ELIMINAR {len(bancos_alvo)} bases de dados permanentemente?\n\n({lista_preview})\n\nEscreva EXCLUIR para confirmar:"

        confirmacao = simpledialog.askstring("PERIGO", msg)
        
        if confirmacao == "EXCLUIR":
            threading.Thread(target=self.thread_drop_bancos, args=(bancos_alvo,), daemon=True).start()
        elif confirmacao is not None:
            messagebox.showwarning("Cancelado", "Palavra de confirmação inválida.")

    def thread_drop_bancos(self, bancos):
        try:
            self.after(0, lambda: self.status.config(text=f"A eliminar {len(bancos)} banco(s)..."))
            conn_str = f"DRIVER={ODBC_DRIVER_RESTORE};SERVER={SERVIDOR};UID={USUARIO};PWD={SENHA};DATABASE=master"
            with pyodbc.connect(conn_str, autocommit=True) as conn:
                cursor = conn.cursor()
                for db in bancos:
                    cursor.execute(f"ALTER DATABASE [{db}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE")
                    cursor.execute(f"DROP DATABASE [{db}]")
            
            self.after(0, lambda: messagebox.showinfo("Sucesso", "Bases de dados eliminadas com sucesso!"))
            self.carregar_banco_atual_sql()
            self.after(0, lambda: self.status.config(text="Pronto."))
        except Exception as e:
            self.after(0, lambda err=str(e): messagebox.showerror("Erro", err))

def validar_caminhos():
    global CAMINHO_DO_ERP_CLIENTE, CAMINHO_DO_MAX_ATUALIZA
    CAMINHO_DO_ERP_CLIENTE = os.path.join(PASTA_DO_SISTEMA, NOME_EXE_CLIENTE)
    CAMINHO_DO_MAX_ATUALIZA = os.path.join(PASTA_DO_SISTEMA, NOME_EXE_ATUALIZADOR)
    
    erros = []
    if not os.path.isdir(PASTA_DO_SISTEMA): erros.append(f"Pasta Sistema: {PASTA_DO_SISTEMA}")
    if not os.path.isdir(PASTA_DAS_VERSOES): erros.append(f"Pasta Versões: {PASTA_DAS_VERSOES}")
    if not os.path.isdir(CAMINHO_BASE_MAX_BACKUP): erros.append(f"Pasta Backups: {CAMINHO_BASE_MAX_BACKUP}")
    if not os.path.exists(CAMINHO_DO_INI): erros.append(f"Ficheiro INI: {CAMINHO_DO_INI}")
    if not os.path.exists(CAMINHO_DO_7ZIP_EXE): erros.append(f"7-Zip: {CAMINHO_DO_7ZIP_EXE}")
    return erros

if __name__ == "__main__":
    carregar_ou_criar_configuracoes()
    app = GerenciadorMaxApp()
    
    # SETUP INICIAL: Abre a tela de configuração imediatamente na primeira execução
    if PRIMEIRA_EXECUCAO:
        cw = ConfigWindow(app, is_first_run=True)
        app.wait_window(cw)
        if not cw.salvo: 
            sys.exit()
        carregar_ou_criar_configuracoes() 
        
    # VERIFICAÇÃO DE ERROS
    while True:
        erros = validar_caminhos()
        if not erros: break
        messagebox.showerror("Configuração Necessária", "Caminhos inválidos:\n" + "\n".join(erros), parent=app)
        cw = ConfigWindow(app)
        app.wait_window(cw)
        if not cw.salvo: sys.exit()
        carregar_ou_criar_configuracoes() 
        
    app.iniciar_interface()
    app.mainloop()