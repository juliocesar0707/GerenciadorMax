import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
from tkinter import filedialog
import ttkbootstrap as bstrap
from ttkbootstrap.constants import *
import os
import pyodbc
import threading
import queue
import sys
import subprocess
import configparser
import time

try:
    from pywinauto.application import Application
except ImportError:
    print("Erro: Biblioteca 'pywinauto' não encontrada.")
    print("Por favor, instale com: pip install pywinauto")
    sys.exit()

# --- (CONFIGURAÇÃO GERAL) ---
CONFIG_FILE_NAME = 'gerenciador_config.ini'

# --- Variáveis Globais de Configuração (serão preenchidas ao iniciar) ---
PASTA_DO_SISTEMA = None
NOME_EXE_CLIENTE = None
NOME_EXE_ATUALIZADOR = None
PASTA_DAS_VERSOES = None
CAMINHO_DO_INI = None
CAMINHO_DO_7ZIP_EXE = None
INI_SECTION = None
INI_KEY = None
SERVIDOR = None
USUARIO = None
SENHA = None
ODBC_DRIVER_RESTORE = None
SQL_DRIVER_LISTA = None
SQL_SERVER_INSTANCE = None

# --- Caminhos derivados (serão preenchidos ao iniciar) ---
CAMINHO_DO_ERP_CLIENTE = None
CAMINHO_DO_MAX_ATUALIZA = None
CAMINHO_BASE_MAX_BACKUP = None
# --- FIM DAS CONFIGURAÇÕES ---

def carregar_ou_criar_configuracoes():
    """
    Lê o 'gerenciador_config.ini'. Se não existir, cria um com valores padrão.
    Popula as variáveis de configuração globais.
    """
    global PASTA_DO_SISTEMA, NOME_EXE_CLIENTE, NOME_EXE_ATUALIZADOR, PASTA_DAS_VERSOES
    global CAMINHO_DO_INI, CAMINHO_DO_7ZIP_EXE, INI_SECTION, INI_KEY, SERVIDOR
    global USUARIO, SENHA, ODBC_DRIVER_RESTORE, SQL_DRIVER_LISTA, SQL_SERVER_INSTANCE
    
    config = configparser.ConfigParser()
    
    if not os.path.exists(CONFIG_FILE_NAME):
        # --- Se o .ini não existe, cria um com os valores padrão ---
        print(f"DEBUG: Arquivo {CONFIG_FILE_NAME} não encontrado. Criando...")
        config['CAMINHOS'] = {
            '; --- ATENÇÃO: Use barras normais (/) ou barras invertidas duplicadas (\\\\) ---': '',
            '; 1. A pasta principal onde o sistema está instalado': '',
            'PASTA_DO_SISTEMA': r'D:\Max',
            '; 2. A pasta onde você guarda os arquivos .RAR das versões': '',
            'PASTA_DAS_VERSOES': r'D:\Max\Versões',
            '; 3. Caminho para o arquivo .ini de configuração do Max': '',
            'CAMINHO_DO_INI': r'D:\Max\max.ini',
            '; 4. Caminho para o executável do 7-Zip': '',
            'CAMINHO_DO_7ZIP_EXE': r'C:\Program Files (x86)\7-Zip\7z.exe'
        }
        config['EXECUTAVEIS'] = {
            '; Nomes dos arquivos dentro da PASTA_DO_SISTEMA': '',
            'NOME_EXE_CLIENTE': 'MAX_manager2.exe',
            'NOME_EXE_ATUALIZADOR': 'MAX_Atualiza.exe'
        }
        config['CONFIG_INI_MAX'] = {
            "; Seção e Chave dentro do 'max.ini' que o launcher deve alterar": '',
            'INI_SECTION': 'CON',
            'INI_KEY': 'Initial catalog'
        }
        config['SQL_LAUDO'] = {
            '; Configuração para listar os bancos de dados no Launcher': '',
            'SQL_DRIVER_LISTA': '{ODBC Driver 17 for SQL Server}',
            'SQL_SERVER_INSTANCE': 'localhost'
        }
        config['SQL_RESTORE'] = {
            '; Credenciais para a ferramenta de Restore de Banco': '',
            'SERVIDOR': 'localhost',
            'USUARIO': 'sa',
            'SENHA': 'macro01',
            'ODBC_DRIVER_RESTORE': '{ODBC Driver 17 for SQL Server}'
        }
        try:
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Configuração Inicial", 
                                f"O arquivo de configuração '{CONFIG_FILE_NAME}' foi criado.\n\n"
                                "Por favor, verifique os caminhos e credenciais dentro dele "
                                "antes de executar o programa novamente.")
            root.destroy()
            print("DEBUG: Arquivo de configuração criado. Encerrando para o usuário configurar.")
            sys.exit()
            
        except Exception as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erro Crítico", f"Não foi possível criar o arquivo de configuração '{CONFIG_FILE_NAME}':\n{e}")
            root.destroy()
            print(f"DEBUG: Erro crítico ao criar config: {e}")
            sys.exit()
    
    # --- Se o .ini JÁ EXISTE, lê os valores ---
    try:
        print(f"DEBUG: Lendo {CONFIG_FILE_NAME}...")
        config.read(CONFIG_FILE_NAME, encoding='utf-8')
        
        # Lendo [CAMINHOS]
        PASTA_DO_SISTEMA = config.get('CAMINHOS', 'PASTA_DO_SISTEMA')
        PASTA_DAS_VERSOES = config.get('CAMINHOS', 'PASTA_DAS_VERSOES')
        CAMINHO_DO_INI = config.get('CAMINHOS', 'CAMINHO_DO_INI')
        CAMINHO_DO_7ZIP_EXE = config.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE')

        # Lendo [EXECUTAVEIS]
        NOME_EXE_CLIENTE = config.get('EXECUTAVEIS', 'NOME_EXE_CLIENTE')
        NOME_EXE_ATUALIZADOR = config.get('EXECUTAVEIS', 'NOME_EXE_ATUALIZADOR')

        # Lendo [CONFIG_INI_MAX]
        INI_SECTION = config.get('CONFIG_INI_MAX', 'INI_SECTION')
        INI_KEY = config.get('CONFIG_INI_MAX', 'INI_KEY')

        # Lendo [SQL_LAUDO]
        SQL_DRIVER_LISTA = config.get('SQL_LAUDO', 'SQL_DRIVER_LISTA')
        SQL_SERVER_INSTANCE = config.get('SQL_LAUDO', 'SQL_SERVER_INSTANCE')

        # Lendo [SQL_RESTORE]
        SERVIDOR = config.get('SQL_RESTORE', 'SERVIDOR')
        USUARIO = config.get('SQL_RESTORE', 'USUARIO')
        SENHA = config.get('SQL_RESTORE', 'SENHA')
        ODBC_DRIVER_RESTORE = config.get('SQL_RESTORE', 'ODBC_DRIVER_RESTORE')
        print("DEBUG: Configurações lidas com sucesso.")
        
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Erro ao Ler Configuração", 
                             f"Ocorreu um erro ao ler o '{CONFIG_FILE_NAME}':\n{e}\n\n"
                             f"Verifique se o arquivo não está corrompido ou faltando.\n"
                             f"Você pode deletar o arquivo para que ele seja recriado.")
        root.destroy()
        print(f"DEBUG: Erro ao ler config: {e}")
        sys.exit()


class ConfigWindow(tk.Toplevel):
    """
    Janela pop-up para o usuário corrigir os caminhos no .ini
    quando a validação falha.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Configuração de Caminhos")
        self.geometry("700x300")
        self.resizable(False, False)
        self.salvo = False # Flag para saber se o usuário salvou

        # Traz a janela para frente
        self.transient(parent)
        self.grab_set()

        # Variáveis para os campos de entrada
        self.var_pasta_sistema = tk.StringVar()
        self.var_pasta_versoes = tk.StringVar()
        self.var_arquivo_ini = tk.StringVar()
        self.var_7zip = tk.StringVar()

        # Carrega os valores atuais do .ini
        self.ler_ini_atual()
        
        # Cria os widgets
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=tk.YES)
        
        ttk.Label(main_frame, text="Por favor, corrija os caminhos inválidos:", font=("-weight bold", 12)).pack(pady=(0, 15))

        self.criar_campo(main_frame, "Pasta do Sistema (Ex: D:\\Max)", self.var_pasta_sistema, self.selecionar_pasta_sistema)
        self.criar_campo(main_frame, "Pasta de Versões (Ex: D:\\Max\\Versões)", self.var_pasta_versoes, self.selecionar_pasta_versoes)
        self.criar_campo(main_frame, "Arquivo .ini do Max (Ex: D:\\Max\\max.ini)", self.var_arquivo_ini, self.selecionar_arquivo_ini)
        self.criar_campo(main_frame, "Executável 7-Zip (7z.exe)", self.var_7zip, self.selecionar_arquivo_7zip)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.btn_salvar = bstrap.Button(btn_frame, text="Salvar e Tentar Novamente", command=self.salvar_e_fechar, bootstyle="success")
        self.btn_salvar.pack(side=tk.RIGHT, padx=5)
        
        self.btn_cancelar = bstrap.Button(btn_frame, text="Cancelar", command=self.destroy, bootstyle="secondary-outline")
        self.btn_cancelar.pack(side=tk.RIGHT)

        # Garante que a janela tenha o foco
        self.focus_force()

    def criar_campo(self, parent, label_text, var, cmd):
        """Helper para criar um Label, Entry e Button."""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        label = ttk.Label(frame, text=label_text, width=28, anchor="w")
        label.pack(side=tk.LEFT)
        
        entry = ttk.Entry(frame, textvariable=var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=tk.YES, padx=5)
        
        btn = bstrap.Button(frame, text="Selecionar...", command=cmd, bootstyle="info-outline", width=12)
        btn.pack(side=tk.LEFT)

    def ler_ini_atual(self):
        """Lê os valores do .ini para popular os campos."""
        try:
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_NAME, encoding='utf-8')
            self.var_pasta_sistema.set(config.get('CAMINHOS', 'PASTA_DO_SISTEMA', fallback=''))
            self.var_pasta_versoes.set(config.get('CAMINHOS', 'PASTA_DAS_VERSOES', fallback=''))
            self.var_arquivo_ini.set(config.get('CAMINHOS', 'CAMINHO_DO_INI', fallback=''))
            self.var_7zip.set(config.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', fallback=''))
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o {CONFIG_FILE_NAME}:\n{e}", parent=self)

    def selecionar_pasta_sistema(self):
        caminho = filedialog.askdirectory(title="Selecione a 'Pasta do Sistema' (Ex: D:\\Max)")
        if caminho:
            self.var_pasta_sistema.set(caminho)
            # Tenta preencher os outros campos automaticamente
            if not self.var_pasta_versoes.get() or "Versões" not in self.var_pasta_versoes.get():
                self.var_pasta_versoes.set(os.path.join(caminho, 'Versões'))
            if not self.var_arquivo_ini.get() or "max.ini" not in self.var_arquivo_ini.get():
                self.var_arquivo_ini.set(os.path.join(caminho, 'max.ini'))

    def selecionar_pasta_versoes(self):
        caminho = filedialog.askdirectory(title="Selecione a 'Pasta de Versões'")
        if caminho:
            self.var_pasta_versoes.set(caminho)
            
    def selecionar_arquivo_ini(self):
        caminho = filedialog.askopenfilename(title="Selecione o arquivo 'max.ini'", filetypes=[("Arquivo INI", "*.ini")])
        if caminho:
            self.var_arquivo_ini.set(caminho)

    def selecionar_arquivo_7zip(self):
        caminho = filedialog.askopenfilename(title="Selecione o executável '7z.exe'", filetypes=[("Executável", "*.exe")])
        if caminho:
            self.var_7zip.set(caminho)

    def salvar_e_fechar(self):
        """Salva os novos valores no .ini e fecha a janela."""
        try:
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_NAME, encoding='utf-8')
            
            # Atualiza apenas os valores de caminhos
            config.set('CAMINHOS', 'PASTA_DO_SISTEMA', self.var_pasta_sistema.get())
            config.set('CAMINHOS', 'PASTA_DAS_VERSOES', self.var_pasta_versoes.get())
            config.set('CAMINHOS', 'CAMINHO_DO_INI', self.var_arquivo_ini.get())
            config.set('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', self.var_7zip.get())

            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            
            self.salvo = True # Seta a flag de sucesso
            self.destroy() # Fecha a janela de configuração

        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o {CONFIG_FILE_NAME}:\n{e}", parent=self)


class GerenciadorMaxApp(bstrap.Window):
    """
    Aplicação principal que unifica o Launcher e o Assistente de Restore
    usando um sistema de Abas (Notebook).
    """
    def __init__(self):
        super().__init__(themename="litera")
        self.title("Gerenciador Max (Launcher e Restore)")
        self.geometry("800x700")

        # --- Variáveis de estado (de ambos os scripts) ---
        self.caminho_backup = CAMINHO_BASE_MAX_BACKUP # Definido no __main__
        self.msg_queue = queue.Queue()

        # 1. Criar os widgets principais (Abas e Statusbar)
        self.create_main_layout()

        # 2. Carregar dados iniciais (de ambos os scripts)
        self.load_backup_list()       # Do RestoreApp
        self.popular_lista_versoes()  # Do AppAtualizador
        self.carregar_banco_atual()   # Do AppAtualizador

        # 3. Iniciar o "ouvinte" da fila de restore
        self.process_queue()

    def create_main_layout(self):
        """Cria o layout principal com Abas e a Barra de Status."""
        
        main_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        main_frame.pack(fill=BOTH, expand=YES)
        
        self.notebook = bstrap.Notebook(main_frame, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=YES)

        tab_launcher = bstrap.Frame(self.notebook, padding=15)
        tab_restore = bstrap.Frame(self.notebook, padding=15)

        self.notebook.add(tab_launcher, text="  Launcher / Atualizador  ")
        self.notebook.add(tab_restore, text="  Assistente de Restore  ")

        self.create_launcher_tab(tab_launcher)
        self.create_restore_tab(tab_restore)

        self.status_bar = bstrap.Label(self, text="Pronto.", relief="sunken", anchor="w", padding=(10, 5),
                                       bootstyle="light")
        self.status_bar.pack(side="bottom", fill="x")

    def create_launcher_tab(self, parent_tab):
        """Cria todos os widgets da aba 'Launcher/Atualizador'."""
        
        app_title = bstrap.Label(parent_tab, text="Launcher Maxdata", font=("-weight bold", 18), bootstyle="primary")
        app_title.pack(pady=(0, 5))
        app_subtitle = bstrap.Label(parent_tab, text="Gerencie a versão e o banco de dados do sistema.", bootstyle="secondary")
        app_subtitle.pack(pady=(0, 15))

        self.criar_widgets_db(parent_tab)

        list_frame = bstrap.Labelframe(parent_tab, text="Atualização de Versão", bootstyle="info", padding=10)
        list_frame.pack(fill="both", expand=True, pady=(0, 15))

        self.scrollbar = bstrap.Scrollbar(list_frame, orient="vertical", bootstyle="primary-round")
        self.listbox_versoes = tk.Listbox(list_frame,
                                          yscrollcommand=self.scrollbar.set,
                                          selectmode="single",
                                          height=8,
                                          font=("", 10))
        self.scrollbar.config(command=self.listbox_versoes.yview)
        self.scrollbar.pack(side="right", fill="y", padx=(5, 0))
        self.listbox_versoes.pack(side="left", fill="both", expand=True)

        btn_frame = bstrap.Frame(parent_tab)
        btn_frame.pack(fill="x")

        self.btn_recarregar = bstrap.Button(btn_frame, text="Recarregar Lista", command=self.popular_lista_versoes,
                                            bootstyle="secondary-outline")
        self.btn_recarregar.pack(side="left")

        self.btn_executar_erp = bstrap.Button(btn_frame, text="Executar Sistema", command=self.iniciar_cliente_erp,
                                              bootstyle="primary")
        self.btn_executar_erp.pack(side="right", padx=(5, 0))

        self.btn_atualizar = bstrap.Button(btn_frame, text="Atualizar Versão", command=self.iniciar_atualizacao,
                                           bootstyle="success")
        self.btn_atualizar.pack(side="right")

    def create_restore_tab(self, parent_tab):
        """Cria todos os widgets da aba 'Assistente de Restore'."""
        
        frame1 = ttk.Labelframe(parent_tab, text="1. Escolha o Backup", padding="10")
        frame1.pack(fill=X, expand=NO, pady=5)

        self.lbl_path = ttk.Label(frame1, text=f"Buscando em: {self.caminho_backup}")
        self.lbl_path.pack(fill=X, expand=YES)

        list_frame = ttk.Frame(frame1)
        list_frame.pack(fill=X, expand=YES, pady=5)
        
        scrollbar_restore = ttk.Scrollbar(list_frame, orient=VERTICAL)
        self.listbox_backups = tk.Listbox(list_frame, yscrollcommand=scrollbar_restore.set, height=8, exportselection=False)
        scrollbar_restore.config(command=self.listbox_backups.yview)
        
        scrollbar_restore.pack(side=RIGHT, fill=Y)
        self.listbox_backups.pack(side=LEFT, fill=X, expand=YES)

        self.btn_refresh_restore = bstrap.Button(frame1, text="Atualizar Lista", command=self.load_backup_list, bootstyle="info-outline")
        self.btn_refresh_restore.pack(fill=X, expand=YES, pady=(5, 0))

        frame2 = ttk.Labelframe(parent_tab, text="2. Defina o Novo Nome", padding="10")
        frame2.pack(fill=X, expand=NO, pady=10)

        self.lbl_new_name = ttk.Label(frame2, text="Nome do Novo Banco de Dados:")
        self.lbl_new_name.pack(fill=X, expand=YES)
        
        self.entry_new_name = bstrap.Entry(frame2, bootstyle="primary")
        self.entry_new_name.pack(fill=X, expand=YES, pady=5)

        frame3 = ttk.Labelframe(parent_tab, text="3. Executar e Acompanhar", padding="10")
        frame3.pack(fill=BOTH, expand=YES, pady=5)

        self.btn_start_restore = bstrap.Button(frame3, text="INICIAR RESTAURAÇÃO", command=self.start_restore_thread, bootstyle="success")
        self.btn_start_restore.pack(fill=X, expand=YES, pady=5, ipady=5)
        
        self.progressbar_restore = bstrap.Progressbar(frame3, mode='indeterminate', bootstyle="success-striped")
        self.progressbar_restore.pack(fill=X, expand=YES, pady=5, side=TOP)

        self.log_output = scrolledtext.ScrolledText(frame3, height=10, state='disabled', wrap=tk.WORD, font=("Courier New", 9))
        self.log_output.pack(fill=BOTH, expand=YES, pady=(10, 0))

    # =========================================================================
    # MÉTODOS DO LAUNCHER/ATUALIZADOR (de atualiza_automacao.py)
    # =========================================================================
    
    def criar_widgets_db(self, parent):
        db_frame = bstrap.Labelframe(parent, text="Configuração do Banco de Dados", bootstyle="info", padding=10)
        db_frame.pack(fill="x", pady=(0, 15))

        current_frame = bstrap.Frame(db_frame)
        current_frame.pack(fill="x", padx=5, pady=(5, 10))
        
        bstrap.Label(current_frame, text="Banco de Dados Atual:").pack(side="left")
        
        self.lbl_current_db = bstrap.Label(current_frame, text="Carregando...", font=("-weight bold"), bootstyle="primary")
        self.lbl_current_db.pack(side="left", padx=5)

        new_frame = bstrap.Frame(db_frame)
        new_frame.pack(fill="x", padx=5, pady=(0, 5))

        bstrap.Label(new_frame, text="Mudar para:").pack(side="left")

        self.combo_new_db = bstrap.Combobox(new_frame, bootstyle="info", state="readonly")
        self.combo_new_db.pack(side="left", fill="x", expand=True, padx=(5, 10))

        self.btn_change_db = bstrap.Button(new_frame, text="Salvar", command=self.mudar_banco_dados, bootstyle="info-outline")
        self.btn_change_db.pack(side="right")

    def atualizar_status(self, texto):
        self.status_bar.config(text=texto)

    def ler_config_ini(self):
        try:
            config = configparser.ConfigParser()
            if not os.path.exists(CAMINHO_DO_INI):
                raise FileNotFoundError(f"Arquivo não encontrado: {CAMINHO_DO_INI}\n\nVerifique 'gerenciador_config.ini'.")
            config.read(CAMINHO_DO_INI)
            database_name = config.get(INI_SECTION, INI_KEY)
            return database_name
        except Exception as e:
            self.atualizar_status(f"Erro ao ler {CAMINHO_DO_INI}")
            messagebox.showerror("Erro de Arquivo .ini", f"Não foi possível ler o 'max.ini':\n{e}\n\nVerifique as variáveis INI_SECTION/INI_KEY em 'gerenciador_config.ini'.")
            return "ERRO"

    def escrever_config_ini(self, novo_banco):
        try:
            config = configparser.ConfigParser()
            config.read(CAMINHO_DO_INI)
            if not config.has_section(INI_SECTION):
                config.add_section(INI_SECTION)
            config.set(INI_SECTION, INI_KEY, novo_banco)
            with open(CAMINHO_DO_INI, 'w') as configfile:
                config.write(configfile)
            return True
        except PermissionError:
            msg = "Erro: Sem permissão para escrever no 'max.ini'.\nTente executar o launcher como Administrador."
            self.atualizar_status("Erro: Sem permissão para salvar o .ini.")
            messagebox.showerror("Erro de Permissão", msg)
            return False
        except Exception as e:
            msg = f"Ocorreu um erro ao salvar o 'max.ini':\n{e}"
            self.atualizar_status("Erro ao salvar .ini.")
            messagebox.showerror("Erro Inesperado", msg)
            return False

    def listar_bancos_de_dados_sql(self):
        bancos_sistema = ['master', 'tempdb', 'model', 'msdb']
        try:
            conn_str = (
                f'DRIVER={SQL_DRIVER_LISTA};'
                f'SERVER={SQL_SERVER_INSTANCE};'
                f'DATABASE=master;'
                f'Trusted_Connection=yes;'
            )
            
            bancos_usuario = []
            with pyodbc.connect(conn_str) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sys.databases;")
                for row in cursor.fetchall():
                    if row.name not in bancos_sistema:
                        bancos_usuario.append(row.name)
            return sorted(bancos_usuario)
        except pyodbc.Error as ex:
            sqlstate = ex.args[0]
            if sqlstate == '28000':
                msg = "Erro de Autenticação.\nO script tentou usar 'Autenticação do Windows' e falhou."
            elif sqlstate == '08001':
                msg = f"Erro de Conexão.\nNão foi possível conectar ao '{SQL_SERVER_INSTANCE}'.\nO serviço do SQL Server está rodando?"
            elif sqlstate == 'IM002':
                msg = f"Erro de Driver ODBC.\nDriver não encontrado: {SQL_DRIVER_LISTA}\n\nInstale o driver ODBC ou corrija 'gerenciador_config.ini'."
            else:
                msg = f"Erro de pyodbc não catalogado:\n{ex}"
            self.atualizar_status("Erro: Falha ao listar bancos SQL.")
            messagebox.showerror("Erro de Conexão SQL", msg)
            return None
        except Exception as e:
            self.atualizar_status("Erro: Falha ao listar bancos SQL.")
            messagebox.showerror("Erro Inesperado (SQL)", f"Erro ao listar bancos: {e}")
            return None

    def carregar_banco_atual(self):
        nome_banco_atual = self.ler_config_ini()
        if "ERRO" in nome_banco_atual:
            self.lbl_current_db.config(text="Erro!", bootstyle="danger")
            self.combo_new_db.config(state="disabled")
            self.btn_change_db.config(state="disabled")
            return
        
        self.lbl_current_db.config(text=nome_banco_atual)
        self.atualizar_status(f"Conectado ao banco: {nome_banco_atual}")

        self.atualizar_status("Buscando lista de bancos no SQL Server...")
        lista_de_bancos = self.listar_bancos_de_dados_sql()
        
        if lista_de_bancos is None:
            self.atualizar_status("Erro ao listar bancos. Verifique a conexão SQL.")
            self.combo_new_db.config(state="disabled")
            self.btn_change_db.config(state="disabled")
            self.combo_new_db.set(nome_banco_atual)
            return

        self.combo_new_db['values'] = lista_de_bancos
        
        if nome_banco_atual in lista_de_bancos:
            self.combo_new_db.set(nome_banco_atual)
        elif lista_de_bancos:
            self.combo_new_db.set(lista_de_bancos[0])
        else:
            self.combo_new_db.set("Nenhum banco encontrado")
            self.combo_new_db.config(state="disabled")
            self.btn_change_db.config(state="disabled")

    def mudar_banco_dados(self):
        novo_banco = self.combo_new_db.get().strip()
        
        if not novo_banco or "Nenhum" in novo_banco:
            messagebox.showwarning("Seleção Inválida", "Por favor, selecione um banco de dados válido.")
            return

        banco_atual = self.lbl_current_db.cget("text")
        if novo_banco == banco_atual:
             messagebox.showinfo("Sem Mudanças", "Este já é o banco de dados atual.")
             return

        if not messagebox.askyesno("Confirmar Mudança", 
                                   f"Você tem certeza que deseja mudar o banco de dados de:\n\n"
                                   f"DE: {banco_atual}\n"
                                   f"PARA: {novo_banco}\n\n"
                                   f"O arquivo '{os.path.basename(CAMINHO_DO_INI)}' será modificado."):
            return

        self.btn_change_db.config(state="disabled")
        
        if self.escrever_config_ini(novo_banco):
            nome_lido = self.ler_config_ini()
            self.lbl_current_db.config(text=nome_lido)
            self.atualizar_status(f"Banco de dados alterado para: {novo_banco}")
            messagebox.showinfo("Sucesso", f"Banco de dados alterado para '{novo_banco}' com sucesso!")
        
        self.btn_change_db.config(state="normal")
            
    def popular_lista_versoes(self):
        self.listbox_versoes.delete(0, "end")
        try:
            arquivos = os.listdir(PASTA_DAS_VERSOES)
            versoes_filtradas = [f for f in arquivos if f.lower().endswith('.rar')]
            versoes_filtradas.sort(reverse=True) 

            if not versoes_filtradas:
                self.listbox_versoes.insert(0, "Nenhuma versão (.rar) encontrada.")
                self.listbox_versoes.config(state="disabled")
                self.btn_atualizar.config(state="disabled")
                self.btn_executar_erp.config(bootstyle="primary")
            else:
                self.listbox_versoes.config(state="normal")
                self.btn_atualizar.config(state="normal")
                self.btn_executar_erp.config(bootstyle="info") 
                for versao in versoes_filtradas:
                    self.listbox_versoes.insert("end", f"  {versao}")
        except FileNotFoundError:
            self.atualizar_status("Erro: Pasta de versões não encontrada!")
            messagebox.showerror("Erro de Configuração", 
                                 f"A pasta de versões não foi encontrada:\n{PASTA_DAS_VERSOES}\n\n"
                                 f"Verifique o caminho em '{CONFIG_FILE_NAME}'.")
            self.listbox_versoes.insert(0, "Erro ao carregar pasta.")
            self.listbox_versoes.config(state="disabled")
            self.btn_atualizar.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro ao ler a pasta: {e}")

    def iniciar_cliente_erp(self):
        indices = self.listbox_versoes.curselection()
        if not indices:
            self._callback_lancar_cliente()
            return

        versao_escolhida = self.listbox_versoes.get(indices[0]).strip()
        if not messagebox.askyesno("Confirmar Extração", 
                                   f"Deseja ATUALIZAR OS ARQUIVOS do sistema com a versão:\n\n{versao_escolhida}\n\n"
                                   "ANTES de executá-lo?"):
            self._callback_lancar_cliente()
            return
            
        self.iniciar_extracao_thread(versao_escolhida, self._callback_lancar_cliente)

    def iniciar_atualizacao(self):
        indices_selecionados = self.listbox_versoes.curselection()
        if not indices_selecionados:
            messagebox.showwarning("Seleção Inválida", "Por favor, selecione uma versão da lista para ATUALIZAR.")
            return
            
        versao_escolhida = self.listbox_versoes.get(indices_selecionados[0]).strip()
        if not messagebox.askyesno("Confirmar Atualização Completa", 
                                   f"O script irá agora:\n\n"
                                   f"1. EXTRAIR os arquivos da versão: {versao_escolhida}\n"
                                   f"2. ATUALIZAR o banco de dados com '{NOME_EXE_ATUALIZADOR}'.\n\n"
                                   "Não use o mouse ou teclado during o processo. Continuar?"):
            return

        self.iniciar_extracao_thread(versao_escolhida, self._callback_lancar_atualizador)

    def iniciar_extracao_thread(self, arquivo_versao_escolhida, on_finish_callback):
        self.desabilitar_botoes_launcher()
        
        thread = threading.Thread(
            target=self._thread_extrair, 
            args=(arquivo_versao_escolhida, on_finish_callback),
            daemon=True
        )
        thread.start()

    def _thread_extrair(self, arquivo_versao_escolhida, on_finish_callback):
        caminho_completo_versao = os.path.join(PASTA_DAS_VERSOES, arquivo_versao_escolhida)
        
        try:
            self.after(0, self.atualizar_status, f"Iniciando extração de {arquivo_versao_escolhida}...")
            comando_7zip = [CAMINHO_DO_7ZIP_EXE, 'x', caminho_completo_versao, f'-o{PASTA_DO_SISTEMA}', '-y']

            subprocess.run(
                comando_7zip, 
                check=True,
                capture_output=True,
                text=True,
                encoding='latin-1',
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.after(0, self.atualizar_status, "Extração concluída com sucesso.")
            
            if on_finish_callback:
                self.after(0, on_finish_callback)
            else:
                self.after(0, self.reabilitar_botoes_launcher)

        except FileNotFoundError:
            msg_erro = f"Erro: '7z.exe' não encontrado.\nVerifique: {CAMINHO_DO_7ZIP_EXE}\n\n(Edite o caminho em '{CONFIG_FILE_NAME}'.)"
            self.after(0, self.on_automacao_erro, msg_erro)
        except subprocess.CalledProcessError as e:
            msg_erro = f"Erro ao extrair o arquivo:\n{e.stderr}\n\nVerifique permissões em {PASTA_DO_SISTEMA}"
            self.after(0, self.on_automacao_erro, msg_erro)
        except Exception as e:
            msg_erro = f"Erro inesperado na extração:\n{e}"
            self.after(0, self.on_automacao_erro, msg_erro)

    def _callback_lancar_cliente(self):
        self.atualizar_status(f"Iniciando {NOME_EXE_CLIENTE}...")
        try:
            subprocess.Popen([CAMINHO_DO_ERP_CLIENTE], cwd=PASTA_DO_SISTEMA) 
            self.atualizar_status("Cliente ERP iniciado.")
            self.wm_state('iconic') 
        except FileNotFoundError:
            msg = f"Erro: Executável do ERP não encontrado.\nCaminho: {CAMINHO_DO_ERP_CLIENTE}\n\nVerifique 'gerenciador_config.ini'."
            self.atualizar_status("Erro: Cliente ERP não encontrado.")
            messagebox.showerror("Erro de Configuração", msg)
        except Exception as e:
            msg = f"Erro inesperado ao iniciar o ERP:\n{e}"
            self.atualizar_status("Erro ao iniciar cliente ERP.")
            messagebox.showerror("Erro na Execução", msg)
        
        self.reabilitar_botoes_launcher()

    def _callback_lancar_atualizador(self):
        indices = self.listbox_versoes.curselection()
        if not indices:
             self.on_automacao_erro("Erro: Perdeu-se a seleção da versão.")
             return
             
        versao_escolhida = self.listbox_versoes.get(indices[0]).strip()
        caminho_completo_versao = os.path.join(PASTA_DAS_VERSOES, versao_escolhida)
        
        comando_automacao = [CAMINHO_DO_MAX_ATUALIZA, caminho_completo_versao]

        try:
            self.after(0, self.atualizar_status, f"Iniciando {NOME_EXE_ATUALIZADOR}...")
            app = Application(backend="win32").start(
                ' '.join(f'"{c}"' for c in comando_automacao),
                work_dir=PASTA_DO_SISTEMA 
            )

            # TITULO_DA_JANELA_MAX_ATUALIZA = "Max Atualiza" 
            # NOME_DO_BOTAO_EXECUTAR = "&Executar F9"
            
            self.after(0, self.atualizar_status, "Atualizador iniciado. Automação manual necessária.")
            messagebox.showinfo("Automação Incompleta", 
                                f"O '{NOME_EXE_ATUALIZADOR}' foi iniciado.\n\n"
                                "A automação do clique (pywinauto) está desativada. "
                                "Por favor, realize a atualização manualmente e feche o atualizador.\n\n"
                                "(Para ativar, defina as variáveis de automação no código.)")
            self.after(0, self.reabilitar_botoes_launcher)

        except Exception as e:
            erro_msg = f"Erro na automação com pywinauto:\n{e}\n\n"
            erro_msg += "VERIFIQUE AS VARIÁVEIS 'TITULO_DA_JANELA...' e 'NOME_DO_BOTAO...'."
            self.after(0, self.on_automacao_erro, erro_msg)

    def desabilitar_botoes_launcher(self):
        self.btn_executar_erp.config(state="disabled")
        self.btn_atualizar.config(state="disabled")
        self.btn_recarregar.config(state="disabled")
        self.btn_change_db.config(state="disabled")
    
    def reabilitar_botoes_launcher(self):
        self.btn_executar_erp.config(state="normal")
        self.btn_recarregar.config(state="normal")
        
        if self.combo_new_db.cget("state") != "disabled":
            self.btn_change_db.config(state="normal")
        if self.listbox_versoes.cget("state") == "normal":
            self.btn_atualizar.config(state="normal")

    def on_automacao_sucesso(self, saida_processo):
        self.atualizar_status("Operação concluída com sucesso!")
        self.reabilitar_botoes_launcher()
        messagebox.showinfo("Sucesso", saida_processo)

    def on_automacao_erro(self, mensagem_erro):
        self.atualizar_status("Erro durante a operação.")
        self.reabilitar_botoes_launcher()
        messagebox.showerror("Erro na Automação/Extração", mensagem_erro)

    # =========================================================================
    # MÉTODOS DO ASSISTENTE DE RESTORE (de restore_database.py)
    # =========================================================================

    def log_message(self, message):
        self.log_output.config(state='normal')
        self.log_output.insert(tk.END, message + "\n")
        self.log_output.config(state='disabled')
        self.log_output.see(tk.END)

    def load_backup_list(self):
        self.listbox_backups.delete(0, tk.END)
        try:
            arquivos_max = [f for f in os.listdir(self.caminho_backup) if f.upper().endswith('.MAX')]
            if not arquivos_max:
                self.listbox_backups.insert(tk.END, "Nenhum arquivo .MAX encontrado.")
            else:
                for f in arquivos_max:
                    self.listbox_backups.insert(tk.END, f)
        except FileNotFoundError:
            self.listbox_backups.insert(tk.END, f"Erro: Pasta '{self.caminho_backup}' não existe.")
        except Exception as e:
            self.listbox_backups.insert(tk.END, f"Erro ao ler pasta: {e}")

    def start_restore_thread(self):
        try:
            selected_index = self.listbox_backups.curselection()[0]
            backup_file_name = self.listbox_backups.get(selected_index)
            if not backup_file_name.upper().endswith('.MAX'):
                messagebox.showerror("Erro", "Por favor, selecione um arquivo de backup válido.")
                return
        except IndexError:
            messagebox.showerror("Erro de Seleção", "Por favor, selecione um arquivo de backup da lista.")
            return

        new_db_name = self.entry_new_name.get().strip()
        if not new_db_name:
            messagebox.showerror("Nome Inválido", "Por favor, digite um nome para o novo banco.")
            return

        self.btn_start_restore.config(state='disabled', text="Restaurando...")
        self.btn_refresh_restore.config(state='disabled')
        self.listbox_backups.config(state='disabled')
        self.entry_new_name.config(state='disabled')
        self.progressbar_restore.start()
        self.log_output.config(state='normal')
        self.log_output.delete('1.0', tk.END)
        self.log_output.config(state='disabled')

        self.log_message(f"--- Iniciando Restore ---")
        self.log_message(f"Backup: {backup_file_name}")
        self.log_message(f"Novo Banco: {new_db_name}")
        self.atualizar_status(f"Iniciando restauração do banco '{new_db_name}'...")
        
        threading.Thread(
            target=self.run_restore_logic, 
            args=(backup_file_name, new_db_name),
            daemon=True
        ).start()

    def process_queue(self):
        try:
            message = self.msg_queue.get_nowait()
            
            if message == "__DONE__":
                self.on_restore_complete(success=True)
            elif message.startswith("__ERROR__"):
                self.log_message(message)
                self.on_restore_complete(success=False)
            else:
                self.log_message(message)
                
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue)

    def on_restore_complete(self, success=True):
        self.progressbar_restore.stop()
        self.btn_start_restore.config(state='normal', text="INICIAR RESTAURAÇÃO")
        self.btn_refresh_restore.config(state='normal')
        self.listbox_backups.config(state='normal')
        self.entry_new_name.config(state='normal')
        
        if success:
            self.log_message("\n--- PROCESSO CONCLUÍDO COM SUCESSO! ---")
            self.atualizar_status("Restauração concluída com sucesso.")
            messagebox.showinfo("Sucesso", "Banco de dados restaurado com sucesso!")
            self.carregar_banco_atual() 
        else:
            self.log_message("\n--- OCORREU UM ERRO. VERIFIQUE O LOG. ---")
            self.atualizar_status("Falha na restauração. Verifique o log.")
            messagebox.showerror("Falha", "A restauração falhou. Verifique o log para detalhes.")

    def run_restore_logic(self, nome_arquivo_backup, nome_banco_novo):
        try:
            caminho_backup_completo = os.path.join(self.caminho_backup, nome_arquivo_backup)
            self.msg_queue.put(f"Arquivo de backup: {caminho_backup_completo}")

            self.msg_queue.put(f"Procurando diretório de dados em: {PASTA_DO_SISTEMA}")
            i = 1
            while True:
                novo_nome_pasta_dados = f"dados{i}"
                caminho_dados_novo = os.path.join(PASTA_DO_SISTEMA, novo_nome_pasta_dados)
                if not os.path.exists(caminho_dados_novo):
                    self.msg_queue.put(f"Diretório '{caminho_dados_novo}' está livre.")
                    break
                self.msg_queue.put(f"Diretório '{caminho_dados_novo}' já existe...")
                i += 1

            self.msg_queue.put(f"Criando diretório: {caminho_dados_novo}")
            os.makedirs(caminho_dados_novo)
            self.msg_queue.put("Diretório criado.")

            novo_arquivo_mdf = os.path.join(caminho_dados_novo, f"{nome_banco_novo}.mdf")
            novo_arquivo_ldf = os.path.join(caminho_dados_novo, f"{nome_banco_novo}_log.ldf")

            connection_string = (
                f"DRIVER={ODBC_DRIVER_RESTORE};"
                f"SERVER={SERVIDOR};"
                f"DATABASE=master;"
                f"UID={USUARIO};"
                f"PWD={SENHA};"
                f"Trusted_Connection=no;"
            )
            
            conn = None
            cursor = None
            try:
                conn = pyodbc.connect(connection_string, autocommit=True)
                cursor = conn.cursor()
                self.msg_queue.put("Conexão com 'master' estabelecida.")

                self.msg_queue.put("Buscando nomes lógicos no arquivo de backup...")
                sql_filelist = "RESTORE FILELISTONLY FROM DISK = ?"
                cursor.execute(sql_filelist, (caminho_backup_completo,))
                
                files = cursor.fetchall()
                logical_data_name = None
                logical_log_name = None
                for file_info in files:
                    if file_info.Type == 'D':
                        logical_data_name = file_info.LogicalName
                    elif file_info.Type == 'L':
                        logical_log_name = file_info.LogicalName
                
                if not logical_data_name or not logical_log_name:
                    raise Exception("Não foi possível encontrar nomes lógicos de dados/log no backup.")

                self.msg_queue.put(f"Nome lógico dos dados: {logical_data_name}")
                self.msg_queue.put(f"Nome lógico do log: {logical_log_name}")

                self.msg_queue.put(f"Iniciando a restauração... (Isso pode demorar)")
                
                sql_restore = f"""
                RESTORE DATABASE [{nome_banco_novo}]
                FROM DISK = '{caminho_backup_completo}'
                WITH 
                    MOVE '{logical_data_name}' TO '{novo_arquivo_mdf}',
                    MOVE '{logical_log_name}' TO '{novo_arquivo_ldf}',
                    NOUNLOAD, 
                    REPLACE,
                    STATS = 10
                """
                
                cursor.execute(sql_restore)
                while cursor.nextset():
                    pass

                self.msg_queue.put("\n-------------------------------------------")
                self.msg_queue.put(f"Sucesso! Banco '{nome_banco_novo}' restaurado.")
                self.msg_queue.put(f"Pasta de dados: {caminho_dados_novo}")
                self.msg_queue.put("-------------------------------------------")
                self.msg_queue.put("__DONE__") 

            except pyodbc.Error as ex:
                sqlstate = ex.args[0]
                if "42000" in sqlstate:
                    self.msg_queue.put("\n[ERRO DE PERMISSÃO SQL]")
                    self.msg_queue.put("O usuário SQL não tem permissão para RESTORE ou...")
                    self.msg_queue.put(f"A CONTA DE SERVIÇO do SQL Server não tem permissão para ler/escrever na pasta {PASTA_DO_SISTEMA}.")
                self.msg_queue.put(f"\nErro de pyodbc: {ex}")
                raise
            
            except Exception as e:
                 raise

            finally:
                if cursor:
                    cursor.close()
                if conn:
                    conn.close()
                    self.msg_queue.put("Conexão fechada.")
        
        except Exception as e:
            self.msg_queue.put(f"\nUm erro fatal ocorreu na thread: {e}")
            self.msg_queue.put("__ERROR__")


def validar_caminhos():
    """
    Valida os caminhos carregados do .ini.
    Retorna uma lista de erros, ou uma lista vazia se tudo estiver OK.
    """
    global CAMINHO_DO_ERP_CLIENTE, CAMINHO_DO_MAX_ATUALIZA, CAMINHO_BASE_MAX_BACKUP
    
    print("DEBUG: Iniciando validação de caminhos...")
    
    # 1. Define os caminhos derivados
    CAMINHO_DO_ERP_CLIENTE = os.path.join(PASTA_DO_SISTEMA, NOME_EXE_CLIENTE)
    CAMINHO_DO_MAX_ATUALIZA = os.path.join(PASTA_DO_SISTEMA, NOME_EXE_ATUALIZADOR)
    CAMINHO_BASE_MAX_BACKUP = os.path.join(PASTA_DO_SISTEMA, 'backup')
    
    # 2. Validações de caminho
    caminhos_invalidos = []
    
    if not os.path.isdir(PASTA_DO_SISTEMA):
        caminhos_invalidos.append(f"Pasta do Sistema: {PASTA_DO_SISTEMA}")
    else:
        if not os.path.exists(CAMINHO_DO_ERP_CLIENTE):
            caminhos_invalidos.append(f"Cliente ERP: {CAMINHO_DO_ERP_CLIENTE}")
        if not os.path.exists(CAMINHO_DO_MAX_ATUALIZA):
            caminhos_invalidos.append(f"Max Atualiza: {CAMINHO_DO_MAX_ATUALIZA}")
            
    if not os.path.isdir(PASTA_DAS_VERSOES):
        caminhos_invalidos.append(f"Pasta de Versões: {PASTA_DAS_VERSOES}")
        
    if not os.path.exists(CAMINHO_DO_INI):
        caminhos_invalidos.append(f"Arquivo .ini: {CAMINHO_DO_INI}")
        
    if not os.path.exists(CAMINHO_DO_7ZIP_EXE):
        caminhos_invalidos.append(f"7-Zip: {CAMINHO_DO_7ZIP_EXE}")
        
    if not os.path.isdir(CAMINHO_BASE_MAX_BACKUP):
         caminhos_invalidos.append(f"Pasta de Backup: {CAMINHO_BASE_MAX_BACKUP}")
    
    if caminhos_invalidos:
         print(f"DEBUG: Validação falhou. Erros: {caminhos_invalidos}")
    else:
         print("DEBUG: Validação OK.")
         
    return caminhos_invalidos


# --- Ponto de entrada da aplicação ---
if __name__ == "__main__":
    
    # Cria uma root 'invisível' para as caixas de diálogo iniciais
    root_inicial = tk.Tk()
    root_inicial.withdraw()

    while True:
        # 1. Carrega as configurações do 'gerenciador_config.ini'
        carregar_ou_criar_configuracoes()

        # 2. Valida os caminhos
        caminhos_invalidos = validar_caminhos()

        if not caminhos_invalidos:
            # 3. Sucesso! Sai do loop e continua
            print("DEBUG: Validação final OK. Saindo do loop.")
            break
        
        # 4. Falha na validação. Mostra o erro e abre a janela de configuração
        msg_erro = "Erro: Um ou mais caminhos não foram encontrados.\n\n"
        msg_erro += "\n".join(caminhos_invalidos)
        msg_erro += f"\n\nPor favor, corrija os caminhos."
        
        print("DEBUG: Erro de validação, mostrando messagebox...")
        
        # --- MUDANÇA DE TESTE AQUI ---
        messagebox.showerror("Erro de Configuração (VERSÃO NOVA)", msg_erro, parent=root_inicial)
        # -----------------------------
        
        # Cria e espera a janela de configuração fechar
        print("DEBUG: Messagebox fechado, chamando ConfigWindow...")
        config_app = ConfigWindow(root_inicial)
        root_inicial.wait_window(config_app) 
        print("DEBUG: ConfigWindow fechada.")
        
        if not config_app.salvo:
            # Se o usuário fechou a janela de config sem salvar, encerra o programa
            print("DEBUG: Usuário cancelou. Encerrando.")
            sys.exit()
        
        print("DEBUG: Usuário salvou. Reiniciando loop de validação...")
        # Se o usuário salvou, o loop 'while True' recomeça e vai revalidar os caminhos
        
    # Destrói a root inicial que não é mais necessária
    root_inicial.destroy()

    # 5. Inicia a Aplicação Principal (só chega aqui se a validação passar)
    print("DEBUG: Iniciando GerenciadorMaxApp principal...")
    app = GerenciadorMaxApp()
    app.mainloop()