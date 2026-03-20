import tkinter as tk
from tkinter import W, END
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
import shutil

try:
    from pywinauto.application import Application
except ImportError:
    pass

# =============================================================================
# CONFIGURAÇÕES E VARIÁVEIS GLOBAIS
# =============================================================================
CONFIG_FILE_NAME = 'gerenciador_config.ini'
SQL_QUERY_VERSAO = "select cofMaxAtualizaVersao from config" 

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

CAMINHO_DO_ERP_CLIENTE = None
CAMINHO_DO_MAX_ATUALIZA = None
CAMINHO_BASE_MAX_BACKUP = None

# =============================================================================
# FUNÇÕES DE CONFIGURAÇÃO
# =============================================================================
def carregar_ou_criar_configuracoes(root_para_aviso=None):
    global PASTA_DO_SISTEMA, NOME_EXE_CLIENTE, NOME_EXE_ATUALIZADOR, PASTA_DAS_VERSOES
    global CAMINHO_DO_INI, CAMINHO_DO_7ZIP_EXE, INI_SECTION, INI_KEY, SERVIDOR
    global USUARIO, SENHA, ODBC_DRIVER_RESTORE, SQL_DRIVER_LISTA, SQL_SERVER_INSTANCE
    
    config = configparser.ConfigParser()
    
    # Se não existir, cria o padrão
    if not os.path.exists(CONFIG_FILE_NAME):
        config['CAMINHOS'] = {
            'PASTA_DO_SISTEMA': r'D:\Max',
            'PASTA_DAS_VERSOES': r'D:\Max\Versões',
            'CAMINHO_DO_INI': r'D:\Max\max.ini',
            'CAMINHO_DO_7ZIP_EXE': r'C:\Program Files (x86)\7-Zip\7z.exe'
        }
        config['EXECUTAVEIS'] = {
            'NOME_EXE_CLIENTE': 'MAX_manager2.exe',
            'NOME_EXE_ATUALIZADOR': 'MAX_Atualiza.exe'
        }
        config['CONFIG_INI_MAX'] = {'INI_SECTION': 'CON', 'INI_KEY': 'Initial catalog'}
        config['SQL_LAUDO'] = {
            'SQL_DRIVER_LISTA': '{ODBC Driver 17 for SQL Server}',
            'SQL_SERVER_INSTANCE': 'localhost'
        }
        config['SQL_RESTORE'] = {
            'SERVIDOR': 'localhost', 'USUARIO': 'sa', 'SENHA': 'macro01',
            'ODBC_DRIVER_RESTORE': '{ODBC Driver 17 for SQL Server}'
        }
        try:
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f: config.write(f)
            messagebox.showinfo("Configuração", f"Arquivo '{CONFIG_FILE_NAME}' criado.", parent=root_para_aviso)
            sys.exit()
        except: sys.exit()
    
    # Lê o arquivo
    try:
        config.read(CONFIG_FILE_NAME, encoding='utf-8')
        PASTA_DO_SISTEMA = config.get('CAMINHOS', 'PASTA_DO_SISTEMA')
        PASTA_DAS_VERSOES = config.get('CAMINHOS', 'PASTA_DAS_VERSOES')
        CAMINHO_DO_INI = config.get('CAMINHOS', 'CAMINHO_DO_INI')
        CAMINHO_DO_7ZIP_EXE = config.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE')
        NOME_EXE_CLIENTE = config.get('EXECUTAVEIS', 'NOME_EXE_CLIENTE')
        NOME_EXE_ATUALIZADOR = config.get('EXECUTAVEIS', 'NOME_EXE_ATUALIZADOR')
        INI_SECTION = config.get('CONFIG_INI_MAX', 'INI_SECTION')
        INI_KEY = config.get('CONFIG_INI_MAX', 'INI_KEY')
        SQL_DRIVER_LISTA = config.get('SQL_LAUDO', 'SQL_DRIVER_LISTA')
        SQL_SERVER_INSTANCE = config.get('SQL_LAUDO', 'SQL_SERVER_INSTANCE')
        SERVIDOR = config.get('SQL_RESTORE', 'SERVIDOR')
        USUARIO = config.get('SQL_RESTORE', 'USUARIO')
        SENHA = config.get('SQL_RESTORE', 'SENHA')
        ODBC_DRIVER_RESTORE = config.get('SQL_RESTORE', 'ODBC_DRIVER_RESTORE')
    except Exception as e:
        messagebox.showerror("Erro Config", f"{e}", parent=root_para_aviso); sys.exit()

class ConfigWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configuração de Caminhos")
        self.geometry("700x320")
        self.resizable(False, False)
        self.salvo = False 
        self.transient(parent); self.grab_set()

        self.var_sistema = tk.StringVar()
        self.var_versoes = tk.StringVar()
        self.var_ini = tk.StringVar()
        self.var_7zip = tk.StringVar()

        self.ler_ini_atual()
        
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)
        
        bstrap.Label(main_frame, text="Corrija os caminhos:", font=("bold", 12), bootstyle="danger").pack(pady=(0, 15))

        self.criar_campo(main_frame, "Pasta do Sistema", self.var_sistema, True)
        self.criar_campo(main_frame, "Pasta de Versões", self.var_versoes, True)
        self.criar_campo(main_frame, "Arquivo .ini", self.var_ini, False)
        self.criar_campo(main_frame, "Executável 7-Zip", self.var_7zip, False)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, pady=(20, 0))
        
        bstrap.Button(btn_frame, text="Salvar", command=self.salvar, bootstyle="success").pack(side=RIGHT, padx=5)
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
            self.var_ini.set(c.get('CAMINHOS', 'CAMINHO_DO_INI', fallback=''))
            self.var_7zip.set(c.get('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', fallback=''))
        except: pass

    def salvar(self):
        try:
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_NAME, encoding='utf-8')
            config.set('CAMINHOS', 'PASTA_DO_SISTEMA', self.var_sistema.get())
            config.set('CAMINHOS', 'PASTA_DAS_VERSOES', self.var_versoes.get())
            config.set('CAMINHOS', 'CAMINHO_DO_INI', self.var_ini.get())
            config.set('CAMINHOS', 'CAMINHO_DO_7ZIP_EXE', self.var_7zip.get())
            with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f: config.write(f)
            self.salvo = True; self.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"{e}", parent=self)

# =============================================================================
# APLICAÇÃO PRINCIPAL (PRETO & VERMELHO)
# =============================================================================
class GerenciadorMaxApp(bstrap.Window):
    def __init__(self):
        super().__init__(themename="cyborg")
        self.title("Gerenciador Max (Black & Red)")
        self.geometry("900x780")
        
        # Oculta a janela durante o carregamento inicial
        self.withdraw() 
        self.msg_queue = queue.Queue()

    def iniciar_interface(self):
        self.caminho_backup = CAMINHO_BASE_MAX_BACKUP 
        self.create_layout()
        self.load_backups()
        self.popular_versoes()
        self.carregar_banco_atual()
        self.process_queue()
        # Mostra a janela pronta
        self.deiconify()

    def create_layout(self):
        main = bstrap.Frame(self, padding=10)
        main.pack(fill=BOTH, expand=YES)
        
        self.notebook = bstrap.Notebook(main, bootstyle="danger")
        self.notebook.pack(fill=BOTH, expand=YES)

        tab_launch = bstrap.Frame(self.notebook, padding=15)
        tab_restore = bstrap.Frame(self.notebook, padding=15)
        tab_tools = bstrap.Frame(self.notebook, padding=15)

        self.notebook.add(tab_launch, text="Launcher")
        self.notebook.add(tab_restore, text="Restore")
        self.notebook.add(tab_tools, text="Ferramentas")

        self.setup_launcher(tab_launch)
        self.setup_restore(tab_restore)
        self.setup_tools(tab_tools)

        self.status = bstrap.Label(self, text="Pronto.", relief=FLAT, anchor=W, padding=8, bootstyle="secondary")
        self.status.pack(side=BOTTOM, fill=X)

    # --- ABA 1: LAUNCHER ---
    def setup_launcher(self, parent):
        bstrap.Label(parent, text="Launcher Maxdata", font=("bold", 20), bootstyle="danger").pack(pady=5)

        db_frame = bstrap.Labelframe(parent, text=" Banco de Dados ", bootstyle="danger", padding=15)
        db_frame.pack(fill=X, pady=10)

        f1 = bstrap.Frame(db_frame)
        f1.pack(fill=X, pady=5)
        bstrap.Label(f1, text="Banco Atual (INI):", bootstyle="secondary").pack(side=LEFT)
        self.lbl_db_atual = bstrap.Label(f1, text="...", font=("bold"), bootstyle="danger")
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
        
        bstrap.Button(f2, text="Salvar", command=self.mudar_banco, bootstyle="danger-outline").pack(side=LEFT)

        v_frame = bstrap.Labelframe(parent, text=" Versões Disponíveis ", bootstyle="danger", padding=15)
        v_frame.pack(fill=BOTH, expand=YES, pady=10)

        self.lb_versoes = ttk.Treeview(v_frame, height=8, bootstyle="danger", columns=("versao"), show="headings")
        self.lb_versoes.heading("versao", text="Versões Disponíveis")
        self.lb_versoes.column("versao", width=300, anchor=W)
        sb = bstrap.Scrollbar(v_frame, command=self.lb_versoes.yview, bootstyle="danger-round")
        self.lb_versoes.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_versoes.pack(side=LEFT, fill=BOTH, expand=YES)

        bts = bstrap.Frame(parent)
        bts.pack(fill=X)
        
        bstrap.Button(bts, text="Recarregar", command=self.popular_versoes, bootstyle="info-outline").pack(side=LEFT)
        bstrap.Button(bts, text="EXECUTAR SISTEMA", command=self.lancar_erp, bootstyle="success-outline").pack(side=RIGHT, padx=5)
        bstrap.Button(bts, text="ATUALIZAR VERSÃO", command=self.lancar_atualizacao, bootstyle="danger").pack(side=RIGHT)

    # --- ABA 2: RESTORE ---
    def setup_restore(self, parent):
        f1 = bstrap.Labelframe(parent, text=" 1. Selecione o Backup ", bootstyle="danger", padding=10)
        f1.pack(fill=BOTH, expand=YES)
        
        self.lb_backups = ttk.Treeview(f1, height=10, bootstyle="danger", columns=("backup"), show="headings")
        self.lb_backups.heading("backup", text="Backups Disponíveis")
        self.lb_backups.column("backup", width=300, anchor=tk.W)
        sb = bstrap.Scrollbar(f1, command=self.lb_backups.yview, bootstyle="danger-round")
        self.lb_backups.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_backups.pack(side=LEFT, fill=BOTH, expand=YES)
        bstrap.Button(f1, text="Atualizar Lista", command=self.load_backups, bootstyle="link-danger").pack(fill=X)

        f2 = bstrap.Labelframe(parent, text=" 2. Nome do Novo Banco ", bootstyle="danger", padding=10)
        f2.pack(fill=X, pady=10)
        self.entry_new_db = bstrap.Entry(f2, bootstyle="danger")
        self.entry_new_db.pack(fill=X)

        f3 = bstrap.Frame(parent)
        f3.pack(fill=BOTH, expand=YES)
        
        self.btn_restore = bstrap.Button(f3, text="INICIAR RESTAURAÇÃO", command=self.iniciar_restore_thread, bootstyle="danger")
        self.btn_restore.pack(fill=X, pady=5)
        
        self.progress = bstrap.Progressbar(f3, mode='indeterminate', bootstyle="danger-striped")
        self.progress.pack(fill=X, pady=5)
        
        self.log_txt = scrolledtext.ScrolledText(f3, height=8, state='disabled', font=("Consolas", 9), bg="#111", fg="#0f0")
        self.log_txt.pack(fill=BOTH, expand=YES)

    # --- ABA 3: TOOLS ---
    def setup_tools(self, parent):
        bstrap.Label(parent, text="Gerenciamento de Bancos", font=("bold", 14), bootstyle="danger").pack(pady=10)
        
        flist = bstrap.Labelframe(parent, text=" Bancos SQL ", bootstyle="danger", padding=10)
        flist.pack(fill=BOTH, expand=YES)
        
        self.lb_tools = ttk.Treeview(flist, height=10, bootstyle="danger", columns=("banco"), show="headings")
        self.lb_tools.heading("banco", text="Bancos de Dados")
        self.lb_tools.column("banco", width=300, anchor=tk.W)
        sb = bstrap.Scrollbar(flist, command=self.lb_tools.yview, bootstyle="danger-round")
        self.lb_tools.config(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill=Y)
        self.lb_tools.pack(side=LEFT, fill=BOTH, expand=YES)
        
        fbtn = bstrap.Frame(flist)
        fbtn.pack(side=RIGHT, fill=Y, padx=10)
        
        bstrap.Button(fbtn, text="Atualizar", command=self.atualizar_tools, bootstyle="secondary-outline").pack(fill=X, pady=5)
        bstrap.Button(fbtn, text="EXCLUIR (DROP)", command=self.drop_database, bootstyle="danger").pack(fill=X, pady=20)

    # =========================================================================
    # LÓGICA DO SISTEMA
    # =========================================================================

    def get_versao(self, db):
        if not db or "Nenhum" in db: return "---"
        conn_str = f'DRIVER={SQL_DRIVER_LISTA};SERVER={SQL_SERVER_INSTANCE};DATABASE={db};Trusted_Connection=yes;'
        try:
            with pyodbc.connect(conn_str, timeout=2) as conn:
                cursor = conn.cursor()
                cursor.execute(SQL_QUERY_VERSAO)
                row = cursor.fetchone()
                return str(row[0]) if row else "N/A"
        except: return "---"

    def carregar_banco_atual(self):
        try:
            c = configparser.ConfigParser()
            c.read(CAMINHO_DO_INI)
            atual = c.get(INI_SECTION, INI_KEY)
        except: atual = "ERRO"
        
        if "ERRO" in atual:
            self.lbl_db_atual.config(text="Erro Ler INI")
        else:
            self.lbl_db_atual.config(text=atual)
            self.lbl_versao_sql.config(text=self.get_versao(atual))
        
        bancos = self.listar_sql_dbs()
        self.combo_db['values'] = bancos
        if atual in bancos: self.combo_db.set(atual)
        for i in self.lb_tools.get_children(): self.lb_tools.delete(i)
        for b in bancos: self.lb_tools.insert("", tk.END, values=(b,))

    def preview_version(self, e):
        db = self.combo_db.get()
        v = self.get_versao(db)
        self.status.config(text=f"Banco selecionado: {db} (Versão: {v})")

    def listar_sql_dbs(self):
        conn_str = f'DRIVER={SQL_DRIVER_LISTA};SERVER={SQL_SERVER_INSTANCE};DATABASE=master;Trusted_Connection=yes;'
        try:
            lst = []
            with pyodbc.connect(conn_str) as conn:
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
            with open(CAMINHO_DO_INI, 'w') as f: c.write(f)
            messagebox.showinfo("Sucesso", f"Alterado para: {novo}")
            self.carregar_banco_atual()
        except Exception as e: messagebox.showerror("Erro", f"{e}")

    def popular_versoes(self):
        for i in self.lb_versoes.get_children(): self.lb_versoes.delete(i)
        try:
            for f in sorted([x for x in os.listdir(PASTA_DAS_VERSOES) if x.lower().endswith('.rar')], reverse=True):
                self.lb_versoes.insert("", END, values=(f,))
        except: pass

    def lancar_erp(self):
        try: subprocess.Popen([CAMINHO_DO_ERP_CLIENTE], cwd=PASTA_DO_SISTEMA)
        except Exception as e: messagebox.showerror("Erro", f"{e}")

    def lancar_atualizacao(self):
        sel = self.lb_versoes.selection()
        if not sel: return
        arq = self.lb_versoes.item(sel[0], "values")[0]
        if messagebox.askyesno("Confirmar", f"Atualizar para {arq}?"):
            threading.Thread(target=self.thread_extrair, args=(arq,), daemon=True).start()

    def thread_extrair(self, arq):
        try:
            self.status.config(text="Extraindo...")
            cmd = [CAMINHO_DO_7ZIP_EXE, 'x', os.path.join(PASTA_DAS_VERSOES, arq), f'-o{PASTA_DO_SISTEMA}', '-y']
            subprocess.run(cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.after(0, self.lancar_atualizador_callback)
        except Exception as e:
            # FIX CRÍTICO: Converte o erro em string e passa como argumento padrão
            # Isso evita que o Python "perca" a variável 'e' quando o except termina
            erro_msg = str(e)
            self.after(0, lambda m=erro_msg: messagebox.showerror("Erro Extração", m))

    def lancar_atualizador_callback(self):
        try:
            subprocess.Popen([CAMINHO_DO_MAX_ATUALIZA], cwd=PASTA_DO_SISTEMA)
            self.status.config(text="Atualizador Aberto.")
        except Exception as e: messagebox.showerror("Erro", f"{e}")

    def load_backups(self):
        for i in self.lb_backups.get_children(): self.lb_backups.delete(i)
        try:
            fs = [f for f in os.listdir(self.caminho_backup) if f.upper().endswith((".MAX",".BAK",".ZIP",".RAR"))]
            for f in sorted(fs, key=lambda x: os.path.getmtime(os.path.join(self.caminho_backup, x)), reverse=True):
                self.lb_backups.insert("", tk.END, values=(f,))
        except: pass

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
                self.msg_queue.put("Extraindo arquivo...")
                tmp_dir = os.path.join(self.caminho_backup, f"_tmp_{int(time.time())}")
                os.makedirs(tmp_dir, exist_ok=True)
                
                cmd = [CAMINHO_DO_7ZIP_EXE, 'x', origem, f'-o{tmp_dir}', '-y']
                subprocess.run(cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                encontrado = None
                for root, dirs, files in os.walk(tmp_dir):
                    for f in files:
                        if f.upper().endswith(('.MAX', '.BAK')):
                            encontrado = os.path.join(root, f)
                            break
                if not encontrado: raise Exception("Arquivo .MAX/.BAK não encontrado.")
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
            rows = cur.fetchall()
            ld, ll = None, None
            for r in rows:
                if r.Type == 'D': ld = r.LogicalName
                if r.Type == 'L': ll = r.LogicalName
            
            self.msg_queue.put("Restaurando SQL...")
            sql = f"""RESTORE DATABASE [{dbname}] FROM DISK='{final}'
                      WITH MOVE '{ld}' TO '{mdf}', MOVE '{ll}' TO '{ldf}', REPLACE"""
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
                    self.carregar_banco_atual()
                elif msg == "__ERROR__":
                    self.progress.stop()
                    self.btn_restore.config(state='normal')
                    messagebox.showerror("Erro", "Falha. Veja o log.")
                else:
                    self.log_txt.config(state='normal')
                    self.log_txt.insert(END, msg + "\n")
                    self.log_txt.see(END)
                    self.log_txt.config(state='disabled')
        except: pass
        self.after(100, self.process_queue)

    def atualizar_tools(self):
        dbs = self.listar_sql_dbs()
        for i in self.lb_tools.get_children(): self.lb_tools.delete(i)
        for d in dbs: self.lb_tools.insert("", tk.END, values=(d,))

    def drop_database(self):
        sel = self.lb_tools.selection()
        if not sel: return
        db = self.lb_tools.item(sel[0], "values")[0]
        
        if db == self.lbl_db_atual.cget("text"):
            messagebox.showerror("Bloqueado", "Não apague o banco definido no INI.")
            return

        if messagebox.askyesno("PERIGO", f"DELETAR '{db}' permanentemente?"):
            try:
                conn_str = f"DRIVER={ODBC_DRIVER_RESTORE};SERVER={SERVIDOR};UID={USUARIO};PWD={SENHA};DATABASE=master"
                with pyodbc.connect(conn_str, autocommit=True) as conn:
                    conn.cursor().execute(f"ALTER DATABASE [{db}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [{db}]")
                messagebox.showinfo("Sucesso", "Banco Apagado.")
                self.carregar_banco_atual()
            except Exception as e: messagebox.showerror("Erro", f"{e}")

def validar_caminhos():
    global CAMINHO_DO_ERP_CLIENTE, CAMINHO_DO_MAX_ATUALIZA, CAMINHO_BASE_MAX_BACKUP
    CAMINHO_DO_ERP_CLIENTE = os.path.join(PASTA_DO_SISTEMA, NOME_EXE_CLIENTE)
    CAMINHO_DO_MAX_ATUALIZA = os.path.join(PASTA_DO_SISTEMA, NOME_EXE_ATUALIZADOR)
    CAMINHO_BASE_MAX_BACKUP = os.path.join(PASTA_DO_SISTEMA, 'backup')
    
    erros = []
    if not os.path.isdir(PASTA_DO_SISTEMA): erros.append(f"Pasta Sistema: {PASTA_DO_SISTEMA}")
    if not os.path.isdir(PASTA_DAS_VERSOES): erros.append(f"Pasta Versões: {PASTA_DAS_VERSOES}")
    if not os.path.exists(CAMINHO_DO_7ZIP_EXE): erros.append(f"7-Zip: {CAMINHO_DO_7ZIP_EXE}")
    return erros

if __name__ == "__main__":
    app = GerenciadorMaxApp()
    while True:
        carregar_ou_criar_configuracoes(root_para_aviso=app)
        erros = validar_caminhos()
        if not erros: break
        messagebox.showerror("Configuração Necessária", "Caminhos inválidos:\n" + "\n".join(erros), parent=app)
        cw = ConfigWindow(app)
        app.wait_window(cw)
        if not cw.salvo: sys.exit()
    app.iniciar_interface()
    app.mainloop()