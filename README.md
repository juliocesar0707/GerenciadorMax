# GerenciadorMax

Ferramenta de desktop para o dia a dia de suporte do ERP **Maxdata (MAX)**: abre e
atualiza o sistema do cliente, troca o banco ativo, restaura backups no SQL Server
e baixa versões e backups direto da Cloud Maxdata — tudo em uma janela só, sem
alternar entre o SQL Server Management Studio, o Explorador de Arquivos e o
navegador.

Escrito em Python 3.12 com Tkinter/ttkbootstrap, empacotado como um único `.exe`
para rodar na máquina do cliente sem instalação de dependências.

---

## O que faz

A janela principal tem três colunas e um painel retrátil da nuvem.

**Manager (coluna 1)** — lista os arquivos `.rar` de versão presentes na pasta de
versões, com campo de busca.

- `▶ Abrir Sistema` — com uma versão selecionada, extrai o `.rar` sobre a pasta do
  sistema e abre o Manager. Sem seleção, compara a versão do executável com a
  versão gravada no banco e oferece a atualização se estiverem diferentes.
- `⚡ Atualizar Sistema` — extrai a versão selecionada e abre o `MAX_Atualiza.exe`.

Antes de qualquer extração o app verifica se o `MAX_manager2.exe` ou o
`MAX_Atualiza.exe` estão abertos e avisa — o 7-Zip não consegue sobrescrever um
executável em uso, e o erro que ele devolve sozinho não explica isso.

**Info Base (coluna 2)** — mostra o banco ativo lido do `max.ini`, a instância SQL
e a versão do sistema gravada no banco.

- Troca a instância SQL Server (detectadas via registro do Windows).
- Troca o banco ativo do cliente, gravando a chave no `max.ini`.
- Exclui bancos de dados — múltipla seleção, com o banco atualmente em uso
  protegido e confirmação por digitação da palavra `EXCLUIR`.

**Restaurador (coluna 3)** — lista os backups locais (`.max`, `.bak`, `.zip`,
`.rar`) ordenados do mais recente para o mais antigo.

Ao restaurar, o app extrai o arquivo se estiver compactado, localiza o `.MAX`/`.BAK`
lá dentro, cria uma pasta `dados{N}` na pasta do sistema, descobre os nomes lógicos
via `RESTORE FILELISTONLY` e executa o `RESTORE DATABASE` com `MOVE`. O progresso
aparece em um log na própria coluna.

**Cloud Maxdata (painel lateral)** — navegador WebDAV do Nextcloud com abas
separadas para Versões e Backups. Navega nas pastas, baixa o arquivo selecionado
para a pasta local correspondente e atualiza a lista local ao terminar. As
listagens ficam em cache por 2 minutos, já que o `PROPFIND` do Nextcloud é lento.

---

## Requisitos

| Item | Observação |
|---|---|
| Windows | O app usa `win32api`, o registro do Windows e `CREATE_NO_WINDOW` |
| Python 3.12+ | Somente para desenvolvimento — o `.exe` empacotado não precisa |
| 7-Zip | `7z.exe`, normalmente em `C:\Program Files\7-Zip\` |
| ODBC Driver 17 for SQL Server | Pacote da Microsoft, não vem pelo pip |
| Acesso ao SQL Server | Autenticação Windows para consultas; usuário/senha SQL para RESTORE e DROP |

---

## Instalação (desenvolvimento)

```powershell
git clone https://github.com/juliocesar0707/GerenciadorMax.git
cd GerenciadorMax

python -m venv .venv
.venv\Scripts\Activate.ps1

pip install -r requirements.txt      # execução
pip install -r requirements-dev.txt  # execução + PyInstaller

python gerenciadorMaxApp.py
```

---

## Configuração

Na primeira execução o app detecta os caminhos padrão (`C:\Max` ou `D:\Max`, o
7-Zip nos dois `Program Files`) e abre a tela de setup para confirmação. Depois
disso, tudo fica em `gerenciador_config.ini`, **ao lado do executável** — não no
diretório de trabalho, para que abrir o `.exe` a partir de outra pasta não faça o
app procurar a configuração no lugar errado.

Os caminhos são validados a cada inicialização; se algum não existir, o app abre a
tela de correção antes de montar a interface. O botão `⚙ Configurações` no rodapé
edita todos os campos.

| Seção | Conteúdo |
|---|---|
| `CAMINHOS` | Pasta do sistema, versões, backups, `max.ini` e `7z.exe` |
| `EXECUTAVEIS` | Nomes do executável do cliente e do atualizador |
| `SQL_LAUDO` | Driver ODBC e instância usados nas consultas de leitura |
| `SQL_RESTORE` | Servidor, usuário, senha e driver usados no RESTORE/DROP |
| `CONFIG_INI_MAX` | Seção e chaves lidas dentro do `max.ini` do cliente |
| `CLOUD` | URL, usuário e senha do Nextcloud |

### Sobre as senhas

As senhas são gravadas no `.ini` com **ofuscação base64** (prefixo `b64:`), o que
protege apenas contra leitura casual do arquivo — **não é criptografia** e não
protege contra quem tenha acesso à máquina. Valores gravados em texto puro por
versões anteriores continuam sendo lidos normalmente.

O `gerenciador_config.ini` está no `.gitignore` e **nunca deve ser versionado**.

---

## Testes

A suíte não depende de pytest nem da máquina real: cada teste monta uma árvore
temporária com a estrutura que o app espera (`Max/`, `Versoes/`, `backup/`,
`max.ini`, um `7z.exe` falso) e redireciona o caminho do arquivo de configuração,
de modo que nenhum teste toque no `gerenciador_config.ini` do usuário.

```powershell
python tests/run_all.py
```

| Arquivo | Cobre |
|---|---|
| `test_config.py` | Ofuscação, ancoragem do `.ini` na pasta do app, tela de setup |
| `test_launcher.py` | Fluxos de Abrir Sistema e Atualizar Sistema, com o 7-Zip e o `Popen` interceptados |
| `test_nuvem.py` | Painel WebDAV: falha de rede, pasta vazia, troca de credenciais |
| `test_sevenzip.py` | Diagnóstico das falhas do 7-Zip e detecção de executável em uso |
| `test_ui.py` | Montagem da janela, filtros das três listas, barra de progresso |

Um teste do `test_sevenzip.py` roda o 7-Zip de verdade e se pula sozinho quando
ele não está instalado.

---

## Build

```powershell
pyinstaller gerenciadorMaxApp.spec
```

O resultado é `dist/gerenciadorMaxApp.exe` — arquivo único, sem console. O `.spec`
declara os `hiddenimports` do `ttkbootstrap` e do `pywin32`, que o PyInstaller não
detecta sozinho.

---

## Estrutura

```
gerenciadorMaxApp.py   Entry point: logging, carga da config, validação, mainloop
app_config.py          AppConfig (dataclass singleton), ofuscação, constantes
ui_app.py              Janela principal: layout das 3 colunas, painel da nuvem, ações
ui_config_window.py    Tela de setup inicial e correção de caminhos
ui_theme.py            Paleta e estilos derivados da tela de login do MaxManager
ui_widgets.py          RoundedButton — botão de cantos arredondados desenhado em Canvas
sql_service.py         Conexões ODBC, listagem, RESTORE e DROP
ini_service.py         Leitura tolerante do max.ini (5 codificações + fallback regex)
sevenzip.py            Execução do 7-Zip com diagnóstico legível e detecção de arquivo travado
webdav_client.py       Cliente WebDAV do Nextcloud com cache
tests/                 Suíte sem pytest — python tests/run_all.py
```

Operações lentas (SQL, rede, 7-Zip) rodam em threads daemon; a atualização da
interface volta pela thread principal via `after()`, e o log do restore trafega por
uma `queue.Queue` lida a cada 100 ms.

---

## Log

`gerenciador_max.log`, ao lado do executável — rotativo, 1 MB × 3 arquivos, nível
DEBUG. O console recebe INFO e acima. É o primeiro lugar a olhar quando algo falha
na máquina do cliente.
