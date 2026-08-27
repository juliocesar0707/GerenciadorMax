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
- **Gera backup** do banco selecionado (ou do banco ativo, se nada estiver
  selecionado) direto na pasta de backups, com data e hora no nome do arquivo.
  Backups sucessivos do mesmo banco não se sobrescrevem, e o `.bak` já aparece
  na lista do Restaurador ao terminar.
- Exclui bancos de dados — múltipla seleção, com o banco atualmente em uso
  protegido e confirmação por digitação da palavra `EXCLUIR`.

Quando o SQL Server não responde, a lista de bancos não fica apenas vazia: o
cabeçalho passa a marcar "sem conexão" e a barra de status explica o motivo
(login recusado, servidor inacessível, driver ausente).

**Restaurador (coluna 3)** — lista os backups locais (`.max`, `.bak`, `.zip`,
`.rar`) ordenados do mais recente para o mais antigo.

Ao selecionar um backup, o nome do banco é sugerido a partir do nome do arquivo
(`MAX-Manager_FORTUP_10082026.MAX` → `FORTUP`); um nome digitado à mão nunca é
sobrescrito pela sugestão. Se o nome escolhido já pertencer a um banco existente,
o app pergunta antes de continuar — o `RESTORE` roda com `REPLACE`.

O restore extrai o arquivo se estiver compactado, localiza o `.MAX`/`.BAK` lá
dentro, cria uma pasta `dados{N}` na pasta do sistema, descobre os nomes lógicos
via `RESTORE FILELISTONLY` e executa o `RESTORE DATABASE` com `MOVE`. A barra de
progresso mostra o **percentual real** publicado pelo SQL Server em
`sys.dm_exec_requests` — sem a permissão `VIEW SERVER STATE` ela simplesmente
segue indeterminada. As mensagens de etapa aparecem no log da própria coluna.

**Cloud Maxdata (painel lateral)** — navegador WebDAV do Nextcloud com abas
separadas para Versões e Backups, mostrando nome, tamanho e data de modificação
de cada arquivo. Navega nas pastas, baixa o selecionado para a pasta local
correspondente e atualiza a lista local ao terminar.

O download escreve em um `.part` e só renomeia ao concluir, então uma queda de
conexão não deixa para trás um arquivo truncado com o nome definitivo. Se o
arquivo já existir localmente, o app pergunta antes de baixar de novo.

### Por que a nuvem responde rápido

O `PROPFIND` do Nextcloud custa caro, e quatro decisões atacam essa espera:

| Decisão | Efeito |
|---|---|
| Conexão HTTP persistente | O `urllib` abre socket novo e manda `Connection: close` a cada chamada; aqui a conexão fica de pé e uma queda por ociosidade é refeita sozinha. Vários cliques, um handshake TLS. |
| `PROPFIND` dirigido | Pede só `resourcetype`, `getcontentlength` e `getlastmodified`. Sem corpo, o pedido equivale a `allprop` e o servidor monta todas as propriedades de todos os filhos. |
| Cache compartilhado por caminho | As duas abas veem a mesma pasta e só diferem no filtro de extensão — o cache guarda a listagem crua e as duas aproveitam a mesma busca. |
| Exibir do cache e atualizar por baixo | Pasta já visitada aparece na hora; a atualização acontece em silêncio. Se ela falhar, o que está na tela continua valendo, com o aviso na barra de status. |

Além disso, a raiz da nuvem é listada em segundo plano no arranque do app, de
modo que o painel normalmente abre já preenchido. Só o que nunca foi visitado
paga a ida ao servidor.

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

As seções **SQL Restore** e **Cloud Nuvem** têm um botão `Testar conexão` que usa
os valores digitados na hora, sem precisar salvar antes — útil justamente quando
se está corrigindo credenciais.

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
```

Operações lentas (SQL, rede, 7-Zip) rodam em threads daemon; a atualização da
interface volta pela thread principal via `after()`, e o log do restore trafega por
uma `queue.Queue` lida a cada 100 ms.

---

## Log

`gerenciador_max.log`, ao lado do executável — rotativo, 1 MB × 3 arquivos, nível
DEBUG. O console recebe INFO e acima. É o primeiro lugar a olhar quando algo falha
na máquina do cliente.
