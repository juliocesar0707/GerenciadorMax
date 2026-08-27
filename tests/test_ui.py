"""Montagem da janela principal, filtros e widgets que ja quebraram."""
import apoio
from apoio import AmbienteFalso, criar_app, linhas

import ui_app


def test_layout_monta_com_todos_os_widgets():
    """Cobre atributos cuja ausencia ja causou AttributeError em producao."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        for atributo in ("btn_restore", "var_busca_banco", "var_busca_versao",
                         "var_busca_backup", "cloud_panel", "lb_nuvem_versoes",
                         "lb_nuvem_backups", "lbl_caminho_versoes",
                         "lbl_caminho_backups", "log_wrap", "progress"):
            assert hasattr(app, atributo), f"falta {atributo}"
    finally:
        apoio.destruir(app); amb.fechar()


def test_filtros_das_tres_listas():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._all_versoes = ["2.4.152.101_Max_Manager.RAR", "2.4.151.110_Max_Manager.RAR"]
        app._all_backups = ["MAX-Manager_FORTUP_10082026.MAX", "outro.zip"]
        app._all_dbs = ["Max_FortupFiscal", "Max_Teste"]
        app._filtrar_versoes(); app._filtrar_backups(); app._filtrar_bancos()
        app.update()

        assert len(linhas(app.lb_versoes)) == 2
        assert len(linhas(app.lb_backups)) == 2
        assert len(linhas(app.lb_tools)) == 2

        app.var_busca_versao.set("152")
        app.var_busca_banco.set("fortup")
        app.var_busca_backup.set("zip")
        app.update()

        assert len(linhas(app.lb_versoes)) == 1, "filtro de versoes nao filtrou"
        assert len(linhas(app.lb_tools)) == 1, "filtro de bancos nao filtrou"
        assert len(linhas(app.lb_backups)) == 1, "filtro de backups nao filtrou"
    finally:
        apoio.destruir(app); amb.fechar()


def test_atualizar_ui_sql_preenche_sem_estourar():
    """Este caminho quebrava por var_busca_banco inexistente."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._atualizar_ui_sql(["Max_A", "Max_B"], "Max_A", "2.4.152.101")
        app.update()

        assert app.lbl_db_atual.cget("text") == "Max_A"
        assert app.lbl_versao_sql.cget("text") == "2.4.152.101"
        assert app.combo_db.get() == "Max_A"
        assert len(linhas(app.lb_tools)) == 2
    finally:
        apoio.destruir(app); amb.fechar()


def test_painel_da_nuvem_abre_e_fecha():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app.webdav_versoes.url = ""      # evita qualquer requisicao real
        app.webdav_backups.url = ""

        # a janela fica withdrawn nos testes, entao winfo_ismapped e sempre 0;
        # winfo_manager diz se o widget esta empacotado no layout.
        app._toggle_cloud(); app.update()
        assert app._cloud_visivel, "estado nao mudou"
        assert app.cloud_panel.winfo_manager() == "pack", "painel nao entrou no layout"

        app._toggle_cloud(); app.update()
        assert not app._cloud_visivel, "nao fechou"
        assert app.cloud_panel.winfo_manager() == "", "painel continuou no layout"
    finally:
        apoio.destruir(app); amb.fechar()


def test_barra_de_progresso_so_aparece_durante_o_restore():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        assert app.progress.winfo_manager() == "", "progresso visivel em repouso"

        app._mostrar_progresso(True); app.update()
        assert app.progress.winfo_manager() == "pack", "progresso nao apareceu"
        assert str(app.btn_restore.cget("state")) == "disabled"

        app._mostrar_progresso(False); app.update()
        assert app.progress.winfo_manager() == "", "progresso nao sumiu"
        assert str(app.btn_restore.cget("state")) == "normal"
    finally:
        apoio.destruir(app); amb.fechar()


def test_janela_de_configuracoes_abre_com_todos_os_campos():
    """bstrap.Toplevel recebe `title` no 1o posicional: passar o pai ali
    impedia esta janela de abrir."""
    from app_config import CONFIG_SECTIONS_UI

    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._abrir_configuracoes()
        app.update()

        esperados = sum(len(chaves) for _, chaves, _ in CONFIG_SECTIONS_UI)
        assert len(app.cfg_vars) == esperados, f"{len(app.cfg_vars)} de {esperados}"

        for filho in app.winfo_children():
            if filho.winfo_class() == "Toplevel":
                filho.destroy()
        app.update()
    finally:
        apoio.destruir(app); amb.fechar()


def test_sql_fora_do_ar_nao_se_confunde_com_lista_vazia():
    """Antes, falha de conexao e "nenhum banco" davam a mesma tela vazia."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._atualizar_ui_sql([], "Max_Teste", "---",
                              erro="Login recusado: usuario ou senha do SQL incorretos.")
        app.update()

        assert "sem conexao" in app.lb_tools.heading("banco")["text"].replace("ã", "a")
        assert "Login recusado" in app.status.cget("text")

        # e volta ao normal quando a conexao volta
        app._atualizar_ui_sql(["Max_A"], "Max_A", "2.4.152.101")
        app.update()
        assert app.lb_tools.heading("banco")["text"] == "Bases de Dados"
    finally:
        apoio.destruir(app); amb.fechar()


def test_selecionar_backup_sugere_o_nome_do_banco():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._all_backups = ["MAX-Manager_FORTUP_10082026.MAX", "Max_CLIENTE_01092026.zip"]
        app._filtrar_backups()
        app.update()

        filhos = app.lb_backups.get_children()
        app.lb_backups.selection_set(filhos[0])
        app.update()
        assert app.entry_new_db.get() == "FORTUP", app.entry_new_db.get()

        # trocar de backup troca a sugestao
        app.lb_backups.selection_set(filhos[1])
        app.update()
        assert app.entry_new_db.get() == "CLIENTE", app.entry_new_db.get()
    finally:
        apoio.destruir(app); amb.fechar()


def test_nome_digitado_a_mao_nao_e_sobrescrito():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._all_backups = ["MAX-Manager_FORTUP_10082026.MAX", "Max_CLIENTE_01092026.zip"]
        app._filtrar_backups()
        app.update()

        from tkinter import END
        app.entry_new_db.delete(0, END)
        app.entry_new_db.insert(0, "NOME_MEU")

        app.lb_backups.selection_set(app.lb_backups.get_children()[0])
        app.update()

        assert app.entry_new_db.get() == "NOME_MEU", "descartou o que foi digitado"
    finally:
        apoio.destruir(app); amb.fechar()


def test_barra_troca_de_indeterminada_para_percentual():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._mostrar_progresso(True); app.update()
        assert str(app.progress.cget("mode")) == "indeterminate"

        app._progresso_percentual(42.0); app.update()
        assert str(app.progress.cget("mode")) == "determinate"
        assert float(app.progress.cget("value")) == 42.0
        assert "42%" in app.status.cget("text"), app.status.cget("text")

        # a operacao seguinte recomeca indeterminada
        app._mostrar_progresso(False)
        app._mostrar_progresso(True); app.update()
        assert str(app.progress.cget("mode")) == "indeterminate"
    finally:
        apoio.destruir(app); amb.fechar()


def test_percentual_chega_pela_fila_do_restore():
    """O SQL reporta o progresso numa thread; a barra so pode ser tocada na
    thread da UI, entao o valor viaja pela mesma fila do log."""
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        app._mostrar_progresso(True)
        app._publicar_percentual(73.4)
        app.process_queue()
        app.update()

        assert str(app.progress.cget("mode")) == "determinate"
        assert round(float(app.progress.cget("value"))) == 73
    finally:
        apoio.destruir(app); amb.fechar()


def test_backup_sem_banco_valido_apenas_avisa():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        vistos = apoio.silenciar_dialogos(ui_app)
        chamados = []
        app.sql.backup_database = lambda *a, **k: chamados.append(a)

        app.lbl_db_atual.config(text="INI NAO ENCONTRADO")
        app.update()
        app._gerar_backup()
        app.update()

        assert not chamados, "nao deveria tentar backup sem banco valido"
        assert vistos["info"], "deveria avisar"
    finally:
        apoio.restaurar_dialogos(ui_app)
        apoio.destruir(app); amb.fechar()


def test_listas_locais_leem_do_disco():
    amb = AmbienteFalso()
    app = criar_app(amb)
    try:
        amb.criar_versoes("a_152.rar", "b_151.rar", "ignorar.txt")
        amb.criar_backups("cliente_1.MAX", "cliente_2.zip", "ignorar.doc")

        app._popular_versoes_locais()
        app._load_backups()
        app.update()

        assert sorted(app._all_versoes) == ["a_152.rar", "b_151.rar"], app._all_versoes
        assert sorted(app._all_backups) == ["cliente_1.MAX", "cliente_2.zip"], app._all_backups
    finally:
        apoio.destruir(app); amb.fechar()
