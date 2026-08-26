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
