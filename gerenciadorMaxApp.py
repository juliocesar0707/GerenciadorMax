"""GerenciadorMax — Entry point do aplicativo.

Configura logging, carrega configurações, valida caminhos e inicia a UI.
"""

import logging
import sys
import os
from logging.handlers import RotatingFileHandler


def setup_logging():
    """Configura o sistema de logging com arquivo rotativo e console."""
    log_format = '%(asctime)s [%(levelname)s] %(name)s: %(message)s'

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # Console handler (INFO+)
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter(log_format))
    root_logger.addHandler(console)

    # File handler (DEBUG+, rotating 1MB x 3)
    # Caminho absoluto: empacotado, o diretório de trabalho pode ser
    # qualquer um, e o log acabaria espalhado fora da pasta do app.
    try:
        from app_config import diretorio_base
        file_handler = RotatingFileHandler(
            os.path.join(diretorio_base(), 'gerenciador_max.log'),
            maxBytes=1024 * 1024,
            backupCount=3,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter(log_format))
        root_logger.addHandler(file_handler)
    except OSError:
        pass  # Sem permissão de escrita, usa apenas console


def main():
    """Função principal do GerenciadorMax."""
    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("=== GerenciadorMax iniciando ===")

    from app_config import AppConfig
    from ui_config_window import ConfigWindow
    from ui_app import GerenciadorMaxApp
    from tkinter import messagebox

    # 1. Carregar configurações
    config = AppConfig.get()
    config.carregar()

    # 2. Criar janela principal (oculta inicialmente)
    app = GerenciadorMaxApp(config)

    # 3. Setup inicial na primeira execução
    if config.primeira_execucao:
        cw = ConfigWindow(app, config, is_first_run=True)
        app.wait_window(cw)
        if not cw.salvo:
            logger.info("Usuário cancelou o setup inicial")
            sys.exit()
        config.carregar()

    # 4. Validar caminhos (loop até corrigir)
    while True:
        erros = config.validar_caminhos()
        if not erros:
            break
        logger.warning("Caminhos inválidos: %s", erros)
        messagebox.showerror(
            "Configuração Necessária",
            "Caminhos inválidos:\n" + "\n".join(erros),
            parent=app
        )
        cw = ConfigWindow(app, config)
        app.wait_window(cw)
        if not cw.salvo:
            logger.info("Usuário cancelou a correção de caminhos")
            sys.exit()
        config.carregar()

    # 5. Iniciar interface
    logger.info("Configuração válida. Iniciando interface...")
    app.iniciar_interface()
    app.mainloop()
    logger.info("=== GerenciadorMax encerrado ===")


if __name__ == "__main__":
    main()