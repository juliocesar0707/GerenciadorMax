"""Serviço de leitura/escrita robusta de arquivos INI (max.ini)."""

import configparser
import os
import re
import logging

logger = logging.getLogger(__name__)


class IniService:
    """Manipulação robusta de arquivos INI com suporte a múltiplas codificações."""

    ENCODINGS = ['utf-8-sig', 'utf-8', 'cp1252', 'latin-1', 'utf-16']

    @staticmethod
    def criar_parser():
        """Cria um RawConfigParser configurado para ler INI do Windows corretamente.

        Desabilita tratamento de ';' como comentário inline (comum em connection strings)
        e preserva maiúsculas/minúsculas nas chaves.
        """
        c = configparser.RawConfigParser(
            comment_prefixes=(),
            inline_comment_prefixes=(),
            delimiters=('=',),
            strict=False
        )
        c.optionxform = str
        return c

    @classmethod
    def ler_arquivo(cls, caminho):
        """Lê um arquivo INI tentando múltiplas codificações.

        Returns:
            configparser.RawConfigParser com o conteúdo do arquivo.
        """
        for enc in cls.ENCODINGS:
            try:
                c = cls.criar_parser()
                c.read(caminho, encoding=enc)
                if c.sections():
                    return c
            except Exception:
                continue

        logger.warning("Não foi possível ler INI '%s' com nenhuma codificação", caminho)
        return cls.criar_parser()

    @staticmethod
    def get_value(config, section, key):
        """Obtém valor do INI de forma case-insensitive.

        Returns:
            str ou None se não encontrado.
        """
        if not config.has_section(section):
            return None
        for opt in config.options(section):
            if opt.strip().lower() == key.strip().lower():
                return config.get(section, opt).strip()
        return None

    @staticmethod
    def set_value(config, section, key, value):
        """Define um valor no INI, preservando a capitalização original se a chave já existir."""
        if not key:
            return
        if not config.has_section(section):
            config.add_section(section)
        
        existing_key = key
        for opt in config.options(section):
            if opt.strip().lower() == key.strip().lower():
                existing_key = opt
                break
                
        config.set(section, existing_key, str(value))

    @staticmethod
    def salvar(config, caminho):
        """Salva o ConfigParser no arquivo."""
        try:
            with open(caminho, 'w', encoding='windows-1252') as f:
                config.write(f, space_around_delimiters=False)
            logger.debug("INI salvo: %s", caminho)
        except OSError as e:
            logger.error("Erro ao salvar INI '%s': %s", caminho, e)
            raise

    @classmethod
    def ler_banco_e_servidor(cls, caminho, section, key_db, key_server):
        """Lê banco de dados e servidor do max.ini com fallback de regex.

        Returns:
            Tupla (banco, servidor) — valores podem ser None.
        """
        banco = None
        servidor = None

        if not os.path.exists(caminho):
            logger.warning("Arquivo INI não encontrado: %s", caminho)
            return "INI NÃO ENCONTRADO", None

        # Estratégia 1: ConfigParser robusto
        try:
            config = cls.ler_arquivo(caminho)
            banco = cls.get_value(config, section, key_db)
            servidor = cls.get_value(config, section, key_server)
        except Exception as e:
            logger.warning("ConfigParser falhou no INI: %s", e)

        # Estratégia 2: Fallback com regex (parsing manual)
        if not banco:
            logger.debug("Tentando fallback com regex para INI '%s'", caminho)
            try:
                content = None
                for enc in cls.ENCODINGS:
                    try:
                        with open(caminho, 'r', encoding=enc) as f:
                            content = f.read()
                        break
                    except (UnicodeDecodeError, UnicodeError):
                        continue

                if content:
                    match = re.search(
                        rf'{re.escape(key_db)}\s*=\s*([^\r\n;]+)',
                        content, re.IGNORECASE
                    )
                    if match:
                        banco = match.group(1).strip()

                    match_srv = re.search(
                        rf'{re.escape(key_server)}\s*=\s*([^\r\n;]+)',
                        content, re.IGNORECASE
                    )
                    if match_srv:
                        servidor = match_srv.group(1).strip()
            except Exception as e:
                logger.warning("Fallback regex falhou: %s", e)

        if not banco:
            return "ERRO LER INI", servidor

        return banco, servidor
