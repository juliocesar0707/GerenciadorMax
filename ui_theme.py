"""Paleta e estilos visuais do GerenciadorMax.

Tema claro derivado da tela de login do MaxManager: fundo cinza, cartões
brancos, rótulos em cinza médio, botão grafite e o vermelho do logo Maxdata
como cor de acento/destaque.
"""

THEME_NAME = "gerenciadormax"

# --- Paleta base (tela de login do MaxManager) ---
BG = "#D9DADD"          # cinza do fundo da janela
CARD = "#FFFFFF"        # cartões brancos
SURFACE = "#F5F6F8"     # listas e campos
BORDER = "#D8DBE1"      # bordas suaves
BORDER_FORTE = "#C6CAD3"

GRAFITE = "#3A3F51"     # botão "Confirmar login"
GRAFITE_HOVER = "#4B5166"
GRAFITE_PRESS = "#2F3342"

VERMELHO = "#D0202B"    # o "maX" do logo
VERMELHO_HOVER = "#E03A44"
VERMELHO_PRESS = "#B01822"

FG = "#2B2F3A"          # texto principal
FG_MUTED = "#6E7481"    # rótulos secundários
FG_BODY = "#3C414D"     # texto de listas

DISABLED_BG = "#E4E6EA"
DISABLED_FG = "#A6ABB6"

FONT_FAMILY = "Segoe UI"

# Variantes usadas por ui_widgets.RoundedButton
VARIANTES_BOTAO = {
    "primary": {
        "bg": GRAFITE, "hover": GRAFITE_HOVER, "pressed": GRAFITE_PRESS,
        "fg": "#FFFFFF", "border": GRAFITE,
        "disabled": DISABLED_BG, "disabled_fg": DISABLED_FG,
    },
    "danger": {
        "bg": VERMELHO, "hover": VERMELHO_HOVER, "pressed": VERMELHO_PRESS,
        "fg": "#FFFFFF", "border": VERMELHO,
        "disabled": DISABLED_BG, "disabled_fg": DISABLED_FG,
    },
    "outline": {
        "bg": CARD, "hover": "#F0F1F4", "pressed": "#E4E6EA",
        "fg": GRAFITE, "border": BORDER_FORTE,
        "disabled": DISABLED_BG, "disabled_fg": DISABLED_FG,
    },
    "outline-danger": {
        "bg": CARD, "hover": "#FCEEEF", "pressed": "#F7DDDF",
        "fg": VERMELHO, "border": VERMELHO,
        "disabled": DISABLED_BG, "disabled_fg": DISABLED_FG,
    },
}

USER_THEME = {
    "type": "light",
    "colors": {
        "primary": GRAFITE,
        "secondary": FG_MUTED,
        "success": "#2E9E5B",
        "info": "#3A7BD5",
        "warning": "#C8891F",
        "danger": VERMELHO,
        "light": SURFACE,
        "dark": GRAFITE,
        "bg": BG,
        "fg": FG,
        "selectbg": GRAFITE,
        "selectfg": "#FFFFFF",
        "border": BORDER,
        "inputfg": FG,
        "inputbg": CARD,
        "active": "#EDEEF1",
    },
}


def registrar_tema():
    """Registra o tema no ttkbootstrap. Deve rodar ANTES de criar a Window."""
    from ttkbootstrap.themes.user import USER_THEMES
    USER_THEMES[THEME_NAME] = USER_THEME


def aplicar_estilos(style):
    """Cria os estilos derivados: cartões brancos, campos e tabelas claras."""

    # --- Superfícies ---
    style.configure("Card.TFrame", background=CARD, borderwidth=0)
    style.configure("Surface.TFrame", background=SURFACE, borderwidth=0)

    # --- Campos de texto (brancos, borda suave, como no login) ---
    style.configure(
        "Campo.TEntry",
        fieldbackground=CARD,
        foreground=FG,
        insertcolor=FG,
        bordercolor=BORDER_FORTE,
        lightcolor=BORDER_FORTE,
        darkcolor=BORDER_FORTE,
        borderwidth=1,
        padding=8,
    )
    style.map(
        "Campo.TEntry",
        bordercolor=[("focus", GRAFITE)],
        lightcolor=[("focus", GRAFITE)],
        darkcolor=[("focus", GRAFITE)],
    )

    # --- Rótulos sobre cartão branco ---
    style.configure("Card.TLabel", background=CARD, foreground=FG)
    style.configure("CardMuted.TLabel", background=CARD, foreground=FG_MUTED,
                    font=(FONT_FAMILY, 9))
    style.configure("CardValue.TLabel", background=CARD, foreground=GRAFITE,
                    font=(FONT_FAMILY, 12, "bold"))
    style.configure("CardWarn.TLabel", background=CARD, foreground=VERMELHO,
                    font=(FONT_FAMILY, 12, "bold"))
    style.configure("CardTitle.TLabel", background=CARD, foreground=FG,
                    font=(FONT_FAMILY, 12, "bold"))

    # --- Rótulos direto sobre o fundo cinza ---
    style.configure("Muted.TLabel", background=BG, foreground=FG_MUTED,
                    font=(FONT_FAMILY, 9))
    style.configure("ColTitle.TLabel", background=BG, foreground=FG,
                    font=(FONT_FAMILY, 17, "bold"))

    # --- Tabelas claras ---
    style.configure(
        "Claro.Treeview",
        background=CARD,
        fieldbackground=CARD,
        foreground=FG_BODY,
        bordercolor=CARD,
        lightcolor=CARD,
        darkcolor=CARD,
        borderwidth=0,
        relief="flat",
        rowheight=26,
        font=(FONT_FAMILY, 9),
    )
    style.layout("Claro.Treeview", [
        ("Claro.Treeview.treearea", {"sticky": "nswe"}),
    ])
    style.configure(
        "Claro.Treeview.Heading",
        background=SURFACE,
        foreground=FG_MUTED,
        bordercolor=BORDER,
        lightcolor=SURFACE,
        darkcolor=SURFACE,
        relief="flat",
        borderwidth=0,
        padding=8,
        font=(FONT_FAMILY, 9, "bold"),
    )
    style.map(
        "Claro.Treeview",
        background=[("selected", GRAFITE)],
        foreground=[("selected", "#FFFFFF")],
    )
    style.map(
        "Claro.Treeview.Heading",
        background=[("active", "#EDEEF1")],
        foreground=[("active", FG)],
    )

    # --- Barra de progresso do restore ---
    style.configure(
        "Restore.Horizontal.TProgressbar",
        background=GRAFITE,
        troughcolor=SURFACE,
        bordercolor=SURFACE,
        lightcolor=GRAFITE,
        darkcolor=GRAFITE,
        borderwidth=0,
        thickness=6,
    )


def ajustar_estilos_derivados(style):
    """Corrige os estilos que o ttkbootstrap gera sozinho.

    O ttkbootstrap constrói os estilos de scrollbar/combobox sob demanda, na
    criação do primeiro widget que os usa, e nesse momento sobrescreve o que
    tivéssemos configurado antes. Por isso estes ajustes rodam ao FIM do
    layout, e não junto com `aplicar_estilos`:

    - scrollbar: a cor sai de um ajuste de HSV sobre o bootstyle, o que
      destoa do cinza discreto que o tema claro pede;
    - combobox readonly: o ttkbootstrap pinta o campo inteiro com a cor
      primária, deixando a caixa chapada.
    """
    for nome in ("secondary.Round.Vertical.TScrollbar",
                 "secondary.Vertical.TScrollbar"):
        try:
            style.configure(
                nome,
                troughcolor=SURFACE,
                background=BORDER_FORTE,
                bordercolor=SURFACE,
                lightcolor=BORDER_FORTE,
                darkcolor=BORDER_FORTE,
                arrowcolor=FG_MUTED,
                borderwidth=0,
            )
            style.map(nome, background=[("active", FG_MUTED),
                                        ("pressed", GRAFITE)])
        except Exception:  # estilo ainda não construído
            continue

    for nome in ("dark.TCombobox", "primary.TCombobox", "TCombobox"):
        try:
            style.map(
                nome,
                fieldbackground=[("readonly", CARD), ("disabled", SURFACE)],
                background=[("readonly", CARD), ("disabled", SURFACE)],
                selectbackground=[("readonly", CARD)],
                selectforeground=[("readonly", FG)],
                foreground=[("readonly", FG), ("disabled", FG_MUTED)],
                bordercolor=[("focus", GRAFITE), ("readonly", BORDER_FORTE)],
                arrowcolor=[("readonly", FG_MUTED), ("active", GRAFITE)],
            )
        except Exception:
            continue
