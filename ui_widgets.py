"""Widgets customizados do GerenciadorMax.

O ttk não desenha cantos arredondados: os temas usam elementos retangulares e
`borderwidth` não vira raio. Para chegar ao botão do MaxManager (retângulo de
cantos arredondados, preenchimento sólido, texto branco em negrito) o caminho é
desenhar num Canvas — é o que `RoundedButton` faz.
"""

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

import ui_theme


def _pontos_retangulo_arredondado(x1, y1, x2, y2, r):
    """Pontos de um retângulo arredondado para `create_polygon(smooth=True)`.

    Cada canto repete o vértice para que a spline "puxe" a curva até ele,
    resultando num arco em vez de um chanfro.
    """
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + r, y1,
        x2 - r, y1, x2, y1,
        x2, y1 + r,
        x2, y2 - r, x2, y2,
        x2 - r, y2,
        x1 + r, y2, x1, y2,
        x1, y2 - r,
        x1, y1 + r, x1, y1,
    ]


def cor_da_superficie(widget):
    """Descobre a cor de fundo do container, para o Canvas não vazar um retângulo.

    Widgets ttk não expõem `background` via `cget`, então a cor vem do estilo.
    """
    try:
        estilo = widget.cget("style") or widget.winfo_class()
        cor = ttk.Style().lookup(estilo, "background")
        if cor:
            return cor
    except tk.TclError:
        pass
    try:
        return widget.cget("background")
    except tk.TclError:
        return ui_theme.BG


class RoundedButton(tk.Canvas):
    """Botão retangular de cantos arredondados, desenhado em Canvas.

    Reproduz a API que o resto da UI usa (`config(state=...)`,
    `config(text=...)`, geometria via pack), então substitui um
    `ttk.Button` sem alterar quem o chama.
    """

    def __init__(self, parent, text="", command=None, variant="primary",
                 radius=9, padx=18, pady=11, font=None, surface=None, **kwargs):
        self._surface = surface or cor_da_superficie(parent)
        super().__init__(
            parent, highlightthickness=0, bd=0, bg=self._surface,
            takefocus=0, **kwargs
        )

        self._text = text
        self._command = command
        self._cores = ui_theme.VARIANTES_BOTAO[variant]
        self._radius = radius
        self._padx = padx
        self._pady = pady
        self._font = font or tkfont.Font(
            family=ui_theme.FONT_FAMILY, size=9, weight="bold"
        )

        self._hover = False
        self._pressed = False
        self._enabled = True

        # Texto pode ser multilinha (a aba vertical da nuvem é um caso):
        # medir só a string inteira daria uma largura enorme e altura de 1 linha.
        medidor = tkfont.Font(font=self._font)
        linhas = text.split("\n") or [""]
        largura = max(medidor.measure(l) for l in linhas) + 2 * padx
        altura = medidor.metrics("linespace") * len(linhas) + 2 * pady
        self.configure(width=largura, height=altura)

        self.bind("<Configure>", self._redesenhar)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    # --- estado visual ---
    def _cor_atual(self):
        if not self._enabled:
            return self._cores["disabled"], self._cores["disabled_fg"]
        if self._pressed:
            return self._cores["pressed"], self._cores["fg"]
        if self._hover:
            return self._cores["hover"], self._cores["fg"]
        return self._cores["bg"], self._cores["fg"]

    def _redesenhar(self, event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1 or h <= 1:
            return

        fundo, texto = self._cor_atual()
        borda = self._cores.get("border") or fundo

        self.create_polygon(
            _pontos_retangulo_arredondado(1, 1, w - 1, h - 1, self._radius),
            smooth=True, splinesteps=36,
            fill=fundo, outline=borda, width=1,
        )
        self.create_text(
            w / 2, h / 2, text=self._text, fill=texto,
            font=self._font, justify="center",
        )

    # --- eventos ---
    def _on_enter(self, _=None):
        if self._enabled:
            self._hover = True
            self.configure(cursor="hand2")
            self._redesenhar()

    def _on_leave(self, _=None):
        self._hover = False
        self._pressed = False
        self._redesenhar()

    def _on_press(self, _=None):
        if self._enabled:
            self._pressed = True
            self._redesenhar()

    def _on_release(self, event=None):
        if not (self._enabled and self._pressed):
            return
        self._pressed = False
        self._redesenhar()
        # Só dispara se o cursor ainda estiver sobre o botão, como um ttk.Button
        dentro = (0 <= event.x <= self.winfo_width()
                  and 0 <= event.y <= self.winfo_height())
        if dentro and self._command:
            self._command()

    # --- API compatível com ttk.Button ---
    def configure(self, cnf=None, **kw):
        """Aceita `state` e `text` como um ttk.Button; o resto vai ao Canvas."""
        if cnf:
            kw.update(cnf)

        redesenhar = False
        if "state" in kw:
            self._enabled = kw.pop("state") not in ("disabled", tk.DISABLED)
            if not self._enabled:
                self._hover = self._pressed = False
            redesenhar = True
        if "text" in kw:
            self._text = kw.pop("text")
            redesenhar = True
        if "command" in kw:
            self._command = kw.pop("command")

        resultado = super().configure(**kw) if kw else None
        if redesenhar:
            self._redesenhar()
        return resultado

    config = configure

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "state":
            return "normal" if self._enabled else "disabled"
        return super().cget(key)

    def invoke(self):
        """Dispara o comando — usado pelos testes."""
        if self._enabled and self._command:
            return self._command()
