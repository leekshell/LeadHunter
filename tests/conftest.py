"""conftest: permite importar App.py em ambientes sem tkinter/display.

O App.py é uma aplicação GUI (customtkinter). Em sandboxes headless sem o
pacote tkinter do Python, instalamos um stub dinâmico de `tkinter` apenas o
suficiente para o import do módulo funcionar. Nada disso afeta ambientes com
tkinter real (Linux desktop/Windows/macOS).
"""

import sys
import types


class _DummyWidget:
    """Base genérica: subclassável, aceita qualquer argumento."""

    def __init__(self, *args, **kwargs):
        pass

    def __getattr__(self, name):
        def _noop(*args, **kwargs):
            return None
        return _noop


def _make_module(name, is_package=False):
    mod = types.ModuleType(name)

    def _mod_getattr(attr):
        if attr.startswith("__"):
            if attr == "__path__" and is_package:
                return []
            raise AttributeError(attr)
        if attr == "TclError":
            return type("TclError", (Exception,), {})
        return type(attr, (_DummyWidget,), {})

    mod.__getattr__ = _mod_getattr
    return mod


def _install_fake_tkinter():
    try:
        import tkinter  # noqa: F401
        return  # tkinter real disponível: nada a fazer
    except ImportError:
        pass

    tk = _make_module("tkinter", is_package=True)
    constants = _make_module("tkinter.constants")
    # constantes reais do tkinter.constants (valores irrelevantes p/ import)
    _consts = {
        "N": "n", "S": "s", "E": "e", "W": "w", "NW": "nw", "NE": "ne",
        "SW": "sw", "SE": "se", "NS": "ns", "EW": "ew", "NSEW": "nsew",
        "CENTER": "center", "TOP": "top", "BOTTOM": "bottom", "LEFT": "left",
        "RIGHT": "right", "BOTH": "both", "X": "x", "Y": "y", "NONE": "none",
        "HORIZONTAL": "horizontal", "VERTICAL": "vertical", "NORMAL": "normal",
        "DISABLED": "disabled", "ACTIVE": "active", "RAISED": "raised",
        "FLAT": "flat", "SUNKEN": "sunken", "GROOVE": "groove", "RIDGE": "ridge",
        "SOLID": "solid", "END": "end", "ANCHOR": "anchor", "TRUE": 1,
        "FALSE": 0, "WORD": "word", "CHAR": "char", "INSERT": "insert",
        "SEL": "sel", "FIRST": "first", "LAST": "last", "DOTBOX": "dotbox",
        "UNDERLINE": "underline", "NUMERIC": "numeric", "UNITS": "units",
        "PAGES": "pages", "MOVETO": "moveto", "SCROLL": "scroll",
        "COMMAND": "command", "ON": 1, "OFF": 0, "YES": 1, "NO": 0,
        "ARC": "arc", "CHORD": "chord", "PIESLICE": "pieslice",
        "BASELINE": "baseline", "CURRENT": "current", "HIDDEN": "hidden",
        "ROUND": "round", "PROJECTING": "projecting", "MITER": "miter",
        "BEVEL": "bevel", "BUTT": "butt", "OUTSIDE": "outside",
        "INSIDE": "inside", "EXTENDED": "extended", "BROWSE": "browse",
        "SINGLE": "single", "MULTIPLE": "multiple",
    }
    for _name, _value in _consts.items():
        setattr(constants, _name, _value)
    constants.__all__ = list(_consts)
    filedialog = _make_module("tkinter.filedialog")
    font_mod = _make_module("tkinter.font")
    ttk = _make_module("tkinter.ttk")
    messagebox = _make_module("tkinter.messagebox")

    tk.constants = constants
    tk.filedialog = filedialog
    tk.font = font_mod
    tk.ttk = ttk
    tk.messagebox = messagebox

    sys.modules["tkinter"] = tk
    sys.modules["tkinter.constants"] = constants
    sys.modules["tkinter.filedialog"] = filedialog
    sys.modules["tkinter.font"] = font_mod
    sys.modules["tkinter.ttk"] = ttk
    sys.modules["tkinter.messagebox"] = messagebox


_install_fake_tkinter()
