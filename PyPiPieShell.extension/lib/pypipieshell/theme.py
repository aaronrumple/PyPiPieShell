# -*- coding: utf-8 -*-
"""Revit light/dark theme detection and the PyPiPie Shell palettes.

* Dark is only reported on Revit 2024+, where the UI is genuinely dark.
  Revit 2021-2023 expose a "Dark" setting but keep a light ribbon and window
  chrome, so the shell stays light there (same rule pyRevit 7 uses).
* The chrome colours are pyRevit 7's own palette, and the shell also publishes
  them under pyRevit's resource keys (``pyRevit<Name>Brush``) so pyRevit's
  Theme.xaml control styles pick up the same values.
* Editor colours follow pyRevit 7's Python Shell (``#222933`` editor surface
  in dark mode).
"""
import clr

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import SystemColors  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402


FIRST_DARK_UI_YEAR = 2024


def revit_year(uiapp):
    try:
        return int(uiapp.Application.VersionNumber)
    except Exception:
        return 0


def is_revit_dark(uiapp=None):
    """True when Revit's active UI theme is Dark (Revit 2024+ only)."""
    try:
        if uiapp is not None:
            year = revit_year(uiapp)
            if 0 < year < FIRST_DARK_UI_YEAR:
                return False
        from Autodesk.Revit.UI import UIThemeManager, UITheme
        return UIThemeManager.CurrentTheme == UITheme.Dark
    except Exception:
        return False


def subscribe_theme_changed(uiapp, callback):
    """Subscribe ``callback()`` to ``UIApplication.ThemeChanged`` (2024+).

    Returns an ``unsubscribe()`` function (a no-op when unsupported).
    """
    if uiapp is None:
        return lambda: None

    def handler(sender, args):
        try:
            callback()
        except Exception:
            pass

    try:
        uiapp.ThemeChanged += handler
    except Exception:
        return lambda: None

    def unsubscribe():
        try:
            uiapp.ThemeChanged -= handler
        except Exception:
            pass

    return unsubscribe


def _c(hexstr):
    hexstr = hexstr.lstrip("#")
    if len(hexstr) == 6:
        hexstr = "FF" + hexstr
    v = int(hexstr, 16)
    return Color.FromArgb((v >> 24) & 0xFF, (v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF)


# pyRevit 7 palette (pyrevit/forms/_ipy.py) -> published as pyRevit<Key>Brush
PYREVIT_LIGHT = {
    "WindowBackground": "#F3F4F6", "WindowForeground": "#000000",
    "ChromeBackground": "#E8E8E8", "ControlBackground": "#FFFFFF",
    "ButtonBackground": "#F0F0F0", "ContrastBackground": "#2C3E50",
    "ContrastForeground": "#FFFFFF", "ControlBorder": "#CCCCCC",
    "ControlHover": "#E5E5E5", "ControlPressed": "#D9D9D9",
    "ControlOverlayHover": "#0C000000", "ControlOverlayPressed": "#18000000",
    "DisabledForeground": "#A0A0A0", "SubtleForeground": "#696969",
    "DangerForeground": "#B42318", "DangerBackground": "#FBD5D5",
    "WarningForeground": "#9A6300", "WarningBackground": "#FDF0D5",
    "SuccessBackground": "#D4EFD8", "Icon": "#000000",
    "SelectionBackground": "#CCE4F7", "SelectionForeground": "#000000",
    "ScrollBarThumb": "#CCCCCC", "ScrollBarThumbHover": "#AAAAAA",
}

PYREVIT_DARK = {
    "WindowBackground": "#2E3440", "WindowForeground": "#ECF0F1",
    "ChromeBackground": "#39424F", "ControlBackground": "#222933",
    "ButtonBackground": "#222933", "ContrastBackground": "#ECF0F1",
    "ContrastForeground": "#2C3E50", "ControlBorder": "#454F61",
    "ControlHover": "#454F61", "ControlPressed": "#525E73",
    "ControlOverlayHover": "#28FFFFFF", "ControlOverlayPressed": "#37FFFFFF",
    "DisabledForeground": "#7F8C8D", "SubtleForeground": "#95A5A6",
    "DangerForeground": "#FF6B61", "DangerBackground": "#5C2E2E",
    "WarningForeground": "#F0AD4E", "WarningBackground": "#5C482E",
    "SuccessBackground": "#2C4C33", "Icon": "#ECF0F1",
    "SelectionBackground": "#33506E", "SelectionForeground": "#ECF0F1",
    "ScrollBarThumb": "#454F61", "ScrollBarThumbHover": "#525E73",
}

# Shell-specific colours -> published as PPS.<Key>
SHELL_LIGHT = {
    "EditorBackground": "#FFFFFF", "EditorForeground": "#1E1E1E",
    "LineNumbers": "#999999", "CurrentLine": "#F3F6FA", "CurrentLineBorder": "#E4EAF2",
    "Selection": "#ADD6FF", "SelectionBorder": "#ADD6FF", "Caret": "#000000",
    "OutputBackground": "#FBFBFC", "OutputText": "#1E1E1E",
    "OutputError": "#C42B1C", "OutputEcho": "#6B7685", "OutputResult": "#0B6A0B",
    "OutputInfo": "#2F6FB3", "Prompt": "#2F6FB3",
    "ErrorLine": "#33E51400", "FindMatch": "#66FFD54F", "BracketMatch": "#40007ACC",
    "PopupBackground": "#FFFFFF", "PopupForeground": "#1E1E1E",
    "PopupBorder": "#C8C8C8", "PopupSelection": "#CCE8FF",
    "StatusBackground": "#FFFFFF", "Accent": "#F39C12", "RunBackground": "#2E7D32",
    "RunForeground": "#FFFFFF",
    "FlashSuccess": "#3CB371", "FlashError": "#E5533D", "FlashWarning": "#DAA520",
    "FlashInfo": "#4169E1",
    "KindClass": "#C27C0E", "KindMethod": "#7B3FC4", "KindProperty": "#0E7490",
    "KindField": "#0E7490", "KindEnum": "#B45309", "KindNamespace": "#6B7280",
    "KindKeyword": "#1D4ED8", "KindVariable": "#2563EB", "KindEvent": "#B91C1C",
    "KindModule": "#6B7280", "KindSnippet": "#6B7280",
}

SHELL_DARK = {
    "EditorBackground": "#222933", "EditorForeground": "#D4D4D4",
    "LineNumbers": "#858585", "CurrentLine": "#2A323E", "CurrentLineBorder": "#313A47",
    "Selection": "#264F78", "SelectionBorder": "#264F78", "Caret": "#AEAFAD",
    "OutputBackground": "#1C222B", "OutputText": "#D4D4D4",
    "OutputError": "#F48771", "OutputEcho": "#8C99A8", "OutputResult": "#89D185",
    "OutputInfo": "#75BEFF", "Prompt": "#9CDCFE",
    "ErrorLine": "#40F14C4C", "FindMatch": "#66B58900", "BracketMatch": "#50569CD6",
    "PopupBackground": "#252D38", "PopupForeground": "#D4D4D4",
    "PopupBorder": "#454F61", "PopupSelection": "#094771",
    "StatusBackground": "#222933", "Accent": "#F39C12", "RunBackground": "#2E7D32",
    "RunForeground": "#FFFFFF",
    "FlashSuccess": "#2E8B57", "FlashError": "#C0392B", "FlashWarning": "#B8860B",
    "FlashInfo": "#3A5FCD",
    "KindClass": "#E5C07B", "KindMethod": "#C586C0", "KindProperty": "#4EC9B0",
    "KindField": "#4EC9B0", "KindEnum": "#F0AD4E", "KindNamespace": "#9DA5B4",
    "KindKeyword": "#569CD6", "KindVariable": "#9CDCFE", "KindEvent": "#F48771",
    "KindModule": "#9DA5B4", "KindSnippet": "#9DA5B4",
}


class Palette(object):
    """Resolved colours/brushes for one theme."""

    def __init__(self, dark):
        self.dark = bool(dark)
        self.pyrevit = dict((k, _c(v)) for k, v in (PYREVIT_DARK if dark else PYREVIT_LIGHT).items())
        self.shell = dict((k, _c(v)) for k, v in (SHELL_DARK if dark else SHELL_LIGHT).items())
        self._brushes = {}

    def color(self, key):
        if key in self.shell:
            return self.shell[key]
        return self.pyrevit[key]

    def brush(self, key):
        b = self._brushes.get(key)
        if b is None:
            b = SolidColorBrush(self.color(key))
            b.Freeze()
            self._brushes[key] = b
        return b

    @property
    def name(self):
        return "Dark" if self.dark else "Light"


def apply_resources(element, palette):
    """Publish the palette as resources on ``element`` (DynamicResource-able).

    Keys: ``pyRevit<Name>Brush`` / ``pyRevit<Name>Color`` (pyRevit 7 compatible)
    and ``PPS.<Name>`` for the shell's own colours.
    """
    res = element.Resources
    for name in palette.pyrevit:
        res["pyRevit" + name + "Color"] = palette.pyrevit[name]
        res["pyRevit" + name + "Brush"] = palette.brush(name)
    for name in palette.shell:
        res["PPS." + name] = palette.brush(name)
    res["pyRevitIsDarkTheme"] = palette.dark
    # system brushes used by templates that do not know pyRevit's keys
    res[SystemColors.WindowBrushKey] = palette.brush("ControlBackground")
    res[SystemColors.WindowTextBrushKey] = palette.brush("WindowForeground")
    res[SystemColors.ControlBrushKey] = palette.brush("WindowBackground")
    res[SystemColors.ControlTextBrushKey] = palette.brush("WindowForeground")
    res[SystemColors.HighlightBrushKey] = palette.brush("Selection")
    res[SystemColors.HighlightTextBrushKey] = palette.brush("EditorForeground")
    res[SystemColors.InactiveSelectionHighlightBrushKey] = palette.brush("Selection")
    res[SystemColors.InactiveSelectionHighlightTextBrushKey] = palette.brush("EditorForeground")
    res[SystemColors.GrayTextBrushKey] = palette.brush("SubtleForeground")
