# -*- coding: utf-8 -*-
"""Hosts for the PyPiPie Shell view: modal window, modeless window, dockable pane."""
import os

import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

import System  # noqa: E402
from System import AppDomain  # noqa: E402
from System.Windows import WindowState, Visibility  # noqa: E402

from pyrevit import forms  # noqa: E402

from pypipieshell import settings as settings_mod
from pypipieshell.execution import DirectExecutor, ExternalEventExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
WINDOW_XAML = os.path.join(HERE, "resources", "ShellWindow.xaml")
PANEL_XAML = os.path.join(HERE, "resources", "ShellPanel.xaml")

PANEL_ID = "7e6cb7e3-f2f5-4f7d-8db3-4c3b3c62d941"
MODELESS_KEY = "PyPiPieShell.ModelessWindow"


def _controller_class():
    # imported lazily so loading the dockable pane at startup stays cheap
    from pypipieshell.view import ShellController
    return ShellController


# ---------------------------------------------------------------------------
# windows
# ---------------------------------------------------------------------------
class ShellWindow(forms.WPFWindow):
    """Modal or modeless window hosting the shell."""

    resolve_theme = True

    def __init__(self, uiapp, modal):
        # handle_esc=False: Esc is used by the editors (close popups, clear REPL)
        forms.WPFWindow.__init__(self, WINDOW_XAML, handle_esc=False, set_owner=True)
        self.modal = modal
        self.uiapp = uiapp
        executor = DirectExecutor(uiapp) if modal else ExternalEventExecutor()
        controller_cls = _controller_class()
        self.controller = controller_cls(uiapp, executor, "Modal" if modal else "Modeless", host=self)
        self.HostContent.Content = self.controller.root
        self._restore_placement()
        self.Closing += self._on_closing
        self.Closed += self._on_closed
        self.Loaded += lambda s, e: self.controller.focus_editor()
        self.on_theme_applied(self.controller.palette)

    # called by the controller
    def on_title_changed(self, name, modified):
        mode = "modal" if self.modal else "modeless"
        self.Title = u"PyPiPie Shell — {0}{1}  ({2})".format(name, u" ●" if modified else u"", mode)

    def on_theme_applied(self, palette):
        self.Background = palette.brush("WindowBackground")
        self.Foreground = palette.brush("WindowForeground")
        try:
            self._apply_dark_titlebar()  # pyRevit 7 WPFWindow: themed caption bar
        except Exception:
            pass

    def _restore_placement(self):
        s = self.controller.settings
        geo = s.get("window_geometry") or {}
        try:
            w, h = float(geo.get("w", 0)), float(geo.get("h", 0))
            if w >= self.MinWidth and h >= self.MinHeight:
                self.Width, self.Height = w, h
            if "x" in geo and "y" in geo:
                x, y = float(geo["x"]), float(geo["y"])
                vs = System.Windows.SystemParameters
                if (vs.VirtualScreenLeft - 50 <= x <= vs.VirtualScreenLeft + vs.VirtualScreenWidth - 100 and
                        vs.VirtualScreenTop - 50 <= y <= vs.VirtualScreenTop + vs.VirtualScreenHeight - 100):
                    self.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual
                    self.Left, self.Top = x, y
            if geo.get("max"):
                self.WindowState = WindowState.Maximized
        except Exception:
            pass

    def _on_closing(self, sender, args):
        try:
            b = self.RestoreBounds if self.WindowState != WindowState.Normal else None
            geo = {"max": self.WindowState == WindowState.Maximized}
            if b is not None and not b.IsEmpty:
                geo.update({"x": b.Left, "y": b.Top, "w": b.Width, "h": b.Height})
            else:
                geo.update({"x": self.Left, "y": self.Top, "w": self.ActualWidth, "h": self.ActualHeight})
            self.controller.settings["window_geometry"] = geo
        except Exception:
            pass

    def _on_closed(self, sender, args):
        self.controller.dispose()
        if not self.modal:
            try:
                if AppDomain.CurrentDomain.GetData(MODELESS_KEY) is self:
                    AppDomain.CurrentDomain.SetData(MODELESS_KEY, None)
            except Exception:
                pass


def _window_is_open(win):
    if win is None:
        return False
    try:
        from System.Windows.Interop import WindowInteropHelper
        return WindowInteropHelper(win).Handle != System.IntPtr.Zero and win.IsLoaded
    except Exception:
        return False


def show_modeless(uiapp):
    """Show (or re-activate) the single modeless shell window."""
    existing = AppDomain.CurrentDomain.GetData(MODELESS_KEY)
    if _window_is_open(existing):
        if existing.WindowState == WindowState.Minimized:
            existing.WindowState = WindowState.Normal
        existing.Activate()
        return existing
    win = ShellWindow(uiapp, modal=False)
    AppDomain.CurrentDomain.SetData(MODELESS_KEY, win)
    win.Show()
    return win


def show_modal(uiapp):
    win = ShellWindow(uiapp, modal=True)
    win.ShowDialog()
    return win


# ---------------------------------------------------------------------------
# dockable pane
# ---------------------------------------------------------------------------
class ShellPanel(forms.WPFPanel):
    """Revit dockable pane.  Registered from the extension's startup.py.

    The ExternalEvent is created here (startup = valid API context); the
    editor UI is built the first time the pane becomes visible.
    """

    panel_id = PANEL_ID
    panel_title = "PyPiPie Shell"
    panel_source = PANEL_XAML
    resolve_theme = True

    def __init__(self):
        forms.WPFPanel.__init__(self)
        self.controller = None
        self._executor = ExternalEventExecutor()
        try:
            from pyrevit import HOST_APP
            self._uiapp = HOST_APP.uiapp
        except Exception:
            self._uiapp = None
        self.IsVisibleChanged += self._on_visible_changed

    def _on_visible_changed(self, sender, args):
        if self.IsVisible and self.controller is None:
            self._build()

    def _build(self):
        try:
            controller_cls = _controller_class()
            self.controller = controller_cls(self._uiapp, self._executor, "Docked", host=self)
            self.HostContent.Content = self.controller.root
            self.PlaceholderText.Visibility = Visibility.Collapsed
            self.on_theme_applied(self.controller.palette)
        except Exception:
            import traceback
            self.PlaceholderText.Text = u"PyPiPie Shell failed to load:\n\n" + traceback.format_exc()

    def on_theme_applied(self, palette):
        self.Background = palette.brush("WindowBackground")
        self.Foreground = palette.brush("WindowForeground")

    def on_title_changed(self, name, modified):
        pass


def register_dockable_pane():
    """Call from the extension's startup.py."""
    try:
        if forms.is_registered_dockable_panel(ShellPanel):
            return None
    except AttributeError:
        pass
    try:
        return forms.register_dockable_panel(ShellPanel, default_visible=False)
    except TypeError:
        # pyRevit < 7 has no default_visible argument
        return forms.register_dockable_panel(ShellPanel)


def show_docked():
    forms.open_dockable_panel(PANEL_ID)


def launch(mode, uiapp):
    """Open the shell in ``Modeless``, ``Modal`` or ``Docked`` mode."""
    from pypipieshell import avalon
    try:
        avalon.load()
    except avalon.AvalonEditMissing as ex:
        choice = forms.alert("PyPiPie Shell needs the AvalonEdit editor control.",
                             sub_msg=str(ex), title="PyPiPie Shell",
                             options=["Open Help", "Close"])
        if choice == "Open Help":
            from pypipieshell import helpdoc
            helpdoc.open_help()
        return None
    return _launch(mode, uiapp)


def _launch(mode, uiapp):
    s = settings_mod.Settings()
    if mode not in ("Modeless", "Modal", "Docked"):
        mode = s["launch_mode"] or "Modeless"
    if mode == "Docked":
        try:
            if uiapp.ActiveUIDocument is None:
                forms.alert("The docked shell needs an open Revit model.\n\n"
                            "Opening the modeless window instead.", title="PyPiPie Shell")
                return show_modeless(uiapp)
            show_docked()
            return None
        except Exception as ex:
            forms.alert("The PyPiPie Shell pane is not registered in this session.\n\n"
                        "Make sure the extension's startup.py calls "
                        "pypipieshell.register_dockable_pane() and reload pyRevit.\n\n{0}".format(ex),
                        title="PyPiPie Shell")
            return show_modeless(uiapp)
    if mode == "Modal":
        return show_modal(uiapp)
    return show_modeless(uiapp)
