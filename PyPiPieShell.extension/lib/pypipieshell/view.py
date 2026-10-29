# -*- coding: utf-8 -*-
"""ShellController - the PyPiPie Shell UI, independent of its host.

The same controller is placed in a modal window, a modeless window or the
Revit dockable pane (see :mod:`pypipieshell.hosts`).  The host supplies the
``UIApplication`` and an executor (direct for modal, ExternalEvent otherwise).
"""
import os
import re
import textwrap

import clr

from pypipieshell import avalon
avalon.load()

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")

import System  # noqa: E402
from System import Action, TimeSpan  # noqa: E402
from System.Windows import Clipboard, MessageBox, MessageBoxButton, MessageBoxResult, MessageBoxImage  # noqa: E402
from System.Windows.Controls import ContextMenu, MenuItem, Separator  # noqa: E402
from System.Windows.Controls.Primitives import PlacementMode  # noqa: E402
from System.Windows.Input import Key, ModifierKeys, Keyboard  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Windows.Media import SolidColorBrush, Color  # noqa: E402
from System.Windows.Threading import DispatcherTimer  # noqa: E402
from System.Xml import XmlReader  # noqa: E402
from Microsoft.Win32 import OpenFileDialog, SaveFileDialog  # noqa: E402

from pypipieshell import pyparse, theme, editors, completion, stubs, settings as settings_mod  # noqa: E402
from pypipieshell.execution import (ShellSession, RunRequest, EDITOR_FILENAME,  # noqa: E402
                                    REPL_FILENAME, EscInterrupter)

HERE = os.path.dirname(os.path.abspath(__file__))
VIEW_XAML = os.path.join(HERE, "resources", "ShellView.xaml")
TXN_MODES = ["auto", "group", "none"]


def load_view():
    reader = XmlReader.Create(VIEW_XAML)
    try:
        return XamlReader.Load(reader)
    finally:
        reader.Close()


def _has(mods, flag):
    try:
        return (int(mods) & int(flag)) != 0
    except Exception:
        return False


def _engine_label():
    try:
        import platform
        return "IronPython " + platform.python_version()
    except Exception:
        return "IronPython"


class ShellController(object):
    def __init__(self, uiapp, executor, mode_label, host=None):
        self.uiapp = uiapp
        self.executor = executor
        self.mode_label = mode_label
        self.host = host
        self.settings = settings_mod.Settings()
        self.session = ShellSession()
        try:
            self.session.refresh_context(uiapp)
        except Exception:
            pass
        self.current_file = None
        self.palette = theme.Palette(theme.is_revit_dark(uiapp))
        self.history = settings_mod.load_history()
        self.history_index = len(self.history)
        self.history_draft = u""
        self._pending = 0
        self._disposed = False
        self._find_cache_key = None

        self.root = load_view()
        self._bind_names()
        theme.apply_resources(self.root, self.palette)

        year = theme.revit_year(uiapp)
        self.revit_year = year
        self.stub_index = stubs.StubIndex(self._stub_folders())
        self.engine = completion.CompletionEngine(lambda: self.session.scope, self.stub_index)

        # editors
        self.editor = editors.create_editor()
        self.py = editors.PythonEditor(self.editor, self.engine, self.settings, self.palette, folding=True)
        self.py.on_status = self.set_status
        self.py.on_zoom = self._apply_zoom
        self.EditorHost.Content = self.editor

        self.output_editor = editors.create_editor(read_only=True, line_numbers=False)
        self.output = editors.OutputView(self.output_editor, self.settings, self.palette)
        self.OutputHost.Content = self.output_editor

        self.repl_editor = editors.create_editor(line_numbers=False)
        self.repl_editor.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Hidden
        try:
            self.repl_editor.Options.HighlightCurrentLine = False
        except Exception:
            pass
        self.repl = editors.PythonEditor(self.repl_editor, self.engine, self.settings, self.palette,
                                         folding=False, is_repl=True)
        self.repl.on_zoom = self._apply_zoom
        self.ReplHost.Content = self.repl_editor

        self._wire()
        self._init_timers()
        self.apply_settings()
        self._apply_theme_to_editors()
        self._restore_buffer()
        self._unsubscribe_theme = theme.subscribe_theme_changed(uiapp, self._on_theme_changed)

        self.StatusMode.Text = u"{0} · {1}".format(mode_label, _engine_label())
        self._update_txn_label()
        self._update_caret()
        self._update_file_labels()
        self._update_run_enabled()
        self.output.append(u"PyPiPie Shell — {0}, Revit {1}, AvalonEdit {2}, {3} theme\n".format(
            mode_label, year or "?", avalon.version(), self.palette.name), "info")
        self.output.append(u"doc, uidoc, uiapp, app, selection, active_view, DB, UI, revit, forms, script "
                           u"are available. F5 runs the editor, the >>> line runs statements.\n", "info")

    # ------------------------------------------------------------------ setup
    def _bind_names(self):
        names = ["EditorHost", "OutputHost", "ReplHost", "ReplPrompt", "RunButton", "RunSelectionButton",
                 "StopButton", "NewButton", "OpenButton", "SaveButton", "SaveAsButton", "UndoButton",
                 "RedoButton", "CommentButton", "UncommentButton", "IndentButton", "OutdentButton",
                 "DuplicateButton", "DeleteLineButton", "FindToggleButton", "TxnModeCombo", "SettingsButton",
                 "HelpButton", "CopyOutputButton", "ResetSessionButton", "ClearOutputButton", "FindBar",
                 "FindInputBox", "FindPrevButton", "FindButton", "FindCountText", "ReplaceInputBox",
                 "ReplaceButton", "ReplaceAllButton", "MatchCaseBox", "WholeWordBox", "RegexBox",
                 "GotoInputBox", "GotoButton", "StatusBorder", "StatusText", "StatusMode", "StatusTxn",
                 "StatusCaret", "StatusDirty", "StatusFile"]
        for n in names:
            el = self.root.FindName(n)
            if el is None:
                raise Exception("ShellView.xaml is missing element '{0}'".format(n))
            setattr(self, n, el)

    def _stub_folders(self):
        folders = list(self.settings["stub_paths"] or [])
        for f in stubs.detect_stub_folders(self.revit_year):
            if f not in folders:
                folders.append(f)
        return folders

    def _wire(self):
        b = self
        b.RunButton.Click += lambda s, e: self.run_editor(selection_only=False)
        b.RunSelectionButton.Click += lambda s, e: self.run_editor(selection_only=True)
        b.StopButton.Click += self._on_stop
        b.NewButton.Click += lambda s, e: self.new_file()
        b.OpenButton.Click += lambda s, e: self.open_file()
        b.SaveButton.Click += lambda s, e: self.save_file()
        b.SaveAsButton.Click += lambda s, e: self.save_file(save_as=True)
        b.UndoButton.Click += lambda s, e: self._editor_cmd(lambda: self.editor.Undo())
        b.RedoButton.Click += lambda s, e: self._editor_cmd(lambda: self.editor.Redo())
        b.CommentButton.Click += lambda s, e: self._editor_cmd(self.py.comment, "Commented")
        b.UncommentButton.Click += lambda s, e: self._editor_cmd(self.py.uncomment, "Uncommented")
        b.IndentButton.Click += lambda s, e: self._editor_cmd(self.py.indent, "Indented")
        b.OutdentButton.Click += lambda s, e: self._editor_cmd(self.py.outdent, "Outdented")
        b.DuplicateButton.Click += lambda s, e: self._editor_cmd(self.py.duplicate, "Duplicated")
        b.DeleteLineButton.Click += lambda s, e: self._editor_cmd(self.py.delete_lines, "Deleted line(s)")
        b.FindToggleButton.Click += lambda s, e: self.toggle_find_bar()
        b.SettingsButton.Click += self._on_settings_click
        b.HelpButton.Click += lambda s, e: self.open_help()
        b.CopyOutputButton.Click += self._on_copy_output
        b.ResetSessionButton.Click += self._on_reset_session
        b.ClearOutputButton.Click += lambda s, e: self.clear_output()
        b.TxnModeCombo.SelectionChanged += self._on_txn_changed

        b.FindButton.Click += lambda s, e: self.find_next(forward=True)
        b.FindPrevButton.Click += lambda s, e: self.find_next(forward=False)
        b.ReplaceButton.Click += lambda s, e: self.replace_one()
        b.ReplaceAllButton.Click += lambda s, e: self.replace_all()
        b.GotoButton.Click += lambda s, e: self.goto_from_box()
        b.FindInputBox.TextChanged += self._on_find_text_changed
        for box in (b.MatchCaseBox, b.WholeWordBox, b.RegexBox):
            box.Checked += self._on_find_text_changed
            box.Unchecked += self._on_find_text_changed
        b.FindInputBox.PreviewKeyDown += self._on_find_key
        b.ReplaceInputBox.PreviewKeyDown += self._on_replace_key
        b.GotoInputBox.PreviewKeyDown += self._on_goto_key

        self.editor.TextChanged += self._on_editor_text_changed
        self.editor.TextArea.Caret.PositionChanged += lambda s, e: self._update_caret()
        self.editor.TextArea.SelectionChanged += lambda s, e: self._update_caret()

        self.repl_editor.TextArea.PreviewKeyDown += self._on_repl_key
        self.repl_editor.TextChanged += lambda s, e: self._update_repl_prompt()
        self.output_editor.MouseDoubleClick += self._on_output_double_click

        self.root.PreviewKeyDown += self._on_root_key
        self.root.Unloaded += self._on_unloaded

    def _init_timers(self):
        self._autosave_timer = DispatcherTimer()
        self._autosave_timer.Interval = TimeSpan.FromSeconds(2)
        self._autosave_timer.Tick += self._on_autosave_tick

        self._find_timer = DispatcherTimer()
        self._find_timer.Interval = TimeSpan.FromMilliseconds(250)
        self._find_timer.Tick += self._on_find_timer

        self._wait_timer = DispatcherTimer()
        self._wait_timer.Interval = TimeSpan.FromMilliseconds(700)
        self._wait_timer.Tick += self._on_wait_tick

        self._flash_timer = DispatcherTimer()
        self._flash_timer.Interval = TimeSpan.FromMilliseconds(60)
        self._flash_timer.Tick += self._on_flash_tick
        self._flash_steps = []
        self._flash_index = 0

    # ------------------------------------------------------------------ theme
    def _on_theme_changed(self):
        self.apply_theme(theme.is_revit_dark(self.uiapp))

    def apply_theme(self, dark):
        self.palette = theme.Palette(dark)
        theme.apply_resources(self.root, self.palette)
        self._apply_theme_to_editors()
        if self.host is not None and hasattr(self.host, "on_theme_applied"):
            try:
                self.host.on_theme_applied(self.palette)
            except Exception:
                pass
        self.set_status(u"{0} theme".format(self.palette.name))

    def _apply_theme_to_editors(self):
        self.py.close_popups()
        self.repl.close_popups()
        self.py.apply_theme(self.palette)
        self.repl.apply_theme(self.palette)
        self.repl_editor.Background = self.palette.brush("EditorBackground")
        self.output.apply_theme(self.palette)
        self.StatusBorder.Background = self.palette.brush("StatusBackground")

    # --------------------------------------------------------------- settings
    def apply_settings(self):
        s = self.settings
        self.py.apply_settings()
        self.repl.apply_settings()
        self.output_editor.FontFamily = self.editor.FontFamily
        self.output_editor.FontSize = self.editor.FontSize
        mode = (s["transaction_mode"] or "auto").lower()
        if mode not in TXN_MODES:
            mode = "auto"
        self.TxnModeCombo.SelectedIndex = TXN_MODES.index(mode)
        self.FindBar.Visibility = System.Windows.Visibility.Visible if s["find_bar_visible"] \
            else System.Windows.Visibility.Collapsed
        self.ReplPrompt.FontFamily = self.editor.FontFamily
        self.ReplPrompt.FontSize = self.editor.FontSize

    def _apply_zoom(self, size):
        self.settings["font_size"] = size
        for ed in (self.editor, self.repl_editor, self.output_editor):
            ed.FontSize = size
        self.ReplPrompt.FontSize = size
        self.settings.save()
        self.set_status(u"Font size {0:g}".format(size))

    def _on_txn_changed(self, sender, args):
        idx = self.TxnModeCombo.SelectedIndex
        if 0 <= idx < len(TXN_MODES):
            self.settings["transaction_mode"] = TXN_MODES[idx]
            self.settings.save()
        self._update_txn_label()

    def _update_txn_label(self):
        self.StatusTxn.Text = u"Txn: " + (self.settings["transaction_mode"] or "auto")

    def _toggle_setting(self, key, label):
        self.settings[key] = not bool(self.settings[key])
        self.settings.save()
        self.apply_settings()
        self.set_status(u"{0}: {1}".format(label, "on" if self.settings[key] else "off"))

    def _on_settings_click(self, sender, args):
        menu = ContextMenu()
        editors.style_popup(menu, self.palette)

        def check(label, key):
            mi = MenuItem()
            mi.Header = label
            mi.IsCheckable = True
            mi.IsChecked = bool(self.settings[key])
            mi.Click += lambda s, e: self._toggle_setting(key, label)
            menu.Items.Add(mi)

        def action(label, fn, enabled=True):
            mi = MenuItem()
            mi.Header = label
            mi.IsEnabled = enabled
            mi.Click += lambda s, e: fn()
            menu.Items.Add(mi)
            return mi

        check("Word wrap", "word_wrap")
        check("Show whitespace", "show_whitespace")
        check("Line numbers", "show_line_numbers")
        check("Highlight current line", "highlight_current_line")
        check("Code folding", "folding")
        check("Find bar", "find_bar_visible")
        menu.Items.Add(Separator())
        check("Autocomplete while typing", "complete_as_you_type")
        check("Hover tips", "hover_tips")
        esc_ok = EscInterrupter().available
        check("Esc interrupts running scripts" + ("" if esc_ok else " (unavailable)"), "esc_interrupt")
        check("Restore last buffer", "restore_buffer")
        menu.Items.Add(Separator())
        action(u"Font size +   (Ctrl+=)", lambda: self.py.zoom(1))
        action(u"Font size −   (Ctrl+−)", lambda: self.py.zoom(-1))
        action(u"Font size reset (Ctrl+0)", lambda: self._apply_zoom(13.0))
        menu.Items.Add(Separator())
        n = self.stub_index.module_count
        state = u"{0} modules".format(n) if self.stub_index.is_ready else u"indexing..."
        action(u"Revit API stubs: {0}".format(state), self.show_stub_info)
        action(u"Build Revit {0} stubs with pyRevit's stubs builder".format(self.revit_year or ""),
               self.build_stubs)
        action(u"Add stub folder...", self.add_stub_folder)
        action(u"Remove added stub folders", self.clear_stub_folders, enabled=bool(self.settings["stub_paths"]))
        menu.Items.Add(Separator())
        launch = MenuItem()
        launch.Header = u"Open the shell next time as"
        for m in ("Modeless", "Modal", "Docked"):
            sub = MenuItem()
            sub.Header = m
            sub.IsCheckable = True
            sub.IsChecked = (self.settings["launch_mode"] == m)
            sub.Click += (lambda mm: (lambda s, e: self._set_launch_mode(mm)))(m)
            launch.Items.Add(sub)
        menu.Items.Add(launch)
        action(u"Open settings folder", lambda: os.startfile(settings_mod.data_dir()))
        menu.Items.Add(Separator())
        action(u"Help   (F1)", self.open_help)
        self.SettingsButton.ContextMenu = menu
        menu.PlacementTarget = self.SettingsButton
        menu.Placement = PlacementMode.Bottom
        menu.IsOpen = True

    def _set_launch_mode(self, mode):
        self.settings["launch_mode"] = mode
        self.settings.save()
        self.set_status(u"Next launch: {0} (Shift+Click the button to change)".format(mode))

    # ------------------------------------------------------------------ stubs
    def show_stub_info(self):
        folders = self.stub_index.folders
        if folders:
            self.output.append(u"Stub folders:\n" + u"".join(u"  {0}\n".format(f) for f in folders), "info")
        else:
            self.output.append(u"No stub folders. Use Settings > Build Revit stubs, or Add stub folder.\n", "info")
        self.output.scroll_to_end()

    def add_stub_folder(self):
        try:
            from pyrevit import forms
            folder = forms.pick_folder(title="Folder containing .pyi stubs")
        except Exception:
            folder = None
        if not folder:
            return
        paths = list(self.settings["stub_paths"] or [])
        if folder not in paths:
            paths.append(folder)
        self.settings["stub_paths"] = paths
        self.settings.save()
        self.stub_index.set_folders(self._stub_folders())
        self.set_status(u"Indexing stubs in " + folder)

    def clear_stub_folders(self):
        self.settings["stub_paths"] = []
        self.settings.save()
        self.stub_index.set_folders(self._stub_folders())
        self.set_status(u"Added stub folders removed")

    def build_stubs(self):
        self.set_status(u"Building Revit API stubs (this can take a minute)...")
        self.output.append(u"Building Revit {0} stubs with pyrevit-stubsbuilder...\n".format(self.revit_year), "info")
        dispatcher = self.root.Dispatcher

        def done(ok, message, dest):
            def ui():
                if ok:
                    self.output.append(u"Stubs written to {0}\n".format(dest), "info")
                    self.stub_index.set_folders(self._stub_folders())
                    self.set_status(u"Revit API stubs built")
                    self.flash("success")
                else:
                    self.output.append(u"Stub build failed: {0}\n".format(message), "err")
                    self.set_status(u"Stub build failed")
                    self.flash("error")
                self.output.scroll_to_end()
            dispatcher.BeginInvoke(Action(ui))

        stubs.build_revit_stubs(self.revit_year, done)

    # -------------------------------------------------------------- file ops
    def _confirm_discard(self):
        if not self.editor.IsModified or not self.current_file:
            return True
        r = MessageBox.Show(u"Save changes to {0}?".format(os.path.basename(self.current_file)),
                            u"PyPiPie Shell", MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
        if r == MessageBoxResult.Cancel:
            return False
        if r == MessageBoxResult.Yes:
            return self.save_file()
        return True

    def _load_text(self, text, path):
        self.editor.Document.Text = text
        self.editor.Document.UndoStack.ClearAll()
        self.editor.IsModified = False
        self.current_file = path
        self.settings["last_file"] = path
        self.settings.save()
        self.editor.ScrollToHome()
        self.editor.CaretOffset = 0
        self.py.update_foldings()
        self._update_file_labels()

    def new_file(self):
        if not self._confirm_discard():
            return
        self._load_text(u"", None)
        self.set_status(u"New file")
        self.flash("info")

    def open_file(self, path=None):
        if not self._confirm_discard():
            return
        if path is None:
            dlg = OpenFileDialog()
            dlg.Filter = "Python files (*.py)|*.py|All files (*.*)|*.*"
            if self.current_file:
                dlg.InitialDirectory = os.path.dirname(self.current_file)
            if not dlg.ShowDialog():
                return
            path = dlg.FileName
        try:
            text = settings_mod.read_text_file(path).replace(u"\t", u"    ")
        except Exception as ex:
            self.set_status(u"Open failed: {0}".format(ex))
            self.flash("error")
            return
        self._load_text(text, path)
        self.set_status(u"Opened " + os.path.basename(path))
        self.flash("info")

    def save_file(self, save_as=False):
        path = self.current_file
        if save_as or not path:
            dlg = SaveFileDialog()
            dlg.Filter = "Python files (*.py)|*.py|All files (*.*)|*.*"
            dlg.DefaultExt = ".py"
            dlg.FileName = os.path.basename(path) if path else "untitled.py"
            if path:
                dlg.InitialDirectory = os.path.dirname(path)
            if not dlg.ShowDialog():
                return False
            path = dlg.FileName
        try:
            settings_mod.write_text_file(path, self.editor.Text.replace(u"\t", u"    "))
        except Exception as ex:
            self.set_status(u"Save failed: {0}".format(ex))
            self.flash("error")
            return False
        self.current_file = path
        self.editor.IsModified = False
        self.settings["last_file"] = path
        self.settings.save()
        self._update_file_labels()
        self.set_status(u"Saved " + os.path.basename(path))
        self.flash("success")
        return True

    def _restore_buffer(self):
        if not self.settings["restore_buffer"]:
            return
        text = settings_mod.read_autosave()
        last = self.settings["last_file"]
        if last and not os.path.isfile(last):
            last = None
        if text is None and last:
            try:
                text = settings_mod.read_text_file(last)
            except Exception:
                text = None
        if text is None:
            return
        self.editor.Document.Text = text
        self.editor.Document.UndoStack.ClearAll()
        self.current_file = last
        same = False
        if last:
            try:
                same = settings_mod.read_text_file(last) == text
            except Exception:
                same = False
        self.editor.IsModified = bool(last) and not same
        self.py.update_foldings()
        self._update_file_labels()

    def _on_autosave_tick(self, sender, args):
        self._autosave_timer.Stop()
        if self.settings["restore_buffer"]:
            settings_mod.write_autosave(self.editor.Text)

    def _update_file_labels(self):
        name = os.path.basename(self.current_file) if self.current_file else u"untitled.py"
        self.StatusFile.Text = name
        self.StatusFile.ToolTip = self.current_file or u"Not saved to a file (buffer is autosaved)"
        self.StatusDirty.Text = u"●" if (self.editor.IsModified and self.current_file) else u""
        if self.host is not None and hasattr(self.host, "on_title_changed"):
            try:
                self.host.on_title_changed(name, self.editor.IsModified)
            except Exception:
                pass

    # ------------------------------------------------------------ editor events
    def _on_editor_text_changed(self, sender, args):
        self._update_file_labels()
        self._update_run_enabled()
        self._autosave_timer.Stop()
        self._autosave_timer.Start()
        if self.FindInputBox.Text:
            self._find_timer.Stop()
            self._find_timer.Start()

    def _update_caret(self):
        caret = self.editor.TextArea.Caret
        text = u"Ln {0}, Col {1}".format(caret.Line, caret.Column)
        n = self.editor.SelectionLength
        if n:
            text += u"  ({0} selected)".format(n)
        self.StatusCaret.Text = text

    def _editor_cmd(self, fn, status=None):
        try:
            fn()
        except Exception as ex:
            self.set_status(u"Edit failed: {0}".format(ex))
            return
        self.editor.TextArea.Focus()
        if status:
            self.set_status(status)

    # ---------------------------------------------------------------- keyboard
    def _on_root_key(self, sender, e):
        key = e.Key
        if key == Key.System:
            key = e.SystemKey
        mods = Keyboard.Modifiers
        ctrl = mods == ModifierKeys.Control
        ctrl_shift = mods == (ModifierKeys.Control | ModifierKeys.Shift)
        in_editor = self.editor.TextArea.IsKeyboardFocusWithin
        handled = True
        if key == Key.F5:
            self.run_editor(selection_only=False)
        elif key == Key.Return and ctrl and in_editor:
            self.run_editor(selection_only=True)
        elif key == Key.F1:
            self.open_help()
        elif key == Key.F3:
            self.find_next(forward=not _has(mods, ModifierKeys.Shift))
        elif ctrl and key == Key.S:
            self.save_file()
        elif ctrl_shift and key == Key.S:
            self.save_file(save_as=True)
        elif ctrl and key == Key.O:
            self.open_file()
        elif ctrl and key == Key.N:
            self.new_file()
        elif ctrl and key == Key.F:
            self.focus_find(self.FindInputBox)
        elif ctrl and key == Key.H:
            self.focus_find(self.ReplaceInputBox)
        elif ctrl and key == Key.G:
            self.focus_find(self.GotoInputBox)
        elif ctrl and key == Key.L:
            self.clear_output()
        elif ctrl and key == Key.E:
            self.repl_editor.TextArea.Focus()
        elif ctrl and key in (Key.OemPlus, Key.Add):
            self.py.zoom(1)
        elif ctrl and key in (Key.OemMinus, Key.Subtract):
            self.py.zoom(-1)
        elif ctrl and key in (Key.D0, Key.NumPad0):
            self._apply_zoom(13.0)
        elif in_editor and ctrl and key in (Key.OemQuestion, Key.Divide):
            self._editor_cmd(self.py.toggle_comment)
        elif in_editor and ctrl and key == Key.K:
            self._editor_cmd(self.py.comment)
        elif in_editor and ctrl and key == Key.U:
            self._editor_cmd(self.py.uncomment)
        elif in_editor and ctrl and key == Key.D:
            self._editor_cmd(self.py.duplicate)
        elif in_editor and ctrl_shift and key == Key.K:
            self._editor_cmd(self.py.delete_lines)
        elif in_editor and mods == ModifierKeys.Alt and key == Key.Up:
            self._editor_cmd(lambda: self.py.move_lines(-1))
        elif in_editor and mods == ModifierKeys.Alt and key == Key.Down:
            self._editor_cmd(lambda: self.py.move_lines(1))
        elif key == Key.Escape and in_editor and self.py.error_marker.segments:
            self.py.error_marker.segments = []
            self.editor.TextArea.TextView.Redraw()
        else:
            handled = False
        if handled:
            e.Handled = True

    # -------------------------------------------------------------------- run
    def _update_run_enabled(self):
        busy = self._pending > 0
        has_code = bool(self.editor.Document.TextLength) and bool(self.editor.Text.strip())
        self.RunButton.IsEnabled = has_code and not busy
        self.RunSelectionButton.IsEnabled = has_code and not busy
        self.StopButton.IsEnabled = busy and self.executor.is_queued

    def _selection_code(self):
        doc = self.editor.Document
        if self.editor.SelectionLength > 0:
            first, last = self.py.selected_line_range()
            sel_text = self.editor.SelectedText
            single_line = u"\n" not in sel_text.strip(u"\n")
            if single_line and first.LineNumber == last.LineNumber:
                # partial line selection: run just that expression/statement
                return sel_text.strip(), first.LineNumber
            text = doc.GetText(first.Offset, last.EndOffset - first.Offset)
        else:
            first = doc.GetLineByOffset(self.editor.CaretOffset)
            text = doc.GetText(first.Offset, first.Length)
        return textwrap.dedent(text), first.LineNumber

    def run_editor(self, selection_only=False):
        if self._pending:
            self.set_status(u"A run is already in progress")
            self.flash("warning")
            return
        if selection_only:
            code, first_line = self._selection_code()
            # pad so tracebacks report editor line numbers
            code = u"\n" * (first_line - 1) + code
            label = u"selection (line {0})".format(first_line)
        else:
            code = self.editor.Text
            label = os.path.basename(self.current_file) if self.current_file else u"untitled.py"
        if not code.strip():
            self.set_status(u"Nothing to run")
            self.flash("warning")
            return
        self.py.error_marker.segments = []
        self.output.append(u"▶ Run {0}\n".format(label), "info")
        self._submit(code, EDITOR_FILENAME, "editor")

    def run_repl(self):
        code = self.repl_editor.Text.rstrip()
        if not code.strip():
            return
        if self._pending:
            self.set_status(u"A run is already in progress")
            self.flash("warning")
            return
        lines = code.split(u"\n")
        echo = u">>> " + lines[0] + u"\n" + u"".join(u"... " + l + u"\n" for l in lines[1:])
        self.output.append(echo, "echo")
        if not self.history or self.history[-1] != code:
            self.history.append(code)
            settings_mod.save_history(self.history)
        self.history_index = len(self.history)
        self.history_draft = u""
        self.repl_editor.Document.Text = u""
        self._submit(code, REPL_FILENAME, "repl")

    def _submit(self, code, filename, kind):
        s = self.settings
        req = RunRequest(code, filename, kind, s["transaction_mode"], bool(s["esc_interrupt"]),
                         self._on_run_output, None)
        req.on_done = lambda result: self._on_run_done(req, result)
        self._pending += 1
        self._update_run_enabled()
        self.set_status(u"Running...")
        self._wait_timer.Start()
        try:
            status = self.executor.submit(self.session, req)
        except Exception as ex:
            self._pending = max(0, self._pending - 1)
            self._wait_timer.Stop()
            self._update_run_enabled()
            self.output.append(u"Could not start run: {0}\n".format(ex), "err")
            self.set_status(u"Run failed to start")
            self.flash("error")
            return
        if status.startswith("rejected"):
            self._pending = max(0, self._pending - 1)
            self._wait_timer.Stop()
            self._update_run_enabled()
            self.output.append(u"Revit refused the run request ({0}). Finish the active command and try again.\n"
                               .format(status), "err")
            self.set_status(u"Revit busy")
            self.flash("warning")

    def _on_wait_tick(self, sender, args):
        self._wait_timer.Stop()
        if self._pending and self.executor.is_queued and self.executor.pending_count:
            self.set_status(u"Waiting for Revit — finish or cancel the active command (Stop cancels the run)")
            self.flash("warning")

    def _on_stop(self, sender, args):
        n = self.executor.cancel_pending()
        if n:
            self.set_status(u"Cancelled queued run")
            self.output.append(u"Run cancelled before it started.\n", "info")

    def _on_run_output(self, text, kind):
        self.output.append(text, kind)

    def _on_run_done(self, req, result):
        self._pending = max(0, self._pending - 1)
        self._wait_timer.Stop()
        completion.clear_value_cache()
        if result.ok:
            msg = u"Done in {0:.2f}s".format(result.elapsed)
            if result.wrapper:
                msg += u" ({0})".format(result.wrapper)
            if req.kind == "editor":
                self.output.append(u"✔ " + msg + u"\n", "result")
            self.set_status(msg)
            self.flash("success")
        elif result.error == "cancelled":
            self.set_status(u"Cancelled")
        else:
            label = u"Interrupted" if result.interrupted else u"Failed: {0}".format(result.error)
            self.set_status(label)
            self.flash("error")
            if req.kind == "editor" and result.error_line:
                self.py.mark_error_line(result.error_line)
        self.output.scroll_to_end()
        self._update_run_enabled()
        try:
            self.session.refresh_context(self.uiapp)
        except Exception:
            pass

    def _on_reset_session(self, sender, args):
        self.session.reset()
        try:
            self.session.refresh_context(self.uiapp)
        except Exception:
            pass
        completion.clear_value_cache()
        self.output.append(u"Session reset.\n", "info")
        self.set_status(u"Session reset")

    def clear_output(self):
        self.output.clear()
        self.set_status(u"Output cleared")

    def _on_copy_output(self, sender, args):
        try:
            Clipboard.SetText(self.output_editor.Text or u" ")
            self.set_status(u"Output copied")
        except Exception:
            pass

    def _on_output_double_click(self, sender, e):
        loc = pyparse.traceback_location(self.output.line_text_at_caret())
        if loc and loc[0] == EDITOR_FILENAME:
            self.py.mark_error_line(loc[1])

    # ------------------------------------------------------------------- REPL
    def _update_repl_prompt(self):
        multi = u"\n" in self.repl_editor.Text
        self.ReplPrompt.Text = u"..." if multi else u">>>"

    def _on_repl_key(self, sender, e):
        mods = Keyboard.Modifiers
        repl = self.repl
        popup_open = repl.completion_window is not None
        if e.Key == Key.Return and not popup_open:
            if mods == ModifierKeys.Control:
                e.Handled = True
                self.run_repl()
                return
            if _has(mods, ModifierKeys.Shift):
                return  # newline
            caret_at_end = self.repl_editor.CaretOffset == self.repl_editor.Document.TextLength
            text = self.repl_editor.Text
            if caret_at_end and pyparse.is_complete_statement(text):
                e.Handled = True
                self.run_repl()
            return
        if e.Key == Key.Escape and not popup_open:
            if repl.insight_window is not None:
                return
            e.Handled = True
            self.repl_editor.Document.Text = u""
            return
        if e.Key in (Key.Up, Key.Down) and not popup_open and int(mods) == 0:
            caret = self.repl_editor.TextArea.Caret
            doc = self.repl_editor.Document
            if e.Key == Key.Up and caret.Line == 1:
                e.Handled = True
                self._history_move(-1)
            elif e.Key == Key.Down and caret.Line == doc.LineCount:
                e.Handled = True
                self._history_move(1)

    def _history_move(self, step):
        if not self.history:
            return
        if self.history_index == len(self.history):
            self.history_draft = self.repl_editor.Text
        idx = max(0, min(len(self.history), self.history_index + step))
        if idx == self.history_index:
            return
        self.history_index = idx
        text = self.history[idx] if idx < len(self.history) else self.history_draft
        self.repl_editor.Document.Text = text
        self.repl_editor.CaretOffset = len(text)

    # -------------------------------------------------------------- find bar
    def toggle_find_bar(self):
        self.settings["find_bar_visible"] = not bool(self.settings["find_bar_visible"])
        self.settings.save()
        self.apply_settings()
        if self.settings["find_bar_visible"]:
            self.focus_find(self.FindInputBox)
        else:
            self.py.set_find_matches([])
            self.editor.TextArea.Focus()

    def focus_find(self, box):
        if not self.settings["find_bar_visible"]:
            self.settings["find_bar_visible"] = True
            self.apply_settings()
        if box is self.FindInputBox and self.editor.SelectionLength and u"\n" not in self.editor.SelectedText:
            self.FindInputBox.Text = self.editor.SelectedText
        box.Focus()
        box.SelectAll()

    def _pattern(self):
        needle = self.FindInputBox.Text or u""
        if not needle:
            return None
        flags = re.M
        if not self.MatchCaseBox.IsChecked:
            flags |= re.I
        src = needle if self.RegexBox.IsChecked else re.escape(needle)
        if self.WholeWordBox.IsChecked:
            src = r"\b(?:" + src + r")\b"
        try:
            return re.compile(src, flags)
        except re.error as ex:
            self.FindCountText.Text = u"bad regex"
            self.FindCountText.ToolTip = str(ex)
            return None

    def _matches(self, pattern):
        text = self.editor.Text
        out = []
        for m in pattern.finditer(text):
            if m.end() == m.start():
                continue
            out.append((m.start(), m.end() - m.start()))
            if len(out) >= 5000:
                break
        return out

    def _on_find_text_changed(self, sender, args):
        self._find_timer.Stop()
        self._find_timer.Start()

    def _on_find_timer(self, sender, args):
        self._find_timer.Stop()
        self._refresh_find_highlights()

    def _refresh_find_highlights(self):
        pattern = self._pattern()
        if pattern is None:
            self.py.set_find_matches([])
            if not (self.FindInputBox.Text and self.RegexBox.IsChecked):
                self.FindCountText.Text = u""
            return []
        matches = self._matches(pattern)
        self.py.set_find_matches(matches)
        self._update_find_count(matches)
        return matches

    def _update_find_count(self, matches):
        if not matches:
            self.FindCountText.Text = u"no results"
            return
        sel = self.editor.SelectionStart
        idx = 0
        for i, (s, l) in enumerate(matches):
            if s == sel:
                idx = i + 1
                break
        self.FindCountText.Text = (u"{0} of {1}".format(idx, len(matches)) if idx
                                   else u"{0} found".format(len(matches)))

    def find_next(self, forward=True):
        pattern = self._pattern()
        if pattern is None:
            self.set_status(u"Find text is empty")
            self.focus_find(self.FindInputBox)
            return False
        matches = self._refresh_find_highlights()
        if not matches:
            self.set_status(u"Text not found")
            self.flash("warning")
            return False
        sel_start = self.editor.SelectionStart
        sel_end = sel_start + self.editor.SelectionLength
        target = None
        if forward:
            for s, l in matches:
                if s >= sel_end and not (s == sel_start and l == self.editor.SelectionLength):
                    target = (s, l)
                    break
            if target is None:
                target = matches[0]
        else:
            for s, l in reversed(matches):
                if s < sel_start:
                    target = (s, l)
                    break
            if target is None:
                target = matches[-1]
        self.editor.Select(target[0], target[1])
        line = self.editor.Document.GetLineByOffset(target[0]).LineNumber
        self.editor.ScrollToLine(line)
        self._update_find_count(matches)
        self.set_status(u"Found")
        return True

    def replace_one(self):
        pattern = self._pattern()
        if pattern is None:
            return
        sel = self.editor.SelectedText
        m = pattern.match(sel) if sel else None
        if m is None or m.end() != len(sel):
            self.find_next(True)
            return
        repl_text = self.ReplaceInputBox.Text or u""
        try:
            new = m.expand(repl_text) if self.RegexBox.IsChecked else repl_text
        except Exception as ex:
            self.set_status(u"Replace failed: {0}".format(ex))
            return
        start = self.editor.SelectionStart
        self.editor.Document.Replace(start, len(sel), new)
        self.editor.Select(start + len(new), 0)
        self.find_next(True)
        self.set_status(u"Replaced")
        self.flash("success")

    def replace_all(self):
        pattern = self._pattern()
        if pattern is None:
            return
        repl_text = self.ReplaceInputBox.Text or u""
        text = self.editor.Text
        try:
            if self.RegexBox.IsChecked:
                new, count = pattern.subn(repl_text, text)
            else:
                new, count = pattern.subn(lambda m: repl_text, text)
        except Exception as ex:
            self.set_status(u"Replace failed: {0}".format(ex))
            return
        if count:
            caret = self.editor.CaretOffset
            doc = self.editor.Document
            doc.Replace(0, doc.TextLength, new)
            self.editor.CaretOffset = min(caret, doc.TextLength)
            self.set_status(u"Replaced {0} occurrence(s)".format(count))
            self.flash("success")
        else:
            self.set_status(u"Text not found")
            self.flash("warning")
        self._refresh_find_highlights()

    def goto_from_box(self):
        raw = (self.GotoInputBox.Text or u"").strip()
        try:
            n = int(raw)
        except ValueError:
            self.set_status(u"Enter a line number")
            self.flash("warning")
            return
        if n < 1 or n > self.editor.Document.LineCount:
            self.set_status(u"Line out of range (1-{0})".format(self.editor.Document.LineCount))
            self.flash("warning")
            return
        self.py.goto_line(n)
        self.set_status(u"Line {0}".format(n))

    def _on_find_key(self, sender, e):
        if e.Key == Key.Return:
            e.Handled = True
            self.find_next(forward=not _has(Keyboard.Modifiers, ModifierKeys.Shift))
        elif e.Key == Key.Escape:
            e.Handled = True
            self.editor.TextArea.Focus()

    def _on_replace_key(self, sender, e):
        if e.Key == Key.Return:
            e.Handled = True
            if _has(Keyboard.Modifiers, ModifierKeys.Control):
                self.replace_all()
            else:
                self.replace_one()
        elif e.Key == Key.Escape:
            e.Handled = True
            self.editor.TextArea.Focus()

    def _on_goto_key(self, sender, e):
        if e.Key == Key.Return:
            e.Handled = True
            self.goto_from_box()
        elif e.Key == Key.Escape:
            e.Handled = True
            self.editor.TextArea.Focus()

    # ----------------------------------------------------------------- status
    def set_status(self, text):
        self.StatusText.Text = text

    def flash(self, kind="info"):
        key = {"success": "FlashSuccess", "error": "FlashError", "warning": "FlashWarning"}.get(kind, "FlashInfo")
        c1 = self.palette.color(key)
        c2 = self.palette.color("StatusBackground")
        steps = [c1] * 8
        n = 14
        for i in range(n):
            t = float(i) / (n - 1)
            steps.append(Color.FromArgb(255,
                                        int(c1.R + (c2.R - c1.R) * t),
                                        int(c1.G + (c2.G - c1.G) * t),
                                        int(c1.B + (c2.B - c1.B) * t)))
        self._flash_steps = steps
        self._flash_index = 0
        self._flash_timer.Start()

    def _on_flash_tick(self, sender, args):
        if self._flash_index >= len(self._flash_steps):
            self._flash_timer.Stop()
            self.StatusBorder.Background = self.palette.brush("StatusBackground")
            self.StatusText.Foreground = self.palette.brush("WindowForeground")
            return
        c = self._flash_steps[self._flash_index]
        self.StatusBorder.Background = SolidColorBrush(c)
        bright = (c.R * 299 + c.G * 587 + c.B * 114) / 1000
        self.StatusText.Foreground = SolidColorBrush(Color.FromRgb(0, 0, 0) if bright > 150
                                                     else Color.FromRgb(255, 255, 255))
        self._flash_index += 1

    # ------------------------------------------------------------------- misc
    def open_help(self):
        from pypipieshell import helpdoc
        err = helpdoc.open_help()
        if err:
            self.set_status(u"Help: {0}".format(err))
            self.flash("warning")
        else:
            self.set_status(u"Help opened in your browser")

    def focus_editor(self):
        self.editor.TextArea.Focus()

    def _on_unloaded(self, sender, args):
        # dockable panes unload when hidden; keep state, just close popups
        try:
            self.py.close_popups()
            self.repl.close_popups()
        except Exception:
            pass

    def dispose(self):
        if self._disposed:
            return
        self._disposed = True
        try:
            if self.settings["restore_buffer"]:
                settings_mod.write_autosave(self.editor.Text)
            self.settings.save()
        except Exception:
            pass
        for t in (self._autosave_timer, self._find_timer, self._wait_timer, self._flash_timer):
            try:
                t.Stop()
            except Exception:
                pass
        try:
            self._unsubscribe_theme()
        except Exception:
            pass
        try:
            self.py.close_popups()
            self.repl.close_popups()
        except Exception:
            pass
        try:
            self.executor.dispose()
        except Exception:
            pass
