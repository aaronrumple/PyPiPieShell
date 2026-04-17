# -*- coding: utf-8 -*-
#Note:Error line 2024...
import clr
import sys
import os
import traceback

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Xml")
clr.AddReference("System.Windows.Forms")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System import TimeSpan
from System.IO import StringReader, StreamReader, StreamWriter
from System.Xml import XmlReader
from System.Windows import Thickness
from System.Windows.Markup import XamlReader
from System.Windows.Documents import Paragraph, Run
from System.Windows.Media import Brushes, SolidColorBrush, ColorConverter
from System.Windows.Input import Key, ModifierKeys, Keyboard
from System.Windows.Threading import DispatcherTimer
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult
from System.Windows.Controls import ListBox
from System.Windows.Controls.Primitives import Popup
from Autodesk.Revit.UI import UIThemeManager


XAML = r"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PyPiPie Mini IDE"
    Width="1180"
    Height="820"
    MinWidth="900"
    MinHeight="620"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="3*" />
            <RowDefinition Height="6" />
            <RowDefinition Height="2*" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="42" />
        </Grid.RowDefinitions>

        <!-- INPUT -->
        <Border x:Name="InputBorder"
                Grid.Row="0"
                BorderBrush="#FFBFBFBF"
                BorderThickness="1"
                CornerRadius="4"
                Padding="0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="56"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <TextBox x:Name="LineNumbersBox"
                         Grid.Column="0"
                         IsReadOnly="True"
                         IsTabStop="False"
                         Focusable="False"
                         BorderThickness="0"
                         Background="#FFF3F3F3"
                         Foreground="#FF666666"
                         FontFamily="Consolas"
                         FontSize="13"
                         VerticalScrollBarVisibility="Hidden"
                         HorizontalScrollBarVisibility="Hidden"
                         TextAlignment="Right"
                         Padding="6,6,4,6"
                         AcceptsReturn="True"
                         AcceptsTab="False"
                         TextWrapping="NoWrap"/>

                <TextBox x:Name="InputBox"
                         Grid.Column="1"
                         AcceptsReturn="True"
                         AcceptsTab="True"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         TextWrapping="NoWrap"
                         FontFamily="Consolas"
                         FontSize="13"
                         Padding="6"
                         SpellCheck.IsEnabled="False"/>
            </Grid>
        </Border>

        <!-- SPLITTER -->
        <GridSplitter Grid.Row="1"
                      Height="6"
                      HorizontalAlignment="Stretch"
                      VerticalAlignment="Stretch"
                      Background="#FFD9D9D9"
                      ResizeDirection="Rows"
                      ResizeBehavior="PreviousAndNext"/>

        <!-- OUTPUT -->
        <Border x:Name="OutputBorder"
                Grid.Row="2"
                BorderBrush="#FFBFBFBF"
                BorderThickness="1"
                CornerRadius="4"
                Padding="0"
                Margin="0,4,0,4">
            <RichTextBox x:Name="OutputBox"
                         IsReadOnly="True"
                         IsDocumentEnabled="False"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Disabled"
                         FontFamily="Consolas"
                         FontSize="13"
                         Padding="8"
                         BorderThickness="0"/>
        </Border>

        <!-- TOOLBAR -->
        <Border x:Name="TopBarBorder"
                Grid.Row="3"
                BorderBrush="#FFBFBFBF"
                BorderThickness="1"
                CornerRadius="4"
                Padding="6"
                Margin="0,4,0,4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Button x:Name="NewButton"         Grid.Column="0"  Width="64" Height="22" Content="New"/>
                <Button x:Name="OpenButton"        Grid.Column="2"  Width="64" Height="22" Content="Open"/>
                <Button x:Name="SaveButton"        Grid.Column="4"  Width="64" Height="22" Content="Save"/>
                <Button x:Name="SaveAsButton"      Grid.Column="6"  Width="72" Height="22" Content="Save As"/>

                <Button x:Name="FindButton"        Grid.Column="8"  Width="64" Height="22" Content="Find"/>
                <Button x:Name="GotoButton"        Grid.Column="10" Width="64" Height="22" Content="Go To"/>

                <Button x:Name="CommentButton"     Grid.Column="12" Width="28" Height="22" Content="#"/>
                <Button x:Name="UncommentButton"   Grid.Column="14" Width="28" Height="22" Content="-"/>
                <Button x:Name="IndentButton"      Grid.Column="16" Width="28" Height="22" Content="&gt;"/>

                <Button x:Name="OutdentButton"     Grid.Column="18" Width="28" Height="22" Content="&lt;"/>
                <Button x:Name="DuplicateButton"   Grid.Column="20" Width="72" Height="22" Content="Duplicate"/>
                <Button x:Name="DeleteLineButton"  Grid.Column="22" Width="86" Height="22" Content="Delete Lines"/>

                <TextBlock x:Name="FileLabel"
                           Grid.Column="24"
                           VerticalAlignment="Center"
                           Foreground="#FF444444"
                           FontFamily="Consolas"
                           Text="untitled.py"/>
            </Grid>
        </Border>

        <!-- STATUS -->
        <Border x:Name="StatusBorder"
                Grid.Row="4"
                BorderBrush="#FFBFBFBF"
                BorderThickness="1"
                CornerRadius="4"
                Padding="6,2,6,2"
                Margin="0,0,0,6"
                Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock x:Name="StatusLeft"
                           Grid.Column="0"
                           VerticalAlignment="Center"
                           FontFamily="Consolas"
                           Foreground="#FF444444"
                           Text="Ready"/>

                <TextBlock x:Name="StatusCaret"
                           Grid.Column="1"
                           VerticalAlignment="Center"
                           FontFamily="Consolas"
                           Foreground="#FF444444"
                           Text="Ln 1, Col 1"/>

                <TextBlock x:Name="StatusDirty"
                           Grid.Column="3"
                           VerticalAlignment="Center"
                           FontFamily="Consolas"
                           Foreground="#FF444444"
                           Text="Saved"/>

                <TextBlock x:Name="StatusFile"
                           Grid.Column="5"
                           VerticalAlignment="Center"
                           FontFamily="Consolas"
                           Foreground="#FF444444"
                           Text="untitled.py"/>
            </Grid>
        </Border>

        <!-- BOTTOM BUTTONS -->
        <Grid x:Name="BottomBarGrid" Grid.Row="5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" />
                <ColumnDefinition Width="8" />
                <ColumnDefinition Width="Auto" />
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="Auto" />
                <ColumnDefinition Width="8" />
                <ColumnDefinition Width="Auto" />
            </Grid.ColumnDefinitions>

            <Button x:Name="ClearOutputButton"
                    Grid.Column="0"
                    Width="110"
                    Height="28"
                    Content="Clear Output"/>

            <Button x:Name="RunButton"
                    Grid.Column="4"
                    Width="90"
                    Height="28"
                    Content="Run"/>

            <Button x:Name="CloseButton"
                    Grid.Column="6"
                    Width="90"
                    Height="28"
                    Content="Close"/>
        </Grid>
    </Grid>
</Window>
"""


class RevitContext(object):
    def __init__(self, doc=None, uidoc=None, uiapp=None, app=None):
        self.doc = doc
        self.uidoc = uidoc
        self.uiapp = uiapp
        self.app = app


class ShellSessionState(object):
    def __init__(self):
        self.scope = None
        self.input_text = ""
        self.filepath = None
        self.is_dirty = False
        self.history = []
        self.history_index = -1
        self.last_find = None
        self.output_entries = []
        self.status_text = "Ready"


def append_state_output(state, text, is_error=False):
    state.output_entries.append((text, is_error))


class StateRedirector(object):
    def __init__(self, state, is_error=False):
        self.state = state
        self.is_error = is_error

    def write(self, text):
        if text is None:
            return
        s = str(text)
        if not s:
            return
        append_state_output(self.state, s, self.is_error)

    def flush(self):
        pass


class ShellWindow(object):
    def __init__(self, state=None):
        self.state = state if state is not None else ShellSessionState()

        self.scope = self.state.scope if self.state.scope is not None else {}
        self._current_paragraph = None
        self._current_is_error = False
        self._filepath = self.state.filepath
        self._is_dirty = self.state.is_dirty
        self._history = list(self.state.history)
        self._history_index = self.state.history_index
        self._last_find = self.state.last_find
        self._suspend_dirty = False
        self._save_restore_brush = None
        self._button_handlers = []
        self._completion_popup = None
        self._completion_list = None

        self.run_requested = False
        self.run_code = None

        sr = StringReader(XAML)
        xr = XmlReader.Create(sr)
        self.window = XamlReader.Load(xr)

        self.new_button = self.window.FindName("NewButton")
        self.open_button = self.window.FindName("OpenButton")
        self.save_button = self.window.FindName("SaveButton")
        self.save_as_button = self.window.FindName("SaveAsButton")
        self.find_button = self.window.FindName("FindButton")
        self.goto_button = self.window.FindName("GotoButton")
        self.comment_button = self.window.FindName("CommentButton")
        self.uncomment_button = self.window.FindName("UncommentButton")
        self.indent_button = self.window.FindName("IndentButton")
        self.outdent_button = self.window.FindName("OutdentButton")
        self.duplicate_button = self.window.FindName("DuplicateButton")
        self.delete_line_button = self.window.FindName("DeleteLineButton")
        self.clear_output_button = self.window.FindName("ClearOutputButton")
        self.run_button = self.window.FindName("RunButton")
        self.close_button = self.window.FindName("CloseButton")

        self.file_label = self.window.FindName("FileLabel")
        self.status_left = self.window.FindName("StatusLeft")
        self.status_caret = self.window.FindName("StatusCaret")
        self.status_dirty = self.window.FindName("StatusDirty")
        self.status_file = self.window.FindName("StatusFile")
        self.status_border = self.window.FindName("StatusBorder")

        self.top_bar_border = self.window.FindName("TopBarBorder")
        self.input_border = self.window.FindName("InputBorder")
        self.output_border = self.window.FindName("OutputBorder")
        self.bottom_bar_grid = self.window.FindName("BottomBarGrid")

        self.input_box = self.window.FindName("InputBox")
        self.line_box = self.window.FindName("LineNumbersBox")
        self.output_box = self.window.FindName("OutputBox")

        self._status_brush = SolidColorBrush(ColorConverter.ConvertFromString("White"))
        self.status_border.Background = self._status_brush

        self._status_timer = DispatcherTimer()
        self._status_timer.Interval = TimeSpan.FromSeconds(2)
        self._status_timer.Tick += self._handle_status_timer_tick

        self._save_flash_timer = DispatcherTimer()
        self._save_flash_timer.Interval = TimeSpan.FromMilliseconds(220)
        self._save_flash_timer.Tick += self._handle_save_flash_timer_tick

        self._normal_output_brush = Brushes.Black
        self._error_output_brush = Brushes.Red

        self._configure_output()

        if not self.scope:
            self._seed_scope()
            self.state.scope = self.scope
        else:
            self._refresh_live_revit_context()

        self._bind_events()
        self._apply_revit_theme()
        self._restore_from_state()
        self._update_line_numbers()
        self._refresh_title()
        self._update_status(self.state.status_text or "Ready")
        self._update_caret_status()

    def _get_revit_objects(self):
        uiapp = None
        uidoc = None
        doc = None
        app = None

        try:
            uiapp = __revit__
        except:
            uiapp = None

        try:
            if uiapp is not None:
                uidoc = uiapp.ActiveUIDocument
        except:
            uidoc = None

        try:
            if uidoc is not None:
                doc = uidoc.Document
        except:
            doc = None

        try:
            if uiapp is not None:
                app = uiapp.Application
        except:
            app = None

        return uiapp, uidoc, doc, app

    def _refresh_live_revit_context(self):
        uiapp, uidoc, doc, app = self._get_revit_objects()

        self.scope["__revit__"] = uiapp
        self.scope["uiapp"] = uiapp
        self.scope["uidoc"] = uidoc
        self.scope["doc"] = doc
        self.scope["app"] = app

        try:
            if "revit" in self.scope:
                self.scope["revit"].uiapp = uiapp
                self.scope["revit"].uidoc = uidoc
                self.scope["revit"].doc = doc
                self.scope["revit"].app = app
        except:
            pass

    def _sync_state(self):
        try:
            self.state.scope = self.scope
        except:
            pass

        try:
            self.state.input_text = self.input_box.Text or ""
        except:
            pass

        self.state.filepath = self._filepath
        self.state.is_dirty = self._is_dirty
        self.state.history = list(self._history)
        self.state.history_index = self._history_index
        self.state.last_find = self._last_find

        try:
            self.state.status_text = self.status_left.Text
        except:
            pass

    def _restore_from_state(self):
        try:
            self.input_box.Text = self.state.input_text or ""
        except:
            pass
        self._restore_output_from_state()

    def _restore_output_from_state(self):
        self.output_box.Document.Blocks.Clear()
        self._current_paragraph = None
        self._current_is_error = False

        for text, is_error in self.state.output_entries:
            self.append_stream_text(text, is_error, store=False)

    def _configure_output(self):
        self.output_box.Document.Blocks.Clear()
        self._current_paragraph = None
        self._current_is_error = False

    def _seed_scope(self):
        uiapp, uidoc, doc, app = self._get_revit_objects()
        revit_ctx = RevitContext(doc=doc, uidoc=uidoc, uiapp=uiapp, app=app)

        self.scope["__builtins__"] = __builtins__
        self.scope["sys"] = sys
        self.scope["clr"] = clr
        self.scope["os"] = os

        self.scope["__revit__"] = uiapp
        self.scope["uiapp"] = uiapp
        self.scope["uidoc"] = uidoc
        self.scope["doc"] = doc
        self.scope["app"] = app
        self.scope["revit"] = revit_ctx

        exec("import sys", self.scope, self.scope)
        exec("import clr", self.scope, self.scope)
        exec("import os", self.scope, self.scope)
        exec("from Autodesk.Revit.DB import *", self.scope, self.scope)

    def _bind_events(self):
        self.new_button.Click += self._handle_new_click
        self.open_button.Click += self._handle_open_click
        self.save_button.Click += self._handle_save_click
        self.save_as_button.Click += self._handle_save_as_click
        self.find_button.Click += self._handle_find_click
        self.goto_button.Click += self._handle_goto_click
        self.comment_button.Click += self._handle_comment_click
        self.uncomment_button.Click += self._handle_uncomment_click
        self.indent_button.Click += self._handle_indent_click
        self.outdent_button.Click += self._handle_outdent_click
        self.duplicate_button.Click += self._handle_duplicate_click
        self.delete_line_button.Click += self._handle_delete_line_click
        self.clear_output_button.Click += self._handle_clear_output_click
        self.run_button.Click += self._handle_run_click
        self.close_button.Click += self._handle_close_click

        self.input_box.TextChanged += self._handle_input_changed
        self.input_box.SelectionChanged += self._handle_input_view_changed
        self.input_box.SizeChanged += self._handle_input_view_changed
        self.input_box.KeyUp += self._handle_input_view_changed
        self.input_box.MouseWheel += self._handle_input_view_changed
        self.input_box.PreviewKeyDown += self._handle_input_key_down

    def _handle_new_click(self, sender, args):
        self.on_new()

    def _handle_open_click(self, sender, args):
        self.on_open()

    def _handle_save_click(self, sender, args):
        self.on_save()

    def _handle_save_as_click(self, sender, args):
        self.on_save_as()

    def _handle_find_click(self, sender, args):
        self.on_find()

    def _handle_goto_click(self, sender, args):
        self.on_goto_line()

    def _handle_comment_click(self, sender, args):
        self.comment()

    def _handle_uncomment_click(self, sender, args):
        self.uncomment()

    def _handle_indent_click(self, sender, args):
        self.indent()

    def _handle_outdent_click(self, sender, args):
        self.outdent()

    def _handle_duplicate_click(self, sender, args):
        self.duplicate_line()

    def _handle_delete_line_click(self, sender, args):
        self.delete_line()

    def _handle_clear_output_click(self, sender, args):
        self.clear_output()

    def _handle_run_click(self, sender, args):
        self.on_run()

    def _handle_close_click(self, sender, args):
        self.on_close()

    def _handle_input_changed(self, sender, args):
        self.on_input_changed()

    def _handle_input_view_changed(self, sender, args):
        self.on_input_view_changed()

    def _handle_input_key_down(self, sender, args):
        self.on_input_key_down(args)

    def _handle_status_timer_tick(self, sender, args):
        self._status_timer.Stop()
        self._set_status_color("White")

    def _handle_save_flash_timer_tick(self, sender, args):
        self._save_flash_timer.Stop()
        try:
            if self._save_restore_brush is not None:
                self.save_button.Background = self._save_restore_brush
        except:
            pass

    def show(self):
        self.window.ShowDialog()

    def _current_filename(self):
        if self._filepath:
            return os.path.basename(self._filepath)
        return "untitled.py"

    def _refresh_title(self):
        star = "*" if self._is_dirty else ""
        name = self._current_filename()
        self.window.Title = "PyPiPie Mini IDE - {0}{1}".format(name, star)
        self.file_label.Text = name
        self.status_file.Text = name
        self.status_dirty.Text = "Modified" if self._is_dirty else "Saved"

    def _set_dirty(self, value):
        self._is_dirty = bool(value)
        self._refresh_title()

    def _make_brush(self, color_name):
        return SolidColorBrush(ColorConverter.ConvertFromString(color_name))

    def _is_revit_dark_mode(self):
        try:
            theme = UIThemeManager.CurrentTheme
            theme_name = UIThemeManager.GetThemeName(theme)
            return str(theme_name).lower() == "dark"
        except:
            return False

    def _style_button_revit_dark(self, btn):
        try:
            normal = self._make_brush("#FF3F4A59")
            hover = self._make_brush("#FF4A5566")
            down = self._make_brush("#FF35404D")
            border = self._make_brush("#FF566273")

            btn.Background = normal
            btn.Foreground = Brushes.White
            btn.BorderBrush = border

            def on_enter(sender, args):
                try:
                    btn.Background = hover
                except:
                    pass

            def on_leave(sender, args):
                try:
                    btn.Background = normal
                except:
                    pass

            def on_down(sender, args):
                try:
                    btn.Background = down
                except:
                    pass

            def on_up(sender, args):
                try:
                    btn.Background = hover
                except:
                    pass

            btn.MouseEnter += on_enter
            btn.MouseLeave += on_leave
            btn.PreviewMouseDown += on_down
            btn.PreviewMouseUp += on_up

            self._button_handlers.append((btn, on_enter, on_leave, on_down, on_up))
        except:
            pass

    def _apply_revit_theme(self):
        is_dark = self._is_revit_dark_mode()

        input_bg = self._make_brush("White")
        output_bg = self._make_brush("#FFF0F0F0")
        line_bg = self._make_brush("#FFF3F3F3")
        light_border = self._make_brush("#FFBFBFBF")
        dark_text = Brushes.Black
        muted_fg = self._make_brush("#FF666666")

        self.input_border.Background = input_bg
        self.input_border.BorderBrush = light_border
        self.input_box.Background = input_bg
        self.input_box.Foreground = dark_text
        self.input_box.CaretBrush = dark_text
        self.input_box.BorderBrush = input_bg

        self.line_box.Background = line_bg
        self.line_box.Foreground = muted_fg
        self.line_box.BorderBrush = line_bg

        self.output_border.Background = output_bg
        self.output_border.BorderBrush = light_border
        self.output_box.Background = output_bg
        self.output_box.Foreground = dark_text
        self.output_box.BorderBrush = output_bg

        self._normal_output_brush = Brushes.Black
        self._error_output_brush = Brushes.Red

        if is_dark:
            title_bg = self._make_brush("#FF253040")
            frame_bg = self._make_brush("#FF4B5666")
            frame_border = self._make_brush("#FF627084")
            file_fg = Brushes.White

            try:
                self.window.Background = title_bg
            except:
                pass

            self.top_bar_border.Background = frame_bg
            self.top_bar_border.BorderBrush = frame_border
            try:
                self.bottom_bar_grid.Background = frame_bg
            except:
                pass

            for btn in [
                self.new_button, self.open_button, self.save_button, self.save_as_button,
                self.find_button, self.goto_button, self.comment_button, self.uncomment_button,
                self.indent_button, self.outdent_button, self.duplicate_button, self.delete_line_button,
                self.clear_output_button, self.run_button, self.close_button
            ]:
                self._style_button_revit_dark(btn)

            self.file_label.Foreground = file_fg
        else:
            panel_bg = self._make_brush("#FFF7F7F7")
            panel_border = self._make_brush("#FFBFBFBF")

            self.top_bar_border.Background = panel_bg
            self.top_bar_border.BorderBrush = panel_border
            try:
                self.bottom_bar_grid.Background = self._make_brush("Transparent")
            except:
                pass

            try:
                self.window.Background = self._make_brush("White")
            except:
                pass

            for btn in [
                self.new_button, self.open_button, self.save_button, self.save_as_button,
                self.find_button, self.goto_button, self.comment_button, self.uncomment_button,
                self.indent_button, self.outdent_button, self.duplicate_button, self.delete_line_button,
                self.clear_output_button, self.run_button, self.close_button
            ]:
                try:
                    btn.ClearValue(btn.BackgroundProperty)
                    btn.ClearValue(btn.ForegroundProperty)
                    btn.ClearValue(btn.BorderBrushProperty)
                except:
                    pass

            self.file_label.Foreground = self._make_brush("#FF444444")

    def _stop_status_timer(self):
        try:
            self._status_timer.Stop()
        except:
            pass

    def _set_status_color(self, color_name):
        try:
            self._status_brush.Color = ColorConverter.ConvertFromString(color_name)
        except:
            pass

    def _flash_status(self, color_name, auto_clear=True):
        self._stop_status_timer()
        self._set_status_color(color_name)
        if auto_clear:
            try:
                self._status_timer.Start()
            except:
                pass

    def _flash_save_button(self):
        try:
            self._save_flash_timer.Stop()
            self._save_restore_brush = self.save_button.Background
            self.save_button.Background = self._make_brush("#FF98FB98")
            self._save_flash_timer.Start()
        except:
            pass

    def _update_status(self, text):
        self.status_left.Text = text
        self.state.status_text = text
        msg = str(text or "")

        if msg.startswith("Saved"):
            self._flash_status("PaleGreen", True)
        elif msg.startswith("Run failed") or msg.startswith("Error") or msg.startswith("Unhandled") or msg.startswith("Traceback"):
            self._flash_status("MistyRose", False)
        elif msg.startswith("Run completed") or msg.startswith("Opened") or msg.startswith("New file") or msg.startswith("Output cleared"):
            self._flash_status("LightGoldenrodYellow", True)
        else:
            self._stop_status_timer()
            self._set_status_color("White")

    def _update_caret_status(self):
        text = self.input_box.Text or ""
        idx = self.input_box.CaretIndex

        line = text.count("\n", 0, idx) + 1
        last_nl = text.rfind("\n", 0, idx)
        if last_nl == -1:
            col = idx + 1
        else:
            col = idx - last_nl

        self.status_caret.Text = "Ln {0}, Col {1}".format(line, col)

    def _new_paragraph(self, is_error=False):
        p = Paragraph()
        p.Margin = Thickness(0)
        self.output_box.Document.Blocks.Add(p)
        self._current_paragraph = p
        self._current_is_error = is_error
        return p

    def _ensure_paragraph(self, is_error=False):
        if self._current_paragraph is None:
            return self._new_paragraph(is_error)
        if self._current_is_error != is_error:
            return self._new_paragraph(is_error)
        return self._current_paragraph

    def append_stream_text(self, text, is_error=False, store=True):
        if text is None:
            return

        if store:
            append_state_output(self.state, text, is_error)

        s = str(text).replace("\r\n", "\n").replace("\r", "\n")
        parts = s.split("\n")

        for i, part in enumerate(parts):
            if part:
                p = self._ensure_paragraph(is_error)
                run = Run(part)
                run.Foreground = self._error_output_brush if is_error else self._normal_output_brush
                p.Inlines.Add(run)

            if i < len(parts) - 1:
                self._new_paragraph(is_error)

        self.output_box.ScrollToEnd()

    def append_traceback(self, text):
        if text is None:
            return
        if self._current_paragraph is not None:
            self._new_paragraph(True)
        self.append_stream_text(str(text), True)
        self._update_status("Run failed")
        self._sync_state()

    def clear_output(self):
        self.output_box.Document.Blocks.Clear()
        self._current_paragraph = None
        self._current_is_error = False
        self.state.output_entries = []
        self._update_status("Output cleared")
        self._sync_state()

    def _big_message_box(self, message, title="Message", buttons="OKCancel", icon="info"):
        is_dark = self._is_revit_dark_mode()

        bg = "#FF2B2B2B" if is_dark else "White"
        fg = "White" if is_dark else "Black"
        btn_bg = "#FF3F4A59" if is_dark else "#FFF3F3F3"
        btn_fg = "White" if is_dark else "Black"
        border = "#FF566273" if is_dark else "#FFBFBFBF"

        xaml = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{0}"
        Width="460"
        Height="220"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        Background="{2}">
    <Grid Margin="14" Background="{2}">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="56"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border x:Name="IconBorder"
                    Grid.Column="0"
                    Width="40"
                    Height="40"
                    CornerRadius="20"
                    Background="#FFEAEAEA"
                    HorizontalAlignment="Left"
                    VerticalAlignment="Top">
                <TextBlock x:Name="IconText"
                           FontSize="24"
                           FontWeight="Bold"
                           HorizontalAlignment="Center"
                           VerticalAlignment="Center"
                           Text="i"/>
            </Border>

            <TextBlock x:Name="Msg"
                       Grid.Column="1"
                       TextWrapping="Wrap"
                       FontSize="16"
                       FontFamily="Segoe UI"
                       Foreground="{3}"
                       VerticalAlignment="Center"
                       Margin="10,0,0,0"
                       Text="{1}" />
        </Grid>

        <StackPanel Grid.Row="1"
                    Orientation="Horizontal"
                    HorizontalAlignment="Right"
                    Margin="0,14,0,0">
            <Button x:Name="YesBtn" Width="90" Height="30" Margin="0,0,8,0"
                    Background="{4}" Foreground="{5}" BorderBrush="{6}">Yes</Button>
            <Button x:Name="NoBtn" Width="90" Height="30" Margin="0,0,8,0"
                    Background="{4}" Foreground="{5}" BorderBrush="{6}">No</Button>
            <Button x:Name="CancelBtn" Width="90" Height="30"
                    Background="{4}" Foreground="{5}" BorderBrush="{6}">Cancel</Button>
        </StackPanel>
    </Grid>
</Window>
""".format(
            title.replace('"', "'"),
            message.replace('"', "'"),
            bg, fg, btn_bg, btn_fg, border
        )

        sr = StringReader(xaml)
        xr = XmlReader.Create(sr)
        win = XamlReader.Load(xr)

        yes_btn = win.FindName("YesBtn")
        no_btn = win.FindName("NoBtn")
        cancel_btn = win.FindName("CancelBtn")
        icon_border = win.FindName("IconBorder")
        icon_text = win.FindName("IconText")

        result = {"value": None}

        def yes(sender, args):
            result["value"] = str(yes_btn.Content)
            win.Close()

        def no(sender, args):
            result["value"] = str(no_btn.Content)
            win.Close()

        def cancel(sender, args):
            result["value"] = str(cancel_btn.Content)
            win.Close()

        def on_key_down(sender, args):
            try:
                k = args.Key

                if k == Key.Escape:
                    result["value"] = str(cancel_btn.Content)
                    win.Close()
                    args.Handled = True
                    return

                if k == Key.Enter or k == Key.Return:
                    if yes_btn.IsVisible and yes_btn.IsEnabled:
                        result["value"] = str(yes_btn.Content)
                        win.Close()
                        args.Handled = True
                        return

                if k == Key.Y and yes_btn.IsVisible and str(yes_btn.Content) == "Yes":
                    result["value"] = "Yes"
                    win.Close()
                    args.Handled = True
                    return

                if k == Key.N and no_btn.IsVisible and str(no_btn.Content) == "No":
                    result["value"] = "No"
                    win.Close()
                    args.Handled = True
                    return

                if k == Key.O and yes_btn.IsVisible and str(yes_btn.Content) == "OK":
                    result["value"] = "OK"
                    win.Close()
                    args.Handled = True
                    return

                if k == Key.C and cancel_btn.IsVisible and str(cancel_btn.Content) == "Cancel":
                    result["value"] = "Cancel"
                    win.Close()
                    args.Handled = True
                    return
            except:
                pass

        yes_btn.Click += yes
        no_btn.Click += no
        cancel_btn.Click += cancel
        win.KeyDown += on_key_down

        try:
            win.Owner = self.window
        except:
            pass

        icon_name = str(icon or "info").lower()
        if icon_name == "warning":
            icon_text.Text = "!"
            icon_text.Foreground = Brushes.DarkOrange
            icon_border.Background = Brushes.LemonChiffon
        elif icon_name == "error":
            icon_text.Text = "×"
            icon_text.Foreground = Brushes.DarkRed
            icon_border.Background = Brushes.MistyRose
        elif icon_name == "question":
            icon_text.Text = "?"
            icon_text.Foreground = Brushes.SteelBlue
            icon_border.Background = Brushes.AliceBlue
        else:
            icon_text.Text = "i"
            icon_text.Foreground = Brushes.SeaGreen
            icon_border.Background = Brushes.Honeydew

        if buttons == "OK":
            yes_btn.Content = "OK"
            no_btn.Visibility = 2
            cancel_btn.Visibility = 2
            yes_btn.Focus()
        elif buttons == "OKCancel":
            yes_btn.Content = "OK"
            no_btn.Visibility = 2
            cancel_btn.Content = "Cancel"
            yes_btn.Focus()
        elif buttons == "YesNoCancel":
            yes_btn.Content = "Yes"
            no_btn.Content = "No"
            cancel_btn.Content = "Cancel"
            yes_btn.Focus()
        else:
            yes_btn.Focus()

        win.ShowDialog()
        return result["value"]

    def _confirm_discard_changes(self):
        if not self._is_dirty:
            return True

        choice = self._big_message_box(
            "You have unsaved changes.\n\nYes = Save\nNo = Discard changes\nCancel = Stay here",
            "Unsaved Changes",
            "YesNoCancel",
            "warning"
        )

        if choice == "Yes":
            return self.on_save()
        elif choice == "No":
            return True
        else:
            return False

    def _pick_open_file(self):
        try:
            dlg = OpenFileDialog()
            dlg.Filter = "Python Files (*.py)|*.py|All Files (*.*)|*.*"
            dlg.Title = "Open Python File"
            dlg.Multiselect = False
            if dlg.ShowDialog() == DialogResult.OK:
                return dlg.FileName
        except Exception:
            self.append_traceback(traceback.format_exc())
        return None

    def _pick_save_file(self):
        try:
            dlg = SaveFileDialog()
            dlg.Filter = "Python Files (*.py)|*.py|All Files (*.*)|*.*"
            dlg.Title = "Save Python File"
            dlg.DefaultExt = "py"
            dlg.AddExtension = True
            if self._filepath:
                try:
                    dlg.FileName = os.path.basename(self._filepath)
                    dlg.InitialDirectory = os.path.dirname(self._filepath)
                except:
                    pass
            if dlg.ShowDialog() == DialogResult.OK:
                return dlg.FileName
        except Exception:
            self.append_traceback(traceback.format_exc())
        return None

    def _read_text_file(self, path):
        try:
            sr = StreamReader(path)
            try:
                return sr.ReadToEnd()
            finally:
                sr.Close()
        except Exception:
            self.append_traceback(traceback.format_exc())
            return None

    def _write_text_file(self, path, text):
        try:
            sw = StreamWriter(path, False)
            try:
                sw.Write(text)
            finally:
                sw.Close()
            return True
        except Exception:
            self.append_traceback(traceback.format_exc())
            return False

    def on_new(self):
        if not self._confirm_discard_changes():
            return
        self._suspend_dirty = True
        try:
            self.input_box.Text = ""
        finally:
            self._suspend_dirty = False
        self._filepath = None
        self._set_dirty(False)
        self._update_status("New file")
        self.input_box.Focus()
        self._sync_state()

    def on_open(self):
        if not self._confirm_discard_changes():
            return

        path = self._pick_open_file()
        if not path:
            return

        content = self._read_text_file(path)
        if content is None:
            return

        self._suspend_dirty = True
        try:
            self.input_box.Text = content
            self.input_box.CaretIndex = len(self.input_box.Text)
        finally:
            self._suspend_dirty = False

        self._filepath = path
        self._set_dirty(False)
        self._update_status("Opened {0}".format(os.path.basename(path)))
        self.input_box.Focus()
        self._sync_state()

    def on_save(self):
        if self._filepath is None:
            return self.on_save_as()

        ok = self._write_text_file(self._filepath, self.input_box.Text or "")
        if ok:
            self._set_dirty(False)
            self._update_status("Saved {0}".format(os.path.basename(self._filepath)))
            self._flash_save_button()
            self._sync_state()
            return True
        self._sync_state()
        return False

    def on_save_as(self):
        path = self._pick_save_file()
        if not path:
            self._sync_state()
            return False

        ok = self._write_text_file(path, self.input_box.Text or "")
        if ok:
            self._filepath = path
            self._set_dirty(False)
            self._update_status("Saved {0}".format(os.path.basename(path)))
            self._flash_save_button()
            self._sync_state()
            return True
        self._sync_state()
        return False

    def on_find(self):
        query = self._prompt("Find", self._last_find or "")
        if query is None:
            self._sync_state()
            return
        if query == "":
            self._update_status("Find cancelled")
            self._sync_state()
            return

        self._last_find = query
        text = self.input_box.Text or ""
        start = self.input_box.SelectionStart + self.input_box.SelectionLength
        idx = text.find(query, start)

        if idx == -1 and start > 0:
            idx = text.find(query, 0)

        if idx == -1:
            self._update_status("'{0}' not found".format(query))
            self._sync_state()
            return

        self.input_box.Focus()
        self.input_box.Select(idx, len(query))
        self._update_status("Found '{0}'".format(query))
        self._update_caret_status()
        self._sync_state()

    def on_goto_line(self):
        value = self._prompt("Go To Line", "")
        if value is None or value == "":
            self._sync_state()
            return

        try:
            target = int(value)
        except:
            self._update_status("Invalid line number")
            self._sync_state()
            return

        text = self.input_box.Text or ""
        if target <= 1:
            idx = 0
        else:
            idx = 0
            current = 1
            while current < target:
                next_nl = text.find("\n", idx)
                if next_nl == -1:
                    idx = len(text)
                    break
                idx = next_nl + 1
                current += 1

        self.input_box.Focus()
        self.input_box.Select(idx, 0)
        self._update_status("Moved to line {0}".format(target))
        self._update_caret_status()
        self._sync_state()

    def _prompt(self, title, default_text):
        xaml = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{0}"
        Width="380"
        Height="150"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="{0}" FontSize="15" FontFamily="Segoe UI"/>
        <TextBox x:Name="PromptBox" Grid.Row="2" MinHeight="26" Text="{1}" FontSize="15"/>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="OkButton" Width="80" Height="26" Margin="0,0,8,0">OK</Button>
            <Button x:Name="CancelButton" Width="80" Height="26">Cancel</Button>
        </StackPanel>
    </Grid>
</Window>
""".format(title.replace('"', "'"), default_text.replace('"', "'"))

        sr = StringReader(xaml)
        xr = XmlReader.Create(sr)
        win = XamlReader.Load(xr)
        box = win.FindName("PromptBox")
        ok = win.FindName("OkButton")
        cancel = win.FindName("CancelButton")

        result = {"value": None}

        def do_ok(sender, args):
            result["value"] = box.Text
            win.Close()

        def do_cancel(sender, args):
            result["value"] = None
            win.Close()

        def on_key_down(sender, args):
            if args.Key == Key.Escape:
                result["value"] = None
                win.Close()
                args.Handled = True
                return
            if args.Key == Key.Enter or args.Key == Key.Return:
                result["value"] = box.Text
                win.Close()
                args.Handled = True
                return

        ok.Click += do_ok
        cancel.Click += do_cancel
        win.KeyDown += on_key_down

        try:
            win.Owner = self.window
        except:
            pass

        win.ShowDialog()
        return result["value"]

    def _get_line_bounds_from_selection(self):
        text = self.input_box.Text or ""
        sel_start = self.input_box.SelectionStart
        sel_len = self.input_box.SelectionLength
        sel_end = sel_start + sel_len

        line_start = text.rfind("\n", 0, sel_start)
        if line_start == -1:
            block_start = 0
        else:
            block_start = line_start + 1

        if sel_len == 0:
            line_end = text.find("\n", sel_start)
            if line_end == -1:
                block_end = len(text)
            else:
                block_end = line_end
        else:
            if sel_end > 0 and sel_end <= len(text) and text[sel_end - 1] == "\n":
                search_from = sel_end - 1
            else:
                search_from = sel_end

            line_end = text.find("\n", search_from)
            if line_end == -1:
                block_end = len(text)
            else:
                block_end = line_end

        return block_start, block_end

    def _modify_selected_lines(self, transform_func):
        text = self.input_box.Text or ""
        block_start, block_end = self._get_line_bounds_from_selection()

        selected_block = text[block_start:block_end]
        lines = selected_block.split("\n")
        new_lines = transform_func(lines)
        new_block = "\n".join(new_lines)

        new_text = text[:block_start] + new_block + text[block_end:]

        self._suspend_dirty = True
        try:
            self.input_box.Text = new_text
            self.input_box.Focus()
            new_caret = block_start + len(new_block)
            self.input_box.Select(new_caret, 0)
        finally:
            self._suspend_dirty = False

        self._set_dirty(True)

    def _comment_lines(self, lines):
        result = []
        for line in lines:
            stripped = line.lstrip(" \t")
            indent_len = len(line) - len(stripped)
            indent = line[:indent_len]

            if stripped == "":
                result.append(line)
            else:
                result.append(indent + "# " + stripped)
        return result

    def _uncomment_lines(self, lines):
        result = []
        for line in lines:
            stripped = line.lstrip(" \t")
            indent_len = len(line) - len(stripped)
            indent = line[:indent_len]

            if stripped.startswith("# "):
                result.append(indent + stripped[2:])
            elif stripped.startswith("#"):
                result.append(indent + stripped[1:])
            else:
                result.append(line)
        return result

    def _indent_lines(self, lines):
        return [("    " + line) if line != "" else "    " for line in lines]

    def _outdent_lines(self, lines):
        result = []
        for line in lines:
            if line.startswith("    "):
                result.append(line[4:])
            elif line.startswith("\t"):
                result.append(line[1:])
            elif line.startswith("   "):
                result.append(line[3:])
            elif line.startswith("  "):
                result.append(line[2:])
            elif line.startswith(" "):
                result.append(line[1:])
            else:
                result.append(line)
        return result

    def comment(self):
        self._modify_selected_lines(self._comment_lines)
        self._update_status("Commented selection")
        self._sync_state()

    def uncomment(self):
        self._modify_selected_lines(self._uncomment_lines)
        self._update_status("Uncommented selection")
        self._sync_state()

    def indent(self):
        self._modify_selected_lines(self._indent_lines)
        self._update_status("Indented selection")
        self._sync_state()

    def outdent(self):
        self._modify_selected_lines(self._outdent_lines)
        self._update_status("Outdented selection")
        self._sync_state()

    def duplicate_line(self):
        text = self.input_box.Text or ""
        block_start, block_end = self._get_line_bounds_from_selection()
        selected = text[block_start:block_end]
        insert_text = selected
        if block_end < len(text):
            insert_text = selected + "\n"
        else:
            if selected != "":
                insert_text = "\n" + selected

        new_text = text[:block_end] + insert_text + text[block_end:]

        self._suspend_dirty = True
        try:
            self.input_box.Text = new_text
            self.input_box.Focus()
            self.input_box.Select(block_end, len(insert_text))
        finally:
            self._suspend_dirty = False

        self._set_dirty(True)
        self._update_status("Duplicated line(s)")
        self._sync_state()

    def delete_line(self):
        text = self.input_box.Text or ""
        block_start, block_end = self._get_line_bounds_from_selection()

        if block_end < len(text) and text[block_end:block_end + 1] == "\n":
            block_end += 1
        elif block_start > 0 and block_start <= len(text):
            if text[block_start - 1:block_start] == "\n":
                block_start -= 1

        new_text = text[:block_start] + text[block_end:]

        self._suspend_dirty = True
        try:
            self.input_box.Text = new_text
            self.input_box.Focus()
            self.input_box.Select(min(block_start, len(new_text)), 0)
        finally:
            self._suspend_dirty = False

        self._set_dirty(True)
        self._update_status("Deleted lines")
        self._sync_state()

    def _normalize_paste(self, text):
        if not text:
            return text

        text = text.replace("\t", "    ")
        lines = text.split("\n")

        min_indent = None
        for line in lines:
            stripped = line.lstrip()
            if stripped:
                indent = len(line) - len(stripped)
                if min_indent is None or indent < min_indent:
                    min_indent = indent

        if min_indent is None:
            return text

        stripped_lines = [line[min_indent:] for line in lines]

        caret = self.input_box.CaretIndex
        text_full = self.input_box.Text or ""

        line_start = text_full.rfind("\n", 0, caret)
        if line_start == -1:
            line_start = 0
        else:
            line_start += 1

        current_line = text_full[line_start:caret]

        base_indent = ""
        for ch in current_line:
            if ch in (" ", "\t"):
                base_indent += ch
            else:
                break

        final_lines = []
        for i, line in enumerate(stripped_lines):
            if i == 0:
                final_lines.append(line)
            else:
                final_lines.append(base_indent + line)

        return "\n".join(final_lines)

    def _safe_dir(self, obj):
        try:
            return dir(obj)
        except:
            return []

    def _get_token_before_caret(self):
        text = self.input_box.Text or ""
        caret = self.input_box.CaretIndex
        left = text[:caret]

        valid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_."
        i = len(left) - 1

        while i >= 0 and left[i] in valid:
            i -= 1

        return left[i + 1:]

    def _resolve_dotted_target(self, expr):
        if not expr:
            return None

        parts = expr.split(".")
        base_name = parts[0]

        if base_name in self.scope:
            obj = self.scope[base_name]
        else:
            return None

        for part in parts[1:]:
            if not part:
                break
            try:
                obj = getattr(obj, part)
            except:
                return None

        return obj

    def _get_autocomplete_items(self, token):
        results = []

        try:
            import Autodesk.Revit.DB as RDB
            results.extend(dir(RDB))
        except:
            pass

        try:
            import Autodesk.Revit.UI as RUI
            results.extend(dir(RUI))
        except:
            pass

        try:
            results.extend(dir(__builtins__))
        except:
            pass

        try:
            results.extend(list(self.scope.keys()))
        except:
            pass

        if "." in token:
            left, right = token.rsplit(".", 1)
            obj = self._resolve_dotted_target(left)
            if obj is not None:
                return sorted([x for x in self._safe_dir(obj) if x.startswith(right)])

        token = token or ""
        return sorted([x for x in set(results) if x.startswith(token)])

    def _close_completion(self):
        try:
            if self._completion_popup is not None:
                self._completion_popup.IsOpen = False
        except:
            pass

    def _insert_completion(self, value):
        token = self._get_token_before_caret()
        text = self.input_box.Text or ""
        caret = self.input_box.CaretIndex

        start = caret - len(token)
        if start < 0:
            start = 0

        new_text = text[:start] + value + text[caret:]

        self._suspend_dirty = True
        try:
            self.input_box.Text = new_text
            self.input_box.CaretIndex = start + len(value)
            self.input_box.Focus()
        finally:
            self._suspend_dirty = False

        self._set_dirty(True)
        self._close_completion()
        self._sync_state()

    def _show_completion(self, items):
        self._close_completion()

        if not items:
            return

        popup = Popup()
        popup.PlacementTarget = self.input_box
        popup.StaysOpen = False
        popup.AllowsTransparency = True

        lb = ListBox()
        lb.Width = 260
        lb.MaxHeight = 220

        for item in items[:30]:
            lb.Items.Add(item)

        lb.SelectedIndex = 0
        popup.Child = lb
        popup.IsOpen = True

        def on_double_click(sender, args):
            try:
                if lb.SelectedItem is not None:
                    self._insert_completion(str(lb.SelectedItem))
            except:
                pass

        lb.MouseDoubleClick += on_double_click

        self._completion_popup = popup
        self._completion_list = lb

    def trigger_completion(self):
        token = self._get_token_before_caret()
        if not token:
            self._close_completion()
            return

        items = self._get_autocomplete_items(token)
        self._show_completion(items)

    def on_run(self):
        code = self.input_box.Text
        if not code or not code.strip():
            self.append_stream_text("[no code to run]\n", True)
            self._update_status("No code to run")
            self._sync_state()
            return

        self.run_requested = True
        self.run_code = code
        self._sync_state()
        self.window.Close()

    def on_close(self):
        if not self._confirm_discard_changes():
            return
        self._sync_state()
        self.window.Close()

    def on_input_changed(self):
        self._close_completion()
        self._update_line_numbers()
        self._sync_line_number_scroll()
        self._update_caret_status()
        if not self._suspend_dirty:
            self._set_dirty(True)
        self._sync_state()

    def on_input_view_changed(self):
        self._sync_line_number_scroll()
        self._update_caret_status()
        self._sync_state()

    def _update_line_numbers(self):
        text = self.input_box.Text or ""
        line_count = text.count("\n") + 1
        self.line_box.Text = "\n".join([str(i) for i in range(1, line_count + 1)])

    def _sync_line_number_scroll(self):
        try:
            first_visible = self.input_box.GetFirstVisibleLineIndex()
            if first_visible >= 0:
                self.line_box.ScrollToLine(first_visible)
        except:
            pass

    def on_input_key_down(self, args):
        mods = Keyboard.Modifiers

        if self._completion_popup is not None and self._completion_popup.IsOpen:
            if args.Key == Key.Down:
                if self._completion_list is not None and self._completion_list.SelectedIndex < self._completion_list.Items.Count - 1:
                    self._completion_list.SelectedIndex += 1
                args.Handled = True
                return

            if args.Key == Key.Up:
                if self._completion_list is not None and self._completion_list.SelectedIndex > 0:
                    self._completion_list.SelectedIndex -= 1
                args.Handled = True
                return

            if args.Key == Key.Return or args.Key == Key.Enter or args.Key == Key.Tab:
                if self._completion_list is not None and self._completion_list.SelectedItem is not None:
                    self._insert_completion(str(self._completion_list.SelectedItem))
                    args.Handled = True
                    return

            if args.Key == Key.Escape:
                self._close_completion()
                args.Handled = True
                return

        if mods == ModifierKeys.Control and args.Key == Key.Space:
            self.trigger_completion()
            args.Handled = True
            return

        if args.Key == Key.Return or args.Key == Key.Enter:
            text = self.input_box.Text or ""
            caret = self.input_box.CaretIndex

            line_start = text.rfind("\n", 0, caret)
            if line_start == -1:
                line_start = 0
            else:
                line_start += 1

            line = text[line_start:caret]
            stripped = line.strip()

            indent = ""
            for ch in line:
                if ch in (" ", "\t"):
                    indent += ch
                else:
                    break

            dedent_starters = ("elif", "else", "except", "finally")
            for kw in dedent_starters:
                if stripped.startswith(kw):
                    indent = indent[:-4] if len(indent) >= 4 else ""
                    break

            if stripped.endswith(":"):
                indent += "    "

            dedent_keywords = ("return", "pass", "break", "continue", "raise")
            for kw in dedent_keywords:
                if stripped.startswith(kw):
                    indent = indent[:-4] if len(indent) >= 4 else ""
                    break

            openers = "([{"
            closers = ")]}"
            open_count = sum(line.count(c) for c in openers)
            close_count = sum(line.count(c) for c in closers)

            if open_count > close_count:
                indent += "    "
            elif stripped.startswith(tuple(closers)):
                indent = indent[:-4] if len(indent) >= 4 else ""

            insert = "\n" + indent

            self._suspend_dirty = True
            try:
                self.input_box.Text = text[:caret] + insert + text[caret:]
                self.input_box.CaretIndex = caret + len(insert)
            finally:
                self._suspend_dirty = False

            self._set_dirty(True)
            self._sync_state()
            args.Handled = True
            return

        if args.Key == Key.Tab and mods == ModifierKeys.None:
            if self.input_box.SelectionLength == 0:
                caret = self.input_box.CaretIndex
                text = self.input_box.Text or ""

                self._suspend_dirty = True
                try:
                    self.input_box.Text = text[:caret] + "    " + text[caret:]
                    self.input_box.CaretIndex = caret + 4
                finally:
                    self._suspend_dirty = False

                self._set_dirty(True)
                self._sync_state()
            else:
                self.indent()

            args.Handled = True
            return

        if args.Key == Key.Tab and mods == ModifierKeys.Shift:
            self.outdent()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.V:
            try:
                from System.Windows import Clipboard
                pasted = Clipboard.GetText()

                if pasted:
                    normalized = self._normalize_paste(pasted)

                    caret = self.input_box.CaretIndex
                    text = self.input_box.Text or ""

                    self._suspend_dirty = True
                    try:
                        self.input_box.Text = text[:caret] + normalized + text[caret:]
                        self.input_box.CaretIndex = caret + len(normalized)
                    finally:
                        self._suspend_dirty = False

                    self._set_dirty(True)
                    self._sync_state()
                    args.Handled = True
                    return
            except:
                pass

        if mods == ModifierKeys.Control and args.Key == Key.Return:
            self.on_run()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.S:
            self.on_save()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.O:
            self.on_open()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.N:
            self.on_new()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.F:
            self.on_find()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.G:
            self.on_goto_line()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.A:
            self.input_box.SelectAll()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.D:
            self.duplicate_line()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.L:
            self.delete_line()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.OemQuestion:
            self.comment()
            args.Handled = True
            return

        if mods == (ModifierKeys.Control | ModifierKeys.Shift) and args.Key == Key.OemQuestion:
            self.uncomment()
            args.Handled = True
            return

        if mods == ModifierKeys.Control and args.Key == Key.Subtract:
            self.uncomment()
            args.Handled = True
            return

        if args.Key == Key.Up and mods == ModifierKeys.Alt:
            if self._history:
                self._history_index = max(0, self._history_index - 1)
                self._suspend_dirty = True
                try:
                    self.input_box.Text = self._history[self._history_index]
                    self.input_box.CaretIndex = len(self.input_box.Text)
                finally:
                    self._suspend_dirty = False
                self._set_dirty(True)
                self._update_status("History previous")
                self._sync_state()
            args.Handled = True
            return

        if args.Key == Key.Down and mods == ModifierKeys.Alt:
            if self._history:
                self._history_index = min(len(self._history) - 1, self._history_index + 1)
                self._suspend_dirty = True
                try:
                    self.input_box.Text = self._history[self._history_index]
                    self.input_box.CaretIndex = len(self.input_box.Text)
                finally:
                    self._suspend_dirty = False
                self._set_dirty(True)
                self._update_status("History next")
                self._sync_state()
            args.Handled = True
            return


session = ShellSessionState()

while True:
    shell = ShellWindow(session)
    shell.show()

    session = shell.state

    if not shell.run_requested:
        break

    code = shell.run_code or ""

    uiapp = None
    uidoc = None
    doc = None
    app = None

    try:
        uiapp = __revit__
    except:
        uiapp = None

    try:
        if uiapp is not None:
            uidoc = uiapp.ActiveUIDocument
    except:
        uidoc = None

    try:
        if uidoc is not None:
            doc = uidoc.Document
    except:
        doc = None

    try:
        if uiapp is not None:
            app = uiapp.Application
    except:
        app = None

    session.scope["__revit__"] = uiapp
    session.scope["uiapp"] = uiapp
    session.scope["uidoc"] = uidoc
    session.scope["doc"] = doc
    session.scope["app"] = app

    try:
        if "revit" in session.scope:
            session.scope["revit"].uiapp = uiapp
            session.scope["revit"].uidoc = uidoc
            session.scope["revit"].doc = doc
            session.scope["revit"].app = app
    except:
        pass

    old_stdout = sys.stdout
    old_stderr = sys.stderr
    sys.stdout = StateRedirector(session, False)
    sys.stderr = StateRedirector(session, True)

    try:
        append_state_output(session, "\n--- Run ---\n", False)
        exec(code, session.scope, session.scope)
        append_state_output(session, "[done]\n", False)
        session.status_text = "Run completed"
        session.history.append(code)
        session.history_index = len(session.history)
    except Exception:
        append_state_output(session, traceback.format_exc(), True)
        session.status_text = "Run failed"
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr