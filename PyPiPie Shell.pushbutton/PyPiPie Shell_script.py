# -*- coding: utf-8 -*-
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
from System.Windows.Media.Animation import ColorAnimation
from System.Windows.Input import Key, ModifierKeys, Keyboard
from System.Windows.Threading import DispatcherTimer
from System.Windows import MessageBoxResult
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult


XAML = r"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PyPiPie Shell"
    Width="1180"
    Height="820"
    MinWidth="900"
    MinHeight="620"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="3*" />
            <RowDefinition Height="6" />
            <RowDefinition Height="2*" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="42" />
        </Grid.RowDefinitions>

        <!-- TOP BAR -->
        <Border Grid.Row="0" BorderBrush="#FFBFBFBF" BorderThickness="1" CornerRadius="4" Padding="6" Margin="0,0,0,4">
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

                <Button x:Name="NewButton"      Grid.Column="0"  Width="64" Height="22" Content="New"/>
                <Button x:Name="OpenButton"     Grid.Column="2"  Width="64" Height="22" Content="Open"/>
                <Button x:Name="SaveButton"     Grid.Column="4"  Width="64" Height="22" Content="Save"/>
                <Button x:Name="SaveAsButton"   Grid.Column="6"  Width="72" Height="22" Content="Save As"/>

                <Button x:Name="FindButton"     Grid.Column="8"  Width="64" Height="22" Content="Find"/>
                <Button x:Name="GotoButton"     Grid.Column="10" Width="64" Height="22" Content="Go To"/>

                <Button x:Name="CommentButton"   Grid.Column="12" Width="28" Height="22" Content="#"/>
                <Button x:Name="UncommentButton" Grid.Column="14" Width="28" Height="22" Content="-"/>
                <Button x:Name="IndentButton"    Grid.Column="16" Width="28" Height="22" Content="&gt;"/>

                <Button x:Name="OutdentButton"   Grid.Column="18" Width="28" Height="22" Content="&lt;"/>
                <Button x:Name="DuplicateButton" Grid.Column="20" Width="72" Height="22" Content="Duplicate"/>
                <Button x:Name="DeleteLineButton" Grid.Column="22" Width="86" Height="22" Content="Delete Lines"/>

                <TextBlock x:Name="FileLabel"
                           Grid.Column="24"
                           VerticalAlignment="Center"
                           Foreground="#FF444444"
                           FontFamily="Consolas"
                           Text="untitled.py"/>
            </Grid>
        </Border>

        <!-- INPUT -->
        <Border Grid.Row="1"
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
        <GridSplitter Grid.Row="2"
                      Height="6"
                      HorizontalAlignment="Stretch"
                      VerticalAlignment="Stretch"
                      Background="#FFD9D9D9"
                      ResizeDirection="Rows"
                      ResizeBehavior="PreviousAndNext"/>

        <!-- OUTPUT -->
        <Border Grid.Row="3"
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

        <!-- STATUS -->
        <Border x:Name="StatusBorder"
                Grid.Row="4"
                BorderBrush="#FFBFBFBF"
                BorderThickness="1"
                CornerRadius="4"
                Padding="6,2,6,2"
                Margin="0,0,0,6"
                Background="Transparent">
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
        <Grid Grid.Row="5" Margin="0,0,0,0">
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


class StreamRedirector(object):
    def __init__(self, shell, is_error=False):
        self.shell = shell
        self.is_error = is_error

    def write(self, text):
        if text is None:
            return
        s = str(text)
        if not s:
            return
        self.shell.append_stream_text(s, self.is_error)

    def flush(self):
        pass


class ShellWindow(object):
    def __init__(self):
        self.scope = {}
        self._current_paragraph = None
        self._current_is_error = False
        self._filepath = None
        self._is_dirty = False
        self._history = []
        self._history_index = -1
        self._last_find = None
        self._suspend_dirty = False

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

        self.input_box = self.window.FindName("InputBox")
        self.line_box = self.window.FindName("LineNumbersBox")
        self.output_box = self.window.FindName("OutputBox")

        self._status_brush = SolidColorBrush(ColorConverter.ConvertFromString("Transparent"))
        self.status_border.Background = self._status_brush

        self._status_timer = DispatcherTimer()
        self._status_timer.Interval = TimeSpan.FromSeconds(2)
        self._status_timer.Tick += self._handle_status_timer_tick

        self._configure_output()
        self._seed_scope()
        self._bind_events()
        self._update_line_numbers()
        self._refresh_title()
        self._update_status("Ready")
        self._update_caret_status()

    def _configure_output(self):
        self.output_box.Document.Blocks.Clear()
        self._current_paragraph = None
        self._current_is_error = False

    def _seed_scope(self):
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

        revit_ctx = RevitContext(doc=doc, uidoc=uidoc, uiapp=uiapp, app=app)

        self.scope["__builtins__"] = __builtins__
        self.scope["sys"] = sys
        self.scope["clr"] = clr

        self.scope["__revit__"] = uiapp
        self.scope["uiapp"] = uiapp
        self.scope["uidoc"] = uidoc
        self.scope["doc"] = doc
        self.scope["app"] = app
        self.scope["revit"] = revit_ctx

        exec("import sys", self.scope, self.scope)
        exec("import clr", self.scope, self.scope)
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
        self._animate_status_to("Transparent")

    def show(self):
        self.window.ShowDialog()

    def _current_filename(self):
        if self._filepath:
            return os.path.basename(self._filepath)
        return "untitled.py"

    def _refresh_title(self):
        star = "*" if self._is_dirty else ""
        name = self._current_filename()
        self.window.Title = "PyPiPie Shell - {0}{1}".format(name, star)
        self.file_label.Text = name
        self.status_file.Text = name
        self.status_dirty.Text = "Modified" if self._is_dirty else "Saved"

    def _set_dirty(self, value):
        self._is_dirty = bool(value)
        self._refresh_title()

    def _stop_status_timer(self):
        try:
            self._status_timer.Stop()
        except:
            pass

    def _set_status_color_immediate(self, color_name):
        try:
            self._status_brush.Color = ColorConverter.ConvertFromString(color_name)
        except:
            pass

    def _animate_status_to(self, color_name):
        try:
            target = ColorConverter.ConvertFromString(color_name)
            anim = ColorAnimation()
            anim.To = target
            anim.Duration = TimeSpan.FromMilliseconds(300)
            self._status_brush.BeginAnimation(SolidColorBrush.ColorProperty, anim)
        except:
            try:
                self._set_status_color_immediate(color_name)
            except:
                pass

    def _flash_status(self, color_name, auto_clear=True):
        self._stop_status_timer()
        self._animate_status_to(color_name)
        if auto_clear:
            try:
                self._status_timer.Start()
            except:
                pass

    def _update_status(self, text):
        self.status_left.Text = text
        msg = str(text or "")

        self._stop_status_timer()

        if msg.startswith("Saved"):
            self._flash_status("PaleGreen", True)
        elif msg.startswith("Run failed") or msg.startswith("Error") or msg.startswith("Unhandled") or msg.startswith("Traceback"):
            self._flash_status("MistyRose", False)
        elif msg.startswith("Run completed") or msg.startswith("Opened") or msg.startswith("New file") or msg.startswith("Output cleared"):
            self._flash_status("LightGoldenrodYellow", True)
        else:
            self._animate_status_to("Transparent")

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

    def append_stream_text(self, text, is_error=False):
        if text is None:
            return

        s = str(text).replace("\r\n", "\n").replace("\r", "\n")
        parts = s.split("\n")

        for i, part in enumerate(parts):
            if part:
                p = self._ensure_paragraph(is_error)
                run = Run(part)
                run.Foreground = Brushes.Red if is_error else Brushes.Black
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

    def clear_output(self):
        self.output_box.Document.Blocks.Clear()
        self._current_paragraph = None
        self._current_is_error = False
        self._update_status("Output cleared")

    def _big_message_box(self, message, title="Message", buttons="OKCancel", icon="info"):
        xaml = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{0}"
        Width="460"
        Height="220"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="14">
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
                       VerticalAlignment="Center"
                       Margin="10,0,0,0"
                       Text="{1}" />
        </Grid>

        <StackPanel Grid.Row="1"
                    Orientation="Horizontal"
                    HorizontalAlignment="Right"
                    Margin="0,14,0,0">
            <Button x:Name="YesBtn" Width="90" Height="30" Margin="0,0,8,0">Yes</Button>
            <Button x:Name="NoBtn" Width="90" Height="30" Margin="0,0,8,0">No</Button>
            <Button x:Name="CancelBtn" Width="90" Height="30">Cancel</Button>
        </StackPanel>
    </Grid>
</Window>
""".format(title.replace('"', "'"), message.replace('"', "'"))

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

    def on_save(self):
        if self._filepath is None:
            return self.on_save_as()

        ok = self._write_text_file(self._filepath, self.input_box.Text or "")
        if ok:
            self._set_dirty(False)
            self._update_status("Saved {0}".format(os.path.basename(self._filepath)))
            return True
        return False

    def on_save_as(self):
        path = self._pick_save_file()
        if not path:
            return False

        ok = self._write_text_file(path, self.input_box.Text or "")
        if ok:
            self._filepath = path
            self._set_dirty(False)
            self._update_status("Saved {0}".format(os.path.basename(path)))
            return True
        return False

    def on_find(self):
        query = self._prompt("Find", self._last_find or "")
        if query is None:
            return
        if query == "":
            self._update_status("Find cancelled")
            return

        self._last_find = query
        text = self.input_box.Text or ""
        start = self.input_box.SelectionStart + self.input_box.SelectionLength
        idx = text.find(query, start)

        if idx == -1 and start > 0:
            idx = text.find(query, 0)

        if idx == -1:
            self._update_status("'{0}' not found".format(query))
            return

        self.input_box.Focus()
        self.input_box.Select(idx, len(query))
        self._update_status("Found '{0}'".format(query))
        self._update_caret_status()

    def on_goto_line(self):
        value = self._prompt("Go To Line", "")
        if value is None or value == "":
            return

        try:
            target = int(value)
        except:
            self._update_status("Invalid line number")
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

    def _prompt(self, title, default_text):
        xaml = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{0}"
        Width="360"
        Height="140"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="{0}" Margin="0,0,0,0"/>
        <TextBox x:Name="PromptBox" Grid.Row="2" MinHeight="24" Text="{1}" />
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="OkButton" Width="70" Height="24" Margin="0,0,8,0">OK</Button>
            <Button x:Name="CancelButton" Width="70" Height="24">Cancel</Button>
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
            win.DialogResult = True
            win.Close()

        def do_cancel(sender, args):
            win.DialogResult = False
            win.Close()

        ok.Click += do_ok
        cancel.Click += do_cancel
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
            self.input_box.Select(block_start, len(new_block))
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

    def uncomment(self):
        self._modify_selected_lines(self._uncomment_lines)
        self._update_status("Uncommented selection")

    def indent(self):
        self._modify_selected_lines(self._indent_lines)
        self._update_status("Indented selection")

    def outdent(self):
        self._modify_selected_lines(self._outdent_lines)
        self._update_status("Outdented selection")

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

    def on_run(self):
        code = self.input_box.Text
        if not code or not code.strip():
            self.append_stream_text("[no code to run]\n", True)
            self._update_status("No code to run")
            return

        code = code.replace("\r\n", "\n").replace("\r", "\n")

        self.append_stream_text("\n--- Run ---\n", False)

        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = StreamRedirector(self, False)
        sys.stderr = StreamRedirector(self, True)

        try:
            exec(code, self.scope, self.scope)
            self.append_stream_text("[done]\n", False)
            self._update_status("Run completed")
            self._history.append(code)
            self._history_index = len(self._history)
        except Exception:
            self.append_traceback(traceback.format_exc())
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    def on_close(self):
        if not self._confirm_discard_changes():
            return
        self.window.Close()

    def on_input_changed(self):
        self._update_line_numbers()
        self._sync_line_number_scroll()
        self._update_caret_status()
        if not self._suspend_dirty:
            self._set_dirty(True)

    def on_input_view_changed(self):
        self._sync_line_number_scroll()
        self._update_caret_status()

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

        if args.Key == Key.Tab and mods == ModifierKeys.None:
            self.indent()
            args.Handled = True
            return

        if args.Key == Key.Tab and mods == ModifierKeys.Shift:
            self.outdent()
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
            args.Handled = True
            return


shell = ShellWindow()
shell.show()