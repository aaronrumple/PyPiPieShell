# -*- coding: utf-8 -*-
"""AvalonEdit editors for PyPiPie Shell.

* :func:`create_editor` builds a themed ``TextEditor``.
* :class:`PythonEditor` adds Python behaviour to one: highlighting,
  smart indent, folding, completion (``.``, typing, Ctrl+Space), signature
  help (``(`` / Ctrl+Shift+Space), hover tips, line commands with a single
  undo step each, find-match and error-line markers.
* :class:`OutputView` is the read-only, colour-coded output console.
"""
import os
import bisect

import clr

from pypipieshell import avalon
avalon.load()

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")

from System import Action, TimeSpan  # noqa: E402
from System.Collections.Generic import List  # noqa: E402
from System.Windows import (Thickness, TextWrapping, FontWeights, Rect,  # noqa: E402
                            HorizontalAlignment, VerticalAlignment, CornerRadius,
                            FontStyles)
from System.Windows.Controls import (StackPanel, TextBlock, Border, Orientation,  # noqa: E402
                                     ToolTip, ScrollBarVisibility)
from System.Windows.Controls.Primitives import PlacementMode  # noqa: E402
from System.Windows.Documents import Run  # noqa: E402
from System.Windows.Input import Key, ModifierKeys, Keyboard  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Windows.Media import FontFamily, Pen, Typeface  # noqa: E402
from System.Windows.Threading import DispatcherTimer  # noqa: E402
from System.Xml import XmlReader  # noqa: E402

from ICSharpCode.AvalonEdit import TextEditor  # noqa: E402
from ICSharpCode.AvalonEdit.CodeCompletion import (CompletionWindow,  # noqa: E402
                                                   InsightWindow, ICompletionData)
from ICSharpCode.AvalonEdit.Document import TextSegment  # noqa: E402
from ICSharpCode.AvalonEdit.Folding import FoldingManager, NewFolding  # noqa: E402
from ICSharpCode.AvalonEdit.Highlighting import HighlightingManager  # noqa: E402
from ICSharpCode.AvalonEdit.Highlighting.Xshd import HighlightingLoader  # noqa: E402
from ICSharpCode.AvalonEdit.Rendering import (DocumentColorizingTransformer,  # noqa: E402
                                              IBackgroundRenderer, KnownLayer,
                                              BackgroundGeometryBuilder, VisualLineElement)

from pypipieshell import pyparse  # noqa: E402
from pypipieshell.completion import KINDS  # noqa: E402

RESOURCES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources")
_HIGHLIGHTING = {}


def python_highlighting(dark):
    """Load (and cache) the light or dark Python XSHD definition."""
    key = "dark" if dark else "light"
    if key not in _HIGHLIGHTING:
        path = os.path.join(RESOURCES, "Python-Dark.xshd" if dark else "Python-Light.xshd")
        reader = XmlReader.Create(path)
        try:
            _HIGHLIGHTING[key] = HighlightingLoader.Load(reader, HighlightingManager.Instance)
        finally:
            reader.Close()
    return _HIGHLIGHTING[key]


def _font_family(name):
    return FontFamily(name or "Consolas")


def create_editor(read_only=False, line_numbers=True):
    ed = TextEditor()
    ed.IsReadOnly = read_only
    ed.ShowLineNumbers = line_numbers
    ed.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
    ed.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
    ed.Padding = Thickness(4, 2, 2, 2)
    opts = ed.Options
    opts.ConvertTabsToSpaces = True
    opts.IndentationSize = 4
    opts.EnableHyperlinks = False
    opts.EnableEmailHyperlinks = False
    opts.AllowScrollBelowDocument = False
    for name, value in (("HighlightCurrentLine", not read_only),
                        ("CutCopyWholeLine", True),
                        ("EnableRectangularSelection", True),
                        ("EnableTextDragDrop", True)):
        try:
            setattr(opts, name, value)
        except Exception:
            pass
    return ed


def apply_editor_theme(ed, palette, settings, output=False):
    ed.Background = palette.brush("OutputBackground" if output else "EditorBackground")
    ed.Foreground = palette.brush("OutputText" if output else "EditorForeground")
    ed.LineNumbersForeground = palette.brush("LineNumbers")
    ed.FontFamily = _font_family(settings["font_family"])
    ed.FontSize = float(settings["font_size"])
    ta = ed.TextArea
    ta.SelectionBrush = palette.brush("Selection")
    ta.SelectionBorder = Pen(palette.brush("SelectionBorder"), 1.0)
    ta.SelectionForeground = None
    try:
        ta.SelectionCornerRadius = 2.0
    except Exception:
        pass
    ta.Caret.CaretBrush = palette.brush("Caret")
    tv = ta.TextView
    try:
        tv.CurrentLineBackground = palette.brush("CurrentLine")
        tv.CurrentLineBorder = Pen(palette.brush("CurrentLineBorder"), 1.0)
    except Exception:
        pass
    try:
        tv.LinkTextForegroundBrush = palette.brush("OutputInfo")
    except Exception:
        pass
    # folding margin markers
    try:
        from ICSharpCode.AvalonEdit.Folding import FoldingMargin
        for margin in ta.LeftMargins:
            if isinstance(margin, FoldingMargin):
                margin.FoldingMarkerBrush = palette.brush("LineNumbers")
                margin.FoldingMarkerBackgroundBrush = palette.brush("EditorBackground")
                margin.SelectedFoldingMarkerBrush = palette.brush("EditorForeground")
                margin.SelectedFoldingMarkerBackgroundBrush = palette.brush("EditorBackground")
    except Exception:
        pass
    tv.Redraw()


# ---------------------------------------------------------------------------
# popup styling (completion list, insight, tooltips)
# ---------------------------------------------------------------------------
_POPUP_XAML = u"""
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Style TargetType="ListBoxItem">
    <Setter Property="Padding" Value="2,1,6,1"/>
    <Setter Property="Foreground" Value="{DynamicResource PPS.PopupForeground}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ListBoxItem">
          <Border x:Name="Bd" Background="Transparent" Padding="{TemplateBinding Padding}" CornerRadius="2">
            <ContentPresenter/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsSelected" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{DynamicResource PPS.PopupSelection}"/>
            </Trigger>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource PPS.PopupBorder}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style TargetType="ToolTip">
    <Setter Property="Background" Value="{DynamicResource PPS.PopupBackground}"/>
    <Setter Property="Foreground" Value="{DynamicResource PPS.PopupForeground}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource PPS.PopupBorder}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="8,6"/>
    <Setter Property="MaxWidth" Value="640"/>
  </Style>
</ResourceDictionary>
"""


def style_popup(window_or_element, palette):
    """Apply palette + popup styles to a completion/insight window or tooltip."""
    from pypipieshell import theme
    theme.apply_resources(window_or_element, palette)
    try:
        rd = XamlReader.Parse(_POPUP_XAML)
        window_or_element.Resources.MergedDictionaries.Add(rd)
    except Exception:
        pass
    try:
        window_or_element.Background = palette.brush("PopupBackground")
        window_or_element.Foreground = palette.brush("PopupForeground")
        window_or_element.BorderBrush = palette.brush("PopupBorder")
        window_or_element.BorderThickness = Thickness(1)
    except Exception:
        pass


def doc_element(text, palette, mono_lines=None, font_family=None):
    """Tooltip body: signature lines in monospace, the rest as prose."""
    panel = StackPanel()
    panel.MaxWidth = 620
    if not text:
        return panel
    lines = text.split("\n")
    blank = lines.index("") if "" in lines else len(lines)
    head = lines[:blank] if mono_lines is None else lines[:mono_lines]
    rest = lines[len(head):]
    for i, ln in enumerate(head):
        tb = TextBlock()
        tb.Text = ln
        tb.FontFamily = _font_family(font_family)
        tb.TextWrapping = TextWrapping.Wrap
        tb.Foreground = palette.brush("PopupForeground")
        if i == 0 and len(head) > 1:
            tb.Foreground = palette.brush("SubtleForeground")
            tb.FontSize = 11
        panel.Children.Add(tb)
    prose = "\n".join(rest).strip()
    if prose:
        tb = TextBlock()
        tb.Text = prose
        tb.TextWrapping = TextWrapping.Wrap
        tb.Margin = Thickness(0, 6, 0, 0)
        tb.Foreground = palette.brush("PopupForeground")
        panel.Children.Add(tb)
    return panel


class CompletionData(ICompletionData):
    """AvalonEdit adapter for :class:`completion.Item`."""

    def __init__(self, item, editor):
        self._item = item
        self._editor = editor
        self._content = None

    @property
    def Text(self):
        return self._item.name

    @property
    def Image(self):
        return None

    @property
    def Priority(self):
        return 0.0

    @property
    def Content(self):
        if self._content is None:
            self._content = self._build_content()
        return self._content

    @property
    def Description(self):
        text = self._item.description(self._editor.engine.stubs if self._editor.engine else None)
        if not text:
            return None
        return doc_element(text, self._editor.palette, font_family=self._editor.settings["font_family"])

    def Complete(self, textArea, completionSegment, insertionRequestEventArgs):
        textArea.Document.Replace(completionSegment, self._item.name)

    def _build_content(self):
        pal = self._editor.palette
        glyph, color_key, _ = KINDS.get(self._item.kind, ("?", "KindSnippet", 5))
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        badge = Border()
        badge.Width = 16
        badge.Height = 16
        badge.CornerRadius = CornerRadius(3)
        badge.Margin = Thickness(0, 0, 6, 0)
        badge.Background = pal.brush(color_key)
        gl = TextBlock()
        gl.Text = glyph
        gl.FontSize = 10
        gl.FontWeight = FontWeights.Bold
        gl.Foreground = pal.brush("EditorBackground")
        gl.HorizontalAlignment = HorizontalAlignment.Center
        gl.VerticalAlignment = VerticalAlignment.Center
        badge.Child = gl
        sp.Children.Add(badge)
        name = TextBlock()
        name.Text = self._item.name
        name.VerticalAlignment = VerticalAlignment.Center
        name.Foreground = pal.brush("PopupForeground")
        sp.Children.Add(name)
        if self._item.detail:
            det = TextBlock()
            det.Text = u"  " + self._item.detail
            det.FontSize = 11
            det.VerticalAlignment = VerticalAlignment.Center
            det.Foreground = pal.brush("SubtleForeground")
            sp.Children.Add(det)
        return sp


# ---------------------------------------------------------------------------
# renderers
# ---------------------------------------------------------------------------
class SegmentHighlighter(IBackgroundRenderer):
    """Draws a background for document segments (find matches, error line)."""

    def __init__(self, full_line=False):
        self.brush = None
        self.full_line = full_line
        self.segments = []

    @property
    def Layer(self):
        return KnownLayer.Selection

    def Draw(self, textView, drawingContext):
        if not self.segments or self.brush is None:
            return
        try:
            textView.EnsureVisualLines()
        except Exception:
            return
        doc = textView.Document
        if doc is None:
            return
        n = doc.TextLength
        for start, length in self.segments:
            if start > n:
                continue
            seg = TextSegment()
            seg.StartOffset = start
            seg.Length = max(0, min(length, n - start))
            if self.full_line:
                line = doc.GetLineByOffset(start)
                seg.StartOffset = line.Offset
                seg.Length = max(line.Length, 0)
            for r in BackgroundGeometryBuilder.GetRectsForSegment(textView, seg, self.full_line):
                if self.full_line:
                    r = Rect(0, r.Top, max(textView.ActualWidth, r.Right), r.Height)
                drawingContext.DrawRectangle(self.brush, None, r)
                if self.full_line:
                    break


class OutputColorizer(DocumentColorizingTransformer):
    """Foreground colours for output spans ``(start, end, kind)``."""

    def __init__(self):
        self.starts = []
        self.spans = []
        self.brushes = {}

    def reset(self):
        self.starts = []
        self.spans = []

    def add(self, start, end, kind):
        if kind == "out" or end <= start:
            return
        self.starts.append(start)
        self.spans.append((start, end, kind))

    def shift(self, delta):
        spans = [(s - delta, e - delta, k) for (s, e, k) in self.spans if e - delta > 0]
        self.spans = [(max(0, s), e, k) for (s, e, k) in spans]
        self.starts = [s for (s, _, _) in self.spans]

    def ColorizeLine(self, line):
        if not self.spans:
            return
        ls = line.Offset
        le = line.EndOffset
        if le <= ls:
            return
        i = bisect.bisect_right(self.starts, le) - 1
        while i >= 0:
            s, e, kind = self.spans[i]
            if e <= ls:
                # spans are appended in order and don't overlap
                break
            a = max(s, ls)
            b = min(e, le)
            if a < b:
                brush = self.brushes.get(kind)
                if brush is not None:
                    italic = kind == "echo"

                    def paint(el, brush=brush, italic=italic):
                        props = el.TextRunProperties
                        props.SetForegroundBrush(brush)
                        if italic:
                            try:
                                tf = props.Typeface
                                props.SetTypeface(Typeface(tf.FontFamily, FontStyles.Italic,
                                                           tf.Weight, tf.Stretch))
                            except Exception:
                                pass
                    self.ChangeLinePart(a, b, Action[VisualLineElement](paint))
            i -= 1


# ---------------------------------------------------------------------------
# Python editor behaviour
# ---------------------------------------------------------------------------
_IDENT = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")


class PythonEditor(object):
    """Adds Python editing features to an AvalonEdit ``TextEditor``."""

    def __init__(self, editor, engine, settings, palette, folding=True, is_repl=False):
        self.editor = editor
        self.engine = engine
        self.settings = settings
        self.palette = palette
        self.is_repl = is_repl
        self.completion_window = None
        self.insight_window = None
        self._hover_tip = None
        self.on_status = None  # callable(text)

        self.find_marker = SegmentHighlighter(full_line=False)
        self.error_marker = SegmentHighlighter(full_line=True)
        tv = editor.TextArea.TextView
        tv.BackgroundRenderers.Add(self.error_marker)
        tv.BackgroundRenderers.Add(self.find_marker)

        self.folding_manager = None
        self._fold_timer = None
        if folding:
            self.folding_manager = FoldingManager.Install(editor.TextArea)
            self._fold_timer = DispatcherTimer()
            self._fold_timer.Interval = TimeSpan.FromMilliseconds(600)
            self._fold_timer.Tick += self._on_fold_tick

        ta = editor.TextArea
        ta.TextEntering += self._on_text_entering
        ta.TextEntered += self._on_text_entered
        ta.PreviewKeyDown += self._on_preview_key_down
        editor.TextChanged += self._on_text_changed
        tv.MouseHover += self._on_mouse_hover
        tv.MouseHoverStopped += self._on_mouse_hover_stopped
        editor.PreviewMouseWheel += self._on_mouse_wheel
        self.apply_theme(palette)

    # -- theme --------------------------------------------------------------
    def apply_theme(self, palette):
        self.palette = palette
        apply_editor_theme(self.editor, palette, self.settings)
        self.editor.SyntaxHighlighting = python_highlighting(palette.dark)
        self.find_marker.brush = palette.brush("FindMatch")
        self.error_marker.brush = palette.brush("ErrorLine")
        self.editor.TextArea.TextView.Redraw()

    def apply_settings(self):
        s = self.settings
        ed = self.editor
        ed.FontFamily = _font_family(s["font_family"])
        ed.FontSize = float(s["font_size"])
        if not self.is_repl:
            ed.WordWrap = bool(s["word_wrap"])
            ed.ShowLineNumbers = bool(s["show_line_numbers"])
            try:
                ed.Options.HighlightCurrentLine = bool(s["highlight_current_line"])
            except Exception:
                pass
        ed.Options.ShowSpaces = bool(s["show_whitespace"])
        ed.Options.ShowTabs = bool(s["show_whitespace"])
        if self.folding_manager is not None:
            if s["folding"]:
                self.update_foldings()
            else:
                self.folding_manager.Clear()

    # -- folding ------------------------------------------------------------
    def _on_text_changed(self, sender, args):
        if self.error_marker.segments:
            self.error_marker.segments = []
            self.editor.TextArea.TextView.InvalidateLayer(KnownLayer.Selection)
        if self._fold_timer is not None:
            self._fold_timer.Stop()
            self._fold_timer.Start()

    def _on_fold_tick(self, sender, args):
        self._fold_timer.Stop()
        self.update_foldings()

    def update_foldings(self):
        if self.folding_manager is None or not self.settings["folding"]:
            return
        try:
            folds = List[NewFolding]()
            for start, end, title in pyparse.compute_foldings(self.editor.Text):
                f = NewFolding(start, end)
                f.Name = title
                folds.Add(f)
            self.folding_manager.UpdateFoldings(folds, -1)
        except Exception:
            pass

    # -- typing -------------------------------------------------------------
    def _on_text_entering(self, sender, e):
        cw = self.completion_window
        if cw is not None and e.Text and len(e.Text) > 0:
            if e.Text[0] not in _IDENT:
                cw.Close()

    def _on_text_entered(self, sender, e):
        text = e.Text
        if not text:
            return
        ch = text[-1]
        if ch == "\n" or ch == "\r":
            self._smart_newline()
            return
        if ch == ".":
            self.show_completion(auto=True)
        elif ch in _IDENT:
            if self.completion_window is None and self.settings["complete_as_you_type"]:
                if self._typed_identifier_length() >= 2 and not ch.isdigit():
                    self.show_completion(auto=True, bare_min=2)
        elif ch == "(":
            self.show_signature_help()
        elif ch == ",":
            if self.insight_window is not None:
                self.show_signature_help()
        elif ch == ")":
            if self.insight_window is not None:
                self.insight_window.Close()
        elif ch == ":":
            self._maybe_dedent_block_keyword()

    def _typed_identifier_length(self):
        doc = self.editor.Document
        off = self.editor.CaretOffset
        i = off
        while i > 0 and doc.GetCharAt(i - 1) in _IDENT:
            i -= 1
        if i > 0 and doc.GetCharAt(i - 1) == ".":
            return 0  # handled by the '.' trigger
        return off - i

    def _smart_newline(self):
        """Indent after ':' / open bracket, dedent after return/pass/..."""
        doc = self.editor.Document
        caret = self.editor.CaretOffset
        line = doc.GetLineByOffset(caret)
        prev = line.PreviousLine
        if prev is None:
            return
        prev_text = doc.GetText(prev.Offset, prev.Length)
        want = pyparse.indent_for_next_line(prev_text, 4)
        cur_text = doc.GetText(line.Offset, line.Length)
        cur_ws = pyparse.leading_ws(cur_text)
        if cur_ws != want:
            caret_in_ws = caret - line.Offset <= len(cur_ws)
            doc.Replace(line.Offset, len(cur_ws), want)
            if caret_in_ws:
                self.editor.CaretOffset = line.Offset + len(want)

    def _maybe_dedent_block_keyword(self):
        """Typing ``else:`` / ``elif x:`` / ``except:`` / ``finally:`` dedents."""
        doc = self.editor.Document
        line = doc.GetLineByOffset(self.editor.CaretOffset)
        text = doc.GetText(line.Offset, line.Length)
        stripped = text.strip()
        word = stripped.split(" ")[0].split(":")[0]
        if word not in ("else", "elif", "except", "finally"):
            return
        ws = pyparse.leading_ws(text)
        # find the matching block opener above
        prev = line.PreviousLine
        while prev is not None:
            pt = doc.GetText(prev.Offset, prev.Length)
            if pt.strip():
                pws = pyparse.leading_ws(pt)
                head = pt.strip().split(" ")[0].split(":")[0]
                if len(pws) < len(ws) and head in ("if", "elif", "for", "while", "try", "except", "else", "with"):
                    if pws != ws:
                        doc.Replace(line.Offset, len(ws), pws)
                    return
                if len(pws) <= len(ws) and head in ("if", "elif", "for", "while", "try", "except"):
                    return
            prev = prev.PreviousLine

    def _on_preview_key_down(self, sender, e):
        mods = Keyboard.Modifiers
        if e.Key == Key.Space and mods == ModifierKeys.Control:
            e.Handled = True
            self.show_completion(auto=False)
        elif e.Key == Key.Space and mods == (ModifierKeys.Control | ModifierKeys.Shift):
            e.Handled = True
            self.show_signature_help(manual=True)
        elif e.Key == Key.Escape:
            if self.insight_window is not None:
                self.insight_window.Close()
                e.Handled = True
            self._close_hover()

    def _on_mouse_wheel(self, sender, e):
        if Keyboard.Modifiers == ModifierKeys.Control:
            e.Handled = True
            self.zoom(1 if e.Delta > 0 else -1)

    def zoom(self, step):
        size = float(self.settings["font_size"]) + step
        size = max(8.0, min(32.0, size))
        self.settings["font_size"] = size
        if self.on_zoom is not None:
            self.on_zoom(size)
        else:
            self.editor.FontSize = size

    on_zoom = None

    # -- completion ---------------------------------------------------------
    def show_completion(self, auto=False, bare_min=0):
        if self.engine is None:
            return
        if self.completion_window is not None:
            return
        try:
            res = self.engine.complete(self.editor.Text, self.editor.CaretOffset)
        except Exception as ex:
            if self.on_status:
                self.on_status("Completion error: {0}".format(ex))
            return
        if res is None:
            return
        start, partial, items = res
        if not items:
            return
        if auto and bare_min and len(partial) < bare_min:
            return
        low = partial.lower()
        if partial:
            items = [it for it in items if low in it.name.lower()]
            if not items:
                return
            if auto and len(items) == 1 and items[0].name == partial:
                return

        def sort_key(it):
            n = it.name
            starts = 0 if n.lower().startswith(low) else 1
            under = 1 if n.startswith("_") else 0
            return (starts, under, n.lower())

        items.sort(key=sort_key)
        cw = CompletionWindow(self.editor.TextArea)
        style_popup(cw, self.palette)
        cw.Width = 420
        cw.CloseWhenCaretAtBeginning = True
        try:
            cw.CompletionList.IsFiltering = True
        except Exception:
            pass
        lb = cw.CompletionList.ListBox
        lb.Background = self.palette.brush("PopupBackground")
        lb.Foreground = self.palette.brush("PopupForeground")
        lb.BorderThickness = Thickness(0)
        lb.FontFamily = _font_family(self.settings["font_family"])
        lb.FontSize = max(11.0, float(self.settings["font_size"]) - 1)
        data = cw.CompletionList.CompletionData
        for it in items[:3000]:
            data.Add(CompletionData(it, self))
        cw.StartOffset = start
        if partial:
            cw.CompletionList.SelectItem(partial)
        cw.Closed += self._on_completion_closed
        self.completion_window = cw
        cw.Show()

    def _on_completion_closed(self, sender, args):
        self.completion_window = None

    # -- signature help -----------------------------------------------------
    def show_signature_help(self, manual=False):
        if self.engine is None:
            return
        try:
            res = self.engine.signature_help(self.editor.Text, self.editor.CaretOffset)
        except Exception:
            res = None
        if res is None:
            if self.insight_window is not None:
                self.insight_window.Close()
            return
        sigs, arg_index, paren, doc = res
        content = self._signature_element(sigs, arg_index, doc)
        iw = self.insight_window
        if iw is None:
            iw = InsightWindow(self.editor.TextArea)
            style_popup(iw, self.palette)
            iw.StartOffset = paren + 1
            iw.Closed += self._on_insight_closed
            self.insight_window = iw
            iw.Content = content
            iw.Show()
        else:
            iw.Content = content

    def _on_insight_closed(self, sender, args):
        self.insight_window = None

    def _signature_element(self, sigs, arg_index, doc):
        pal = self.palette
        panel = StackPanel()
        panel.MaxWidth = 700
        mono = _font_family(self.settings["font_family"])
        for sig in sigs[:8]:
            tb = TextBlock()
            tb.FontFamily = mono
            tb.TextWrapping = TextWrapping.Wrap
            tb.Foreground = pal.brush("PopupForeground")
            self._fill_signature(tb, sig, arg_index)
            panel.Children.Add(tb)
        if len(sigs) > 8:
            more = TextBlock()
            more.Text = u"(+{0} more)".format(len(sigs) - 8)
            more.Foreground = pal.brush("SubtleForeground")
            panel.Children.Add(more)
        if doc:
            d = TextBlock()
            d.Text = doc
            d.TextWrapping = TextWrapping.Wrap
            d.Margin = Thickness(0, 6, 0, 0)
            d.Foreground = pal.brush("PopupForeground")
            panel.Children.Add(d)
        return panel

    def _fill_signature(self, tb, sig, arg_index):
        """Bold the active parameter of ``name(a, b, c) -> r``."""
        open_i = sig.find("(")
        close_i = sig.rfind(")")
        if open_i < 0 or close_i < open_i:
            tb.Text = sig
            return
        params = []
        depth = 0
        cur = []
        for ch in sig[open_i + 1:close_i]:
            if ch in "[(":
                depth += 1
            elif ch in "])":
                depth -= 1
            if ch == "," and depth == 0:
                params.append("".join(cur))
                cur = []
            else:
                cur.append(ch)
        if cur or params:
            params.append("".join(cur))
        # skip a leading 'self' in python-style stubs
        offset = 1 if params and params[0].strip().startswith("self") else 0
        tb.Inlines.Add(Run(sig[:open_i + 1]))
        for i, p in enumerate(params):
            r = Run(p.strip() if i == 0 else u" " + p.strip())
            if i - offset == arg_index:
                r.FontWeight = FontWeights.Bold
                r.Foreground = self.palette.brush("Accent")
            tb.Inlines.Add(r)
            if i < len(params) - 1:
                tb.Inlines.Add(Run(u","))
        tb.Inlines.Add(Run(sig[close_i:]))

    # -- hover ----------------------------------------------------------------
    def _on_mouse_hover(self, sender, e):
        if not self.settings["hover_tips"] or self.engine is None:
            return
        try:
            pos = self.editor.GetPositionFromPoint(e.GetPosition(self.editor))
            if pos is None:
                return
            offset = self.editor.Document.GetOffset(pos.Line, pos.Column)
            text = self.engine.hover(self.editor.Text, offset)
        except Exception:
            text = None
        if not text:
            return
        tip = ToolTip()
        style_popup(tip, self.palette)
        tip.Content = doc_element(text, self.palette, font_family=self.settings["font_family"])
        tip.PlacementTarget = self.editor
        tip.Placement = PlacementMode.Mouse
        tip.IsOpen = True
        self._hover_tip = tip
        e.Handled = True

    def _on_mouse_hover_stopped(self, sender, e):
        self._close_hover()

    def _close_hover(self):
        if self._hover_tip is not None:
            self._hover_tip.IsOpen = False
            self._hover_tip = None

    def close_popups(self):
        self._close_hover()
        if self.completion_window is not None:
            self.completion_window.Close()
        if self.insight_window is not None:
            self.insight_window.Close()

    # -- line commands (each one undo step) --------------------------------
    def selected_line_range(self):
        doc = self.editor.Document
        sel_start = self.editor.SelectionStart
        sel_end = sel_start + self.editor.SelectionLength
        first = doc.GetLineByOffset(sel_start)
        last = doc.GetLineByOffset(sel_end)
        if self.editor.SelectionLength > 0 and sel_end == last.Offset and last.LineNumber > first.LineNumber:
            last = last.PreviousLine
        return first, last

    def _transform_lines(self, fn, keep_selection=True):
        doc = self.editor.Document
        first, last = self.selected_line_range()
        start = first.Offset
        end = last.EndOffset
        old = doc.GetText(start, end - start)
        lines = old.split("\n")
        # keep \r on CRLF documents
        new = "\n".join(fn(lines))
        if new == old:
            return
        doc.BeginUpdate()
        try:
            doc.Replace(start, end - start, new)
        finally:
            doc.EndUpdate()
        if keep_selection:
            self.editor.Select(start, len(new))

    def toggle_comment(self):
        def fn(lines):
            code = [l for l in lines if l.strip()]
            all_commented = code and all(l.lstrip().startswith("#") for l in code)
            if all_commented:
                return self._uncomment(lines)
            return self._comment(lines)
        self._transform_lines(fn)

    def comment(self):
        self._transform_lines(self._comment)

    def uncomment(self):
        self._transform_lines(self._uncomment)

    @staticmethod
    def _comment(lines):
        code = [l for l in lines if l.strip()]
        if not code:
            return lines
        col = min(len(pyparse.leading_ws(l)) for l in code)
        out = []
        for l in lines:
            if not l.strip():
                out.append(l)
            else:
                out.append(l[:col] + "# " + l[col:])
        return out

    @staticmethod
    def _uncomment(lines):
        out = []
        for l in lines:
            ws = pyparse.leading_ws(l)
            body = l[len(ws):]
            if body.startswith("# "):
                out.append(ws + body[2:])
            elif body.startswith("#"):
                out.append(ws + body[1:])
            else:
                out.append(l)
        return out

    def indent(self):
        self._transform_lines(lambda ls: [("    " + l) if l.strip() else l for l in ls])

    def outdent(self):
        def fn(ls):
            out = []
            for l in ls:
                if l.startswith("\t"):
                    out.append(l[1:])
                else:
                    n = len(l) - len(l.lstrip(" "))
                    out.append(l[min(4, n):])
            return out
        self._transform_lines(fn)

    def duplicate(self):
        doc = self.editor.Document
        if self.editor.SelectionLength > 0 and "\n" not in self.editor.SelectedText:
            s = self.editor.SelectionStart
            t = self.editor.SelectedText
            doc.Insert(s + len(t), t)
            self.editor.Select(s + len(t), len(t))
            return
        first, last = self.selected_line_range()
        block = doc.GetText(first.Offset, last.EndOffset - first.Offset)
        doc.Insert(last.EndOffset, "\n" + block)
        self.editor.Select(last.EndOffset + 1, len(block))

    def delete_lines(self):
        doc = self.editor.Document
        if doc.TextLength == 0:
            return
        first, last = self.selected_line_range()
        start = first.Offset
        end = last.EndOffset + last.DelimiterLength
        if last.DelimiterLength == 0 and first.PreviousLine is not None:
            start = first.PreviousLine.EndOffset
        doc.Remove(start, end - start)

    def move_lines(self, direction):
        doc = self.editor.Document
        first, last = self.selected_line_range()
        if direction < 0 and first.PreviousLine is None:
            return
        if direction > 0 and last.NextLine is None:
            return
        block = doc.GetText(first.Offset, last.EndOffset - first.Offset)
        doc.BeginUpdate()
        try:
            if direction < 0:
                prev = first.PreviousLine
                prev_text = doc.GetText(prev.Offset, prev.Length)
                start = prev.Offset
                doc.Replace(start, last.EndOffset - start, block + "\n" + prev_text)
                self.editor.Select(start, len(block))
            else:
                nxt = last.NextLine
                nxt_text = doc.GetText(nxt.Offset, nxt.Length)
                start = first.Offset
                doc.Replace(start, nxt.EndOffset - start, nxt_text + "\n" + block)
                self.editor.Select(start + len(nxt_text) + 1, len(block))
        finally:
            doc.EndUpdate()

    # -- markers --------------------------------------------------------------
    def mark_error_line(self, line_number):
        doc = self.editor.Document
        if line_number < 1 or line_number > doc.LineCount:
            return
        line = doc.GetLineByNumber(line_number)
        self.error_marker.segments = [(line.Offset, line.Length)]
        self.editor.TextArea.TextView.InvalidateLayer(KnownLayer.Selection)
        self.goto_line(line_number, select=False)

    def set_find_matches(self, segments):
        self.find_marker.segments = segments or []
        self.editor.TextArea.TextView.InvalidateLayer(KnownLayer.Selection)

    def goto_line(self, line_number, select=False):
        doc = self.editor.Document
        line_number = max(1, min(line_number, doc.LineCount))
        line = doc.GetLineByNumber(line_number)
        self.editor.ScrollToLine(line_number)
        if select:
            self.editor.Select(line.Offset, line.Length)
        else:
            self.editor.CaretOffset = line.Offset
        self.editor.TextArea.Focus()


class OutputView(object):
    """Read-only AvalonEdit console with coloured spans."""

    MAX_CHARS = 400000
    TRIM_CHARS = 100000

    def __init__(self, editor, settings, palette):
        self.editor = editor
        self.settings = settings
        self.colorizer = OutputColorizer()
        editor.TextArea.TextView.LineTransformers.Add(self.colorizer)
        editor.WordWrap = True
        self.apply_theme(palette)

    def apply_theme(self, palette):
        self.palette = palette
        apply_editor_theme(self.editor, palette, self.settings, output=True)
        self.colorizer.brushes = {
            "err": palette.brush("OutputError"),
            "echo": palette.brush("OutputEcho"),
            "result": palette.brush("OutputResult"),
            "info": palette.brush("OutputInfo"),
            "prompt": palette.brush("Prompt"),
        }
        self.editor.TextArea.TextView.Redraw()

    def append(self, text, kind="out"):
        if not text:
            return
        doc = self.editor.Document
        start = doc.TextLength
        doc.Insert(start, text)
        self.colorizer.add(start, start + len(text), kind)
        if doc.TextLength > self.MAX_CHARS:
            cut = doc.GetLineByOffset(self.TRIM_CHARS).EndOffset
            doc.Remove(0, cut)
            self.colorizer.shift(cut)

    def scroll_to_end(self):
        self.editor.ScrollToEnd()

    def clear(self):
        self.editor.Document.Text = ""
        self.colorizer.reset()

    def line_text_at_caret(self):
        doc = self.editor.Document
        line = doc.GetLineByOffset(self.editor.CaretOffset)
        return doc.GetText(line.Offset, line.Length)
