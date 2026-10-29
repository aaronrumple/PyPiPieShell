# PyPiPie Shell

A Python editor and interactive REPL for pyRevit, built on AvalonEdit.

- Python syntax highlighting, folding, smart indent, find/replace (regex), go to line
- Revit API autocomplete from live .NET reflection, signature help, hover tips
- Optional Revit API stubs (`.pyi`) for docstrings; can build them with pyRevit's stubs builder
- REPL sharing one session with the editor; `doc`, `uidoc`, `selection`, `active_view` ready to use
- Transaction modes (Auto / Group / None), Esc to interrupt, error-line highlighting
- Opens as a **modal** window, a **modeless** window or a Revit **dockable pane**
- Follows Revit's light/dark theme (Revit 2024+) using pyRevit 7's palette

## Requirements

- pyRevit 7 or newer (IronPython engine). pyRevit 7 ships `ICSharpCode.AvalonEdit.dll`.
  With older pyRevit, put the DLL in `lib/pypipieshell/bin/` (see the README there).
- Revit 2021+. Dark theme on Revit 2024+.

## Install

1. Copy the `PyPiPieShell.extension` folder into a folder pyRevit loads extensions from
   (pyRevit Settings > Custom Extension Directories), or into
   `%APPDATA%\pyRevit\Extensions`.
2. Reload pyRevit, or restart Revit. Restart Revit for the docked pane: Revit only
   registers dockable panes while pyRevit loads.
3. The **PyPiPie Shell** button appears on the **PyPiPie** tab, **Shell** panel.

If the PyPiPie extension that used to contain the shell is also installed, remove its
`PyPiPie Shell.pushbutton`, `lib/pypipieshell`, `lib/pypipieshell_load.py` and the
shell registration in its `startup.py`, so only one copy loads.

## Use

- Click: open in the saved mode (modeless by default).
- Shift+Click: choose Modeless, Modal, Docked, or Help.
- F5 runs the editor, Ctrl+Enter runs the selection, the `>>>` line runs statements.
- F1 or the Help button opens the full help (`lib/pypipieshell/resources/help.html`).

Settings, REPL history, the autosaved buffer and built stubs are stored in
`%APPDATA%\pyRevit\PyPiPieShell`.

## Layout

```
PyPiPieShell.extension/
  startup.py                       registers the dockable pane
  PyPiPie.tab/Shell.panel/PyPiPie Shell.pushbutton/
      PyPiPie Shell_script.py      opens the shell
      config.py                    Shift+Click mode picker
      bundle.yaml, icon.png, help.html
  lib/pypipieshell/                the shell package
      bin/                         optional ICSharpCode.AvalonEdit.dll
      resources/                   XAML, Python XSHD (light/dark), help.html
```

Author: Aaron Rumple, AIA
