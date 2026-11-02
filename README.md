<p align="center"><img src="PyPiPie.tab/Shell.panel/PyPiPie Shell.pushbutton/pypipieshell.png" width="128" alt="PyPiPie Shell"></p>

# PyPiPie Shell

Crafted by: [Aaron Rumple, AIA](https://www.linkedin.com/in/aaronrumple/)

PyPiPie Shell is a [pyRevit](https://github.com/pyrevitlabs/pyRevit) extension for Autodesk Revit: a Python editor and interactive REPL connected to the running Revit session, built on AvalonEdit. Part of the PyPiPie tools.

- **Opens three ways:** a modal window, a modeless window, or a Revit dockable pane.
- **Follows Revit's theme:** dark or light, and switches when Revit does (Revit 2024+).
- **Knows the Revit API:** autocomplete, signature help and hover tips from the running Revit.

## Features

| Area | What it does |
|---|---|
| **Editor** | Python highlighting (light and dark), folding, smart indent, one-step undo for line commands (comment, indent, duplicate, move, delete), zoom, autosaved buffer. |
| **Autocomplete** | Members from live .NET reflection, so they always match your Revit. Chains follow return types (`FilteredElementCollector(doc).OfClass(Wall).FirstElement().`). Variables typed from `x = Type(...)`, `for x in ...`, `with ... as x`. `import` completion. Nothing runs to produce completions. |
| **Signatures and hover** | Every overload shown on `(`, current argument highlighted; hover for type, signatures and docs. |
| **Revit API stubs** | Optional `.pyi` stubs add docstrings. Build them with pyRevit's stubs builder from the settings menu, or add your own folder. |
| **REPL** | The `>>>` line shares one session with the editor: Enter runs complete statements, history persists between sessions. |
| **Running code** | F5 runs the editor, Ctrl+Enter the selection. `doc`, `uidoc`, `uiapp`, `selection`, `active_view`, `DB`, `UI`, `revit`, `forms` ready to use. Transaction modes Auto / Group / None; failed runs roll back and the error line is highlighted. Hold Esc to interrupt. |
| **Find and replace** | Match case, whole word, regex, live match highlighting, go to line. |

### Everywhere

- **Title bar** drawn by the app: icon, file name, **?** help (also **F1**), minimize, maximize/restore and close.
- **Status bar:** credit, status and hints, mode, transaction mode, caret, file and version; flashes on success or failure.
- **Help:** press **F1** or click **?** to open the full HTML guide (`lib/pypipieshell/resources/help.html`).

## Install

### pyRevit Extension Manager / CLI

```
pyrevit extend ui PyPiPieShell https://github.com/aaronrumple/PyPiPieShell.git
```

Then restart Revit (the docked pane can only register while pyRevit loads; a **pyRevit › Reload** is enough for the windows).

### Manual

1. Download or clone this repository.
2. Rename the folder to `PyPiPieShell.extension`.
3. Put it in `%APPDATA%\pyRevit\Extensions`, or add its parent folder under *pyRevit › Settings › Custom Extension Directories*.
4. Restart Revit.

The button is at **PyPiPie tab › Shell › PyPiPie Shell**. Click to open in the saved mode (modeless by default); **Shift+Click** to choose Modeless, Modal or Docked.

**Requirements:** pyRevit 7 or later with the IronPython engine (pyRevit 7 ships `ICSharpCode.AvalonEdit.dll`; with older pyRevit put the DLL in `lib/pypipieshell/bin/`). Revit 2021+; theme following needs Revit 2024 or later.

## Repository layout

```
extension.json                     pyRevit extension manifest
startup.py                         registers the dockable pane
PyPiPie.tab/Shell.panel/PyPiPie Shell.pushbutton/
  PyPiPie Shell_script.py          opens the shell in the saved mode
  config.py                        Shift+Click mode picker
  bundle.yaml, icon.png, icon.dark.png, pypipieshell.png
lib/pypipieshell/
  view.py                          UI controller: toolbar, editor, REPL, find, run, status bar
  hosts.py                         modal/modeless window (custom title bar) and dockable pane
  editors.py                       AvalonEdit setup, completion popup, markers, output console
  completion.py                    autocomplete, signature help and hover resolver
  stubs.py                         stub discovery, lazy index, stubs builder
  execution.py                     session, transactions, Esc interrupt, ExternalEvent queue
  theme.py                         Revit theme detection and palettes
  avalon.py                        finds and loads AvalonEdit
  pyparse.py                       pure-Python parsing helpers
  settings.py, helpdoc.py          settings/history/autosave, opening the help
  resources/                       XAML, Python XSHD (light/dark), app logo, help.html
  bin/                             optional ICSharpCode.AvalonEdit.dll
```

Settings, REPL history, the autosaved buffer and built stubs are stored in `%APPDATA%\pyRevit\PyPiPieShell`.

### Notes for contributors

- **Python files must be pure ASCII:** escape anything else as `\uXXXX`. IronPython's codec handling is unreliable otherwise.
- **Revit writes from the modeless window and pane go through `execution.ExternalEventExecutor`.**
- `pyparse.py` is pure Python and can be tested under CPython.

## License

GPL-3.0. See [LICENSE](LICENSE).

© 2026 [Aaron Rumple, AIA](https://www.linkedin.com/in/aaronrumple/)
