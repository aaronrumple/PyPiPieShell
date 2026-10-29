# -*- coding: utf-8 -*-
"""Locate and load ICSharpCode.AvalonEdit.

Search order (first file that loads wins):

1. an AvalonEdit already loaded in Revit (pyRevit 7's Python Shell, Dynamo...)
2. ``lib/pypipieshell/bin/ICSharpCode.AvalonEdit.dll`` (your own copy)
3. folders of loaded pyRevit / IronPython assemblies (pyRevit 7 deploys
   AvalonEdit into every IronPython engine folder)
4. the whole pyRevit install ``bin`` tree (found from ``pyrevit.HOME_DIR``,
   ``sys.path`` and the default install folders)
5. Dynamo for Revit, which also ships AvalonEdit

Copies matching the running .NET (netfx for Revit <= 2024, netcore/net8 for
Revit 2025+) are tried first.
"""
import os
import sys

import clr
from System import AppDomain, Environment

ASSEMBLY_NAME = "ICSharpCode.AvalonEdit"
DLL_NAME = ASSEMBLY_NAME + ".dll"

_loaded = None
_searched = []


class AvalonEditMissing(Exception):
    pass


def _loaded_assembly():
    for asm in AppDomain.CurrentDomain.GetAssemblies():
        try:
            if asm.GetName().Name == ASSEMBLY_NAME:
                return asm
        except Exception:
            continue
    return None


def _is_netcore():
    try:
        return Environment.Version.Major >= 5
    except Exception:
        return False


def _assembly_dirs():
    """Folders of loaded assemblies that belong to pyRevit or IronPython."""
    dirs = []
    for asm in AppDomain.CurrentDomain.GetAssemblies():
        try:
            if asm.IsDynamic:
                continue
            loc = asm.Location
            if not loc:
                continue
            low = loc.lower()
            name = asm.GetName().Name.lower()
            if "pyrevit" in low or name.startswith(("pyrevit", "ironpython", "microsoft.dynamic",
                                                     "microsoft.scripting")):
                d = os.path.dirname(loc)
                if d not in dirs:
                    dirs.append(d)
        except Exception:
            continue
    return dirs


def _pyrevit_roots():
    """Candidate pyRevit install roots (folders that contain ``bin``)."""
    roots = []

    def add(path):
        if path and os.path.isdir(path) and path not in roots:
            roots.append(path)

    try:
        import pyrevit
        add(getattr(pyrevit, "HOME_DIR", None))
        add(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(pyrevit.__file__)))))
    except Exception:
        pass
    # walk up from sys.path entries and loaded assemblies until a folder with bin/
    for p in list(sys.path) + _assembly_dirs():
        try:
            cur = os.path.abspath(p)
        except Exception:
            continue
        for _ in range(6):
            if os.path.isdir(os.path.join(cur, "bin")) and "pyrevit" in cur.lower():
                add(cur)
                break
            parent = os.path.dirname(cur)
            if parent == cur:
                break
            cur = parent
    for env in ("APPDATA", "PROGRAMDATA", "LOCALAPPDATA", "ProgramFiles"):
        base = os.environ.get(env)
        if base:
            for name in ("pyRevit-Master", "pyRevit", "pyRevit-7"):
                add(os.path.join(base, name))
    return roots


def _walk_for_dll(root, max_depth=6):
    found = []
    if not os.path.isdir(root):
        return found
    base_depth = root.rstrip("\\/").count(os.sep)
    for dirpath, dirs, files in os.walk(root):
        if DLL_NAME in files:
            found.append(os.path.join(dirpath, DLL_NAME))
        depth = dirpath.count(os.sep) - base_depth
        if depth >= max_depth:
            del dirs[:]
        else:
            dirs[:] = [d for d in dirs if d.lower() not in ("site-packages", ".git", "extensions", "docs")]
    return found


def _dynamo_dirs():
    out = []
    try:
        from System.Diagnostics import Process
        revit_dir = os.path.dirname(Process.GetCurrentProcess().MainModule.FileName)
        out.append(os.path.join(revit_dir, "AddIns", "DynamoForRevit"))
        out.append(revit_dir)
    except Exception:
        pass
    return out


def _runtime_rank(path):
    """Sort key: prefer copies built for the running .NET."""
    low = path.lower()
    core_hint = any(t in low for t in ("netcore", "net8", "net6", "net5", "net7", "ipy342"))
    fx_hint = any(t in low for t in ("netfx", "net4", "ipy2712"))
    if _is_netcore():
        return 0 if core_hint else (2 if fx_hint else 1)
    return 0 if fx_hint else (2 if core_hint else 1)


def candidate_paths():
    here = os.path.dirname(os.path.abspath(__file__))
    ordered = [os.path.join(here, "bin", DLL_NAME)]
    near = [os.path.join(d, DLL_NAME) for d in _assembly_dirs()]
    install = []
    for root in _pyrevit_roots():
        install.extend(_walk_for_dll(os.path.join(root, "bin")))
    dynamo = []
    for d in _dynamo_dirs():
        dynamo.extend(_walk_for_dll(d, max_depth=1))
    near.sort(key=_runtime_rank)
    install.sort(key=_runtime_rank)
    ordered += near + install + dynamo
    seen = set()
    out = []
    for p in ordered:
        k = os.path.normcase(os.path.abspath(p))
        if k not in seen:
            seen.add(k)
            out.append(p)
    return out


def load():
    """Load AvalonEdit (once) and return the Assembly."""
    global _loaded
    if _loaded is not None:
        return _loaded
    asm = _loaded_assembly()
    if asm is not None:
        clr.AddReference(asm)
        _loaded = asm
        return asm
    del _searched[:]
    errors = []
    for path in candidate_paths():
        exists = os.path.isfile(path)
        _searched.append(("found  " if exists else "missing") + "  " + path)
        if not exists:
            continue
        try:
            clr.AddReferenceToFileAndPath(path)
            _loaded = _loaded_assembly()
            if _loaded is not None:
                return _loaded
        except Exception as ex:
            errors.append("{0}: {1}".format(path, ex))
    roots = _pyrevit_roots()
    raise AvalonEditMissing(
        "ICSharpCode.AvalonEdit.dll could not be loaded.\n\n"
        "Fix: copy ICSharpCode.AvalonEdit.dll (from pyRevit 7's engine folder, Dynamo, or the "
        "AvalonEdit NuGet package - net462 build for Revit 2024 and older, net6/8 build for 2025+) into:\n"
        "  {0}\n\n"
        "pyRevit roots checked:\n  {1}\n\n"
        "Paths tried:\n  {2}{3}".format(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin"),
            "\n  ".join(roots) or "(none found)",
            "\n  ".join(_searched) or "(no candidate files)",
            ("\n\nLoad errors:\n  " + "\n  ".join(errors)) if errors else ""))


def version():
    try:
        return str(load().GetName().Version)
    except Exception:
        return "?"


def diagnostics():
    """Text report of where AvalonEdit was looked for (for troubleshooting)."""
    lines = [".NET {0}".format(Environment.Version), "pyRevit roots:"]
    lines += ["  " + r for r in _pyrevit_roots()]
    lines.append("Assembly folders:")
    lines += ["  " + d for d in _assembly_dirs()]
    lines.append("Candidates:")
    lines += ["  {0}  {1}".format("found  " if os.path.isfile(p) else "missing", p) for p in candidate_paths()]
    return "\n".join(lines)
