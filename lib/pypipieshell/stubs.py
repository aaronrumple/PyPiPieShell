# -*- coding: utf-8 -*-
"""Revit API stub (.pyi) discovery, indexing and generation.

Where stubs come from
---------------------
pyRevit does not ship ready-made ``.pyi`` files; it ships McNeel's
``pyrevit-stubsbuilder.exe`` (PythonStubsBuilder), which generates them
from the Revit API assemblies.  PyPiPie Shell can run it for the running
Revit version ("Build Revit stubs" in the settings menu) and caches the
result in ``%APPDATA%\\pyRevit\\PyPiPieShell\\stubs\\Revit<year>``.

Any other stub folders (gtalarico ironpython-stubs, RevitAPIStubs,
pyRevit-with-VSCode stubs...) can be added in settings; common install
locations are detected automatically.

How they are used
-----------------
Completion lists themselves come from live .NET reflection (always exact for
the running Revit).  Stubs add what reflection lacks: Python-style
signatures and docstrings for the completion tooltip, signature help and
hover, plus module names for ``import`` completion of modules that are not
loaded yet.

Indexing is lazy: folders are scanned for ``.pyi`` files once (paths only)
on a background thread and each module file is parsed the first time it is
needed.
"""
import os
import threading

from pypipieshell import pyparse
from pypipieshell import settings as _settings

_MAX_STUB_BYTES = 8 * 1024 * 1024


def default_stub_dir(revit_year):
    return os.path.join(_settings.data_dir(), "stubs", "Revit{0}".format(revit_year or "Unknown"))


def _year_tokens(year):
    if not year:
        return []
    y = str(year)
    return [y, "RVT " + y[-2:], "RVT" + y[-2:], "RVT_" + y[-2:], "R" + y[-2:], y[-2:]]


def _looks_like_stub_root(path):
    try:
        names = os.listdir(path)
    except Exception:
        return False
    return "Autodesk" in names or "System" in names or any(n.endswith(".pyi") for n in names)


def detect_stub_folders(revit_year):
    """Best-effort discovery of stub folders already on this machine."""
    found = []
    appdata = os.environ.get("APPDATA") or ""
    home = os.path.expanduser("~")
    roots = [
        default_stub_dir(revit_year),
        os.path.join(appdata, "RevitAPIStubs"),
        os.path.join(appdata, "Revit-API-Stubs"),
        os.path.join(appdata, "RevitAPI stubs"),
        os.path.join(appdata, "pyRevit with vscode"),
        os.path.join(home, "ironpython-stubs"),
    ]
    tokens = _year_tokens(revit_year)
    for root in roots:
        if not os.path.isdir(root):
            continue
        if _looks_like_stub_root(root) and root not in found:
            found.append(root)
            continue
        try:
            subs = [os.path.join(root, d) for d in os.listdir(root)]
        except Exception:
            continue
        subs = [s for s in subs if os.path.isdir(s)]
        # per-version folders: pick the one matching this Revit
        matched = [s for s in subs if any(t.lower() in os.path.basename(s).lower() for t in tokens)]
        for s in matched + [s for s in subs if "stubs" in os.path.basename(s).lower()]:
            if _looks_like_stub_root(s) and s not in found:
                found.append(s)
            else:
                inner = os.path.join(s, "stubs.min")
                if os.path.isdir(inner) and inner not in found:
                    found.append(inner)
    return found


class StubIndex(object):
    """module name -> .pyi path, with lazily parsed modules."""

    def __init__(self, folders=None):
        self._folders = []
        self._paths = {}
        self._parsed = {}
        self._lock = threading.Lock()
        self._ready = threading.Event()
        self.set_folders(folders or [])

    # -- indexing ---------------------------------------------------------
    def set_folders(self, folders):
        self._folders = [f for f in folders if f and os.path.isdir(f)]
        self._ready.clear()
        t = threading.Thread(target=self._scan)
        t.daemon = True
        t.start()

    @property
    def folders(self):
        return list(self._folders)

    @property
    def is_ready(self):
        return self._ready.is_set()

    @property
    def module_count(self):
        return len(self._paths)

    def _scan(self):
        paths = {}
        try:
            for root in self._folders:
                for dirpath, dirs, files in os.walk(root):
                    rel = os.path.relpath(dirpath, root)
                    parts = [] if rel in (".", "") else rel.replace("\\", "/").split("/")
                    if any(p.startswith(".") for p in parts):
                        continue
                    for fn in files:
                        if not fn.endswith(".pyi"):
                            continue
                        stem = fn[:-4]
                        mod_parts = parts if stem == "__init__" else parts + [stem]
                        if not mod_parts:
                            continue
                        name = ".".join(mod_parts)
                        paths.setdefault(name, os.path.join(dirpath, fn))
        except Exception:
            pass
        with self._lock:
            self._paths = paths
            self._parsed = {}
        self._ready.set()

    # -- queries ----------------------------------------------------------
    def module(self, name):
        with self._lock:
            if name in self._parsed:
                return self._parsed[name]
            path = self._paths.get(name)
        if not path:
            return None
        data = None
        try:
            if os.path.getsize(path) <= _MAX_STUB_BYTES:
                data = pyparse.parse_pyi(_settings.read_text_file(path))
        except Exception:
            data = None
        with self._lock:
            self._parsed[name] = data
        return data

    def submodules(self, parent):
        prefix = (parent + ".") if parent else ""
        out = set()
        with self._lock:
            names = list(self._paths.keys())
        for n in names:
            if prefix and not n.startswith(prefix):
                continue
            rest = n[len(prefix):]
            if rest:
                out.add(rest.split(".")[0])
        return sorted(out)

    def module_names(self, module):
        """Top-level names defined by a stub module."""
        data = self.module(module)
        if not data:
            return []
        names = list(data["classes"].keys()) + list(data["functions"].keys()) + list(data["attributes"].keys())
        return sorted(set(names))

    def class_info(self, namespace, class_name):
        data = self.module(namespace)
        if not data:
            return None
        return data["classes"].get(class_name)

    def member_info(self, namespace, class_name, member):
        cls = self.class_info(namespace, class_name)
        if not cls:
            return None
        return cls["members"].get(member)


# ---------------------------------------------------------------------------
# generation with pyRevit's stubs builder
# ---------------------------------------------------------------------------
def find_stubs_builder():
    """Locate ``pyrevit-stubsbuilder.exe`` inside the pyRevit install."""
    roots = []
    try:
        import pyrevit
        home = getattr(pyrevit, "HOME_DIR", None)
        if home:
            roots.append(os.path.join(home, "bin"))
            roots.append(home)
    except Exception:
        pass
    for root in roots:
        if not os.path.isdir(root):
            continue
        base_depth = root.rstrip("\\/").count(os.sep)
        for dirpath, dirs, files in os.walk(root):
            for fn in files:
                if fn.lower() in ("pyrevit-stubsbuilder.exe", "pystubsbuilder.exe"):
                    return os.path.join(dirpath, fn)
            if dirpath.count(os.sep) - base_depth >= 3:
                del dirs[:]
            dirs[:] = [d for d in dirs if d.lower() not in ("site-packages", "extensions", ".git")]
    return None


def revit_install_dir():
    try:
        from System.Diagnostics import Process
        return os.path.dirname(Process.GetCurrentProcess().MainModule.FileName)
    except Exception:
        return None


def build_revit_stubs(revit_year, on_done):
    """Run the stubs builder asynchronously.

    ``on_done(ok, message, dest)`` is called from a worker thread; marshal to
    the UI dispatcher before touching WPF.
    """
    exe = find_stubs_builder()
    if not exe:
        on_done(False, "pyrevit-stubsbuilder.exe was not found in the pyRevit install.", None)
        return
    revit_dir = revit_install_dir()
    if not revit_dir:
        on_done(False, "Could not determine the Revit install folder.", None)
        return
    dest = default_stub_dir(revit_year)
    if not os.path.isdir(dest):
        try:
            os.makedirs(dest)
        except OSError:
            pass
    targets = [os.path.join(revit_dir, n) for n in ("RevitAPI.dll", "RevitAPIUI.dll")]
    targets = [t for t in targets if os.path.isfile(t)]
    if not targets:
        on_done(False, "RevitAPI.dll was not found in " + revit_dir, None)
        return

    def worker():
        try:
            from System.Diagnostics import Process, ProcessStartInfo
            args = ['--dest="{0}"'.format(dest), '--search="{0}"'.format(revit_dir)]
            args += ['"{0}"'.format(t) for t in targets]
            psi = ProcessStartInfo(exe, " ".join(args))
            psi.UseShellExecute = False
            psi.CreateNoWindow = True
            psi.RedirectStandardOutput = True
            psi.RedirectStandardError = True
            proc = Process.Start(psi)
            out = proc.StandardOutput.ReadToEnd()
            err = proc.StandardError.ReadToEnd()
            proc.WaitForExit()
            ok = proc.ExitCode == 0 and not ("error" in (out + err).lower() and "failed" in (out + err).lower())
            msg = (out + "\n" + err).strip()
            on_done(ok, msg or ("exit code {0}".format(proc.ExitCode)), dest)
        except Exception as ex:
            on_done(False, str(ex), dest)

    t = threading.Thread(target=worker)
    t.daemon = True
    t.start()
