# -*- coding: utf-8 -*-
"""Running code for PyPiPie Shell.

Session
    One persistent scope shared by the editor and the REPL, refreshed with
    the Revit context (``doc``, ``uidoc``, ``selection`` ...) before each run.

Executors
    * :class:`DirectExecutor` - modal window.  ``ShowDialog`` keeps the pyRevit
      command (and its API context) alive, so code runs immediately.
    * :class:`ExternalEventExecutor` - modeless window and dockable pane.
      Requests are queued and run by an ``ExternalEvent`` when Revit is
      idle.  A queued request can be cancelled before it starts.

Transactions (``transaction_mode`` setting)
    * ``auto``  - editor runs and REPL statements that do not manage their
      own transactions are wrapped in one ``Transaction`` so model edits
      just work; code that mentions Transaction/TransactionGroup/
      SubTransaction is wrapped in a ``TransactionGroup`` (editor) or left
      alone (REPL, so a transaction can span several statements).
    * ``group`` - always a ``TransactionGroup`` (one undo step, rolled back on
      error).  Code must start its own transactions.
    * ``none``  - no wrapping.

Interrupt
    With ``esc_interrupt`` on, a trace hook polls the Esc key every few
    hundred Python events and raises ``KeyboardInterrupt``.  Long single
    .NET calls cannot be interrupted until they return.
"""
import re
import sys
import time
import traceback

import clr
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import Transaction, TransactionGroup, TransactionStatus  # noqa: E402
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, ExternalEventRequest  # noqa: E402

EDITOR_FILENAME = "<editor>"
REPL_FILENAME = "<repl>"
TRANSACTION_NAME = "PyPiPie Shell"

_TXN_RE = re.compile(r"\b(Transaction|TransactionGroup|SubTransaction)\s*\(|\brevit\.Transaction\b|\brevit\.TransactionGroup\b|\bdb\.Transaction\b")


# ---------------------------------------------------------------------------
# output redirection
# ---------------------------------------------------------------------------
class StreamRedirector(object):
    """File-like object that forwards writes to ``sink(text, kind)``."""

    encoding = "utf-8"
    errors = "replace"

    def __init__(self, sink, kind):
        self._sink = sink
        self._kind = kind

    def write(self, text):
        if text:
            try:
                self._sink(text if isinstance(text, type(u"")) else str(text), self._kind)
            except Exception:
                pass

    def writelines(self, lines):
        for l in lines:
            self.write(l)

    def flush(self):
        pass

    def isatty(self):
        return False


# ---------------------------------------------------------------------------
# Esc interrupt
# ---------------------------------------------------------------------------
_VK_ESCAPE = 0x1B
_key_state = None


def _get_key_state_fn():
    global _key_state
    if _key_state is None:
        try:
            import ctypes
            _key_state = ctypes.windll.user32.GetAsyncKeyState
        except Exception:
            _key_state = False
    return _key_state or None


class EscInterrupter(object):
    """``sys.settrace`` hook that raises KeyboardInterrupt when Esc is held."""

    CHECK_EVERY = 400

    def __init__(self):
        self.fn = _get_key_state_fn()
        self.count = 0
        self.armed_at = 0.0

    @property
    def available(self):
        return self.fn is not None

    def _pressed(self):
        try:
            return (self.fn(_VK_ESCAPE) & 0x8000) != 0
        except Exception:
            return False

    def tracer(self, frame, event, arg):
        self.count += 1
        if self.count >= self.CHECK_EVERY:
            self.count = 0
            # ignore an Esc that was already down when the run started
            if time.time() - self.armed_at > 0.3 and self._pressed():
                raise KeyboardInterrupt("Interrupted with Esc")
        return self.tracer

    def install(self):
        if self.fn is None:
            return False
        self.count = 0
        self.armed_at = time.time()
        try:
            # clear a stale "was pressed since last call" bit
            self.fn(_VK_ESCAPE)
        except Exception:
            pass
        try:
            sys.settrace(self.tracer)
            return True
        except Exception:
            return False

    @staticmethod
    def uninstall():
        try:
            sys.settrace(None)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# session
# ---------------------------------------------------------------------------
class RunRequest(object):
    _next_id = [0]

    def __init__(self, code, filename, kind, transaction_mode, esc_interrupt, on_output, on_done,
                 line_offset=0):
        RunRequest._next_id[0] += 1
        self.id = RunRequest._next_id[0]
        self.code = code
        self.filename = filename
        self.kind = kind  # "editor" | "repl"
        self.transaction_mode = transaction_mode
        self.esc_interrupt = esc_interrupt
        self.on_output = on_output  # (text, kind)
        self.on_done = on_done  # (RunResult)
        self.line_offset = line_offset
        self.cancelled = False


class RunResult(object):
    def __init__(self, ok, elapsed, error=None, error_line=None, interrupted=False, wrapper=None):
        self.ok = ok
        self.elapsed = elapsed
        self.error = error
        self.error_line = error_line
        self.interrupted = interrupted
        self.wrapper = wrapper


class ShellSession(object):
    def __init__(self):
        self.scope = {}
        self.is_running = False
        self._init_scope()

    def _init_scope(self):
        self.scope.clear()
        self.scope["__name__"] = "__main__"
        self.scope["__doc__"] = None
        try:
            import __builtin__ as b
        except ImportError:
            import builtins as b
        self.scope["__builtins__"] = b
        self.scope["sys"] = sys
        self.scope["clr"] = clr
        self.scope["System"] = System
        try:
            import os
            self.scope["os"] = os
        except Exception:
            pass
        try:
            exec("from Autodesk.Revit.DB import *", self.scope, self.scope)
            exec("from Autodesk.Revit.UI import *", self.scope, self.scope)
        except Exception:
            pass
        try:
            from pyrevit import forms, script, revit, DB, UI
            self.scope.update({"forms": forms, "script": script, "revit": revit, "DB": DB, "UI": UI})
        except Exception:
            pass

    def reset(self):
        self._init_scope()

    def refresh_context(self, uiapp):
        s = self.scope
        s["__revit__"] = uiapp
        s["uiapp"] = uiapp
        s["app"] = uiapp.Application if uiapp is not None else None
        uidoc = None
        doc = None
        try:
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document if uidoc is not None else None
        except Exception:
            pass
        s["uidoc"] = uidoc
        s["doc"] = doc
        s["active_view"] = None
        s["selection"] = []
        if doc is not None:
            try:
                s["active_view"] = doc.ActiveView
                ids = uidoc.Selection.GetElementIds()
                s["selection"] = [doc.GetElement(i) for i in ids]
            except Exception:
                pass
        return doc

    # -- compile -------------------------------------------------------------
    @staticmethod
    def _compile(req):
        code = (req.code or "").replace("\r\n", "\n").replace("\r", "\n")
        if req.kind == "repl":
            try:
                return compile(code + "\n", req.filename, "single")
            except SyntaxError:
                return compile(code + "\n", req.filename, "exec")
        return compile(code + "\n", req.filename, "exec")

    @staticmethod
    def _wrapper_for(req, doc):
        if doc is None:
            return None
        mode = req.transaction_mode or "auto"
        if mode == "none":
            return None
        if mode == "group":
            return "group"
        manages_own = _TXN_RE.search(req.code or "") is not None
        if manages_own:
            return "group" if req.kind == "editor" else None
        try:
            if doc.IsReadOnly or doc.IsModifiable:
                # read-only docs can't be edited; an open transaction already exists
                return None
        except Exception:
            pass
        return "transaction"

    # -- run -----------------------------------------------------------------
    def execute(self, uiapp, req):
        """Run ``req`` now (must be in a valid Revit API context)."""
        out = req.on_output
        started = time.time()
        if req.cancelled:
            return RunResult(False, 0.0, error="cancelled")
        self.is_running = True
        doc = self.refresh_context(uiapp)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = StreamRedirector(out, "out")
        sys.stderr = StreamRedirector(out, "err")
        interrupter = EscInterrupter() if req.esc_interrupt else None
        wrapper = None
        tg = None
        tx = None
        result = None
        # IronPython only traces code compiled while a tracer is active, so the
        # Esc hook is installed before compiling.
        traced = interrupter.install() if interrupter is not None else False
        try:
            try:
                codeobj = self._compile(req)
            except SyntaxError as ex:
                line = getattr(ex, "lineno", None)
                out(u"SyntaxError: {0}\n".format(_syntax_message(ex)), "err")
                return RunResult(False, time.time() - started, error="SyntaxError", error_line=line)

            wrapper = self._wrapper_for(req, doc)
            if wrapper == "group":
                tg = TransactionGroup(doc, TRANSACTION_NAME)
                tg.Start()
            elif wrapper == "transaction":
                tx = Transaction(doc, TRANSACTION_NAME)
                tx.Start()

            exec(codeobj, self.scope, self.scope)
            if traced:
                EscInterrupter.uninstall()
                traced = False

            if tx is not None and tx.GetStatus() == TransactionStatus.Started:
                status = tx.Commit()
                if status != TransactionStatus.Committed:
                    out(u"Transaction was not committed ({0}).\n".format(status), "err")
            if tg is not None and tg.HasStarted():
                if _has_open_transaction(doc):
                    out(u"A transaction is still open at the end of the run; "
                        u"the run was rolled back.\n", "err")
                    tg.RollBack()
                else:
                    tg.Assimilate()
            result = RunResult(True, time.time() - started, wrapper=wrapper)
        except KeyboardInterrupt:
            _rollback(tx, tg)
            out(u"KeyboardInterrupt\n", "err")
            result = RunResult(False, time.time() - started, error="KeyboardInterrupt", interrupted=True,
                               wrapper=wrapper)
        except SystemExit:
            _rollback(tx, tg)
            out(u"SystemExit\n", "info")
            result = RunResult(False, time.time() - started, error="SystemExit", wrapper=wrapper)
        except Exception as ex:
            _rollback(tx, tg)
            tb_text = _format_exception()
            out(tb_text, "err")
            hint = _transaction_hint(ex, wrapper)
            if hint:
                out(hint + u"\n", "info")
            line = _error_line(tb_text, req.filename)
            result = RunResult(False, time.time() - started, error=type(ex).__name__, error_line=line,
                               wrapper=wrapper)
        finally:
            if traced:
                EscInterrupter.uninstall()
            sys.stdout, sys.stderr = old_out, old_err
            self.is_running = False
        return result


def _syntax_message(ex):
    msg = getattr(ex, "msg", None) or str(ex)
    line = getattr(ex, "lineno", None)
    text = (getattr(ex, "text", None) or "").rstrip()
    s = msg
    if line:
        s += u" (line {0})".format(line)
    if text:
        s += u"\n    " + text.strip()
    return s


def _has_open_transaction(doc):
    try:
        return doc.IsModifiable
    except Exception:
        return False


def _rollback(tx, tg):
    try:
        if tx is not None and tx.GetStatus() == TransactionStatus.Started:
            tx.RollBack()
    except Exception:
        pass
    try:
        if tg is not None and tg.HasStarted():
            # an inner transaction left open blocks the group rollback
            tg.RollBack()
    except Exception:
        pass


def _format_exception():
    try:
        etype, value, tb = sys.exc_info()
        lines = traceback.format_exception(etype, value, tb)
        # drop the shell's own frame(s) above the user code
        keep = []
        skipping = True
        for ln in lines:
            if skipping and ln.startswith("  File ") and ("<editor>" in ln or "<repl>" in ln):
                skipping = False
            if not skipping or not ln.startswith("  "):
                keep.append(ln)
        return u"".join(keep if not skipping else lines)
    except Exception:
        return traceback.format_exc()


def _error_line(tb_text, filename):
    from pypipieshell import pyparse
    return pyparse.last_traceback_line(tb_text, filename)


def _transaction_hint(ex, wrapper):
    msg = str(ex)
    low = msg.lower()
    if "outside of transaction" in low or "modification of the document is forbidden" in low:
        return (u"Hint: model changes need a transaction. Set Transaction mode to "
                u"'Auto' (wraps the run) or start a Transaction in your code.")
    if wrapper == "transaction" and ("transaction" in low and ("already" in low or "not permitted" in low
                                                               or "cannot start" in low)):
        return (u"Hint: this code (or a library it calls) starts its own transaction "
                u"inside the shell's auto transaction. Set Transaction mode to "
                u"'Group' for this script.")
    return None


# ---------------------------------------------------------------------------
# executors
# ---------------------------------------------------------------------------
class DirectExecutor(object):
    """Runs immediately - for the modal window (API context is live)."""

    is_queued = False

    def __init__(self, uiapp):
        self.uiapp = uiapp

    def submit(self, session, req):
        result = session.execute(self.uiapp, req)
        req.on_done(result)
        return "ran"

    def cancel_pending(self):
        return 0

    @property
    def pending_count(self):
        return 0

    def dispose(self):
        pass


class _RunHandler(IExternalEventHandler):
    def __init__(self, owner):
        self.owner = owner

    def Execute(self, uiapp):
        self.owner._drain(uiapp)

    def GetName(self):
        return "PyPiPie Shell Run"


class ExternalEventExecutor(object):
    """Queues runs for an ExternalEvent (modeless window / dockable pane).

    Must be constructed inside a valid Revit API context (a command or
    startup script) because ``ExternalEvent.Create`` requires one.
    """

    is_queued = True

    def __init__(self):
        self._queue = []
        self._session = None
        self._handler = _RunHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def submit(self, session, req):
        self._session = session
        self._queue.append(req)
        status = self._event.Raise()
        if status not in (ExternalEventRequest.Accepted, ExternalEventRequest.Pending):
            self._queue.remove(req)
            return "rejected:{0}".format(status)
        return "queued"

    def cancel_pending(self):
        n = 0
        for req in self._queue:
            req.cancelled = True
            n += 1
        return n

    @property
    def pending_count(self):
        return len([r for r in self._queue if not r.cancelled])

    def _drain(self, uiapp):
        queue, self._queue = self._queue, []
        for req in queue:
            if req.cancelled:
                try:
                    req.on_done(RunResult(False, 0.0, error="cancelled"))
                except Exception:
                    pass
                continue
            try:
                result = self._session.execute(uiapp, req)
            except Exception as ex:
                result = RunResult(False, 0.0, error=str(ex))
            try:
                req.on_done(result)
            except Exception:
                pass

    def dispose(self):
        self.cancel_pending()
        try:
            self._event.Dispose()
        except Exception:
            pass
