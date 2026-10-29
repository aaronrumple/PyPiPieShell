# -*- coding: utf-8 -*-
"""PyPiPie Shell - a pyRevit Python editor + REPL.

AvalonEdit editor with Python highlighting, Revit API autocomplete,
signature help and hover; persistent REPL; light/dark theme that follows
Revit; opens as a modal window, a modeless window or a dockable pane.

Public API::

    import pypipieshell
    pypipieshell.launch("Modeless", __revit__)   # or "Modal" / "Docked"
    pypipieshell.register_dockable_pane()        # from the extension startup.py
"""

__version__ = "2.0.0"

MODES = ("Modeless", "Modal", "Docked")


def launch(mode, uiapp):
    from pypipieshell import hosts
    return hosts.launch(mode, uiapp)


def register_dockable_pane():
    from pypipieshell import hosts
    return hosts.register_dockable_pane()


# --- compatibility with the 1.x loader (pypipieshell_load.py) ---------------
# 1.x gated the Run button on an Idling heartbeat.  2.x queues runs through an
# ExternalEvent instead, so the heartbeat is no longer needed; these remain so
# an existing Idling hook keeps working.
_LAST_IDLE = [None]


def set_revit_idle_heartbeat():
    import System
    _LAST_IDLE[0] = System.DateTime.UtcNow


def get_revit_idle_heartbeat():
    return _LAST_IDLE[0]
