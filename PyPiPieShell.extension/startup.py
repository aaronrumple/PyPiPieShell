# -*- coding: utf-8 -*-
"""PyPiPie Shell extension startup.

Registers the PyPiPie Shell dockable pane.  Revit only accepts dockable
pane registration while pyRevit loads, so this has to run at startup.
Registration is quiet when it works; on failure an Otto Pilot toast says so
and the traceback goes to the pyRevit log.  The modal and modeless windows
work even if the pane is not registered.

Author: Aaron Rumple, AIA / PyPiPie
"""
import os
import sys
import traceback

from pyrevit import script, forms

logger = script.get_logger()
LOG_PREFIX = "PyPiPie Shell"

LIB_DIR = os.path.join(os.path.dirname(__file__), "lib")
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)


def toast(message):
    """Show a non-modal Otto Pilot notification."""
    try:
        forms.toast(
            "Otto Pilot says... {0}".format(message),
            title="Otto Pilot",
            appid="PyPiPie"
        )
        return True
    except Exception:
        return False


def register_pypipie_shell():
    try:
        import pypipieshell
        pypipieshell.register_dockable_pane()
        logger.debug("[{0}] {1} pane registered".format(
            LOG_PREFIX, getattr(pypipieshell, "__version__", "")))
        return True
    except Exception as ex:
        toast("PyPiPie Shell docked pane is unavailable ({0}). "
              "Modal and modeless windows still work.".format(ex))
        try:
            logger.error("[{0}] pane registration failed\n{1}".format(
                LOG_PREFIX, traceback.format_exc()))
        except Exception:
            pass
        return False


register_pypipie_shell()
