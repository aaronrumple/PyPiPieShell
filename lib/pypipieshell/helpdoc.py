# -*- coding: utf-8 -*-
"""Open the PyPiPie Shell HTML help (no WPF imports, safe from any dialog)."""
import os

HELP_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", "help.html")


def open_help():
    """Open help.html in the default browser. Returns an error string or None."""
    if not os.path.isfile(HELP_FILE):
        return "Help file not found: " + HELP_FILE
    try:
        os.startfile(HELP_FILE)
        return None
    except Exception as ex:
        try:
            from System.Diagnostics import Process
            Process.Start(HELP_FILE)
            return None
        except Exception:
            return str(ex)
