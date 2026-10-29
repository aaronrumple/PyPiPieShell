# -*- coding: utf-8 -*-
"""Shift+Click: choose how PyPiPie Shell opens, then open it."""
__persistentengine__ = True

from pyrevit import forms

import pypipieshell
from pypipieshell import settings as shell_settings

s = shell_settings.Settings()
current = s["launch_mode"] or "Modeless"
choice = forms.CommandSwitchWindow.show(
    ["Modeless", "Modal", "Docked", "Help"],
    message="Open PyPiPie Shell as (current: {0})".format(current),
)
if choice == "Help":
    from pypipieshell import helpdoc
    helpdoc.open_help()
elif choice:
    s["launch_mode"] = choice
    s.save()
    pypipieshell.launch(choice, __revit__)
