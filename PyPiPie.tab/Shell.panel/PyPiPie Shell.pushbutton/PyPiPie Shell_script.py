# -*- coding: utf-8 -*-
"""Open PyPiPie Shell in the saved mode (Shift+Click to choose)."""
# The modeless window and the dockable pane raise ExternalEvents after this
# script returns, so keep this engine alive.
__persistentengine__ = True

import pypipieshell

pypipieshell.launch(None, __revit__)
