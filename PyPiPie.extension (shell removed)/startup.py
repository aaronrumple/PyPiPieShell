# -*- coding: UTF-8 -*-
"""
PyRevit startup dynamic updaters.

Production notes:
- Registers two optional Revit Dynamic Model Updaters:
    1. RoomsUpdater: maintains room net/gross area parameters and IBC helper parameters.
    2. FilledRegionUpdater: maintains filled region area/perimeter parameters.
    3. ViewportGridUpdater: numbers viewports from the titleblock drawing grid
       (lives in PyPiPie_ViewportGridLib; work is deferred to an ExternalEvent).
- Avoids modal dialogs inside IUpdater.Execute. Dynamic updaters can run during normal
  Revit transactions, so failures are logged and skipped instead of interrupting users.
- Uses defensive parameter checks so missing shared parameters do not break startup.

Author: Aaron Rumple, AIA / PyPiPie
"""

import clr
import os
import sys
import traceback

from pyrevit import script, forms

logger = script.get_logger()
EXT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(EXT_DIR, "lib")

if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

# PyPiPie Shell moved to its own extension (PyPiPieShell.extension), which
# registers its dockable pane in its own startup.py.
OPTIONAL_LOADERS = [
    "supersearch_load",
]


def load_optional_modules():
    """Load optional PyPiPie startup modules without blocking the remaining startup work."""
    for module_name in OPTIONAL_LOADERS:
        try:
            module = __import__(module_name, fromlist=["load"])
            module.load()
        except Exception as ex:
            toast("Startup loader {} failed: {}".format(module_name, ex))
            try:
                logger.error("Startup loader {} failed: {}".format(module_name, ex))
            except Exception:
                pass

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System import Guid

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ChangePriority,
    ChangeType,
    Element,
    ElementCategoryFilter,
    ElementClassFilter,
    ElementId,
    FilledRegion,
    GlobalParametersManager,
    IUpdater,
    SpatialElementBoundaryLocation,
    SpatialElementBoundaryOptions,
    SpatialElementGeometryCalculator,
    StorageType,
    SubTransaction,
    UpdaterId,
    UpdaterRegistry,
    UV,
    XYZ,
)


# Parameter loaders must be imported only AFTER the Revit API assemblies have
# been referenced.  Keep the actual import exception so startup/updater logging
# reports the real cause instead of only saying that the library was unavailable.
room_param_lib = None
room_param_lib_import_error = None
filled_region_param_lib = None
filled_region_param_lib_import_error = None
viewport_grid_lib = None
viewport_grid_lib_import_error = None

try:
    import PyPiPie_RoomLib as room_param_lib
except Exception as ex:
    room_param_lib_import_error = ex

try:
    import PyPiPie_FilledRegionLib as filled_region_param_lib
except Exception as ex:
    filled_region_param_lib_import_error = ex

try:
    import PyPiPie_ViewportGridLib as viewport_grid_lib
except Exception as ex:
    viewport_grid_lib_import_error = ex


# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

ROOMS_UPDATER_GUID = "f240012d-2358-409e-a856-fed322cff6da"
FILLED_REGION_UPDATER_GUID = "4fb7de21-925d-41f3-9ad9-42de62c329f6"

ROOM_AREA_GROSS_PARAM = "Room Area Gross"
ROOM_AREA_NET_PARAM = "Room Area Net"

GLOBAL_PARAM_USACE_PROJECT = "USACE Project"
PARAM_IBC_OCCUPANCY_CLASSIFICATION = "W - IBC - Occupancy Classification and Use"
PARAM_IBC_KEY_FUNCTION_BAD_SPELLING = "W IBC Key Funciton of Space"
PARAM_IBC_KEY_FUNCTION = "W IBC Key Function of Space"
PARAM_IBC_KEY_CLASS = "Key IBC Occupancy Classification"
PARAM_IBC_KEY_LOAD = "IBC Key Room Occupancy Factor"
PARAM_IBC_LOAD = "IBC OCCUPANCY LOAD FACTOR"
PARAM_IBC_INCLUDE = "IBC Space Included in Occupant Load"

FILLED_REGION_AREA_PARAM = "Filled Region Area"
FILLED_REGION_GROSS_AREA_PARAM = "Filled Region Gross Area"
FILLED_REGION_PERIMETER_PARAM = "Filled Region Perimeter"
FILLED_REGION_INNER_PERIMETER_PARAM = "Filled Region Inner Perimeter"
FILLED_REGION_OUTER_PERIMETER_PARAM = "Filled Region Outer Perimeter"

AREA_TOLERANCE = 0.000001
ANGLE_Z_TOLERANCE = 0.999999
LOG_PREFIX = "PyPiPie Startup"


# -----------------------------------------------------------------------------
# Revit application handles
# -----------------------------------------------------------------------------

uiapp = __revit__
app = uiapp.Application
addin_id = uiapp.ActiveAddInId


# -----------------------------------------------------------------------------
# Logging and safe helpers
# -----------------------------------------------------------------------------

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


def log(message):
    """Notify through Otto Pilot and retain the pyRevit logger as a fallback."""
    toast(message)
    try:
        logger.info("[{0}] {1}".format(LOG_PREFIX, message))
    except Exception:
        pass


def log_exception(context):
    try:
        log("{0}\n{1}".format(context, traceback.format_exc()))
    except Exception:
        pass


def is_valid_element_id(element_id):
    try:
        return element_id is not None and element_id != ElementId.InvalidElementId
    except Exception:
        return False


def safe_get_element(doc, element_id):
    try:
        if not is_valid_element_id(element_id):
            return None
        return doc.GetElement(element_id)
    except Exception:
        return None


def get_param(element, name):
    if element is None or not name:
        return None
    try:
        return element.LookupParameter(name)
    except Exception:
        return None


def get_first_param(element, names):
    for name in names:
        param = get_param(element, name)
        if param is not None:
            return param
    return None


def can_set_param(param):
    try:
        return param is not None and not param.IsReadOnly
    except Exception:
        return False


def set_double_param(param, value):
    if not can_set_param(param):
        return False
    try:
        if param.StorageType != StorageType.Double:
            return False
        current = param.AsDouble()
        if current is not None and abs(current - float(value)) <= AREA_TOLERANCE:
            return False
        return bool(param.Set(float(value)))
    except Exception:
        return False


def set_integer_param(param, value):
    if not can_set_param(param):
        return False
    try:
        if param.StorageType != StorageType.Integer:
            return False
        int_value = int(value)
        if param.AsInteger() == int_value:
            return False
        return bool(param.Set(int_value))
    except Exception:
        return False


def set_yes_no_param(param, value):
    return set_integer_param(param, 1 if value else 0)


def set_string_param(param, value):
    if not can_set_param(param):
        return False
    try:
        if param.StorageType != StorageType.String:
            return False
        string_value = "" if value is None else str(value)
        current = param.AsString()
        if current == string_value:
            return False
        return bool(param.Set(string_value))
    except Exception:
        return False


def get_param_display_value(param):
    if param is None:
        return ""
    try:
        value = param.AsValueString()
        if value:
            return value
    except Exception:
        pass
    try:
        value = param.AsString()
        if value:
            return value
    except Exception:
        pass
    return ""


def parameter_is_none_choice(param):
    value = get_param_display_value(param).strip().lower()
    return value in ("", "(none)", "<none>", "none")


def get_global_bool(doc, global_parameter_name):
    """Return (exists, value, element_id) for a Boolean-ish global parameter."""
    try:
        param_id = GlobalParametersManager.FindByName(doc, global_parameter_name)
        if not is_valid_element_id(param_id):
            return (False, False, None)

        global_param = doc.GetElement(param_id)
        if global_param is None:
            return (False, False, None)

        value = global_param.GetValue()
        raw_value = value.Value
        return (True, bool(raw_value), param_id)
    except Exception:
        return (False, False, None)


def associate_with_global_parameter(param, global_param_id):
    if not can_set_param(param) or not is_valid_element_id(global_param_id):
        return False
    try:
        associated = param.GetAssociatedGlobalParameter()
        if associated == global_param_id:
            return False
    except Exception:
        pass
    try:
        param.AssociateWithGlobalParameter(global_param_id)
        return True
    except Exception:
        return False


def dissociate_from_global_parameter(param):
    if not can_set_param(param):
        return False
    try:
        associated = param.GetAssociatedGlobalParameter()
        if not is_valid_element_id(associated):
            return False
    except Exception:
        pass
    try:
        param.DissociateFromGlobalParameter()
        return True
    except Exception:
        return False


def unique_existing_elements(doc, element_ids):
    elements = []
    seen = set()
    for element_id in element_ids:
        try:
            int_id = element_id.IntegerValue
        except Exception:
            continue
        if int_id in seen:
            continue
        seen.add(int_id)
        element = safe_get_element(doc, element_id)
        if element is not None:
            elements.append(element)
    return elements


def upward_horizontal_area(solid):
    """Return the sum of upward-facing horizontal face areas for a spatial solid."""
    area = 0.0
    if solid is None:
        return area

    try:
        faces = solid.Faces
    except Exception:
        return area

    for face in faces:
        try:
            normal = face.FaceNormal
        except Exception:
            try:
                bbox = face.GetBoundingBox()
                center_u = (bbox.Min.U + bbox.Max.U) / 2.0
                center_v = (bbox.Min.V + bbox.Max.V) / 2.0
                normal = face.ComputeNormal(UV(center_u, center_v))
            except Exception:
                continue

        try:
            if normal.Z >= ANGLE_Z_TOLERANCE:
                area += face.Area
        except Exception:
            continue

    return area


def safe_subtransaction(doc, work_func, context):
    """Run model changes inside the parent Revit transaction provided to the updater."""
    if doc is None or doc.IsFamilyDocument:
        return False

    transaction = SubTransaction(doc)
    try:
        transaction.Start()
        work_func()
        transaction.Commit()
        return True
    except Exception:
        log_exception(context)
        try:
            transaction.RollBack()
        except Exception:
            pass
        return False


# -----------------------------------------------------------------------------
# Updater parameter loading
# -----------------------------------------------------------------------------

def ensure_updater_parameters(doc, parameter_lib, context, import_error=None):
    """Ensure an updater's shared parameters exist before it writes values.

    Dynamic updater execution occurs inside Revit's parent transaction.  The lib
    modules are expected to detect doc.IsModifiable and avoid opening a top-level
    Transaction in that case.  Failure is logged and the updater skips the write
    rather than interrupting the user with a modal dialog.
    """
    if doc is None or doc.IsFamilyDocument:
        return False
    if parameter_lib is None:
        log("{0}: parameter library could not be imported: {1}".format(context, import_error if import_error is not None else "unknown import error"))
        return False

    try:
        if parameter_lib.parameters_loaded(doc):
            return True
    except Exception:
        log_exception("{0}: unable to check required parameters".format(context))
        return False

    try:
        parameter_lib.ensure_parameters(doc, app)
    except Exception:
        log_exception("{0}: unable to load required parameters".format(context))
        return False

    # Verify the binding actually became available before the updater continues.
    try:
        if parameter_lib.parameters_loaded(doc):
            log("{0}: required parameters loaded.".format(context))
            return True
        missing = parameter_lib.missing_parameters(doc)
        log("{0}: required parameters are still missing: {1}".format(
            context, ", ".join(missing) if missing else "unknown"))
    except Exception:
        log_exception("{0}: unable to verify loaded parameters".format(context))
    return False


# -----------------------------------------------------------------------------
# Rooms updater
# -----------------------------------------------------------------------------

class RoomsUpdater(IUpdater):
    def __init__(self, active_addin_id):
        self.updater_id = UpdaterId(active_addin_id, Guid(ROOMS_UPDATER_GUID))

    def GetUpdaterId(self):
        return self.updater_id

    def GetUpdaterName(self):
        return "PyPiPie Rooms Updater"

    def GetAdditionalInformation(self):
        return "Maintains room net/gross area and IBC helper parameters."

    def GetChangePriority(self):
        return ChangePriority.RoomsSpacesZones

    def Execute(self, updater_data):
        doc = updater_data.GetDocument()
        if doc is None or doc.IsFamilyDocument:
            return

        if not ensure_updater_parameters(doc, room_param_lib, "RoomsUpdater", room_param_lib_import_error):
            return

        element_ids = []
        try:
            element_ids.extend(list(updater_data.GetModifiedElementIds()))
            element_ids.extend(list(updater_data.GetAddedElementIds()))
        except Exception:
            log_exception("RoomsUpdater: unable to read changed element ids")
            return

        rooms = unique_existing_elements(doc, element_ids)
        if not rooms:
            return

        def update_rooms():
            self.update_room_net_gross(doc, rooms)
            self.update_ibc(doc, rooms)

        safe_subtransaction(doc, update_rooms, "RoomsUpdater failed")

    def update_room_net_gross(self, doc, rooms):
        net_options = SpatialElementBoundaryOptions()
        net_options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

        gross_options = SpatialElementBoundaryOptions()
        gross_options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center

        net_calculator = SpatialElementGeometryCalculator(doc, net_options)
        gross_calculator = SpatialElementGeometryCalculator(doc, gross_options)

        for room in rooms:
            try:
                gross_param = get_param(room, ROOM_AREA_GROSS_PARAM)
                net_param = get_param(room, ROOM_AREA_NET_PARAM)

                if gross_param is None and net_param is None:
                    continue

                try:
                    if hasattr(room, "Area") and room.Area <= AREA_TOLERANCE:
                        set_double_param(gross_param, 0.0)
                        set_double_param(net_param, 0.0)
                        continue
                except Exception:
                    pass

                net_geom = net_calculator.CalculateSpatialElementGeometry(room)
                gross_geom = gross_calculator.CalculateSpatialElementGeometry(room)

                net_area = upward_horizontal_area(net_geom.GetGeometry())
                gross_area = upward_horizontal_area(gross_geom.GetGeometry())

                if net_area > AREA_TOLERANCE:
                    set_double_param(net_param, net_area)
                if gross_area > AREA_TOLERANCE:
                    set_double_param(gross_param, gross_area)
            except Exception:
                set_double_param(get_param(room, ROOM_AREA_GROSS_PARAM), 0.0)
                set_double_param(get_param(room, ROOM_AREA_NET_PARAM), 0.0)

    def update_ibc(self, doc, rooms):
        gp_exists, is_usace, gp_id = get_global_bool(doc, GLOBAL_PARAM_USACE_PROJECT)
        if not gp_exists:
            return

        for room in rooms:
            try:
                occupancy_classification = get_param(room, PARAM_IBC_OCCUPANCY_CLASSIFICATION)
                key_function = get_first_param(room, [
                    PARAM_IBC_KEY_FUNCTION_BAD_SPELLING,
                    PARAM_IBC_KEY_FUNCTION,
                ])
                key_class = get_param(room, PARAM_IBC_KEY_CLASS)
                key_load = get_param(room, PARAM_IBC_KEY_LOAD)
                ibc_load = get_param(room, PARAM_IBC_LOAD)
                include_in_load = get_param(room, PARAM_IBC_INCLUDE)

                if include_in_load is None:
                    continue

                occupancy_missing = parameter_is_none_choice(occupancy_classification)
                function_missing = parameter_is_none_choice(key_function)

                if is_usace:
                    associate_with_global_parameter(include_in_load, gp_id)
                else:
                    dissociate_from_global_parameter(include_in_load)
                    set_yes_no_param(include_in_load, not (occupancy_missing or function_missing))

                if occupancy_missing:
                    set_string_param(key_class, "")

                if function_missing:
                    set_double_param(key_load, 0.0)
                    set_double_param(ibc_load, 0.0)
            except Exception:
                log_exception("RoomsUpdater: IBC update skipped for one room")


# -----------------------------------------------------------------------------
# Filled region updater
# -----------------------------------------------------------------------------

class FilledRegionUpdater(IUpdater):
    def __init__(self, active_addin_id):
        self.updater_id = UpdaterId(active_addin_id, Guid(FILLED_REGION_UPDATER_GUID))

    def GetUpdaterId(self):
        return self.updater_id

    def GetUpdaterName(self):
        return "PyPiPie Filled Region Updater"

    def GetAdditionalInformation(self):
        return "Maintains filled region net area, gross area, and perimeter parameters."

    def GetChangePriority(self):
        return ChangePriority.DetailComponents

    def Execute(self, updater_data):
        doc = updater_data.GetDocument()
        if doc is None or doc.IsFamilyDocument:
            return

        if not ensure_updater_parameters(doc, filled_region_param_lib, "FilledRegionUpdater", filled_region_param_lib_import_error):
            return

        element_ids = []
        try:
            element_ids.extend(list(updater_data.GetModifiedElementIds()))
            element_ids.extend(list(updater_data.GetAddedElementIds()))
        except Exception:
            log_exception("FilledRegionUpdater: unable to read changed element ids")
            return

        filled_regions = unique_existing_elements(doc, element_ids)
        if not filled_regions:
            return

        def update_regions():
            self.update_filled_region_data(filled_regions)

        safe_subtransaction(doc, update_regions, "FilledRegionUpdater failed")

    def update_filled_region_data(self, elements):
        for element in elements:
            try:
                self.update_one_filled_region(element)
            except Exception:
                log_exception("FilledRegionUpdater: update skipped for one region")

    def update_one_filled_region(self, element):
        if element is None:
            return

        area_param = None
        try:
            area_param = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        except Exception:
            pass

        if area_param is None:
            return

        try:
            net_area = area_param.AsDouble()
        except Exception:
            return

        if net_area <= AREA_TOLERANCE:
            return

        boundaries = element.GetBoundaries()
        if boundaries is None or boundaries.Count == 0:
            return

        gross_area = self.get_outer_loop_area(boundaries)
        outer_perimeter, inner_perimeter, total_perimeter = self.get_perimeter_values(boundaries)

        set_double_param(get_param(element, FILLED_REGION_AREA_PARAM), net_area)
        set_double_param(get_param(element, FILLED_REGION_GROSS_AREA_PARAM), gross_area)
        set_double_param(get_param(element, FILLED_REGION_OUTER_PERIMETER_PARAM), outer_perimeter)
        set_double_param(get_param(element, FILLED_REGION_INNER_PERIMETER_PARAM), inner_perimeter)
        set_double_param(get_param(element, FILLED_REGION_PERIMETER_PARAM), total_perimeter)

    def _loop_points(self, loop):
        """Return an ordered XY point list from a CurveLoop using native tessellation."""
        points = []
        try:
            for curve in loop:
                try:
                    tess = list(curve.Tessellate())
                except Exception:
                    tess = []
                for point in tess:
                    if not points or points[-1].DistanceTo(point) > 1.0e-9:
                        points.append(point)
        except Exception:
            return []

        if len(points) > 1 and points[0].DistanceTo(points[-1]) <= 1.0e-9:
            points.pop()
        return points

    def _polygon_area(self, points):
        """Return absolute projected XY area in Revit internal square feet."""
        if len(points) < 3:
            return 0.0
        twice_area = 0.0
        for index in range(len(points)):
            p1 = points[index]
            p2 = points[(index + 1) % len(points)]
            twice_area += (p1.X * p2.Y) - (p2.X * p1.Y)
        return abs(twice_area) * 0.5

    def _point_on_segment_xy(self, point, a, b, tolerance=1.0e-8):
        cross = ((point.X - a.X) * (b.Y - a.Y) -
                 (point.Y - a.Y) * (b.X - a.X))
        if abs(cross) > tolerance:
            return False
        dot = ((point.X - a.X) * (b.X - a.X) +
               (point.Y - a.Y) * (b.Y - a.Y))
        if dot < -tolerance:
            return False
        length_sq = ((b.X - a.X) ** 2 + (b.Y - a.Y) ** 2)
        return dot <= length_sq + tolerance

    def _point_in_polygon_xy(self, point, polygon):
        """Ray-cast containment test. Points on the boundary count as inside."""
        if len(polygon) < 3:
            return False
        inside = False
        j = len(polygon) - 1
        for i in range(len(polygon)):
            pi = polygon[i]
            pj = polygon[j]
            if self._point_on_segment_xy(point, pj, pi):
                return True
            if ((pi.Y > point.Y) != (pj.Y > point.Y)):
                x_hit = ((pj.X - pi.X) * (point.Y - pi.Y) /
                         (pj.Y - pi.Y) + pi.X)
                if point.X < x_hit:
                    inside = not inside
            j = i
        return inside

    def _classify_boundary_loops(self, boundaries):
        """Classify every disconnected/nested boundary by geometric nesting depth.

        Even depth (0, 2, ...) is material/outer boundary. Odd depth (1, 3, ...)
        is a hole boundary. This supports multiple disconnected outer islands.
        """
        data = []
        try:
            loops = list(boundaries)
        except Exception:
            loops = []

        for loop in loops:
            points = self._loop_points(loop)
            if len(points) < 3:
                continue
            perimeter = 0.0
            try:
                for curve in loop:
                    perimeter += curve.Length
            except Exception:
                perimeter = 0.0
            data.append({
                "loop": loop,
                "points": points,
                "sample": points[0],
                "area": self._polygon_area(points),
                "perimeter": perimeter,
                "depth": 0,
                "is_outer": True,
            })

        # A loop's nesting depth is the number of larger loops containing it.
        # Area prefilter prevents a coincident/sample-on-edge case from allowing
        # a smaller loop to be treated as the container of a larger loop.
        for item in data:
            depth = 0
            for candidate in data:
                if candidate is item:
                    continue
                if candidate["area"] <= item["area"] + AREA_TOLERANCE:
                    continue
                if self._point_in_polygon_xy(item["sample"], candidate["points"]):
                    depth += 1
            item["depth"] = depth
            item["is_outer"] = (depth % 2 == 0)
        return data

    def get_outer_loop_area(self, boundaries):
        classified = self._classify_boundary_loops(boundaries)
        return sum(item["area"] for item in classified if item["is_outer"])

    def get_perimeter_values(self, boundaries):
        classified = self._classify_boundary_loops(boundaries)
        if not classified:
            return (0.0, 0.0, 0.0)

        outer = sum(item["perimeter"] for item in classified if item["is_outer"])
        inner = sum(item["perimeter"] for item in classified if not item["is_outer"])
        total = outer + inner
        return (outer, inner, total)


# -----------------------------------------------------------------------------
# Registration
# -----------------------------------------------------------------------------

def register_updater(updater_class, element_filter, change_type):
    updater = updater_class(addin_id)
    updater_id = updater.GetUpdaterId()

    try:
        if UpdaterRegistry.IsUpdaterRegistered(updater_id):
            UpdaterRegistry.UnregisterUpdater(updater_id)
    except Exception:
        log_exception("Could not unregister existing updater: {0}".format(updater.GetUpdaterName()))

    try:
        UpdaterRegistry.RegisterUpdater(updater, True)
        UpdaterRegistry.AddTrigger(updater_id, element_filter, change_type)
        log("Registered {0}".format(updater.GetUpdaterName()))
        return True
    except Exception:
        log_exception("Could not register updater: {0}".format(updater.GetUpdaterName()))
        return False


def register_all_updaters():
    room_filter = ElementCategoryFilter(BuiltInCategory.OST_Rooms)
    room_change_type = ChangeType.ConcatenateChangeTypes(
        Element.GetChangeTypeElementAddition(),
        Element.GetChangeTypeAny()
    )
    register_updater(RoomsUpdater, room_filter, room_change_type)

    filled_region_filter = ElementClassFilter(FilledRegion)
    filled_region_change_type = ChangeType.ConcatenateChangeTypes(
        Element.GetChangeTypeElementAddition(),
        Element.GetChangeTypeGeometry()
    )
    register_updater(FilledRegionUpdater, filled_region_filter, filled_region_change_type)

    # Viewport grid numbering registers its own triggers (viewports + titleblocks).
    if viewport_grid_lib is None:
        log("ViewportGridUpdater: library could not be imported: {0}".format(
            viewport_grid_lib_import_error))
    else:
        try:
            updater = viewport_grid_lib.register(addin_id, log=log, log_exception=log_exception)
            log("Registered {0}".format(updater.GetUpdaterName()))
        except Exception:
            log_exception("Could not register updater: PyPiPie Viewport Grid Updater")


# -----------------------------------------------------------------------------
# Startup
# -----------------------------------------------------------------------------

load_optional_modules()

try:
    register_all_updaters()
except Exception:
    log_exception("Startup updater registration failed")
