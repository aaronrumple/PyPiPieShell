# -*- coding: utf-8 -*-
"""Autocomplete, signature help and hover for PyPiPie Shell.

Resolution never executes user code.  An expression in front of the caret is
resolved by:

1. live values already in the shell scope (``doc``, ``uidoc``, variables
   created by earlier runs, imported modules / CLR namespaces);
2. static inference from the source being edited (``x = Foo(...)``,
   ``for x in expr:``, ``with expr as x:``, ``import a.b as x``);
3. .NET reflection over the resolved type: property, field and method
   return types carry the chain forward (``FilteredElementCollector(doc)
   .OfClass(Wall).FirstElement().`` resolves to ``Element``), indexers and
   ``IEnumerable<T>`` give element types for ``x[0]`` / ``for x in``;
4. the Revit API stub index for docstrings and Python-style signatures.

Instance properties are resolved through their declared type, so no Revit
API getter runs on the UI thread outside API context.
"""
import sys

import clr
from System.Reflection import BindingFlags, MemberTypes

from pypipieshell import pyparse

try:
    import __builtin__ as _builtins  # IronPython 2.7
except ImportError:  # IronPython 3.4
    import builtins as _builtins

BF_INSTANCE = BindingFlags.Public | BindingFlags.Instance
BF_STATIC = BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy

CLR_ROOTS = ["System", "Autodesk", "Microsoft", "clr", "pyrevit", "rpw"]

_PY_ALIASES = {
    "System.String": "str", "System.Int32": "int", "System.Int64": "long",
    "System.Double": "float", "System.Single": "float", "System.Boolean": "bool",
    "System.Void": "None", "System.Object": "object", "System.Byte": "byte",
    "System.Int16": "short", "System.UInt32": "uint", "System.Char": "char",
}

# kinds -> (glyph letter, palette key, sort priority)
KINDS = {
    "keyword": ("k", "KindKeyword", 0),
    "variable": ("v", "KindVariable", 3),
    "builtin": ("b", "KindKeyword", 1),
    "module": ("m", "KindModule", 2),
    "namespace": ("n", "KindNamespace", 2),
    "class": ("C", "KindClass", 2),
    "enum": ("E", "KindEnum", 2),
    "enummember": ("e", "KindEnum", 3),
    "method": ("M", "KindMethod", 3),
    "function": ("f", "KindMethod", 3),
    "property": ("P", "KindProperty", 4),
    "field": ("F", "KindField", 4),
    "event": ("!", "KindEvent", 1),
    "word": ("w", "KindSnippet", 0),
}


# ---------------------------------------------------------------------------
# .NET reflection helpers
# ---------------------------------------------------------------------------
def type_name(t, full=False):
    if t is None:
        return "?"
    try:
        if t.IsByRef:
            return type_name(t.GetElementType(), full)
        if t.IsArray:
            return type_name(t.GetElementType(), full) + "[]"
        fn = t.FullName or t.Name
        if fn in _PY_ALIASES:
            return _PY_ALIASES[fn]
        if t.IsGenericType:
            base = t.Name.split("`")[0]
            args = ", ".join(type_name(a, full) for a in t.GetGenericArguments())
            return "{0}[{1}]".format(base, args)
        return fn if full else t.Name
    except Exception:
        return str(t)


def _param_sig(p):
    prefix = ""
    try:
        if p.IsOut:
            prefix = "out "
        elif p.ParameterType.IsByRef:
            prefix = "ref "
    except Exception:
        pass
    default = ""
    try:
        if p.IsOptional:
            default = "=" + repr(p.DefaultValue) if p.DefaultValue is not None else "=None"
    except Exception:
        pass
    return "{0}{1}: {2}{3}".format(prefix, p.Name, type_name(p.ParameterType), default)


def method_signature(m, include_static=True):
    try:
        params = ", ".join(_param_sig(p) for p in m.GetParameters())
    except Exception:
        params = "..."
    if m.MemberType == MemberTypes.Constructor:
        return "{0}({1})".format(m.DeclaringType.Name, params)
    ret = type_name(m.ReturnType)
    static = "static " if (include_static and m.IsStatic) else ""
    return "{0}{1}({2}) -> {3}".format(static, m.Name, params, ret)


def _is_special(m):
    try:
        if m.IsSpecialName:
            return True
    except Exception:
        pass
    name = m.Name
    return name.startswith(("get_", "set_", "add_", "remove_", "op_")) or name in (".ctor", ".cctor")


def all_interfaces(t):
    seen = []
    try:
        for i in t.GetInterfaces():
            seen.append(i)
    except Exception:
        pass
    return seen


_MEMBER_CACHE = {}


def clr_members(t, static):
    """Group public members of ``t`` by name: ``{name: [MemberInfo, ...]}``.

    Cached per (type, static) for the session.
    """
    try:
        key = (t.AssemblyQualifiedName, bool(static))
    except Exception:
        key = None
    if key is not None and key in _MEMBER_CACHE:
        return _MEMBER_CACHE[key]
    groups = _clr_members_uncached(t, static)
    if key is not None:
        if len(_MEMBER_CACHE) > 2000:
            _MEMBER_CACHE.clear()
        _MEMBER_CACHE[key] = groups
    return groups


def clear_value_cache():
    """Forget cached members of Python modules (they may change between runs)."""
    for k in list(_MEMBER_CACHE.keys()):
        if k and k[0] == "py":
            _MEMBER_CACHE.pop(k, None)


def _clr_members_uncached(t, static):
    groups = {}
    flags = BF_STATIC if static else BF_INSTANCE
    types = [t]
    try:
        if t.IsInterface and not static:
            types += all_interfaces(t)
    except Exception:
        pass
    for tt in types:
        try:
            members = tt.GetMembers(flags)
        except Exception:
            continue
        for m in members:
            try:
                mt = m.MemberType
                if mt == MemberTypes.Constructor:
                    continue
                if mt == MemberTypes.Method and _is_special(m):
                    continue
                if mt == MemberTypes.NestedType and not static:
                    continue
                groups.setdefault(m.Name, []).append(m)
            except Exception:
                continue
    if static:
        try:
            for nt in t.GetNestedTypes(BindingFlags.Public):
                groups.setdefault(nt.Name, []).append(nt)
        except Exception:
            pass
    return groups


def member_kind(members, owner=None):
    m = members[0]
    mt = m.MemberType
    if mt == MemberTypes.Method:
        return "method"
    if mt == MemberTypes.Property:
        return "property"
    if mt == MemberTypes.Event:
        return "event"
    if mt == MemberTypes.NestedType or mt == MemberTypes.TypeInfo:
        try:
            return "enum" if m.IsEnum else "class"
        except Exception:
            return "class"
    if mt == MemberTypes.Field:
        try:
            if owner is not None and owner.IsEnum:
                return "enummember"
        except Exception:
            pass
        return "field"
    return "field"


def member_result_type(members):
    """Declared type a member access / call produces."""
    for m in members:
        try:
            mt = m.MemberType
            if mt == MemberTypes.Property:
                return m.PropertyType
            if mt == MemberTypes.Field:
                return m.FieldType
            if mt == MemberTypes.Method:
                if m.ReturnType.FullName != "System.Void":
                    return m.ReturnType
        except Exception:
            continue
    return None


def element_type(t):
    """Element type for indexing / iterating ``t``."""
    if t is None:
        return None
    try:
        if t.IsArray:
            return t.GetElementType()
        cands = [t] + list(all_interfaces(t))
        for c in cands:
            if c.IsGenericType:
                name = c.GetGenericTypeDefinition().FullName or ""
                if name.startswith(("System.Collections.Generic.IEnumerable`1",
                                    "System.Collections.Generic.IList`1",
                                    "System.Collections.Generic.ICollection`1",
                                    "System.Collections.Generic.IReadOnlyList`1",
                                    "System.Collections.Generic.List`1",
                                    "System.Collections.Generic.ISet`1",
                                    "System.Collections.Generic.HashSet`1")):
                    return c.GetGenericArguments()[0]
        prop = t.GetProperty("Item")
        if prop is not None:
            return prop.PropertyType
    except Exception:
        pass
    return None


def _is_pythonic_clr_type(t):
    """Types better described by Python's dir() than by reflection."""
    try:
        if t.IsPrimitive or t.FullName in ("System.String", "System.Decimal"):
            return True
        ns = t.Namespace or ""
        fn = t.FullName or ""
        return ns.startswith(("IronPython", "Microsoft.Scripting")) or "NewTypes" in fn
    except Exception:
        return True


def _clr_type_of_class(obj):
    """System.Type for a Python reference to a .NET class, else None."""
    try:
        if isinstance(obj, type):
            t = clr.GetClrType(obj)
            if t is not None and not _is_pythonic_clr_type(t):
                return t
    except Exception:
        pass
    return None


def _clr_type_of_instance(obj):
    try:
        t = clr.GetClrType(type(obj))
        if t is not None and not _is_pythonic_clr_type(t):
            return t
    except Exception:
        pass
    return None


def _is_namespace(obj):
    try:
        return type(obj).__name__ in ("namespace#", "NamespaceTracker") or \
            "NamespaceTracker" in clr.GetClrType(type(obj)).FullName
    except Exception:
        return False


def _is_module(obj):
    return type(obj).__name__ == "module"


# ---------------------------------------------------------------------------
# resolved references
# ---------------------------------------------------------------------------
class Ref(object):
    def members(self):
        return {}

    def member(self, name):
        return None

    def call(self):
        return None

    def index(self):
        return None

    def iterate(self):
        return self.index()

    def describe(self, stubs=None):
        return None

    def signatures(self, stubs=None):
        return []


class PyRef(Ref):
    """A live Python value (module, CLR namespace, class, or plain object)."""

    def __init__(self, value):
        self.value = value

    def _safe_container(self):
        v = self.value
        return _is_module(v) or _is_namespace(v) or isinstance(v, type)

    def members(self):
        v = self.value
        container_key = None
        if _is_module(v) or _is_namespace(v):
            container_key = ("py", id(v))
            cached = _MEMBER_CACHE.get(container_key)
            if cached is not None:
                return cached
        out = self._members_uncached()
        if container_key is not None:
            _MEMBER_CACHE[container_key] = out
        return out

    def _members_uncached(self):
        v = self.value
        out = {}
        try:
            names = dir(v)
        except Exception:
            return out
        container = self._safe_container()
        cls = type(v)
        for n in names:
            kind = "field"
            try:
                if container:
                    a = getattr(v, n)
                else:
                    a = getattr(cls, n, None)
                    if a is None and hasattr(v, "__dict__") and n in v.__dict__:
                        a = v.__dict__[n]
                if _is_module(a):
                    kind = "module"
                elif _is_namespace(a):
                    kind = "namespace"
                elif isinstance(a, type):
                    t = _clr_type_of_class(a)
                    kind = "enum" if (t is not None and t.IsEnum) else "class"
                elif isinstance(a, property):
                    kind = "property"
                elif callable(a):
                    kind = "method" if not container or isinstance(v, type) else "function"
                elif type(a).__name__ in ("getset_descriptor", "member_descriptor", "reflectedproperty", "reflectedfield"):
                    kind = "property"
            except Exception:
                pass
            out[n] = (kind, None)
        return out

    def member(self, name):
        v = self.value
        try:
            if self._safe_container():
                return from_value(getattr(v, name))
            d = getattr(v, "__dict__", None)
            if isinstance(d, dict) and name in d:
                return from_value(d[name])
        except Exception:
            pass
        return None

    def call(self):
        v = self.value
        if isinstance(v, type):
            return None  # python class instance: unknown members without running
        return None

    def index(self):
        v = self.value
        try:
            if isinstance(v, (list, tuple)) and len(v):
                return from_value(v[0])
        except Exception:
            pass
        return None

    def describe(self, stubs=None):
        v = self.value
        try:
            if _is_module(v):
                return "module " + getattr(v, "__name__", "?")
            if _is_namespace(v):
                return "namespace " + str(v)
            doc = getattr(v, "__doc__", None)
            head = "{0}: {1}".format(type(v).__name__, _short_repr(v))
            if callable(v) and doc:
                return doc.strip()
            return head
        except Exception:
            return None

    def signatures(self, stubs=None):
        v = self.value
        # python functions: real argument spec
        try:
            import inspect
            if inspect.isfunction(v) or inspect.ismethod(v):
                getspec = getattr(inspect, "getfullargspec", None) or inspect.getargspec
                spec = getspec(v)
                args = list(spec.args)
                defaults = list(spec.defaults or [])
                parts = []
                for i, a in enumerate(args):
                    j = i - (len(args) - len(defaults))
                    parts.append(a if j < 0 else u"{0}={1!r}".format(a, defaults[j]))
                if spec.varargs:
                    parts.append("*" + spec.varargs)
                kw = getattr(spec, "varkw", None) or getattr(spec, "keywords", None)
                if kw:
                    parts.append("**" + kw)
                return [u"{0}({1})".format(getattr(v, "__name__", "f"), ", ".join(parts))]
        except Exception:
            pass
        # .NET methods / builtins: IronPython lists overloads in __doc__
        doc = None
        try:
            doc = getattr(v, "__doc__", None)
        except Exception:
            pass
        if not doc:
            return []
        lines = [l.strip() for l in doc.strip().splitlines() if l.strip()]
        sigs = [l for l in lines if "(" in l and ")" in l and not l.endswith(".")][:8]
        return sigs or lines[:2]


class TypeRef(Ref):
    """A .NET class referenced by name (static members, constructors)."""

    def __init__(self, t):
        self.type = t

    def members(self):
        out = {}
        for name, ms in clr_members(self.type, static=True).items():
            out[name] = (member_kind(ms, self.type), ms)
        return out

    def member(self, name):
        ms = clr_members(self.type, static=True).get(name)
        if not ms:
            return None
        kind = member_kind(ms, self.type)
        if kind in ("class", "enum"):
            return TypeRef(ms[0])
        if kind in ("method",):
            return MethodRef(self.type, name, ms)
        if kind == "enummember":
            return InstanceRef(self.type)
        rt = member_result_type(ms)
        return InstanceRef(rt) if rt is not None else None

    def call(self):
        return InstanceRef(self.type)

    def describe(self, stubs=None):
        t = self.type
        kind = "enum" if t.IsEnum else ("interface" if t.IsInterface else "class")
        base = ""
        try:
            if t.BaseType is not None and t.BaseType.FullName not in ("System.Object", "System.Enum", "System.ValueType"):
                base = "({0})".format(type_name(t.BaseType))
        except Exception:
            pass
        text = "{0} {1}{2}".format(kind, t.FullName, base)
        doc = _stub_class_doc(stubs, t)
        return text + ("\n\n" + doc if doc else "")

    def signatures(self, stubs=None):
        try:
            ctors = self.type.GetConstructors()
            return [method_signature(c) for c in ctors][:10]
        except Exception:
            return []


class InstanceRef(Ref):
    """An instance of a .NET type (declared or runtime)."""

    def __init__(self, t, value=None):
        self.type = t
        self.value = value

    def members(self):
        out = {}
        for name, ms in clr_members(self.type, static=False).items():
            out[name] = (member_kind(ms, self.type), ms)
        return out

    def member(self, name):
        ms = clr_members(self.type, static=False).get(name)
        if not ms:
            # static access through an instance also works in IronPython
            return TypeRef(self.type).member(name)
        kind = member_kind(ms, self.type)
        if kind == "method":
            return MethodRef(self.type, name, ms)
        rt = member_result_type(ms)
        return InstanceRef(rt) if rt is not None else None

    def index(self):
        et = element_type(self.type)
        return InstanceRef(et) if et is not None else None

    def describe(self, stubs=None):
        text = "{0} instance".format(type_name(self.type, full=True))
        if self.value is not None:
            text += "\n" + _short_repr(self.value)
        return text


class MethodRef(Ref):
    def __init__(self, owner, name, members):
        self.owner = owner
        self.name = name
        self.methods = [m for m in members if m.MemberType == MemberTypes.Method]

    def call(self):
        rt = member_result_type(self.methods)
        return InstanceRef(rt) if rt is not None else None

    def signatures(self, stubs=None):
        sigs = [method_signature(m) for m in self.methods][:12]
        return sigs

    def describe(self, stubs=None):
        head = "{0}.{1}".format(type_name(self.owner, full=True), self.name)
        sigs = self.signatures(stubs)
        doc = _stub_member_doc(stubs, self.owner, self.name)
        out = head + "\n" + "\n".join(sigs)
        if doc:
            out += "\n\n" + doc
        return out


def from_value(v):
    if v is None:
        return None
    if _is_module(v) or _is_namespace(v):
        return PyRef(v)
    t = _clr_type_of_class(v)
    if t is not None:
        return TypeRef(t)
    t = _clr_type_of_instance(v)
    if t is not None:
        return InstanceRef(t, v)
    return PyRef(v)


def _short_repr(v, limit=200):
    try:
        r = repr(v)
    except Exception:
        return "<?>"
    return r if len(r) <= limit else r[:limit - 3] + "..."


def _stub_class_doc(stubs, t):
    if stubs is None or t is None:
        return None
    try:
        info = stubs.class_info(t.Namespace, t.Name)
        return info.get("doc") if info else None
    except Exception:
        return None


def _stub_member_info(stubs, t, name):
    if stubs is None or t is None:
        return None
    try:
        cur = t
        while cur is not None:
            info = stubs.member_info(cur.Namespace, cur.Name, name)
            if info:
                return info
            cur = cur.BaseType
    except Exception:
        pass
    return None


def _stub_member_doc(stubs, t, name):
    info = _stub_member_info(stubs, t, name)
    return info.get("doc") if info else None


def describe_member(kind, members, owner, stubs, name):
    """Tooltip text for one completion entry."""
    lines = []
    if members:
        try:
            for m in members[:10]:
                mt = m.MemberType
                if mt == MemberTypes.Method:
                    lines.append(method_signature(m))
                elif mt == MemberTypes.Property:
                    acc = []
                    if m.CanRead:
                        acc.append("get")
                    if m.CanWrite:
                        acc.append("set")
                    lines.append("{0}: {1}  {{{2}}}".format(m.Name, type_name(m.PropertyType), "; ".join(acc)))
                elif mt == MemberTypes.Field:
                    if kind == "enummember":
                        lines.append("{0}.{1}".format(owner.Name, m.Name))
                    else:
                        lines.append("{0}: {1}".format(m.Name, type_name(m.FieldType)))
                elif mt == MemberTypes.Event:
                    lines.append("event {0}: {1}".format(m.Name, type_name(m.EventHandlerType)))
                elif mt in (MemberTypes.NestedType, MemberTypes.TypeInfo):
                    lines.append(("enum " if m.IsEnum else "class ") + (m.FullName or m.Name))
            if len(members) > 10:
                lines.append("(+{0} more overloads)".format(len(members) - 10))
        except Exception:
            pass
        if owner is not None:
            lines.insert(0, type_name(owner, full=True))
    info = _stub_member_info(stubs, owner, name) if owner is not None else None
    if info:
        if not members and info.get("sigs"):
            lines.extend(info["sigs"][:6])
        if info.get("doc"):
            lines.append("")
            lines.append(info["doc"])
    return "\n".join(lines) if lines else None


# ---------------------------------------------------------------------------
# completion items (UI-agnostic)
# ---------------------------------------------------------------------------
class Item(object):
    __slots__ = ("name", "kind", "members", "owner", "detail", "_desc", "describer")

    def __init__(self, name, kind, members=None, owner=None, detail=None, describer=None):
        self.name = name
        self.kind = kind
        self.members = members
        self.owner = owner
        self.detail = detail
        self._desc = None
        self.describer = describer

    def description(self, stubs):
        if self._desc is None:
            try:
                if self.describer is not None:
                    self._desc = self.describer() or ""
                else:
                    self._desc = describe_member(self.kind, self.members, self.owner, stubs, self.name) or ""
            except Exception:
                self._desc = ""
        return self._desc or None

    @property
    def priority(self):
        return KINDS.get(self.kind, ("?", "", 5))[2]


def _detail_for(kind, members):
    if not members:
        return None
    try:
        rt = member_result_type(members)
        if kind == "method":
            n = len(members)
            rt_s = type_name(rt) if rt is not None else "None"
            return "-> " + rt_s + (" (+{0})".format(n - 1) if n > 1 else "")
        if kind in ("property", "field") and rt is not None:
            return type_name(rt)
    except Exception:
        pass
    return None


class CompletionEngine(object):
    def __init__(self, scope_provider, stubs=None):
        """``scope_provider()`` returns the live session scope dict."""
        self._scope = scope_provider
        self.stubs = stubs

    # -- resolution -------------------------------------------------------
    def _resolve_module(self, dotted):
        if dotted in sys.modules:
            return PyRef(sys.modules[dotted])
        root = dotted.split(".")[0]
        scope = self._scope() or {}
        if root in scope and len(dotted.split(".")) > 1:
            ref = from_value(scope[root])
            for part in dotted.split(".")[1:]:
                ref = ref.member(part) if ref else None
            if ref is not None:
                return ref
        if root in CLR_ROOTS or root in sys.modules:
            try:
                mod = __import__(root)
                ref = PyRef(mod)
                for part in dotted.split(".")[1:]:
                    ref = ref.member(part) if ref else None
                return ref
            except Exception:
                return None
        return None

    def _resolve_name(self, name, text, offset, depth):
        scope = self._scope() or {}
        if name in scope:
            return from_value(scope[name])
        binding = pyparse.find_binding(text, name, offset) if text else None
        if binding and depth < 6:
            kind, expr = binding
            if kind == "import":
                return self._resolve_module(expr)
            sub = self.resolve(expr, text, offset, depth + 1)
            if sub is None:
                return None
            if kind == "for":
                return sub.iterate()
            return sub
        if hasattr(_builtins, name):
            return from_value(getattr(_builtins, name))
        if name in CLR_ROOTS:
            return self._resolve_module(name)
        return None

    def resolve(self, expr, text=None, offset=None, depth=0):
        segs = pyparse.split_chain(expr)
        if not segs:
            return None
        ref = None
        for i, (name, trailers) in enumerate(segs):
            if i == 0:
                if name == "<str>":
                    ref = PyRef(u"")
                else:
                    ref = self._resolve_name(name, text, offset, depth)
            else:
                ref = ref.member(name) if ref is not None else None
            for tr in trailers:
                if ref is None:
                    break
                ref = ref.call() if tr == "call" else ref.index()
            if ref is None:
                return None
        return ref

    # -- completion -------------------------------------------------------
    def complete(self, text, offset):
        """Return ``(start_offset, partial, [Item])`` or ``None``."""
        before = text[:offset]
        line_start = before.rfind("\n") + 1
        line_before = before[line_start:]
        ctx = pyparse.import_context(line_before)
        if ctx is not None:
            kind, module, partial = ctx
            items = self._import_items(kind, module)
            return offset - len(partial), partial, items
        tail = before[-4000:]
        expr, partial = pyparse.extract_expression(tail)
        if partial is None:
            return None
        if expr is None:
            items = [it for it in self._bare_items(text, offset)
                     if not (it.kind == "word" and it.name == partial)]
            return offset - len(partial), partial, items
        ref = self.resolve(expr, text, line_start)
        if ref is None:
            return None
        owner = getattr(ref, "type", None)
        items = []
        for name, (kind, members) in ref.members().items():
            if name.startswith("__") and not partial.startswith("_"):
                continue
            items.append(Item(name, kind, members, owner, _detail_for(kind, members)))
        return offset - len(partial), partial, items

    def _bare_items(self, text, offset):
        items = []
        seen = set()

        def add(name, kind, describer=None):
            if name not in seen and name:
                seen.add(name)
                items.append(Item(name, kind, describer=describer))

        for k in pyparse.PY_KEYWORDS:
            add(k, "keyword")
        scope = self._scope() or {}
        for n, v in list(scope.items()):
            if n.startswith("__"):
                continue
            ref_kind = "variable"
            if _is_module(v):
                ref_kind = "module"
            elif _is_namespace(v):
                ref_kind = "namespace"
            elif isinstance(v, type):
                ref_kind = "class"
            elif callable(v):
                ref_kind = "function"
            add(n, ref_kind, (lambda vv=v: (from_value(vv) or Ref()).describe(self.stubs)))
        for n in dir(_builtins):
            if not n.startswith("_"):
                add(n, "builtin")
        for w in pyparse.words_in_text(text):
            add(w, "word")
        return items

    def _import_items(self, kind, module):
        items = []
        seen = set()
        if kind == "names":
            ref = self._resolve_module(module)
            if ref is not None:
                for name, (k, members) in ref.members().items():
                    if name.startswith("_"):
                        continue
                    seen.add(name)
                    items.append(Item(name, k, members, getattr(ref, "type", None)))
            if self.stubs is not None:
                for name in self.stubs.module_names(module):
                    if name not in seen:
                        seen.add(name)
                        items.append(Item(name, "class"))
                for name in self.stubs.submodules(module):
                    if name not in seen:
                        seen.add(name)
                        items.append(Item(name, "namespace"))
            items.append(Item("*", "keyword"))
            return items
        # module path
        if module:
            ref = self._resolve_module(module)
            if ref is not None:
                for name, (k, members) in ref.members().items():
                    if k in ("module", "namespace") and name not in seen:
                        seen.add(name)
                        items.append(Item(name, k))
        else:
            for name in CLR_ROOTS + [m.split(".")[0] for m in list(sys.modules.keys())]:
                if name and not name.startswith("_") and name not in seen:
                    seen.add(name)
                    items.append(Item(name, "module"))
        if self.stubs is not None:
            for name in self.stubs.submodules(module):
                if name not in seen:
                    seen.add(name)
                    items.append(Item(name, "namespace"))
        return items

    # -- signature help ---------------------------------------------------
    def signature_help(self, text, offset):
        """Return ``(signatures, arg_index, open_paren_offset, doc)`` or None."""
        before = text[:offset]
        ctx = pyparse.call_context(before[-4000:])
        if ctx is None:
            return None
        callee, arg_index, paren_rel = ctx
        paren_offset = max(0, len(before) - 4000) + paren_rel
        ref = self.resolve(callee, text, before.rfind("\n") + 1)
        if ref is None:
            return None
        sigs = ref.signatures(self.stubs)
        if not sigs:
            return None
        doc = None
        if isinstance(ref, MethodRef):
            doc = _stub_member_doc(self.stubs, ref.owner, ref.name)
        elif isinstance(ref, TypeRef):
            doc = _stub_class_doc(self.stubs, ref.type)
        return sigs, arg_index, paren_offset, doc

    # -- hover --------------------------------------------------------------
    def hover(self, text, offset):
        s, e = pyparse.identifier_at(text, offset)
        if s >= e:
            return None
        expr = text[s:e]
        line_start = text.rfind("\n", 0, s) + 1
        if "." in expr:
            parent, _, name = expr.rpartition(".")
            pref = self.resolve(parent, text, line_start)
            if pref is None:
                return None
            members = pref.members().get(name)
            if members is not None and members[1]:
                return describe_member(members[0], members[1], getattr(pref, "type", None), self.stubs, name)
            ref = pref.member(name)
        else:
            if expr in pyparse.PY_KEYWORDS:
                return None
            ref = self.resolve(expr, text, line_start)
        if ref is None:
            return None
        return ref.describe(self.stubs)
