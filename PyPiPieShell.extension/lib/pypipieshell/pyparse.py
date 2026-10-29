# -*- coding: utf-8 -*-
"""Pure-Python source helpers for PyPiPie Shell.

Nothing in this module touches .NET, WPF or Revit, so it runs (and is unit
tested) under CPython as well as IronPython 2.7 / 3.4.  The editor uses it for:

* finding the expression in front of the caret for autocomplete,
* import-statement completion context,
* signature help (which call is the caret inside, which argument),
* deciding when a REPL entry is a complete statement,
* indentation-based folding,
* light static inference (``x = Foo(...)`` / ``for x in expr:``),
* parsing ``.pyi`` stub files into a member index,
* locating ``File "<editor>", line N`` in tracebacks.
"""

import re

IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
_IDENT_CHARS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")

PY_KEYWORDS = [
    "and", "as", "assert", "break", "class", "continue", "def", "del", "elif",
    "else", "except", "exec", "finally", "for", "from", "global", "if",
    "import", "in", "is", "lambda", "nonlocal", "not", "or", "pass", "print",
    "raise", "return", "try", "while", "with", "yield", "True", "False", "None",
]

DEDENT_AFTER = ("return", "pass", "break", "continue", "raise")


# ---------------------------------------------------------------------------
# string / comment masking
# ---------------------------------------------------------------------------
def mask_strings_and_comments(text):
    """Return ``(masked, in_string)``: ``text`` with string literal bodies and
    comments replaced by spaces (same length), so bracket and colon scanning
    is not fooled, plus whether the text ends inside an unterminated string.
    """
    masked, in_string, _ = _scan(text)
    return masked, in_string


def ends_in_comment(text):
    return _scan(text)[2]


def _scan(text):
    in_comment = False
    out = []
    i = 0
    n = len(text)
    in_str = None  # the quote token: ', ", ''' or """
    while i < n:
        ch = text[i]
        if in_str:
            if ch == "\\" and i + 1 < n:
                out.append("  ")
                i += 2
                continue
            if text.startswith(in_str, i):
                out.append(in_str)
                i += len(in_str)
                in_str = None
                continue
            if ch == "\n" and len(in_str) == 1:
                # single-quoted strings end at a newline (syntax error, but
                # keep the scanner sane)
                in_str = None
                out.append(ch)
                i += 1
                continue
            out.append("\n" if ch == "\n" else " ")
            i += 1
            continue
        if ch == "#":
            j = text.find("\n", i)
            if j == -1:
                j = n
                in_comment = True
            out.append(" " * (j - i))
            i = j
            continue
        if ch in ("'", '"'):
            tri = ch * 3
            if text.startswith(tri, i):
                in_str = tri
                out.append(tri)
                i += 3
            else:
                in_str = ch
                out.append(ch)
                i += 1
            continue
        out.append(ch)
        i += 1
    return "".join(out), in_str is not None, in_comment


def bracket_depth(text):
    masked, _ = mask_strings_and_comments(text)
    depth = 0
    for ch in masked:
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth -= 1
    return depth


# ---------------------------------------------------------------------------
# expression in front of the caret
# ---------------------------------------------------------------------------
def _skip_group_backward(text, end, open_ch, close_ch):
    """``text[end]`` is ``close_ch``; return index of matching ``open_ch`` or -1."""
    depth = 0
    i = end
    while i >= 0:
        ch = text[i]
        if ch == close_ch:
            depth += 1
        elif ch == open_ch:
            depth -= 1
            if depth == 0:
                return i
        i -= 1
    return -1


def extract_expression(before_caret):
    """Split the text in front of the caret into ``(expr, partial)``.

    ``partial`` is the identifier being typed (may be empty).  ``expr`` is the
    dotted expression before the last ``.`` (``None`` when there is no dot)::

        "x = doc.Act"                       -> ("doc", "Act")
        "FilteredElementCollector(doc)."    -> ("FilteredElementCollector(doc)", "")
        "print(wal"                         -> (None, "wal")
    """
    masked, in_string, in_comment = _scan(before_caret)
    if in_string or in_comment:
        return None, None
    i = len(masked) - 1
    while i >= 0 and masked[i] in _IDENT_CHARS:
        i -= 1
    partial = before_caret[i + 1:]
    if partial and partial[0].isdigit():
        return None, None
    if i < 0 or masked[i] != ".":
        return None, partial
    dot = i
    i -= 1
    # walk back over  name, name(...), name[...], chains
    while i >= 0:
        ch = masked[i]
        if ch in _IDENT_CHARS:
            i -= 1
        elif ch == ")":
            j = _skip_group_backward(masked, i, "(", ")")
            if j < 0:
                return None, partial
            i = j - 1
        elif ch == "]":
            j = _skip_group_backward(masked, i, "[", "]")
            if j < 0:
                return None, partial
            i = j - 1
        elif ch == ".":
            if i + 1 < len(masked) and masked[i + 1] == ".":
                return None, partial
            i -= 1
        elif ch in ("'", '"'):
            # string literal receiver: "abc".upper
            q = ch
            j = masked.rfind(q, 0, i)
            if j < 0:
                return None, partial
            i = j - 1
            break
        else:
            break
    expr = before_caret[i + 1:dot].strip()
    if not expr or expr[0] == "." or expr[0].isdigit():
        return None, partial
    return expr, partial


def split_chain(expr):
    """Split ``a.b(x).c[0]`` into ``[("a", []), ("b", ["call"]), ("c", ["index"])]``.

    A leading string literal becomes ``("<str>", [])``.  Returns ``None`` if the
    expression cannot be parsed.
    """
    if expr is None:
        return None
    masked, _ = mask_strings_and_comments(expr)
    segs = []
    i = 0
    n = len(masked)
    if n and masked[0] in ("'", '"'):
        q = masked[0]
        j = masked.find(q, 1)
        if j < 0:
            return None
        segs.append(("<str>", []))
        i = j + 1
        if i < n and masked[i] == ".":
            i += 1
        else:
            return segs if i >= n else None
    while i < n:
        m = IDENT_RE.match(masked, i)
        if not m:
            return None
        name = masked[m.start():m.end()]
        i = m.end()
        trailers = []
        while i < n and masked[i] in "([":
            open_ch = masked[i]
            close_ch = ")" if open_ch == "(" else "]"
            depth = 0
            j = i
            while j < n:
                if masked[j] == open_ch:
                    depth += 1
                elif masked[j] == close_ch:
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            if j >= n:
                return None
            trailers.append("call" if open_ch == "(" else "index")
            i = j + 1
        segs.append((name, trailers))
        if i < n:
            if masked[i] != ".":
                return None
            i += 1
    return segs


def identifier_at(text, offset):
    """Return ``(start, end)`` of the dotted expression touching ``offset``
    (used for hover); ``(offset, offset)`` if none."""
    n = len(text)
    if offset > n:
        offset = n
    s = offset
    while s > 0 and (text[s - 1] in _IDENT_CHARS or text[s - 1] == "."):
        s -= 1
    e = offset
    while e < n and text[e] in _IDENT_CHARS:
        e += 1
    while s < e and text[s] == ".":
        s += 1
    return s, e


# ---------------------------------------------------------------------------
# import completion context
# ---------------------------------------------------------------------------
_FROM_IMPORT_RE = re.compile(r"^\s*from\s+([A-Za-z_][\w\.]*)\s+import\s+(?:\(?\s*)?((?:[\w\s,]*,\s*)?)([A-Za-z_]\w*)?$")
_FROM_RE = re.compile(r"^\s*from\s+([A-Za-z_][\w\.]*\.)?([A-Za-z_]\w*)?$")
_IMPORT_RE = re.compile(r"^\s*import\s+(?:[\w\.]+\s*,\s*)*([A-Za-z_][\w\.]*\.)?([A-Za-z_]\w*)?$")


def import_context(line_before_caret):
    """Return ``("names", module, partial)`` inside ``from module import ...``,
    ``("modules", parent_or_empty, partial)`` while typing a module path,
    or ``None``."""
    m = _FROM_IMPORT_RE.match(line_before_caret)
    if m:
        return ("names", m.group(1), m.group(3) or "")
    m = _FROM_RE.match(line_before_caret)
    if m:
        parent = (m.group(1) or "").rstrip(".")
        return ("modules", parent, m.group(2) or "")
    m = _IMPORT_RE.match(line_before_caret)
    if m:
        parent = (m.group(1) or "").rstrip(".")
        return ("modules", parent, m.group(2) or "")
    return None


# ---------------------------------------------------------------------------
# signature help
# ---------------------------------------------------------------------------
def call_context(before_caret):
    """Find the innermost unclosed call around the caret.

    Returns ``(callee_expr, arg_index, open_paren_offset)`` or ``None``.
    """
    masked, in_string = mask_strings_and_comments(before_caret)
    if in_string:
        return None
    depth = 0
    commas = 0
    i = len(masked) - 1
    # do not cross more than a few lines back
    limit = max(0, len(masked) - 2000)
    while i >= limit:
        ch = masked[i]
        if ch in ")]}":
            depth += 1
        elif ch in "[{":
            if depth == 0:
                return None
            depth -= 1
        elif ch == "(":
            if depth == 0:
                callee, _ = extract_expression(before_caret[:i] + ".")
                if callee is None:
                    return None
                return callee, commas, i
            depth -= 1
        elif ch == "," and depth == 0:
            commas += 1
        elif ch == ":" and depth == 0:
            return None
        i -= 1
    return None


# ---------------------------------------------------------------------------
# REPL
# ---------------------------------------------------------------------------
def is_complete_statement(src):
    """Heuristic: should Enter in the REPL execute ``src`` or add a line?"""
    if not src.strip():
        return True
    lines = src.split("\n")
    masked, in_string = mask_strings_and_comments(src)
    if in_string:
        return False
    if bracket_depth(src) > 0:
        return False
    mlines = masked.split("\n")
    last = mlines[-1].rstrip()
    if last.endswith("\\"):
        return False
    if len(lines) > 1:
        # a block is finished by an empty line
        return lines[-1].strip() == ""
    if last.endswith(":"):
        return False
    first = lines[0].lstrip()
    if first.startswith("@"):
        return False
    return True


# ---------------------------------------------------------------------------
# indentation
# ---------------------------------------------------------------------------
def leading_ws(line):
    return line[:len(line) - len(line.lstrip(" \t"))]


def indent_for_next_line(prev_line, indent_size=4):
    """Indentation for a new line typed after ``prev_line``."""
    base = leading_ws(prev_line).replace("\t", " " * indent_size)
    masked, _ = mask_strings_and_comments(prev_line)
    code = masked.rstrip()
    if code.endswith(":"):
        return base + " " * indent_size
    if bracket_depth(prev_line) > 0:
        return base + " " * indent_size
    stripped = code.strip()
    for kw in DEDENT_AFTER:
        if stripped == kw or stripped.startswith(kw + " ") or stripped.startswith(kw + "("):
            return base[:max(0, len(base) - indent_size)]
    return base


# ---------------------------------------------------------------------------
# folding
# ---------------------------------------------------------------------------
_FOLD_HEADER_RE = re.compile(r"^\s*(def|class|if|elif|else|for|while|try|except|finally|with)\b")


def compute_foldings(text, min_lines=2):
    """Indentation-based folds for compound statements.

    Returns a sorted list of ``(start_offset, end_offset, title)``.  The fold
    starts at the end of the header line so the header stays visible.
    """
    masked, _ = mask_strings_and_comments(text)
    lines = text.split("\n")
    mlines = masked.split("\n")
    offsets = []
    pos = 0
    for ln in lines:
        offsets.append(pos)
        pos += len(ln) + 1

    def indent_of(idx):
        s = mlines[idx]
        return len(s) - len(s.lstrip(" \t"))

    def is_blank(idx):
        return not mlines[idx].strip()

    folds = []
    n = len(lines)
    for i in range(n):
        if is_blank(i):
            continue
        code = mlines[i].rstrip()
        if not code.endswith(":"):
            continue
        m = _FOLD_HEADER_RE.match(lines[i])
        if not m:
            continue
        base = indent_of(i)
        last = i
        j = i + 1
        while j < n:
            if is_blank(j):
                j += 1
                continue
            if indent_of(j) <= base:
                break
            last = j
            j += 1
        if last - i + 1 < min_lines or last == i:
            continue
        start = offsets[i] + len(lines[i])
        end = offsets[last] + len(lines[last])
        title = lines[i].strip()
        if len(title) > 60:
            title = title[:57] + "..."
        folds.append((start, end, " ... "))
    folds.sort(key=lambda f: f[0])
    return folds


# ---------------------------------------------------------------------------
# static inference
# ---------------------------------------------------------------------------
def find_binding(text, name, before_offset=None):
    """Find the last binding of ``name`` before ``before_offset``.

    Returns ``("assign", rhs_expr)``, ``("for", iter_expr)``,
    ``("with", ctx_expr)``, ``("import", dotted)``, or ``None``.
    """
    if before_offset is not None:
        text = text[:before_offset]
    n = re.escape(name)
    patterns = [
        ("assign", re.compile(r"^[ \t]*" + n + r"[ \t]*=(?!=)[ \t]*(.+?)[ \t]*(?:#.*)?$", re.M)),
        ("for", re.compile(r"^[ \t]*for[ \t]+" + n + r"[ \t]+in[ \t]+(.+?)[ \t]*:[ \t]*(?:#.*)?$", re.M)),
        ("with", re.compile(r"^[ \t]*with[ \t]+(.+?)[ \t]+as[ \t]+" + n + r"[ \t]*:", re.M)),
        ("import", re.compile(r"^[ \t]*import[ \t]+([\w\.]+)[ \t]+as[ \t]+" + n + r"\b", re.M)),
        ("fromimport", re.compile(r"^[ \t]*from[ \t]+([\w\.]+)[ \t]+import[ \t]+(?:[\w, \t]*,[ \t]*)?(\w+)[ \t]+as[ \t]+" + n + r"\b", re.M)),
        ("fromimport_plain", re.compile(r"^[ \t]*from[ \t]+([\w\.]+)[ \t]+import[ \t]+[^#\n]*(?<!as )\b" + n + r"\b(?![ \t]+as\b)", re.M)),
    ]
    best = None
    best_pos = -1
    for kind, rx in patterns:
        for m in rx.finditer(text):
            if m.start() >= best_pos:
                if kind == "fromimport":
                    value = m.group(1) + "." + m.group(2)
                    kind_out = "import"
                elif kind == "fromimport_plain":
                    value = m.group(1) + "." + name
                    kind_out = "import"
                else:
                    value = m.group(1)
                    kind_out = kind
                best = (kind_out, value)
                best_pos = m.start()
    if best and best[0] == "assign":
        rhs = best[1]
        # tuple unpacking or multiple targets are not inferred
        if rhs.startswith("="):
            return None
    return best


_WORD_RE = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]{2,}\b")


def words_in_text(text, limit=4000):
    seen = set()
    out = []
    masked, _ = mask_strings_and_comments(text)
    for m in _WORD_RE.finditer(masked):
        w = m.group(0)
        if w not in seen:
            seen.add(w)
            out.append(w)
            if len(out) >= limit:
                break
    return out


# ---------------------------------------------------------------------------
# tracebacks
# ---------------------------------------------------------------------------
_TB_RE = re.compile(r'File "([^"]+)", line (\d+)')


def traceback_location(line):
    m = _TB_RE.search(line or "")
    if not m:
        return None
    return m.group(1), int(m.group(2))


def last_traceback_line(tb_text, filename):
    found = None
    for m in _TB_RE.finditer(tb_text or ""):
        if m.group(1) == filename:
            found = int(m.group(2))
    return found


# ---------------------------------------------------------------------------
# .pyi stub parsing
# ---------------------------------------------------------------------------
_CLASS_RE = re.compile(r"^class\s+([A-Za-z_]\w*)\s*(?:\((.*)\))?\s*:")
_DEF_RE = re.compile(r"^(\s+)?def\s+([A-Za-z_]\w*)\s*\((.*)\)\s*(?:->\s*(.+?))?\s*:\s*(.*)$")
_ATTR_RE = re.compile(r"^(\s+)?([A-Za-z_]\w*)\s*(?::\s*([^=]+?))?\s*(?:=\s*(.*))?$")


def _read_docstring(lines, idx):
    """If ``lines[idx]`` starts a docstring return ``(doc, next_idx)``."""
    if idx >= len(lines):
        return None, idx
    s = lines[idx].strip()
    for q in ('"""', "'''"):
        if s.startswith(q):
            body = s[3:]
            if body.endswith(q) and len(body) >= 3:
                return body[:-3].strip(), idx + 1
            parts = [body]
            j = idx + 1
            while j < len(lines):
                t = lines[j]
                k = t.find(q)
                if k >= 0:
                    parts.append(t[:k])
                    return "\n".join(p.strip() for p in parts).strip(), j + 1
                parts.append(t)
                j += 1
            return "\n".join(parts).strip(), j
    return None, idx


def parse_pyi(text):
    """Parse a stub module.

    Returns ``{"doc": str|None, "classes": {name: cls}, "functions": {name: [sig]},
    "attributes": {name: type}}`` where ``cls`` is
    ``{"bases": str, "doc": str|None, "members": {name: {"kind": k, "sigs": [..], "doc": d, "type": t}}}``.

    ``kind`` is ``method``, ``staticmethod``, ``property`` or ``field``.
    """
    lines = text.replace("\r\n", "\n").split("\n")
    result = {"doc": None, "classes": {}, "functions": {}, "attributes": {}}
    doc, i = _read_docstring(lines, 0)
    result["doc"] = doc
    cur_cls = None
    cur_cls_indent = 0
    decorators = []
    n = len(lines)
    while i < n:
        raw = lines[i]
        stripped = raw.strip()
        if not stripped or stripped.startswith("#"):
            i += 1
            continue
        indent = len(raw) - len(raw.lstrip())
        if cur_cls is not None and indent <= cur_cls_indent:
            cur_cls = None
        if stripped.startswith("@"):
            decorators.append(stripped[1:].split("(")[0].strip())
            i += 1
            continue
        if indent == 0 or (cur_cls is None and indent == 0):
            m = _CLASS_RE.match(stripped)
            if m and indent == 0:
                name = m.group(1)
                cls = result["classes"].setdefault(
                    name, {"bases": (m.group(2) or "").strip(), "doc": None, "members": {}})
                cur_cls = cls
                cur_cls_indent = indent
                decorators = []
                cdoc, nxt = _read_docstring(lines, i + 1)
                if cdoc:
                    cls["doc"] = cdoc
                    i = nxt
                else:
                    i += 1
                continue
        m = _DEF_RE.match(raw.rstrip())
        if m:
            name = m.group(2)
            params = m.group(3).strip()
            ret = (m.group(4) or "").strip()
            sig = "{0}({1})".format(name, params)
            if ret:
                sig += " -> " + ret
            ddoc, nxt = _read_docstring(lines, i + 1)
            if cur_cls is not None and indent > cur_cls_indent:
                kind = "method"
                if "property" in decorators:
                    kind = "property"
                elif "staticmethod" in decorators or "classmethod" in decorators:
                    kind = "staticmethod"
                setter = any(d.endswith(".setter") for d in decorators)
                mem = cur_cls["members"].setdefault(name, {"kind": kind, "sigs": [], "doc": None, "type": None})
                if kind == "property":
                    mem["kind"] = "property"
                    mem["type"] = ret or mem["type"]
                elif not setter and sig not in mem["sigs"]:
                    mem["sigs"].append(sig)
                if ddoc and not mem["doc"]:
                    mem["doc"] = ddoc
            elif indent == 0:
                fn = result["functions"].setdefault(name, {"sigs": [], "doc": None})
                if sig not in fn["sigs"]:
                    fn["sigs"].append(sig)
                if ddoc and not fn["doc"]:
                    fn["doc"] = ddoc
            decorators = []
            # skip the body of a def that is not just "..."
            i = nxt if ddoc else i + 1
            continue
        m = _ATTR_RE.match(raw.rstrip())
        if m and not stripped.startswith(("return", "pass", "...", "raise")):
            name = m.group(2)
            typ = (m.group(3) or "").strip() or None
            if cur_cls is not None and indent > cur_cls_indent:
                mem = cur_cls["members"].setdefault(name, {"kind": "field", "sigs": [], "doc": None, "type": typ})
                if typ and not mem["type"]:
                    mem["type"] = typ
                adoc, nxt = _read_docstring(lines, i + 1)
                if adoc and not mem["doc"]:
                    mem["doc"] = adoc
                    i = nxt
                    decorators = []
                    continue
            elif indent == 0 and (m.group(3) or m.group(4)):
                result["attributes"][name] = typ
        decorators = []
        i += 1
    return result
