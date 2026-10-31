"""Generate ``src/docviews/compat/enum/`` from python-docx.

    python3 tools/gen_enums.py            # needs python-docx installed

The value tables are the OOXML constants themselves (``WD_ALIGN_PARAGRAPH.CENTER
== 1`` is a fact about the file format, not an implementation). Copying the
*names, values and XML tokens* is what lets existing code keep writing
``from docx.enum.text import WD_LINE_SPACING`` after swapping the import to
``docviews.compat.enum``; no python-docx code is copied, and the generated file
carries a header saying where the numbers came from.

Three things are generated:

``_generated.py``
    ``ENUMS`` (member -> value), ``TOKENS`` (member -> XML token), ``BASES``
    (which of python-docx's two base classes the enum derives from), ``ALIASES``
    (names python-docx re-exports for the same class) and ``HOME`` (the module
    each name is importable from).

``text.py``, ``style.py``, ``section.py``, ``table.py``, ``shape.py``, ``dml.py``
    the per-module import path, so ``from docx.enum.text import X`` has a
    one-word answer here too. They are re-exports, never hand-edited.
"""

from __future__ import annotations

import importlib
from pathlib import Path

MODULES = ("docx.enum.text", "docx.enum.style", "docx.enum.section",
           "docx.enum.table", "docx.enum.shape", "docx.enum.dml",
           "docx.enum.base", "docx.shared")

ENUM_DIR = (Path(__file__).resolve().parent.parent / "src" / "docviews"
            / "compat" / "enum")

OUT = ENUM_DIR / "_generated.py"

# docx.enum.<stem> for the module a name lives in
STEMS = {"docx.enum.text": "text", "docx.enum.style": "style",
         "docx.enum.section": "section", "docx.enum.table": "table",
         "docx.enum.shape": "shape", "docx.enum.dml": "dml"}

# python-docx marks a member that has no XML token with this string; None is the
# honest answer here, and it keeps such a member out of the writer's path.
UNMAPPED = "UNMAPPED"


def members(cls) -> dict:
    """``{name: value}`` for an enum class, canonical members only."""
    found: dict = {}
    for name, member in getattr(cls, "__members__", {}).items():
        value = getattr(member, "value", None)
        if isinstance(value, (int, str)):
            found[name] = value
    if not found:                                  # a plain class of constants
        for name in dir(cls):
            if name.startswith("_"):
                continue
            value = getattr(cls, name)
            if isinstance(value, (int, str)):
                found[name] = value
    return found


def tokens(cls) -> dict:
    """``{name: xml_token}`` — python-docx's own member-to-XML mapping."""
    found: dict = {}
    for name, member in getattr(cls, "__members__", {}).items():
        token = getattr(member, "xml_value", None)
        if token == UNMAPPED:
            token = None
        if isinstance(token, str):
            found[name] = token
    return found


def base_of(obj) -> str:
    """Which base the class's members print like.

    python-docx has two documented bases, and a few deprecated classes that
    derive from neither (``WD_BREAK_TYPE``) and so print the standard-library
    way, ``"WD_BREAK_TYPE.PAGE"``. That third shape is recorded as ``PlainEnum``
    so the members here print it too.
    """
    try:
        from docx.enum.base import BaseXmlEnum
    except ImportError:                            # pragma: no cover
        return "BaseEnum"
    if isinstance(obj, type) and issubclass(obj, BaseXmlEnum):
        return "BaseXmlEnum"
    first = next(iter(getattr(obj, "__members__", {}).values()), None)
    if first is not None and str(first) == f"{obj.__name__}.{first.name}":
        return "PlainEnum"
    return "BaseEnum"


def main() -> int:
    enums: dict = {}
    token_map: dict = {}
    bases: dict = {}
    aliases: dict = {}
    home: dict = {}
    for module_name in MODULES:
        module = importlib.import_module(module_name)
        for name in dir(module):
            if name.startswith("_") or name in ("enum",):
                continue
            obj = getattr(module, name)
            defined = getattr(obj, "__module__", "")
            if defined and not defined.startswith("docx.enum") \
                    and not defined.startswith("docx.shared"):
                continue                            # an import, not a definition
            if not (isinstance(obj, type) or hasattr(obj, "__members__")):
                continue
            values = members(obj)
            if not values:
                continue
            home.setdefault(name, module_name)
            if name in enums and enums[name] != values:
                continue
            enums.setdefault(name, values)
            token_map.setdefault(name, tokens(obj))
            bases.setdefault(name, base_of(obj))
            if obj.__name__ != name:                # WD_ALIGN_PARAGRAPH -> …
                aliases[name] = obj.__name__

    lines = [
        '"""Enum members, by value and by XML token — generated, checked in.',
        "",
        "Regenerate with ``python3 tools/gen_enums.py``. The names, values and",
        "XML tokens are the OOXML constants (MIT, python-docx " +
        _python_docx_version() + "); this file contains no python-docx code.",
        '"""',
        "",
        "# name -> {member: value}, exactly what python-docx exposes",
        "ENUMS = {",
    ]
    for name in sorted(enums):
        lines.append(f"    {name!r}: {{")
        for member, value in enums[name].items():
            lines.append(f"        {member!r}: {value!r},")
        lines.append("    },")
    lines.append("}")
    lines.append("")
    lines.append("# name -> {member: xml token}, the attribute value Word's XML wants")
    lines.append("TOKENS = {")
    for name in sorted(token_map):
        if not token_map[name]:
            continue
        lines.append(f"    {name!r}: {{")
        for member, token in token_map[name].items():
            lines.append(f"        {member!r}: {token!r},")
        lines.append("    },")
    lines.append("}")
    lines.append("")
    lines.append("# name -> which of python-docx's two enum bases it derives from")
    lines.append("BASES = {")
    for name in sorted(bases):
        lines.append(f"    {name!r}: {bases[name]!r},")
    lines.append("}")
    lines.append("")
    lines.append("# names python-docx re-exports for the same class")
    lines.append("ALIASES = {")
    for name in sorted(aliases):
        lines.append(f"    {name!r}: {aliases[name]!r},")
    lines.append("}")
    lines.append("")
    lines.append("# which module each name is importable from in python-docx")
    lines.append("HOME = {")
    for name in sorted(home):
        lines.append(f"    {name!r}: {home[name]!r},")
    lines.append("}")
    lines.append("")
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT.name} — {len(enums)} classes, "
          f"{sum(len(v) for v in enums.values())} members, "
          f"{len(aliases)} aliases")

    for module_name, stem in STEMS.items():
        names = sorted(n for n, where in home.items() if where == module_name)
        if not names:
            continue
        body = [f'"""``docx.enum.{stem}`` — python-docx\'s module, re-exported.',
                "",
                "Generated by ``tools/gen_enums.py``; edit that, not this.",
                "",
                "The base classes are here for the same reason they are in",
                "python-docx's module: they are imported there for internal use and",
                "end up importable from it.",
                '"""',
                "",
                "from . import (",
                "    BaseEnum,",
                "    BaseXmlEnum,"]
        for name in names:
            body.append(f"    {name},")
        body.append(")")
        body.append("")
        body.append("__all__ = [")
        for name in names:
            body.append(f"    {name!r},")
        body.append("]")
        body.append("")
        (ENUM_DIR / f"{stem}.py").write_text("\n".join(body), encoding="utf-8")
        print(f"wrote {stem}.py — {len(names)} names")
    return 0


def _python_docx_version() -> str:
    import docx
    return docx.__version__


if __name__ == "__main__":
    raise SystemExit(main())
