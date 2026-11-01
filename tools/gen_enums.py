"""Generate ``src/docviews/compat/enum/`` from python-docx.

    python3 tools/gen_enums.py            # needs python-docx installed

The runtime tables copy python-docx 1.2.0's enum names, values, XML tokens, and
base-class shape so code can keep importing ``WD_LINE_SPACING`` after swapping
the import to ``docviews.compat.enum``. The documented-only ``MIXED`` member is
kept as an explicit Office-API-sourced addition because the official
python-docx enum page lists it but the pinned runtime omits it. No python-docx
implementation code is copied.

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

# Keep python-docx's public sentinel exactly: ``xml_value`` is literally
# ``"UNMAPPED"`` for these members, even though it is not a writable token.
UNMAPPED = "UNMAPPED"

# The published python-docx enum page documents MIXED, but v1.2.0's runtime
# omits it. The Office API defines it as -2. Keep this docs-first addition
# explicit and source-backed rather than pretending it came from the runtime.
DOCUMENTED_ONLY_MEMBERS = {
    "MSO_THEME_COLOR_INDEX": {"MIXED": -2},
}
DOCUMENTED_ONLY_MEMBER_DOCSTRINGS = {
    "MSO_THEME_COLOR_INDEX": {
        "MIXED": "Indicates multiple theme colors are used.",
    },
}
DOCUMENTED_ONLY_VALUE_SOURCES = {
    "MSO_THEME_COLOR_INDEX.MIXED": (
        "https://learn.microsoft.com/en-us/office/vba/api/office.msothemecolorindex"
    ),
}


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

    # Apply documented members that the pinned runtime omits to both the
    # canonical class and its import aliases. Their value and provenance stay
    # separately recorded in the generated data below.
    for name, additions in DOCUMENTED_ONLY_MEMBERS.items():
        for enum_name in enums:
            if aliases.get(enum_name, enum_name) == name:
                for member, value in additions.items():
                    if member in enums[enum_name] and enums[enum_name][member] != value:
                        raise ValueError(
                            f"{enum_name}.{member} conflicts with runtime value")
                    enums[enum_name][member] = value

    lines = [
        '"""Enum members, by value and by XML token — generated, checked in.',
        "",
        "Regenerate with ``python3 tools/gen_enums.py``. Runtime members and XML",
        "tokens come from python-docx " + _python_docx_version() + ". The",
        "documented-only MIXED member is sourced from the Office API; no",
        "python-docx implementation code is copied here.",
        '"""',
        "",
        "# name -> {member: value}; runtime values plus documented-only members",
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
    lines.append("# documented members absent from the pinned runtime")
    lines.append("DOCUMENTED_ONLY_MEMBERS = {")
    for name, additions in DOCUMENTED_ONLY_MEMBERS.items():
        lines.append(f"    {name!r}: {{")
        for member, value in additions.items():
            lines.append(f"        {member!r}: {value!r},")
        lines.append("    },")
    lines.append("}")
    lines.append("DOCUMENTED_ONLY_MEMBER_DOCSTRINGS = {")
    for name, additions in DOCUMENTED_ONLY_MEMBER_DOCSTRINGS.items():
        lines.append(f"    {name!r}: {{")
        for member, docstring in additions.items():
            lines.append(f"        {member!r}: {docstring!r},")
        lines.append("    },")
    lines.append("}")
    lines.append("DOCUMENTED_ONLY_VALUE_SOURCES = {")
    for name, url in DOCUMENTED_ONLY_VALUE_SOURCES.items():
        lines.append(f"    {name!r}: {url!r},")
    lines.append("}")
    lines.append("")
    lines.append("# name -> documented base, or PlainEnum for a direct enum.Enum base")
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
