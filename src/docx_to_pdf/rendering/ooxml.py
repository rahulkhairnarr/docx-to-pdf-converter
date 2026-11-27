"""Small helpers for reading WordprocessingML (the XML inside a .docx) with lxml."""

from __future__ import annotations

from functools import cache

from docx.oxml.ns import nsmap
from lxml.etree import _Element

EMU_PER_PT = 12700

# python-docx's map, plus the prefixes for text boxes, VML pictures and alternate content.
NAMESPACES = {
    **nsmap,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}


@cache
def qn(tag: str) -> str:
    """Clark notation for a prefixed tag: "w:p" -> "{http://...}p"."""
    prefix, local = tag.split(":")
    return f"{{{NAMESPACES[prefix]}}}{local}"


def child(element: _Element | None, *path: str) -> _Element | None:
    """The first element at `path` (tags like "w:pPr") below `element`, or None."""
    for tag in path:
        if element is None:
            return None
        element = element.find(qn(tag))
    return element


def attr(element: _Element | None, name: str = "w:val") -> str | None:
    return None if element is None else element.get(qn(name))


def on_off(element: _Element | None) -> bool | None:
    """A WordprocessingML on/off property: absent is None (inherit), `<w:b/>` is True,
    `<w:b w:val="0"/>` (or "false", "off") is False."""
    if element is None:
        return None
    return attr(element) not in ("0", "false", "off", "none")


def twips_to_pt(value: str | None) -> float | None:
    """Twentieths of a point, the unit of indents and spacing."""
    return None if value is None else int(float(value)) / 20


def half_points_to_pt(value: str | None) -> float | None:
    """Half points, the unit of font sizes."""
    return None if value is None else int(float(value)) / 2
