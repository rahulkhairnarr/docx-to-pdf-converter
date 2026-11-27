"""Resolve the formatting Word would apply to a paragraph or run.

A property is looked up, first hit wins, in: the element's own properties, its
character style, the paragraph's numbering level (paragraph properties only), the
paragraph style and the styles it is based on, then the document defaults.
python-docx reports inherited values as None, so this walks the XML itself.

ponytail: toggle properties (bold in a bold style turning bold off) and table-style
conditional formatting are not applied; the nearest explicit value wins.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml.etree import _Element

from .ooxml import attr, child, half_points_to_pt, on_off, qn, twips_to_pt

DEFAULT_FONT_SIZE = 11.0


@dataclass(frozen=True)
class RunFormat:
    font: str | None
    size: float
    bold: bool
    italic: bool
    underline: bool
    strike: bool
    color: str | None  # "#RRGGBB"
    background: str | None  # "#RRGGBB" or a colour name
    vertical: str | None  # "superscript" | "subscript"
    caps: bool
    hidden: bool


@dataclass(frozen=True)
class TabStop:
    position: float  # points from the left margin
    alignment: str  # "left" | "center" | "right" | "decimal"
    leader: str | None  # "dot" | "hyphen" | "underscore" | "middleDot" | "heavy"


@dataclass(frozen=True)
class Border:
    width: float  # points
    color: str  # "#RRGGBB"
    space: float  # points between the text and the line
    style: str  # w:val: "single", "double", "dashed", "dotted", ...


@dataclass(frozen=True)
class ParagraphFormat:
    alignment: str  # "left" | "center" | "right" | "justify"
    left: float
    right: float
    first_line: float  # negative for a hanging indent
    space_before: float
    space_after: float
    line_spacing: tuple[str, float]  # ("auto", multiple) or ("exact" | "atLeast", points)
    page_break_before: bool
    keep_with_next: bool
    widow_control: bool
    tabs: tuple[TabStop, ...] = ()  # by position
    borders: dict[str, Border] = field(default_factory=dict)  # "top" | "bottom" | "left" | "right"
    shading: str | None = None  # "#RRGGBB" fill behind the paragraph


_ALIGNMENTS = {"center": "center", "right": "right", "end": "right", "both": "justify", "distribute": "justify"}
_HIGHLIGHTS = {"darkBlue": "navy", "darkRed": "maroon", "darkGreen": "darkgreen", "darkYellow": "olive",
               "darkCyan": "teal", "darkMagenta": "purple", "darkGray": "darkgray", "lightGray": "lightgrey"}


class StyleSheet:
    def __init__(self, styles: _Element, theme_fonts: dict[str, str] | None = None) -> None:
        """`theme_fonts` maps "major"/"minor" to the theme's Latin typefaces, which
        `w:asciiTheme="minorHAnsi"` and the like refer to (see `theme_fonts()`)."""
        self._theme_fonts = theme_fonts or {}
        self._by_id = {style.get(qn("w:styleId")): style for style in styles.iterfind(qn("w:style"))}
        self._default_paragraph_style = next(
            (style_id for style_id, style in self._by_id.items()
             if style.get(qn("w:type")) == "paragraph" and on_off_attr(style, "w:default")),
            None,
        )
        self._default_rpr = child(styles, "w:docDefaults", "w:rPrDefault", "w:rPr")
        self._default_ppr = child(styles, "w:docDefaults", "w:pPrDefault", "w:pPr")

    def style_chain(self, style_id: str | None) -> list[_Element]:
        """The style and every style it is based on, nearest first. Cycle-safe."""
        chain: list[_Element] = []
        while style_id and (style := self._by_id.get(style_id)) is not None and style not in chain:
            chain.append(style)
            style_id = attr(child(style, "w:basedOn"))
        return chain

    def paragraph_style_id(self, paragraph: _Element) -> str | None:
        return attr(child(paragraph, "w:pPr", "w:pStyle")) or self._default_paragraph_style

    def paragraph_properties(self, paragraph: _Element, numbering_ppr: _Element | None = None) -> list[_Element]:
        """Paragraph property elements, highest priority first."""
        chain = [child(paragraph, "w:pPr"), numbering_ppr]
        chain += [child(style, "w:pPr") for style in self.style_chain(self.paragraph_style_id(paragraph))]
        chain.append(self._default_ppr)
        return [ppr for ppr in chain if ppr is not None]

    def run_properties(self, paragraph: _Element, run: _Element | None) -> list[_Element]:
        """Run property elements, highest priority first. With `run` None, the
        paragraph's own font (its paragraph mark), used for empty paragraphs and labels."""
        own = child(run, "w:rPr") if run is not None else child(paragraph, "w:pPr", "w:rPr")
        chain = [own]
        chain += [child(style, "w:rPr") for style in self.style_chain(attr(child(own, "w:rStyle")))]
        chain += [child(style, "w:rPr") for style in self.style_chain(self.paragraph_style_id(paragraph))]
        chain.append(self._default_rpr)
        return [rpr for rpr in chain if rpr is not None]

    def run_format(self, paragraph: _Element, run: _Element | None) -> RunFormat:
        chain = self.run_properties(paragraph, run)
        underline = attr(_first(chain, "w:u"))
        vertical = attr(_first(chain, "w:vertAlign"))
        return RunFormat(
            font=self._latin_font(chain),
            size=half_points_to_pt(attr(_first(chain, "w:sz"))) or DEFAULT_FONT_SIZE,
            bold=bool(on_off(_first(chain, "w:b"))),
            italic=bool(on_off(_first(chain, "w:i"))),
            underline=underline is not None and underline != "none",
            strike=bool(on_off(_first(chain, "w:strike")) or on_off(_first(chain, "w:dstrike"))),
            color=_hex(attr(_first(chain, "w:color"))),
            background=_background(_first(chain, "w:highlight"), _first(chain, "w:shd")),
            vertical=vertical if vertical in ("superscript", "subscript") else None,
            caps=bool(on_off(_first(chain, "w:caps")) or on_off(_first(chain, "w:smallCaps"))),
            hidden=bool(on_off(_first(chain, "w:vanish"))),
        )

    def _latin_font(self, chain: list[_Element]) -> str | None:
        """The font for Latin text: from the nearest level that names one. Within one
        w:rFonts a theme reference wins over a plain name (ECMA-376 17.3.2.26); a
        level that only sets other scripts' fonts (w:eastAsia, w:cs) is passed over."""
        for properties in chain:
            fonts = properties.find(qn("w:rFonts"))
            if fonts is None:
                continue
            if theme := attr(fonts, "w:asciiTheme") or attr(fonts, "w:hAnsiTheme"):
                # "minorHAnsi", "majorAscii", ...: the theme's minor (body) or major (heading) font
                return self._theme_fonts.get("major" if theme.startswith("major") else "minor")
            if name := attr(fonts, "w:ascii") or attr(fonts, "w:hAnsi"):
                return name
        return None

    def paragraph_format(self, paragraph: _Element, numbering_ppr: _Element | None = None) -> ParagraphFormat:
        chain = self.paragraph_properties(paragraph, numbering_ppr)
        indent = _merged(chain, "w:ind")
        spacing = _merged(chain, "w:spacing")
        hanging = twips_to_pt(indent.get("hanging"))
        first_line = -hanging if hanging else twips_to_pt(indent.get("firstLine")) or 0.0
        line, rule = spacing.get("line"), spacing.get("lineRule", "auto")
        return ParagraphFormat(
            alignment=_ALIGNMENTS.get(attr(_first(chain, "w:jc")) or "", "left"),
            left=twips_to_pt(indent.get("left", indent.get("start"))) or 0.0,
            right=twips_to_pt(indent.get("right", indent.get("end"))) or 0.0,
            first_line=first_line,
            space_before=twips_to_pt(spacing.get("before")) or 0.0,
            space_after=twips_to_pt(spacing.get("after")) or 0.0,
            line_spacing=("auto", int(line) / 240) if line and rule == "auto"
            else (rule, twips_to_pt(line)) if line else ("auto", 1.0),
            page_break_before=bool(on_off(_first(chain, "w:pageBreakBefore"))),
            keep_with_next=bool(on_off(_first(chain, "w:keepNext"))),
            # Absent means on, as in Word (which writes only w:val="0", to turn it off) and
            # in LibreOffice's DOCX import (StyleSheetTable.cxx sets widows/orphans to 2).
            widow_control=on_off(_first(chain, "w:widowControl")) is not False,
            tabs=_tab_stops(chain),
            borders=_borders(_first(chain, "w:pBdr")),
            shading=_shading(_first(chain, "w:shd")),
        )

    def numbering_reference(self, paragraph: _Element) -> tuple[str, int] | None:
        """(numId, level) from the paragraph or its style; None if not in a list."""
        num_prs = [num_pr for ppr in self.paragraph_properties(paragraph) if (num_pr := child(ppr, "w:numPr")) is not None]
        num_id = attr(_first(num_prs, "w:numId"))
        if not num_id or num_id == "0":  # numId 0 explicitly removes a style's numbering
            return None
        return num_id, int(attr(_first(num_prs, "w:ilvl")) or 0)


def on_off_attr(element: _Element, name: str) -> bool:
    return element.get(qn(name)) in ("1", "true", "on")


def _first(chain: list[_Element], tag: str) -> _Element | None:
    for properties in chain:
        if (found := properties.find(qn(tag))) is not None:
            return found
    return None


def _merged(chain: list[_Element], tag: str) -> dict[str, str]:
    """Attributes of `tag` merged across the chain: Word inherits indent and spacing
    attribute by attribute (a style's left indent survives a direct first-line indent)."""
    merged: dict[str, str] = {}
    for properties in reversed(chain):
        if (found := properties.find(qn(tag))) is not None:
            values = {qname.split("}")[-1]: value for qname, value in found.attrib.items()}
            for one, other in _EXCLUSIVE:  # a nearer first-line indent replaces a hanging one
                if one in values:
                    merged.pop(other, None)
            merged.update(values)
    return merged


_EXCLUSIVE = (("firstLine", "hanging"), ("hanging", "firstLine"))
_TAB_ALIGNMENTS = {"center": "center", "right": "right", "end": "right", "decimal": "decimal"}  # else left


def _tab_stops(chain: list[_Element]) -> tuple[TabStop, ...]:
    """Tab stops accumulate down the chain; a nearer `w:val="clear"` removes an
    inherited stop at the same position."""
    stops: dict[str, TabStop] = {}
    for properties in reversed(chain):
        for tab in properties.iterfind(f"{qn('w:tabs')}/{qn('w:tab')}"):
            position, kind, leader = tab.get(qn("w:pos")), attr(tab), attr(tab, "w:leader")
            if kind == "clear":
                stops.pop(position, None)
            elif position is not None and kind != "bar":  # a bar tab draws a line, it is no stop
                stops[position] = TabStop(twips_to_pt(position) or 0.0, _TAB_ALIGNMENTS.get(kind or "", "left"),
                                          None if leader in (None, "none") else leader)
    return tuple(sorted(stops.values(), key=lambda stop: stop.position))


_SIDES = {"top": "top", "bottom": "bottom", "left": "left", "start": "left", "right": "right", "end": "right"}


def _borders(paragraph_borders: _Element | None) -> dict[str, Border]:
    """The nearest w:pBdr's sides. w:sz is in eighths of a point, w:space in points."""
    found: dict[str, Border] = {}
    for edge in paragraph_borders if paragraph_borders is not None else ():
        side, kind = _SIDES.get(edge.tag.split("}")[-1]), attr(edge)
        if side and kind not in (None, "nil", "none"):
            found[side] = Border(int(attr(edge, "w:sz") or 4) / 8, _hex(attr(edge, "w:color")) or "#000000",
                                 float(attr(edge, "w:space") or 0), kind)
    return found


def theme_fonts(theme: _Element | None) -> dict[str, str]:
    """{"major": typeface, "minor": typeface} from a theme part's font scheme."""
    scheme = child(theme, "a:themeElements", "a:fontScheme")
    found = {}
    for key, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
        latin = child(scheme, tag, "a:latin")
        if latin is not None and (typeface := latin.get("typeface")):  # an unprefixed attribute
            found[key] = typeface
    return found


def _hex(value: str | None) -> str | None:
    if not value or value == "auto" or len(value) != 6:
        return None
    return f"#{value}"


def _shading(shading: _Element | None) -> str | None:
    """A paragraph's fill. A percentage pattern ("pct25": 25% dots of w:color over
    w:fill) is drawn as the colour it averages to."""
    fill, pattern, ink = attr(shading, "w:fill"), attr(shading) or "", _hex(attr(shading, "w:color"))
    background = _hex(fill) or "#FFFFFF"
    if pattern.startswith("pct") and pattern[3:].isdigit() and ink:
        share = int(pattern[3:]) / 100
        mixed = (round(int(ink[i:i + 2], 16) * share + int(background[i:i + 2], 16) * (1 - share)) for i in (1, 3, 5))
        return "#" + "".join(f"{channel:02X}" for channel in mixed)
    return _hex(fill)


def _background(highlight: _Element | None, shading: _Element | None) -> str | None:
    if (name := attr(highlight)) and name != "none":
        return _HIGHLIGHTS.get(name, name.lower())
    return _hex(attr(shading, "w:fill"))
