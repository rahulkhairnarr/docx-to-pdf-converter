"""List labels ("1.", "3.2", "(a)", bullets). Word generates them from numbering.xml
rather than storing them in the text, and clause numbers are content, so they are
computed here: one counter per list (numId) and level, and starting a level resets
the levels below it.

ponytail: counters are kept per numId. Word can also continue a list across numIds
that share an abstract definition. If real documents number wrongly, key them by
abstractNumId unless the num has a startOverride.
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml.etree import _Element

from .ooxml import attr, child, qn

BULLET = "•"  # bullets are usually Symbol/Wingdings codes, which mean nothing in other fonts


@dataclass(frozen=True)
class _Level:
    start: int
    format: str
    text: str
    paragraph_properties: _Element | None
    run_properties: _Element | None


class Numbering:
    def __init__(self, numbering: _Element | None) -> None:
        self._levels: dict[str, dict[int, _Level]] = {}
        self._counters: dict[str, dict[int, int]] = {}
        if numbering is None:
            return
        abstracts = {
            attr(abstract, "w:abstractNumId"): abstract for abstract in numbering.iterfind(qn("w:abstractNum"))
        }
        for num in numbering.iterfind(qn("w:num")):
            abstract = abstracts.get(attr(child(num, "w:abstractNumId")))
            if abstract is None:
                continue
            levels = {int(attr(lvl, "w:ilvl") or 0): _level(lvl) for lvl in abstract.iterfind(qn("w:lvl"))}
            for override in num.iterfind(qn("w:lvlOverride")):
                index = int(attr(override, "w:ilvl") or 0)
                if (replacement := child(override, "w:lvl")) is not None:
                    levels[index] = _level(replacement)
                if (start := attr(child(override, "w:startOverride"))) is not None and index in levels:
                    old = levels[index]
                    levels[index] = _Level(int(start), old.format, old.text, old.paragraph_properties,
                                           old.run_properties)
            self._levels[attr(num, "w:numId")] = levels

    def level_properties(self, num_id: str, level: int) -> tuple[_Element | None, _Element | None]:
        """The level's own (paragraph, run) properties: the list's indent, the label's font."""
        found = self._levels.get(num_id, {}).get(level)
        return (found.paragraph_properties, found.run_properties) if found else (None, None)

    def next_label(self, num_id: str, level: int) -> str | None:
        """Advance the list's counter for `level` and return the label; None if the
        list is undefined or its level has no label."""
        levels = self._levels.get(num_id)
        if not levels or level not in levels:
            return None
        counters = self._counters.setdefault(num_id, {})
        counters[level] = counters.get(level, levels[level].start - 1) + 1
        for deeper in [index for index in counters if index > level]:
            del counters[deeper]
        this = levels[level]
        if this.format == "none":
            return None
        if this.format == "bullet":
            return _bullet(this.text)
        label = this.text
        for index in range(9):
            placeholder = f"%{index + 1}"
            if placeholder in label:
                value = counters.get(index, levels[index].start if index in levels else 1)
                fmt = levels[index].format if index in levels else "decimal"
                label = label.replace(placeholder, format_number(value, fmt))
        return label or None


def format_number(value: int, fmt: str) -> str:
    """`value` in a Word number format. ponytail: formats beyond these (ordinal,
    cardinalText, chicago, other scripts) fall back to decimal."""
    if fmt in ("lowerLetter", "upperLetter"):
        letter = chr(ord("a") + (value - 1) % 26) * ((value - 1) // 26 + 1)  # Word: y, z, aa, bb
        return letter.upper() if fmt == "upperLetter" else letter
    if fmt in ("lowerRoman", "upperRoman"):
        roman = _roman(value)
        return roman.lower() if fmt == "lowerRoman" else roman
    if fmt == "decimalZero":
        return f"{value:02d}"
    return str(value)


def _roman(value: int) -> str:
    numerals = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
    result = ""
    for number, numeral in numerals:
        count, value = divmod(value, number)
        result += numeral * count
    return result


def _bullet(text: str) -> str:
    """Keep a real Unicode bullet ("●", "■", "✓", as docx-js and Word's Unicode lists
    write); a Symbol/Wingdings code becomes the round bullet. Those are private-use
    characters (U+F0xx) or Latin-1 letters ("§", "Ø" are bullet shapes in Wingdings)."""
    if text in ("-", "o", "*", "·") or (len(text) == 1 and ord(text) > 0x2000 and not 0xE000 <= ord(text) <= 0xF8FF):
        return text
    return BULLET


def _level(lvl: _Element) -> _Level:
    return _Level(
        start=int(attr(child(lvl, "w:start")) or 1),
        format=attr(child(lvl, "w:numFmt")) or "decimal",
        text=attr(child(lvl, "w:lvlText")) or "",
        paragraph_properties=child(lvl, "w:pPr"),
        run_properties=child(lvl, "w:rPr"),
    )
