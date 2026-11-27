"""Word's "keep with next": a paragraph stays on the page where the next one starts.

reportlab groups such flowables in a KeepTogether, which moves the whole group to a
new page when it doesn't fit: a heading followed by a 45-row table, or by a long
paragraph, leaves the rest of the page empty (docs/DESIGN.md). Word needs
only the first line or row of the next item on the same page, so `KeptWithNext`
splits the next item in place and moves on only when not even that fits.
"""

from __future__ import annotations

from reportlab.platypus import Flowable
from reportlab.platypus.doctemplate import ActionFlowable


class KeptWithNext(Flowable):
    def __init__(self, first: Flowable, following: Flowable) -> None:
        super().__init__()
        self.first, self.following = first, following
        self.keepWithNext = following.getKeepWithNext()
        self._sizes = ((0.0, 0.0), (0.0, 0.0))

    @property
    def _gap(self) -> float:
        return self.first.getSpaceAfter() + self.following.getSpaceBefore()

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        first = self.first.wrap(availWidth, availHeight)
        following = self.following.wrap(availWidth, availHeight - first[1] - self._gap)
        self._sizes = (first, following)
        self.width, self.height = availWidth, first[1] + self._gap + following[1]
        return self.width, self.height

    def split(self, availWidth: float, availHeight: float) -> list[Flowable]:
        _, height = self.first.wrap(availWidth, availHeight)
        room = availHeight - height - self._gap
        parts = self.following.split(availWidth, room) if room > 0 else []
        if parts:
            return [self.first, *parts]
        frame = getattr(self, "_frame", None)
        if frame is not None and getattr(frame, "_atTop", False):
            return [self.first, self.following]  # can't be kept even on a fresh page: let go
        return []  # moves the pair to the next page

    def getSpaceBefore(self) -> float:
        return self.first.getSpaceBefore()

    def getSpaceAfter(self) -> float:
        return self.following.getSpaceAfter()

    def draw(self) -> None:
        (first_width, first_height), (following_width, following_height) = self._sizes
        self.first.drawOn(self.canv, 0, self.height - first_height, _sW=self.width - first_width)
        self.following.drawOn(self.canv, 0, self.height - first_height - self._gap - following_height,
                              _sW=self.width - following_width)


def keep_with_next(story: list[Flowable]) -> list[Flowable]:
    """Pair each keep-with-next flowable with the one after it (chains nest)."""
    kept: list[Flowable] = []
    for flowable in reversed(story):
        if (kept and flowable.getKeepWithNext() and not isinstance(flowable, ActionFlowable)
                and not isinstance(kept[-1], ActionFlowable)):
            flowable.keepWithNext = 0  # handled here, not by reportlab's KeepTogether
            kept[-1] = KeptWithNext(flowable, kept[-1])
        else:
            kept.append(flowable)
    return kept[::-1]
