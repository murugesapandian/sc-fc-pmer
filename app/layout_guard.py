"""Keeps generated content inside its shape's footprint so a growing
workbook (more tasks, more RAID items, longer descriptions) never pushes
text or table rows into a neighboring shape, footer, or image.

The template's shape geometry is the source of truth: for any text/table
shape we find the nearest sibling shape positioned below it (within the same
horizontal band) and treat that sibling's top edge as a hard floor. Content
is capped to what actually fits above that floor, with a "+N more" note
appended when data has to be truncated, rather than letting PowerPoint
silently overflow one shape into another.
"""
from __future__ import annotations

from pptx.util import Emu, Pt

_MIN_GAP = Pt(6)
_LINE_HEIGHT_FACTOR = 1.25
_DEFAULT_FONT_PT = 14
_MAX_CHARS_DEFAULT = 220


def _x_overlap(a, b) -> bool:
    return a.left < (b.left + b.width) and b.left < (a.left + a.width)


def available_height_below(slide, shape) -> int:
    """Vertical budget, in EMU, from `shape`'s top down to the nearest
    sibling below it that overlaps its horizontal range — text box, table,
    picture, or any other shape type. Falls back to the shape's own current
    height (i.e. don't grow past what the template already allotted) when no
    such sibling is found, or when this shape's own geometry is unset
    (inherited from the layout rather than placed explicitly)."""
    if shape.top is None or shape.left is None or shape.width is None or shape.height is None:
        return shape.height or Emu(0)

    floor = None
    for other in slide.shapes:
        if other.shape_id == shape.shape_id:
            continue
        if other.top is None or other.left is None or other.width is None or other.height is None:
            continue
        if other.top <= shape.top:
            continue
        if not _x_overlap(shape, other):
            continue
        if floor is None or other.top < floor:
            floor = other.top
    if floor is None:
        return shape.height
    return max(shape.height, floor - shape.top - _MIN_GAP)


def _first_font_size_pt(text_frame) -> float:
    for p in text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                return r.font.size.pt
    return _DEFAULT_FONT_PT


def _chars_per_line(shape, font_pt: float) -> int:
    """Rough estimate of how many characters fit on one wrapped visual line
    of this shape, given its width and font size — needed because a long
    bullet doesn't cost 1 line, it costs however many lines word-wrap turns
    it into, and undercounting that is exactly what let long text overflow."""
    width_pt = (shape.width / 12700) if shape.width else 400
    usable_width_pt = max(width_pt - 14, 20)  # ~0.1in left/right text insets
    avg_char_width_pt = max(font_pt * 0.52, 1)  # typical proportional-font average
    return max(10, int(usable_width_pt / avg_char_width_pt))


def _visual_lines(text: str, chars_per_line: int) -> int:
    if not text:
        return 1
    return max(1, -(-len(text) // chars_per_line))  # ceil division


def cap_lines_to_fit(shape, lines: list[str], available_height_emu: int,
                      more_note: str = "more — see the workbook for the full list") -> list[str]:
    """Greedily keeps as many `lines` as fit within `available_height_emu`,
    accounting for word-wrap (each line's *wrapped* height, not just its
    paragraph count) so long bullets can't quietly blow the budget."""
    font_pt = _first_font_size_pt(shape.text_frame)
    line_height_emu = Pt(font_pt * _LINE_HEIGHT_FACTOR)
    budget_lines = max(1, int(available_height_emu // line_height_emu))
    chars_per_line = _chars_per_line(shape, font_pt)

    visible: list[str] = []
    used = 0
    for line in lines:
        cost = _visual_lines(line, chars_per_line)
        if visible and used + cost > budget_lines:
            hidden = len(lines) - len(visible)
            visible.append(f"+{hidden} {more_note}")
            return visible
        visible.append(line)
        used += cost
    return visible


def max_rows_for(available_height_emu: int, header_height_emu: int, row_height_emu: int) -> int:
    usable = available_height_emu - header_height_emu
    if row_height_emu <= 0:
        return 10 ** 6  # unknown row height: don't guess, let caller skip capping
    return max(1, int(usable // row_height_emu))


def cap_rows(rows: list[list], max_rows: int, note_row_factory) -> list[list]:
    """note_row_factory(hidden_count) -> a row (list of cell values) summarizing
    the truncated remainder, shaped to match the table's column count."""
    if len(rows) <= max_rows:
        return rows
    if max_rows <= 1:
        return [note_row_factory(len(rows))]
    visible = rows[: max_rows - 1]
    hidden = len(rows) - len(visible)
    visible.append(note_row_factory(hidden))
    return visible


def truncate_chars(text: str, max_len: int = _MAX_CHARS_DEFAULT) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"
