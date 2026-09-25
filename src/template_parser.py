"""Scan a PPTX for explicit {{text:key}} / {{image:key}} placeholders.

The placeholder itself declares the data type. Type is never inferred from
the value. Malformed or unknown placeholder tokens (e.g. ``{{title}}``,
``{{file:x}}``, ``{{image}}``) are collected and reported, not guessed.
"""
import re

VALID_PATTERN = re.compile(r"\{\{\s*(text|image)\s*:\s*([A-Za-z0-9_]+)\s*\}\}")
ANY_TOKEN_PATTERN = re.compile(r"\{\{.*?\}\}", re.DOTALL)


class MalformedPlaceholderError(ValueError):
    """Raised when a template contains a placeholder that is not
    {{text:key}} or {{image:key}}."""

    def __init__(self, offenders):
        self.offenders = offenders
        shown = "; ".join(sorted(set(offenders)))
        super().__init__(
            "Malformed placeholder(s) (must be {{{{text:key}}}} or "
            f"{{{{image:key}}}}): {shown}"
        )


def extract_placeholders(text):
    """Return [(type, key), ...] for valid placeholders found in text."""
    return [(m.group(1), m.group(2)) for m in VALID_PATTERN.finditer(text)]


def find_malformed(text):
    """Return raw {{...}} tokens that are not valid typed placeholders."""
    return [
        token
        for token in ANY_TOKEN_PATTERN.findall(text)
        if not VALID_PATTERN.fullmatch(token)
    ]


def iter_text_shapes(shapes):
    """Yield shapes that carry a text frame, recursing into groups."""
    for shape in shapes:
        if shape.shape_type == 6:  # GROUP
            yield from iter_text_shapes(shape.shapes)
        elif shape.has_text_frame:
            yield shape


def iter_tables(shapes):
    """Yield tables from graphic frames, recursing into groups."""
    for shape in shapes:
        if shape.shape_type == 6:  # GROUP
            yield from iter_tables(shape.shapes)
        elif getattr(shape, "has_table", False):
            yield shape.table


def iter_cell_text_frames(shapes):
    """Yield the text frame of every table cell, recursing into groups."""
    for table in iter_tables(shapes):
        for row in table.rows:
            for cell in row.cells:
                yield cell.text_frame


class TextPlaceholder:
    """A shape or table cell containing {{text:key}} placeholders."""

    def __init__(self, slide_index, shape, keys):
        self.slide_index = slide_index
        self.shape = shape
        self.keys = keys


class ImagePlaceholder:
    """A shape whose entire text is exactly one {{image:key}}: the image region."""

    def __init__(self, slide_index, shape, key, region):
        self.slide_index = slide_index
        self.shape = shape
        self.key = key
        # (left, top, width, height) in EMU
        self.region = region


def parse_slides(prs, data):
    """Find all text and image placeholders in a presentation.

    A shape is an IMAGE placeholder when its whole text is exactly one
    {{image:key}} token — the key's value is never inspected. Anything
    containing {{text:key}} is a text placeholder (including a text token
    whose value happens to be an image path).

    Table cells are scanned for {{text:key}} placeholders too. An
    {{image:key}} token inside a cell cannot become a region (pictures
    cannot be embedded in table cells); it simply remains unchanged.

    Returns (text_placeholders, image_placeholders, malformed_tokens).
    """
    text_placeholders = []
    image_placeholders = []
    malformed = []

    for slide_index, slide in enumerate(prs.slides):
        for shape in iter_text_shapes(slide.shapes):
            full_text = shape.text_frame.text
            malformed.extend(find_malformed(full_text))

            single_image = re.fullmatch(
                r"\s*\{\{\s*image\s*:\s*([A-Za-z0-9_]+)\s*\}\}\s*", full_text
            )
            if single_image:
                region = (shape.left, shape.top, shape.width, shape.height)
                image_placeholders.append(
                    ImagePlaceholder(slide_index, shape, single_image.group(1), region)
                )
                continue

            keys = [k for t, k in extract_placeholders(full_text) if t == "text"]
            if keys:
                text_placeholders.append(TextPlaceholder(slide_index, shape, keys))

        for table in iter_tables(slide.shapes):
            for row in table.rows:
                for cell in row.cells:
                    full_text = cell.text_frame.text
                    malformed.extend(find_malformed(full_text))
                    keys = [
                        k for t, k in extract_placeholders(full_text) if t == "text"
                    ]
                    if keys:
                        text_placeholders.append(
                            TextPlaceholder(slide_index, cell, keys)
                        )

    return text_placeholders, image_placeholders, malformed
