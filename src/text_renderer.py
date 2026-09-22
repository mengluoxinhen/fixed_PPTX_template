"""Replace {{text:key}} placeholders while preserving run-level styling.

Runs are edited in place (only their text changes), so font family, size,
weight, color, alignment and box geometry from the template are untouched.

When a placeholder is split across multiple runs by PowerPoint, the
replacement value inherits the formatting of the run where it starts.
{{image:key}} tokens are NOT touched here; the image renderer owns them.
"""
from .template_parser import VALID_PATTERN


def _render_paragraph(paragraph, data):
    runs = paragraph.runs
    if not runs:
        return

    full_text = "".join(r.text for r in runs)
    matches = [
        m for m in VALID_PATTERN.finditer(full_text) if m.group(1) == "text"
    ]
    if not matches:
        return

    # Run boundaries: run i covers [starts[i], starts[i+1])
    starts = [0]
    for r in runs:
        starts.append(starts[-1] + len(r.text))

    # Process right-to-left so earlier match offsets stay valid.
    for m in reversed(matches):
        key = m.group(2)
        if key not in data:
            continue  # leave unknown keys untouched for verification
        value = str(data[key])
        start, end = m.start(), m.end()

        first = next(i for i in range(len(runs)) if starts[i] <= start < starts[i + 1])
        last = next(i for i in range(len(runs)) if starts[i] < end <= starts[i + 1])

        prefix = runs[first].text[: start - starts[first]]
        suffix = runs[last].text[end - starts[last]:]

        if first == last:
            runs[first].text = prefix + value + suffix
        else:
            runs[first].text = prefix + value
            for i in range(first + 1, last):
                runs[i].text = ""
            runs[last].text = suffix


def render_text(prs, data):
    """Replace all {{text:key}} placeholders in the presentation.

    Returns the list of keys that were actually replaced.
    """
    from .template_parser import iter_text_shapes

    replaced = []
    for slide in prs.slides:
        for shape in iter_text_shapes(slide.shapes):
            for paragraph in shape.text_frame.paragraphs:
                before = {
                    m.group(2)
                    for m in VALID_PATTERN.finditer(
                        "".join(r.text for r in paragraph.runs)
                    )
                    if m.group(1) == "text"
                }
                _render_paragraph(paragraph, data)
                after = {
                    m.group(2)
                    for m in VALID_PATTERN.finditer(
                        "".join(r.text for r in paragraph.runs)
                    )
                    if m.group(1) == "text"
                }
                replaced.extend(before - after)
    return replaced
