"""Replace {{text:key}} placeholders while preserving run-level styling.

Runs are edited in place (only their text changes), so font family, size,
weight, color, alignment and box geometry from the template are untouched.

The renderer walks a text frame with the SAME text model the parser uses
(``text_frame.text``): runs, ``<a:br>`` (\\v), ``<a:fld>`` cached text and
paragraph separators (\\n). A placeholder therefore renders even when
PowerPoint physically splits it, e.g.::

    {{text:  ->  run "{{text:" + <a:br> or next paragraph + run "key}}"

A line break covered by a placeholder is always accidental, so it is
removed; paragraphs after the first are merged into it (the first
paragraph's alignment/bullet wins). Placeholder text found inside an
<a:fld> is re-created as a plain run; the field element is dropped.

When a placeholder is split across multiple runs, the replacement value
inherits the formatting of the run where it starts. {{image:key}} tokens
are NOT touched here; the image renderer owns them. Text inside
<a:hyperlink> is invisible to python-pptx's .text (and to PowerPoint's
plain-text model), so it is neither parsed nor rendered.
"""
from pptx.oxml.ns import qn

from .template_parser import VALID_PATTERN

A_R = qn("a:r")
A_T = qn("a:t")
A_BR = qn("a:br")
A_FLD = qn("a:fld")
A_P = qn("a:p")
A_PPR = qn("a:pPr")


def _iter_frame_nodes(text_frame):
    """Yield (kind, element, text) for every piece ``text_frame.text`` is
    built from: 'run' | 'break' (\\v) | 'field' (cached a:t) | 'para' (\\n
    before the referenced a:p). Concatenation equals text_frame.text."""
    for i, para in enumerate(text_frame.paragraphs):
        if i:
            yield "para", para._p, "\n"
        for child in para._p:
            if child.tag == A_R:
                yield "run", child, child.findtext(A_T) or ""
            elif child.tag == A_BR:
                yield "break", child, "\v"
            elif child.tag == A_FLD:
                yield "field", child, child.findtext(A_T) or ""


def _make_run(parent, text):
    run = parent.makeelement(A_R, {})
    t = parent.makeelement(A_T, {})
    t.text = text
    run.append(t)
    return run


def _set_run_text(run_el, text):
    t = run_el.find(A_T)
    if t is None:
        t = run_el.makeelement(A_T, {})
        r_pr = run_el.find(qn("a:rPr"))
        run_el.insert(0 if r_pr is None else 1, t)
    t.text = text


def _merge_paragraph_into_prev(p_el):
    """Move p_el's contents into the previous paragraph and drop p_el."""
    prev = p_el.getprevious()
    while prev is not None and prev.tag != A_P:
        prev = prev.getprevious()
    if prev is None:  # no paragraph before it: nothing to merge into
        return
    for child in list(p_el):
        if child.tag != A_PPR:
            prev.append(child)
    p_el.getparent().remove(p_el)


def _node_text(kind, el):
    """Current text of a node (re-read every use: earlier right-to-left
    replacements may have changed it)."""
    if kind == "run" or kind == "field":
        return el.findtext(A_T) or ""
    return "\v" if kind == "break" else "\n"


def _render_frame(frame, data):
    """Render all {{text:key}} in one text frame, single right-to-left
    pass so earlier match offsets stay valid."""
    nodes = list(_iter_frame_nodes(frame))
    full_text = "".join(text for _, _, text in nodes)
    matches = [
        m for m in VALID_PATTERN.finditer(full_text) if m.group(1) == "text"
    ]
    if not matches:
        return

    starts = [0]
    for _, _, text in nodes:
        starts.append(starts[-1] + len(text))

    for m in reversed(matches):
        key = m.group(2)
        if key not in data:
            continue  # leave unknown keys untouched for verification
        value = str(data[key])
        start, end = m.start(), m.end()

        first = next(i for i in range(len(nodes)) if starts[i] <= start < starts[i + 1])
        last = next(i for i in range(len(nodes)) if starts[i] < end <= starts[i + 1])

        kind_f, el_f, _ = nodes[first]
        kind_l, el_l, _ = nodes[last]
        prefix = _node_text(kind_f, el_f)[: start - starts[first]]
        suffix = _node_text(kind_l, el_l)[end - starts[last]:]

        if first == last:
            if kind_f == "run":
                _set_run_text(el_f, prefix + value + suffix)
            else:  # whole token inside a field element
                el_f.addprevious(_make_run(el_f.getparent(), prefix + value + suffix))
                el_f.getparent().remove(el_f)
            continue

        # Fully covered middle nodes: clear runs, remove breaks/fields,
        # merge later paragraphs into the previous one (left-to-right chains
        # correctly because a merged paragraph's new prev is the target).
        for i in range(first + 1, last):
            kind, el, _ = nodes[i]
            if kind == "run":
                _set_run_text(el, "")
            elif kind == "para":
                _merge_paragraph_into_prev(el)
            else:  # break / field swallowed by the placeholder
                el.getparent().remove(el)

        if kind_f == "run":
            _set_run_text(el_f, prefix + value)
        elif kind_f == "field":  # token starts inside a field
            el_f.addprevious(_make_run(el_f.getparent(), prefix + value))
            el_f.getparent().remove(el_f)

        if kind_l == "run":
            _set_run_text(el_l, suffix)
        elif kind_l == "field":  # token ends inside a field
            if suffix:
                el_l.addprevious(_make_run(el_l.getparent(), suffix))
            el_l.getparent().remove(el_l)


def render_text(prs, data):
    """Replace all {{text:key}} placeholders, shapes and table cells alike.

    Returns the list of keys that were actually replaced.
    """
    from .template_parser import iter_cell_text_frames, iter_text_shapes

    def text_keys(frame):
        return {
            m.group(2)
            for m in VALID_PATTERN.finditer(frame.text)
            if m.group(1) == "text"
        }

    replaced = []
    for slide in prs.slides:
        frames = [shape.text_frame for shape in iter_text_shapes(slide.shapes)]
        frames += list(iter_cell_text_frames(slide.shapes))
        for frame in frames:
            before = text_keys(frame)
            _render_frame(frame, data)
            replaced.extend(before - text_keys(frame))
    return replaced
