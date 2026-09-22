"""Automated verification of the render workflow and placeholder rules.

    python verify.py

Part 1 — end-to-end workflow: template/output existence, 3 slides, values
rendered, no leftover {{...}} placeholders, image inserted, unchanged slide
size, zip validity.

Part 2 — placeholder semantics tests (in-memory templates):
  1. {{text:key}} renders text
  2. {{image:key}} renders an image
  3. a .png value via {{text:key}} stays plain text (no type inference)
  4. malformed/unknown placeholder types are rejected
  5. no valid {{text:}}/{{image:}} placeholders remain in the output
  6. cover-crop behavior preserved
  7. text run styling preserved (single-run, inline, and split-across-runs)
"""
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.renderer import render  # noqa: E402
from src.template_parser import MalformedPlaceholderError  # noqa: E402

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "templates" / "sales_report_template.pptx"
OUTPUT = ROOT / "output" / "sales_report.pptx"

EXPECTED_TEXT = [
    "September 2026 Sales Report",
    "East China Regional Performance",
    "Example Technology Co., Ltd.",
    "¥12,580,000",
    "+18.5%",
    "12,580",
    "3,280",
    "Smart Office Terminal X1",
    "A next-generation smart terminal designed for enterprise office scenarios.",
    "¥2,580",
    "+23.6%",
    "Key Business Metrics",
    "Total reported sales: ¥12,580,000 year to date.",
]

ANY_TOKEN = re.compile(r"\{\{.*?\}\}", re.DOTALL)

failures = []


def check(name, ok, detail=""):
    status = "PASS" if ok else "FAIL"
    print(f"[{status}] {name}" + (f" — {detail}" if detail else ""))
    if not ok:
        failures.append(name)


def all_text(prs):
    return "\n".join(
        shape.text_frame.text
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_text_frame
    )


def count_pictures(prs):
    return [
        shape
        for slide in prs.slides
        for shape in slide.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]


def make_single_shape_template(text, style_runs=None):
    """Build a 1-slide template with one text box holding `text`.

    style_runs: optional list of (text, size, bold, color) runs for the
    first paragraph instead of a single default run.
    """
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    p = box.text_frame.paragraphs[0]
    if style_runs:
        for run_text, size, bold, color in style_runs:
            r = p.add_run()
            r.text = run_text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
    else:
        r = p.add_run()
        r.text = text
    return prs


def render_template(prs, data):
    """Render an in-memory template and return the parsed output."""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.pptx"
        src = Path(tmp) / "tpl.pptx"
        prs.save(str(src))
        render(str(src), data, str(out), base_dirs=(str(ROOT),))
        return Presentation(str(out))


def workflow_checks():
    check("template exists", TEMPLATE.exists(), str(TEMPLATE))
    check("output exists", OUTPUT.exists(), str(OUTPUT))
    if not (TEMPLATE.exists() and OUTPUT.exists()):
        return False

    tpl = Presentation(str(TEMPLATE))
    prs = Presentation(str(OUTPUT))
    text = all_text(prs)

    check("output parses as PPTX", True)
    check("3 slides present", len(prs.slides._sldIdLst) == 3,
          f"{len(prs.slides._sldIdLst)} slides")

    missing = [t for t in EXPECTED_TEXT if t not in text]
    check("all expected values rendered", not missing, "; ".join(missing))

    leftovers = sorted(set(ANY_TOKEN.findall(text)))
    check("no {{...}} placeholders remain (check 5)", not leftovers,
          "; ".join(leftovers))

    pictures = count_pictures(prs)
    check("product image inserted (check 2)", len(pictures) >= 1,
          f"{len(pictures)} picture(s)")
    if pictures:
        pic = pictures[0]
        check("image within slide bounds",
              pic.left >= 0 and pic.top >= 0
              and pic.left + pic.width <= prs.slide_width
              and pic.top + pic.height <= prs.slide_height,
              f"width={pic.width / 914400:.2f}in")

    check("slide size unchanged",
          (tpl.slide_width, tpl.slide_height) == (prs.slide_width, prs.slide_height),
          f"{prs.slide_width}x{prs.slide_height} EMU")

    with zipfile.ZipFile(str(OUTPUT)) as zf:
        bad = zf.testzip()
        media = [n for n in zf.namelist() if n.startswith("ppt/media/")]
    check("zip structurally valid", bad is None, f"corrupt entry: {bad}")
    check("image embedded in package", any(n.endswith(".png") for n in media),
          ", ".join(media))
    return True


def semantic_checks():
    data = {"sales": "¥12,580,000", "product_image": "assets/product.png"}

    # 1. {{text:key}} renders text
    out = render_template(make_single_shape_template("{{text:sales}}"), data)
    ok = "¥12,580,000" in all_text(out)
    check("check 1: {{text:key}} replaced with value", ok)

    # 2. {{image:key}} creates an image region
    tpl = make_single_shape_template("{{image:product_image}}")
    box = tpl.slides[0].shapes[0]
    box.left, box.top = Inches(2), Inches(2)
    box.width, box.height = Inches(2), Inches(2)
    out = render_template(tpl, data)
    pics = count_pictures(out)
    check("check 2: {{image:key}} inserts a picture", len(pics) == 1)
    if pics:
        pic = pics[0]
        region_ok = (pic.left, pic.top, pic.width, pic.height) == (
            Inches(2), Inches(2), Inches(2), Inches(2)
        )
        check("image occupies the exact placeholder region", region_ok)
        # 6. cover-crop: 1200x1040 image into a 1:1 square region is wider
        # than the region -> centered left/right crop, no vertical crop.
        crop_x = 1.0 - 1.0 / (1200 / 1040)
        expected_side = crop_x / 2
        approx = abs(pic.crop_left - expected_side) < 1e-4
        symmetric = (abs(pic.crop_left - pic.crop_right) < 1e-9
                     and pic.crop_top == 0.0 and pic.crop_bottom == 0.0)
        check("check 6: cover crop applied, centered, no overflow",
              approx and symmetric,
              f"crop_left={pic.crop_left:.5f} expected≈{expected_side:.5f}")

    # 3. .png value via {{text:key}} is plain text, NOT an image
    out = render_template(make_single_shape_template("{{text:product_image}}"), data)
    as_text = all_text(out).strip() == "assets/product.png"
    check("check 3: .png value rendered as TEXT via {{text:key}}", as_text,
          repr(all_text(out).strip()))
    check("check 3: no image created for text-typed .png value",
          len(count_pictures(out)) == 0)

    # 4. malformed / unknown placeholder types are rejected
    for label, token in [
        ("untyped {{title}}", "{{title}}"),
        ("unknown type {{file:product}}", "{{file:product}}"),
        ("incomplete {{image}}", "{{image}}"),
    ]:
        try:
            render_template(make_single_shape_template(token), data)
            check(f"check 4: rejects {label}", False, "no error raised")
        except MalformedPlaceholderError:
            check(f"check 4: rejects {label}", True)

    # 7. text styling preserved (single-run, inline, split-across-runs)
    tpl = make_single_shape_template(
        None,
        style_runs=[("{{text:sales}}", 20, True, RGBColor(0xFF, 0x00, 0x00))],
    )
    out = render_template(tpl, data)
    run = out.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    check("check 7: single-run style preserved (20pt bold red)",
          run.text == "¥12,580,000" and run.font.size.pt == 20
          and run.font.bold and str(run.font.color.rgb) == "FF0000",
          f"{run.text!r} {run.font.size.pt}pt bold={run.font.bold} {run.font.color.rgb}")

    tpl = make_single_shape_template(
        None,
        style_runs=[
            ("Sales: {{text:", 14, False, RGBColor(0x11, 0x22, 0x33)),
            ("sales", 14, False, RGBColor(0x11, 0x22, 0x33)),
            ("}} year", 14, False, RGBColor(0x11, 0x22, 0x33)),
        ],
    )
    out = render_template(tpl, data)
    para = out.slides[0].shapes[0].text_frame.paragraphs[0]
    joined = "".join(r.text for r in para.runs)
    first = para.runs[0]
    check("check 7: inline variable split across runs replaced correctly",
          joined == "Sales: ¥12,580,000 year", repr(joined))
    check("check 7: value inherits starting run style",
          first.font.size.pt == 14 and str(first.font.color.rgb) == "112233")


def main():
    print("== End-to-end workflow ==")
    if workflow_checks():
        print("\n== Placeholder semantics ==")
        semantic_checks()
    else:
        failures.append("workflow incomplete")

    print()
    if failures:
        print(f"VERIFICATION FAILED: {len(failures)} check(s) failed")
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
