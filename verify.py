"""Automated verification of the render workflow and placeholder rules.

    python verify.py

Part 1 — end-to-end workflow (direct --data mode): template/output existence,
4 slides, values rendered, only the missing asset's placeholder survives,
5 images inserted, unchanged slide size, zip validity.

Part 2 — placeholder semantics tests (in-memory templates):
  1. {{text:key}} renders text
  2. {{image:key}} renders an image
  3. a .png value via {{text:key}} stays plain text (no type inference)
  4. malformed/unknown placeholder types are rejected
  5. no leftover placeholders except intentionally missing assets
  6. cover-crop behavior preserved
  7. text run styling preserved (single-run, inline, and split-across-runs)

Part 3 — repeated placeholders: same (type, key) across slides/shapes/
paragraphs; per-occurrence regions and crops.

Part 4 — UUID workspace: validation, content.json loading, stem-based image
resolution (.png/.jpg/.jpeg/.webp), missing-asset survival, path safety,
explicit cleanup, and the workspace-mode output file.
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
DIRECT_OUTPUT = ROOT / "output" / "direct_mode.pptx"
WORKSPACE_OUTPUT = ROOT / "output" / "workspace_mode.pptx"
WORKSPACE_DIR = ROOT / "workspace" / "550e8400-e29b-41d4-a716-446655440000"

MISSING_IMAGE_TOKEN = "{{image:missing_image}}"

EXPECTED_TEXT = [
    "September 2026 Sales Report",
    "Report: September 2026 Sales Report",
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


def make_boxes_template(slide_texts):
    """Build a template with one text box per entry in slide_texts.

    slide_texts: list (one per slide) of lists of strings. Boxes are laid
    out at 2.2in horizontal steps so sizes never overlap.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for texts in slide_texts:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for i, t in enumerate(texts):
            box = slide.shapes.add_textbox(
                Inches(0.5 + i * 2.2), Inches(1), Inches(2), Inches(1)
            )
            r = box.text_frame.paragraphs[0].add_run()
            r.text = t
    return prs


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


def find_picture(pictures, l, t, w, h):
    """Return the picture whose bounds equal the given region (inches), or None."""
    want = (Inches(l), Inches(t), Inches(w), Inches(h))
    return next((p for p in pictures
                 if (p.left, p.top, p.width, p.height) == want), None)


def workflow_checks():
    """End-to-end checks for direct --data mode output."""
    check("template exists", TEMPLATE.exists(), str(TEMPLATE))
    check("direct-mode output exists (test 14: --data mode works)",
          DIRECT_OUTPUT.exists(), str(DIRECT_OUTPUT))
    if not (TEMPLATE.exists() and DIRECT_OUTPUT.exists()):
        return False

    tpl = Presentation(str(TEMPLATE))
    prs = Presentation(str(DIRECT_OUTPUT))
    text = all_text(prs)

    check("output parses as PPTX", True)
    check("4 slides present", len(prs.slides._sldIdLst) == 4,
          f"{len(prs.slides._sldIdLst)} slides")

    missing = [t for t in EXPECTED_TEXT if t not in text]
    check("all expected values rendered", not missing, "; ".join(missing))

    leftovers = sorted(set(ANY_TOKEN.findall(text)))
    check("only the missing asset's placeholder survives in direct mode",
          leftovers == [MISSING_IMAGE_TOKEN], "; ".join(leftovers))
    check("W10/§5: {{image:missing_image}} left unchanged",
          MISSING_IMAGE_TOKEN in text)

    check("R1/R5: repeated {{{{text:title}}}} rendered on all 3 slides",
          text.count("September 2026 Sales Report") == 3,
          f"{text.count('September 2026 Sales Report')} occurrences")

    pictures = count_pictures(prs)
    # product_image x2 (slides 1+3) + product x2 + logo x1 (slide 4)
    check("R3/R7: all 5 image occurrences inserted", len(pictures) == 5,
          f"{len(pictures)} picture(s)")
    if len(pictures) == 5:
        check("R3: slide-1 square image region honored exactly",
              find_picture(pictures, 10.9, 4.95, 1.6, 1.6) is not None)
        check("R3: slide-3 landscape image region honored exactly",
              find_picture(pictures, 7.1, 1.9, 5.3, 4.6) is not None)
        sq = find_picture(pictures, 0.9, 1.9, 2.6, 2.6)
        tall = find_picture(pictures, 7.9, 1.9, 1.2, 3.4)
        logo = find_picture(pictures, 3.9, 1.9, 3.6, 1.2)
        check("W: slide-4 stem-key regions filled (product x2 + logo)",
              sq is not None and tall is not None and logo is not None)
        cover_pic = find_picture(pictures, 10.9, 4.95, 1.6, 1.6)
        product_pic = find_picture(pictures, 7.1, 1.9, 5.3, 4.6)
        if cover_pic and product_pic:
            # 1200x1040 image: square 1.6x1.6 region -> side crop (≈0.0667);
            # 5.3x4.6 region -> tiny side crop (≈0.0007). Per-occurrence.
            check("R3: same key, per-occurrence independent crops",
                  cover_pic.crop_left > 0.05 and cover_pic.crop_top == 0
                  and 0 < product_pic.crop_left < 0.01,
                  f"cover crop_left={cover_pic.crop_left:.4f}, "
                  f"product crop_left={product_pic.crop_left:.4f}")
        if sq and tall and logo:
            check("T16: crop direction per occurrence on slide 4",
                  sq.crop_left > 0 and sq.crop_top == 0
                  and tall.crop_left > sq.crop_left
                  and logo.crop_top > 0 and logo.crop_left == 0,
                  f"square={sq.crop_left:.4f} tall={tall.crop_left:.4f} "
                  f"logo crop_top={logo.crop_top:.4f}")

    check("slide size unchanged",
          (tpl.slide_width, tpl.slide_height) == (prs.slide_width, prs.slide_height),
          f"{prs.slide_width}x{prs.slide_height} EMU")

    with zipfile.ZipFile(str(DIRECT_OUTPUT)) as zf:
        bad = zf.testzip()
        media = [n for n in zf.namelist() if n.startswith("ppt/media/")]
    check("zip structurally valid", bad is None, f"corrupt entry: {bad}")
    check("png + jpg embedded in package",
          any(n.endswith(".png") for n in media)
          and any(n.endswith(".jpg") for n in media),
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


def repetition_checks():
    from src.template_parser import parse_slides

    title = "2026年销售报告"
    data = {"title": title, "product_image": "assets/product.png"}

    # R1 + R5: same text placeholder repeated across slides
    tpl = make_boxes_template([
        ["{{text:title}}"],
        [f"Report: {{{{text:title}}}}"],
        ["{{text:title}}", "{{text:title}} | {{text:title}}"],
    ])
    out = render_template(tpl, data)
    out_text = all_text(out)
    # 1 (slide1) + 1 inline (slide2) + 1 + 2 (slide3, incl. same-paragraph repeat)
    check("R1/R5: text key repeated across slides all rendered",
          out_text.count(title) == 5 and "{{text:title}}" not in out_text,
          f"{out_text.count(title)}/5 occurrences")
    check("R2: same-paragraph repeat renders both copies",
          list(out.slides)[2].shapes[1].text_frame.text == f"{title} | {title}")
    check("R6: no valid placeholders remain in repeated-render output",
          not ANY_TOKEN.search(out_text))

    # R3: same image placeholder on multiple slides, different sizes
    tpl = make_boxes_template([["{{image:product_image}}"], ["{{image:product_image}}"]])
    s0, s1 = list(tpl.slides)
    box0, box1 = list(s0.shapes)[0], list(s1.shapes)[0]
    box0.left, box0.top, box0.width, box0.height = Inches(1), Inches(2), Inches(4), Inches(1)
    box1.left, box1.top, box1.width, box1.height = Inches(7), Inches(3), Inches(2), Inches(2)
    out = render_template(tpl, data)
    pics = count_pictures(out)
    check("R3: one image key on 2 slides -> 2 pictures", len(pics) == 2)
    if len(pics) == 2:
        check("R3: each occurrence keeps its own region",
              (pics[0].left, pics[0].width, pics[0].height) ==
              (Inches(1), Inches(4), Inches(1))
              and (pics[1].left, pics[1].width, pics[1].height) ==
              (Inches(7), Inches(2), Inches(2)))
        # 4:1 region crops top/bottom; 1:1 region crops sides
        check("R9: cover crop computed per occurrence region",
              pics[0].crop_top > 0 and pics[0].crop_left == 0
              and pics[1].crop_left > 0 and pics[1].crop_top == 0,
              f"wide crop_top={pics[0].crop_top:.4f}, "
              f"square crop_left={pics[1].crop_left:.4f}")

    # R4 / binding identity (type, key): same key as text AND as image
    tpl = make_boxes_template([["{{text:product_image}}", "{{image:product_image}}"]])
    box = list(tpl.slides)[0].shapes[1]
    box.left, box.top, box.width, box.height = Inches(8), Inches(2), Inches(3), Inches(2)
    out = render_template(tpl, data)
    pics = count_pictures(out)
    text = all_text(out)
    check("R4: {{text:product_image}} renders the path as literal text",
          "assets/product.png" in text)
    check("R4: {{image:product_image}} inserts the picture on the same slide",
          len(pics) == 1, f"{len(pics)} picture(s)")

    # R7: parser exposes every occurrence, not just unique keys
    tpl = make_boxes_template([
        ["{{text:title}}"],
        ["{{text:title}}", "{{text:title}}"],
        ["{{image:product_image}}", "{{image:product_image}}"],
    ])
    t_ph, i_ph, malformed = parse_slides(tpl, data)
    text_keys = sorted(k for ph in t_ph for k in ph.keys)
    check("R7: parser returns every text occurrence (3 shapes, 3 keys)",
          len(t_ph) == 3 and text_keys == ["title"] * 3, text_keys)
    check("R7: parser returns every image occurrence as its own region (2)",
          len(i_ph) == 2 and i_ph[0].region != i_ph[1].region
          and [p.key for p in i_ph] == ["product_image"] * 2)
    check("R10: parser still reports no malformed tokens here", malformed == [])


def workspace_checks():
    import json
    import shutil
    import uuid as uuid_mod

    from PIL import Image
    from pptx.enum.shapes import MSO_SHAPE

    from src.workspace_loader import (WorkspaceError, cleanup_workspace,
                                      discover_images, load_workspace,
                                      render_workspace, validate_workspace_dir)

    # W1/T3: valid UUID4 workspace accepted, content.json loaded
    try:
        ws = load_workspace(WORKSPACE_DIR)
    except WorkspaceError as e:
        check("W1: valid UUID4 workspace accepted", False, str(e))
        return
    check("W1/T3: UUID4 workspace accepted, content.json loaded",
          ws["data"].get("title") == "2026年销售报告"
          and ws["data"].get("sales") == "1258万元")
    check("T5: .png asset resolved by stem ({{image:product}})",
          ws["image_paths"].get("product", "").lower().endswith(".png"))
    check("T6/T9: .jpg asset resolved; key carries NO extension",
          ws["image_paths"].get("logo", "").lower().endswith(".jpg")
          and all("." not in k for k in ws["image_paths"]))
    check("SEP: image paths are NOT merged into text data",
          "product" not in ws["data"] and "logo" not in ws["data"])
    check("SEP: text keys are NOT consulted for images",
          "title" not in ws["image_paths"])
    check("Non-image files ignored by discovery",
          "notes_readme" not in ws["image_paths"])
    check("Absent image leaves key out of image_paths (product_image)",
          "product_image" not in ws["image_paths"])
    ws_root = Path(ws["path"])
    check("SEC: every resolved path stays inside the workspace",
          all(Path(p).is_relative_to(ws_root) for p in ws["image_paths"].values()))

    def render_ws_case(tmpdir, content, image_files, template_texts, ext_shapes=None):
        """Create a UUID workspace + in-memory template, render via loader
        glue, and return the parsed output Presentation."""
        d = Path(tmpdir) / str(uuid_mod.uuid4())
        d.mkdir()
        with open(d / "content.json", "w", encoding="utf-8") as f:
            json.dump(content, f, ensure_ascii=False)
        for name, source in image_files.items():
            shutil.copyfile(source, d / name)
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for i, text in enumerate(template_texts):
            if text.startswith("IMG:"):
                shp = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(1), Inches(4), Inches(4), Inches(2))
                shp.text_frame.paragraphs[0].add_run().text = text[4:]
            else:
                box = slide.shapes.add_textbox(
                    Inches(1), Inches(1 + i), Inches(4), Inches(0.8))
                box.text_frame.paragraphs[0].add_run().text = text
        src = Path(tmpdir) / f"tpl_{d.name[:8]}.pptx"
        out = Path(tmpdir) / f"out_{d.name[:8]}.pptx"
        prs.save(str(src))
        render_workspace(str(src), load_workspace(d), str(out))
        return Presentation(str(out))

    # ---- §7 KEY REGRESSION: same stem in BOTH namespaces, no interference
    with tempfile.TemporaryDirectory() as tmp:
        out = render_ws_case(
            tmp,
            content={"logo": "XX科技有限公司"},
            image_files={"logo.jpg": ROOT / "assets" / "logo.jpg"},
            template_texts=["{{text:logo}}", "IMG:{{image:logo}}"],
        )
        text = all_text(out)
        shapes = list(out.slides[0].shapes)
        pics = count_pictures(out)
        check("SEP-7: content.json text value rendered for {{text:logo}}",
              "XX科技有限公司" in text and "{{text:logo}}" not in text)
        check("SEP-7: image file NOT shadowed by same-key text value",
              len(pics) == 1, f"{len(pics)} picture(s)")
        check("SEP-7: {{image:logo}} placeholder consumed, not textified",
              "{{image:logo}}" not in text and "logo.jpg" not in text)
        if pics:
            check("SEP-7: image got its own region + crop (jpg into 4x2 region)",
                  (pics[0].left, pics[0].top, pics[0].width, pics[0].height)
                  == (Inches(1), Inches(4), Inches(4), Inches(2))
                  and pics[0].crop_top > 0 and pics[0].crop_left == 0)

        # ---- §8: {{image:product}} resolves from each supported extension
        for ext in (".png", ".jpg", ".jpeg"):
            asset = Path(tmp) / f"banner{ext}"
            Image.new("RGB", (300, 100), (200, 30, 30)).save(asset)
            out = render_ws_case(
                tmp, content={}, image_files={f"banner{ext}": asset},
                template_texts=["IMG:{{image:banner}}"],
            )
            check(f"EXT-8: {{{{image:banner}}}} rendered from {ext}",
                  len(count_pictures(out)) == 1)

        # ---- §3: only an unsupported .webp exists -> placeholder unchanged,
        #      rendering succeeds without failing on the unsupported file
        webp = Path(tmp) / "only.webp"
        Image.new("RGB", (300, 100), (10, 20, 30)).save(webp, format="WEBP")
        out = render_ws_case(
            tmp, content={}, image_files={"product.webp": webp},
            template_texts=["IMG:{{image:product}}"],
        )
        webp_text = all_text(out)
        check("WEBP-3: unsupported .webp leaves {{image:product}} unchanged",
              "{{image:product}}" in ANY_TOKEN.findall(webp_text)
              and len(count_pictures(out)) == 0,
              "placeholder survived, no picture, no error")

    # W2: invalid workspace names rejected
    with tempfile.TemporaryDirectory() as tmp:
        for bad in ("test", "123", "abc"):
            d = Path(tmp) / bad
            d.mkdir()
            try:
                validate_workspace_dir(d)
                rejected = False
            except WorkspaceError:
                rejected = True
            check(f"W2: workspace name {bad!r} rejected", rejected)
        d = Path(tmp) / str(uuid_mod.uuid1())
        d.mkdir()
        try:
            validate_workspace_dir(d)
            rejected = False
        except WorkspaceError:
            rejected = True
        check("W2: UUIDv1 folder name rejected (must be UUID4)", rejected)

        # T7: .jpeg discovery + deterministic stem conflict rule
        # T4: .webp is UNSUPPORTED and must be ignored by discovery
        wsd = Path(tmp) / str(uuid_mod.uuid4())
        wsd.mkdir()
        (wsd / "content.json").write_text("{}", encoding="utf-8")
        (wsd / "cover.webp").write_bytes(b"x")
        (wsd / "report.jpeg").write_bytes(b"x")
        (wsd / "hero.png").write_bytes(b"x")
        (wsd / "hero.jpg").write_bytes(b"x")
        found = discover_images(wsd)
        check("T4: .webp NOT discovered (unsupported format)",
              "cover" not in found)
        check("T7: .jpeg discovered", found.get("report", "").endswith(".jpeg"))
        check("T7: stem conflict resolves deterministically (png > jpg)",
              found.get("hero", "").endswith(".png"))

        # T13: explicit cleanup + cleanup safety
        cleanup_dir = wsd.parent / str(uuid_mod.uuid4())
        cleanup_dir.mkdir()
        (cleanup_dir / "content.json").write_text("{}", encoding="utf-8")
        (cleanup_dir / "product.png").write_bytes(b"x")
        load_workspace(cleanup_dir)  # sanity: loadable before cleanup
        cleanup_workspace(cleanup_dir)
        check("T13: cleanup_workspace deletes the UUID folder", not cleanup_dir.exists())
        try:
            cleanup_workspace(Path(tmp) / "abc")
            safe = False
        except WorkspaceError:
            safe = Path(tmp).exists()
        check("T13: cleanup refuses non-UUID4 folders", safe)

    # Workspace-mode output
    if not WORKSPACE_OUTPUT.exists():
        check("workspace-mode output exists", False, str(WORKSPACE_OUTPUT))
        return
    prs = Presentation(str(WORKSPACE_OUTPUT))
    text = all_text(prs)
    tokens = set(ANY_TOKEN.findall(text))

    check("T4: content.json values reached the renderer",
          "2026年销售报告" in text and "1258万元" in text)
    check("T13: repeated {{text:title}} all rendered from workspace value",
          text.count("2026年销售报告") == 3 and "{{text:title}}" not in tokens)
    check("T12: repeated {{image:product}} both rendered",
          "{{image:product}}" not in tokens)
    check("T10: {{image:missing_image}} placeholder left unchanged",
          MISSING_IMAGE_TOKEN in tokens)
    check("T10b: image key absent from workspace -> region unchanged",
          "{{image:product_image}}" in tokens
          and find_picture(pics_all := count_pictures(prs), 10.9, 4.95, 1.6, 1.6) is None
          and find_picture(pics_all, 7.1, 1.9, 5.3, 4.6) is None)
    check("T11: missing text key leaves {{text:customers}} unchanged",
          "{{text:customers}}" in tokens)
    check("No error: render completed with missing assets", True)

    pics = count_pictures(prs)
    check("W: workspace image count == 3 (product x2 + logo x1)",
          len(pics) == 3, f"{len(pics)} picture(s)")
    if len(pics) == 3:
        sq = find_picture(pics, 0.9, 1.9, 2.6, 2.6)
        tall = find_picture(pics, 7.9, 1.9, 1.2, 3.4)
        logo = find_picture(pics, 3.9, 1.9, 3.6, 1.2)
        check("W: each occurrence got its own region",
              sq is not None and tall is not None and logo is not None)
        check("T16: cover crop still per-occurrence in workspace mode",
              sq is not None and tall is not None and logo is not None
              and sq.crop_left > 0 and tall.crop_left > sq.crop_left
              and logo.crop_top > 0)


def main():
    print("== End-to-end workflow (direct --data mode) ==")
    if workflow_checks():
        print("\n== Placeholder semantics ==")
        semantic_checks()
        print("\n== Repeated-placeholder semantics ==")
        repetition_checks()
        print("\n== UUID workspace semantics ==")
        workspace_checks()
    else:
        failures.append("workflow incomplete")

    print()
    if failures:
        print(f"VERIFICATION FAILED: {len(failures)} check(s) failed")
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
