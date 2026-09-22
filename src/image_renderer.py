"""Replace {{image:key}} placeholder shapes with real pictures.

Only the template declares a region as an image: the placeholder type is
explicit, so values are never inspected for extensions or path shape.

The placeholder shape defines the region (position + size). The renderer
deletes the placeholder and inserts the image so that it exactly fills the
region, center-cropped ("cover") to preserve the image's aspect ratio
without overflowing the region bounds.
"""
from PIL import Image


def _cover_crop_factors(img_w, img_h, region_w, region_h):
    """Return (crop_x_total, crop_y_total) fractions for a center cover crop."""
    img_aspect = img_w / img_h
    region_aspect = region_w / region_h

    crop_x = crop_y = 0.0
    if img_aspect > region_aspect:
        # Image wider than region -> trim left/right.
        crop_x = 1.0 - region_aspect / img_aspect
    elif img_aspect < region_aspect:
        # Image taller than region -> trim top/bottom.
        crop_y = 1.0 - img_aspect / region_aspect
    return crop_x, crop_y


def render_images(prs, data, resolve_path, placeholders=None):
    """Replace each {{image:key}} placeholder with its picture.

    resolve_path(value) must return an existing image file path.
    Returns list of (key, used_path). Keys missing from data are skipped
    so the placeholder remains visible.
    """
    if placeholders is None:
        from .template_parser import parse_slides

        _, placeholders, _ = parse_slides(prs, data)

    used = []
    for ph in placeholders:
        if ph.key not in data:
            continue
        path = resolve_path(data[ph.key])
        left, top, width, height = ph.region
        slide = prs.slides[ph.slide_index]

        with Image.open(path) as im:
            img_w, img_h = im.size
        crop_x, crop_y = _cover_crop_factors(img_w, img_h, width, height)

        # Remove the placeholder shape, then insert the picture in its region.
        ph.shape._element.getparent().remove(ph.shape._element)
        pic = slide.shapes.add_picture(path, left, top, width=width, height=height)

        pic.crop_left = crop_x / 2
        pic.crop_right = crop_x / 2
        pic.crop_top = crop_y / 2
        pic.crop_bottom = crop_y / 2

        used.append((ph.key, path))

    return used
