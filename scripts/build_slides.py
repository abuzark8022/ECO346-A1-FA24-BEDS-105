#!/usr/bin/env python3
"""build_slides.py -- Question 6.

Builds slides/FA24-BEDS-105.pptx: at most four slides showing the variant, the
makeover (default and redesign side by side), and the key findings.

Each figure is inserted as an SVG picture, exactly as PowerPoint does with
Insert > Pictures > .svg: the slide stores the vector SVG and a PNG raster
fallback, joined through the asvg:svgBlip extension.  The PNG is generated on
the fly by rasterising the SVG in figures/.

Run from the repository root:  python scripts/build_slides.py
"""

import os

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.opc.package import Part
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

ACCENT = RGBColor(0x1F, 0x3B, 0x6E)
MUTED = RGBColor(0x55, 0x55, 0x55)

_svg_counter = [0]


def rasterise(svg_path, png_path):
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPM
    renderPM.drawToFile(svg2rlg(svg_path), png_path, fmt="PNG")


def add_svg_image_part(slide, svg_path):
    """Store the SVG as an image part and relate it from the slide part.

    Returns the relationship id to reference from the slide's asvg:svgBlip.
    The part name follows PowerPoint's own convention for SVG media:
    ppt/media/imageN.svg with content type image/svg+xml.
    """
    _svg_counter[0] += 1
    package = slide.part.package
    partname = package.next_partname("/ppt/media/image%d.svg")
    with open(svg_path, "rb") as fh:
        blob = fh.read()
    part = Part(partname, "image/svg+xml", package, blob)
    return slide.part.relate_to(part, RT.IMAGE)


def add_svg_picture(slide, svg_path, left, top, width):
    """Insert an SVG picture with a PNG fallback, as PowerPoint does."""
    png_path = os.path.join(os.path.dirname(svg_path),
                            "_" + os.path.splitext(os.path.basename(svg_path))[0] + "_slide.png")
    try:
        rasterise(svg_path, png_path)
        pic = slide.shapes.add_picture(png_path, left, top, width=width)
    finally:
        if os.path.exists(png_path):
            os.remove(png_path)

    blip = pic._element.blipFill.find("{%s}blip" % A)

    svg_rid = add_svg_image_part(slide, svg_path)

    ext_lst = blip.find("{%s}extLst" % A)
    if ext_lst is None:
        ext_lst = etree.SubElement(blip, "{%s}extLst" % A)
    ext = etree.SubElement(ext_lst, "{%s}ext" % A)
    ext.set("uri", SVG_EXT_URI)
    svg_blip = etree.SubElement(ext, "{%s}svgBlip" % ASVG)
    svg_blip.set("{%s}embed" % R, svg_rid)
    return pic


def add_title(slide, text, size=26):
    box = slide.shapes.add_textbox(Inches(0.55), Inches(0.3), Inches(9.0), Inches(0.9))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = ACCENT
    return box


def add_bullets(slide, lines, top, size=17):
    box = slide.shapes.add_textbox(Inches(0.6), top, Inches(8.8), Inches(5.4))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for main, sub in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(6)
        r = p.add_run(); r.text = main
        r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = ACCENT
        if sub:
            p2 = tf.add_paragraph()
            r2 = p2.add_run(); r2.text = sub
            r2.font.size = Pt(size - 4); r2.font.color.rgb = MUTED
    return box


prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# ---------------------------------------------------------------- Slide 1 ---
s = prs.slides.add_slide(blank)
add_title(s, "ECO346 Assignment 1 \u2014 Variant 5: KP urban")
add_bullets(s, [
    ("Registration number", "FA24-BEDS-105"),
    ("Variant", "5 = Khyber Pakhtunkhwa, urban  (prov_code 1, urban_rural 2)"),
    ("Data", "PSLM 2019-20 person and household files; all statistics weighted by hhweight"),
    ("Deliverables", "Two figures in R, a makeover, and the band table reproduced in Python"),
], Inches(1.6), size=18)

# ---------------------------------------------------------------- Slide 2 ---
s = prs.slides.add_slide(blank)
add_title(s, "The makeover: default (left) vs redesign (right)", size=22)
add_svg_picture(s, "figures/default.svg", Inches(0.3), Inches(1.35), Inches(4.5))
add_svg_picture(s, "figures/literacy.svg", Inches(5.2), Inches(1.15), Inches(4.5))
cap = s.shapes.add_textbox(Inches(0.3), Inches(6.5), Inches(9.4), Inches(0.8))
cap.text_frame.word_wrap = True
p = cap.text_frame.paragraphs[0]
p.text = ("Default: raw literate-vs-age points, overplotted.  "
          "Redesign: weighted literacy by age band, two series, direct labels, "
          "no legend, thin cells hollow, widest gap annotated.")
p.font.size = Pt(11); p.font.color.rgb = MUTED

# ---------------------------------------------------------------- Slide 3 ---
s = prs.slides.add_slide(blank)
add_title(s, "Key finding: large gender gap in literacy across older cohorts", size=22)
add_bullets(s, [
    ("Women: 68.1% at ages 15-19 vs 16.2% at ages 55-59",
     "Men: 89.7% at ages 15-19 vs 58.8% at ages 55-59"),
    ("Largest gap at ages 40-44: 47.4 percentage points",
     "Men 77.1% vs women 29.7%"),
    ("Interpretation",
     "Cross-sectional data show generational inequalities; cannot establish causes or aging effects"),
], Inches(1.6), size=17)

# ---------------------------------------------------------------- Slide 4 ---
s = prs.slides.add_slide(blank)
add_title(s, "Education spending: skewed even in logs", size=24)
add_bullets(s, [
    ("Weighted median Rs 9,500; weighted 90th percentile Rs 39,000",
     "Ratio of 90th to 10th percentile: 12.38"),
    ("Weighting shifts median from Rs 10,100 (sample) to Rs 9,500 (population)",
     "Lower-spending students have higher average survey weights (268.7 vs 228.4)"),
    ("Sample",
     "1,060 students analysed of 1,061 (1 excluded outside Rs 100-Rs 1,000,000)"),
], Inches(1.6), size=17)

os.makedirs("slides", exist_ok=True)
out = "slides/FA24-BEDS-105.pptx"
prs.save(out)

# PowerPoint wants a Default extension mapping for svg (not just an Override),
# so that it recognises the vector image.  Add it to [Content_Types].xml.
import zipfile
import shutil

tmp = out + ".tmp"
with zipfile.ZipFile(out, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")
            if 'Extension="svg"' not in text:
                # insert a Default right after the opening <Types ...> element
                open_end = text.index("<Types", text.index("<Types") if "<Types" in text else 0)
                close = text.index(">", open_end) + 1
                text = (text[:close]
                        + '<Default Extension="svg" ContentType="image/svg+xml"/>'
                        + text[close:])
            data = text.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, out)

print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
