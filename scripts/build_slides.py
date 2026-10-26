#!/usr/bin/env python3
"""build_slides.py -- Question 6.

Builds slides/FA24-BEDS-105.pptx: a publication-grade four-slide presentation
strictly conforming to assignment requirements:
- Slide 1: Variant, Data Architecture & Weighting Methodology
- Slide 2: Question 2: The Distribution of Education Spending
- Slide 3: Question 4: The Makeover (Default Scatter vs Redesign)
- Slide 4: Question 3 & 6: Primary Demographic Finding (Gender Literacy Gap)

Embeds SVG figures with vector SVG + PNG raster fallback matching PowerPoint's
native Insert -> Pictures mechanism.

Run from repository root:
    python scripts/build_slides.py
"""

import os
import shutil
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

# ==============================================================================
# Refined Executive Academic Palette (Non-Generic AI Palette)
# Pure crisp white background ensures SVG charts blend seamlessly without borders
# ==============================================================================
BG_COLOR    = RGBColor(0xFF, 0xFF, 0xFF)  # Pure White
CARD_BG     = RGBColor(0xF8, 0xFA, 0xFC)  # Crisp light slate container
CARD_BORDER = RGBColor(0xE2, 0xE8, 0xF0)  # Hairline border
TEXT_MAIN   = RGBColor(0x0F, 0x17, 0x2A)  # Deep Oxford Navy
TEXT_BODY   = RGBColor(0x33, 0x41, 0x55)  # Charcoal Slate
TEXT_MUTED  = RGBColor(0x64, 0x74, 0x8B)  # Cool Slate Gray
LINE_COLOR  = RGBColor(0xCB, 0xD5, 0xE1)  # Divider Gray

# Data-aligned accents (matching figure color codings)
COLOR_MEN   = RGBColor(0x2C, 0x7F, 0xB8)  # Steel Blue (Men)
COLOR_WOMEN = RGBColor(0xD9, 0x5F, 0x02)  # Terracotta / Amber (Women)
COLOR_TEAL  = RGBColor(0x0D, 0x94, 0x88)  # Deep Emerald Teal (Spending / Stats)
COLOR_ALERT = RGBColor(0xB9, 0x1C, 0x1C)  # Ruby Crimson (Peak Gap)
COLOR_NAVY  = RGBColor(0x1E, 0x3A, 0x8A)  # Classic Navy Accent

FONT_TITLE  = "Georgia"
FONT_BODY   = "Calibri"

_svg_counter = [0]


def rasterise(svg_path, png_path):
    from reportlab.graphics import renderPM
    from svglib.svglib import svg2rlg
    renderPM.drawToFile(svg2rlg(svg_path), png_path, fmt="PNG")


def add_svg_image_part(slide, svg_path):
    _svg_counter[0] += 1
    package = slide.part.package
    partname = package.next_partname("/ppt/media/image%d.svg")
    with open(svg_path, "rb") as fh:
        blob = fh.read()
    part = Part(partname, "image/svg+xml", package, blob)
    return slide.part.relate_to(part, RT.IMAGE)


def add_svg_picture(slide, svg_path, left, top, width):
    png_path = os.path.join(
        os.path.dirname(svg_path),
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


def _latin(run, name):
    rPr = run._r.get_or_add_rPr()
    el = rPr.find(qn("a:latin"))
    if el is None:
        el = etree.SubElement(rPr, qn("a:latin"))
    el.set("typeface", name)


def style_run(run, size, color, bold=False, italic=False, font=FONT_BODY):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    _latin(run, font)


def fill_slide(slide, color=BG_COLOR):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def rect(slide, left, top, width, height, fill=None, line=None, line_width=0.75):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_width)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def add_header(slide, kicker, title, subtitle):
    """Adds a unified header frame ensuring zero text collisions."""
    rect(slide, Inches(0.50), Inches(0.34), Inches(0.40), Inches(0.04), COLOR_NAVY)

    box = slide.shapes.add_textbox(Inches(0.50), Inches(0.42), Inches(9.0), Inches(1.05))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    p0 = tf.paragraphs[0]
    p0.clear()
    p0.space_before = Pt(0)
    p0.space_after = Pt(2)
    r0 = p0.add_run()
    r0.text = kicker.upper()
    style_run(r0, 9.5, COLOR_NAVY, bold=True, font=FONT_BODY)

    p1 = tf.add_paragraph()
    p1.space_before = Pt(0)
    p1.space_after = Pt(3)
    r1 = p1.add_run()
    r1.text = title
    style_run(r1, 21, TEXT_MAIN, bold=True, font=FONT_TITLE)

    p2 = tf.add_paragraph()
    p2.space_before = Pt(0)
    p2.space_after = Pt(0)
    r2 = p2.add_run()
    r2.text = subtitle
    style_run(r2, 11.5, TEXT_MUTED, font=FONT_BODY)


def add_footer(slide, current_slide, total_slides=4):
    y = Inches(7.08)
    rect(slide, Inches(0.50), y, Inches(9.0), Inches(0.01), LINE_COLOR)

    box = slide.shapes.add_textbox(Inches(0.50), y + Inches(0.06), Inches(7.0), Inches(0.25))
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "ECO346 Data Visualization · FA24-BEDS-105 · Variant 5 (KP Urban) · PSLM 2019–20"
    style_run(r, 9.5, TEXT_MUTED, font=FONT_BODY)

    box2 = slide.shapes.add_textbox(Inches(7.60), y + Inches(0.06), Inches(1.90), Inches(0.25))
    tf2 = box2.text_frame
    tf2.margin_left = tf2.margin_right = tf2.margin_top = tf2.margin_bottom = 0
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = f"Slide {current_slide} of {total_slides}"
    style_run(r2, 9.5, TEXT_MUTED, font=FONT_BODY)


def add_card(slide, left, top, width, height, stripe_color=None, bg=CARD_BG, border=CARD_BORDER):
    """Draws a card and returns a clean, zero-margin text frame inside it."""
    rect(slide, left, top, width, height, fill=bg, line=border)
    if stripe_color:
        rect(slide, left, top, width, Inches(0.04), fill=stripe_color)
    box = slide.shapes.add_textbox(left + Inches(0.14), top + Inches(0.12),
                                  width - Inches(0.28), height - Inches(0.20))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


# ==============================================================================
# Build Presentation Deck
# ==============================================================================
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

# ------------------------------------------------------------------------------
# SLIDE 1: Variant, Data & Methodology Overview
# ------------------------------------------------------------------------------
s1 = prs.slides.add_slide(blank)
fill_slide(s1)
add_header(s1,
           "ECO346  Assignment 1  ·  Methodology & Overview",
           "Variant 5: Urban Khyber Pakhtunkhwa",
           "PSLM 2019–20 microdata analysis of household education spending and demographic literacy rates.")

card_w = Inches(2.88)
card_h = Inches(3.25)
y_card = Inches(1.65)
xs = [Inches(0.50), Inches(3.56), Inches(6.62)]

# Card 1: Variant Assignment
tf1 = add_card(s1, xs[0], y_card, card_w, card_h, stripe_color=COLOR_NAVY)
p = tf1.paragraphs[0]
r = p.add_run(); r.text = "ASSIGNMENT VARIANT"; style_run(r, 9.5, COLOR_NAVY, bold=True)
p = tf1.add_paragraph(); p.space_before = Pt(3); p.space_after = Pt(8)
r = p.add_run(); r.text = "FA24-BEDS-105"; style_run(r, 16, TEXT_MAIN, bold=True, font=FONT_TITLE)
p = tf1.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("• Registration: FA24-BEDS-105\n"
          "• Last 2 digits: 05 ÷ 8 = 0 remainder 5\n"
          "• Variant 5 = KP Urban\n"
          "  - Province: Khyber Pakhtunkhwa (1)\n"
          "  - Area: Urban (2)\n"
          "• Strict subsetting applied in all R and Python scripts.")
style_run(r, 10.5, TEXT_BODY)

# Card 2: Microdata Samples
tf2 = add_card(s1, xs[1], y_card, card_w, card_h, stripe_color=COLOR_TEAL)
p = tf2.paragraphs[0]
r = p.add_run(); r.text = "MICRODATA SAMPLES"; style_run(r, 9.5, COLOR_TEAL, bold=True)
p = tf2.add_paragraph(); p.space_before = Pt(3); p.space_after = Pt(8)
r = p.add_run(); r.text = "Two Subsets Analysed"; style_run(r, 16, TEXT_MAIN, bold=True, font=FONT_TITLE)
p = tf2.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("• Literacy Sample (Q3–Q5):\n"
          "  - Ages 10–64 with non-missing literacy\n"
          "  - Total sample size: n = 2,357 persons\n"
          "• Students Sample (Q2):\n"
          "  - ed_total > 0: n = 1,061 students\n"
          "  - Analysed range: Rs 100 to Rs 1,000,000\n"
          "  - Excluded: exactly 1 student (Rs 80).")
style_run(r, 10.5, TEXT_BODY)

# Card 3: Weighting & Code
tf3 = add_card(s1, xs[2], y_card, card_w, card_h, stripe_color=COLOR_MEN)
p = tf3.paragraphs[0]
r = p.add_run(); r.text = "SURVEY WEIGHTING & CODE"; style_run(r, 9.5, COLOR_MEN, bold=True)
p = tf3.add_paragraph(); p.space_before = Pt(3); p.space_after = Pt(8)
r = p.add_run(); r.text = "Exact Weighting & Dual Code"; style_run(r, 16, TEXT_MAIN, bold=True, font=FONT_TITLE)
p = tf3.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("• All population rates weighted by hhweight.\n"
          "• Exact cumulative weight quantiles\n"
          "  (assignment definition, not built-in).\n"
          "• R scripts 01–03 + Python script 04.\n"
          "• Automated verification: R and Python\n"
          "  band tables match to 0.00% diff.")
style_run(r, 10.5, TEXT_BODY)

# Bottom Executive Summary Banner
y_sum = Inches(5.08)
tf_sum = add_card(s1, Inches(0.50), y_sum, Inches(9.0), Inches(1.80), stripe_color=COLOR_NAVY)
p = tf_sum.paragraphs[0]
r = p.add_run(); r.text = "CORE ASSIGNMENT FINDINGS AT A GLANCE"; style_run(r, 10, COLOR_NAVY, bold=True)

# 2 Columns inside banner
box_l = s1.shapes.add_textbox(Inches(0.68), y_sum + Inches(0.35), Inches(4.25), Inches(1.35))
tf_l = box_l.text_frame; tf_l.word_wrap = True; tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = tf_l.margin_bottom = 0
p = tf_l.paragraphs[0]
r = p.add_run(); r.text = "Question 2 · Education Spending:\n"; style_run(r, 10.5, TEXT_MAIN, bold=True)
r = p.add_run()
r.text = ("• Population median is Rs 9,500 vs sample median Rs 10,000.\n"
          "• High inequality: P90 (Rs 39,000) is 12.38× P10 (Rs 3,150).\n"
          "• Lower-spending students have higher survey weights,\n"
          "  pulling the weighted population median downward.")
style_run(r, 10, TEXT_BODY)

box_r = s1.shapes.add_textbox(Inches(5.08), y_sum + Inches(0.35), Inches(4.25), Inches(1.35))
tf_r = box_r.text_frame; tf_r.word_wrap = True; tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = tf_r.margin_bottom = 0
p = tf_r.paragraphs[0]
r = p.add_run(); r.text = "Question 3–4 · Literacy & Makeover:\n"; style_run(r, 10.5, TEXT_MAIN, bold=True)
r = p.add_run()
r.text = ("• Makeover: replaced overplotted points with conditional means.\n"
          "• Largest gap is at ages 40–44: 47.4 percentage points\n"
          "  (Men 77.1% vs Women 29.7%), wider than among youth (21.6).\n"
          "• Reflects historical schooling expansion, not aging loss.")
style_run(r, 10, TEXT_BODY)

add_footer(s1, 1)

# ------------------------------------------------------------------------------
# SLIDE 2: Question 2 - The Distribution of Education Spending
# ------------------------------------------------------------------------------
s2 = prs.slides.add_slide(blank)
fill_slide(s2)
add_header(s2,
           "Question 2  ·  Annual Education Spending Per Student",
           "Education Spending is Heavily Right-Skewed",
           "Weighted histogram on log10 scale (Rs 100 to Rs 1,000,000; bin width 0.1; N = 1,060 students in KP urban).")

# Left: Histogram SVG figure
add_svg_picture(s2, "figures/distribution.svg", Inches(0.50), Inches(1.60), Inches(5.20))

# Right: 3 Metric Cards
right_x = Inches(5.88)
right_w = Inches(3.62)

# Card A: Median Comparison
tf_a = add_card(s2, right_x, Inches(1.60), right_w, Inches(1.60), stripe_color=COLOR_TEAL)
p = tf_a.paragraphs[0]
r = p.add_run(); r.text = "MEDIAN SPENDING: WEIGHTED VS SAMPLE"; style_run(r, 9, COLOR_TEAL, bold=True)

# Split numbers inside card A
box_n1 = s2.shapes.add_textbox(right_x + Inches(0.14), Inches(1.88), Inches(1.60), Inches(0.55))
tfn1 = box_n1.text_frame; tfn1.margin_left = tfn1.margin_right = tfn1.margin_top = tfn1.margin_bottom = 0
p = tfn1.paragraphs[0]
r = p.add_run(); r.text = "Rs 9,500\n"; style_run(r, 19, COLOR_TEAL, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "Weighted Population"; style_run(r, 9.5, COLOR_TEAL, bold=True)

box_n2 = s2.shapes.add_textbox(right_x + Inches(1.85), Inches(1.88), Inches(1.60), Inches(0.55))
tfn2 = box_n2.text_frame; tfn2.margin_left = tfn2.margin_right = tfn2.margin_top = tfn2.margin_bottom = 0
p = tfn2.paragraphs[0]
r = p.add_run(); r.text = "Rs 10,000\n"; style_run(r, 19, TEXT_MUTED, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "Unweighted Sample"; style_run(r, 9.5, TEXT_MUTED, bold=True)

box_desc = s2.shapes.add_textbox(right_x + Inches(0.14), Inches(2.55), right_w - Inches(0.28), Inches(0.58))
tfdesc = box_desc.text_frame; tfdesc.word_wrap = True; tfdesc.margin_left = tfdesc.margin_right = tfdesc.margin_top = tfdesc.margin_bottom = 0
p = tfdesc.paragraphs[0]
r = p.add_run()
r.text = "Students below sample median have higher average survey weights (274.9 vs 228.4), pulling the weighted median downward."
style_run(r, 9.5, TEXT_BODY)

# Card B: Inequality Ratio
tf_b = add_card(s2, right_x, Inches(3.32), right_w, Inches(1.60), stripe_color=COLOR_NAVY)
p = tf_b.paragraphs[0]
r = p.add_run(); r.text = "PERCENTILE DISPARITY (P90 / P10)"; style_run(r, 9, COLOR_NAVY, bold=True)

box_r1 = s2.shapes.add_textbox(right_x + Inches(0.14), Inches(3.60), Inches(1.60), Inches(0.55))
tfr1 = box_r1.text_frame; tfr1.margin_left = tfr1.margin_right = tfr1.margin_top = tfr1.margin_bottom = 0
p = tfr1.paragraphs[0]
r = p.add_run(); r.text = "12.38 ×\n"; style_run(r, 20, COLOR_NAVY, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "P90 / P10 Ratio"; style_run(r, 9.5, COLOR_NAVY, bold=True)

box_r2 = s2.shapes.add_textbox(right_x + Inches(1.85), Inches(3.60), Inches(1.60), Inches(0.55))
tfr2 = box_r2.text_frame; tfr2.margin_left = tfr2.margin_right = tfr2.margin_top = tfr2.margin_bottom = 0
p = tfr2.paragraphs[0]
r = p.add_run(); r.text = "Rs 39,000\n"; style_run(r, 18, TEXT_MAIN, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "90th Pct (P10: Rs 3,150)"; style_run(r, 9, TEXT_MUTED)

box_rdesc = s2.shapes.add_textbox(right_x + Inches(0.14), Inches(4.28), right_w - Inches(0.28), Inches(0.55))
tfrdesc = box_rdesc.text_frame; tfrdesc.word_wrap = True; tfrdesc.margin_left = tfrdesc.margin_right = tfrdesc.margin_top = tfrdesc.margin_bottom = 0
p = tfrdesc.paragraphs[0]
r = p.add_run()
r.text = "Stark expenditure gap: top 10% of households spend over 12 times more per student than the bottom 10%."
style_run(r, 9.5, TEXT_BODY)

# Card C: Sample & Exclusion
tf_c = add_card(s2, right_x, Inches(5.04), right_w, Inches(1.82), stripe_color=LINE_COLOR)
p = tf_c.paragraphs[0]
r = p.add_run(); r.text = "SAMPLE ARCHITECTURE & EXCLUSION"; style_run(r, 9, TEXT_MUTED, bold=True)
p = tf_c.add_paragraph(); p.space_before = Pt(3)
r = p.add_run()
r.text = ("• Students with ed_total > 0: n = 1,061\n"
          "• Analysed in histogram (Rs 100–1M): n = 1,060\n"
          "• Outlier excluded: 1 student (spending = Rs 80)\n"
          "• Population weighted mean: Rs 19,146\n"
          "• Log10 scale makes high positive skewness interpretable.")
style_run(r, 10, TEXT_BODY)

add_footer(s2, 2)

# ------------------------------------------------------------------------------
# SLIDE 3: Question 4 - The Makeover (Side by Side)
# ------------------------------------------------------------------------------
s3 = prs.slides.add_slide(blank)
fill_slide(s3)
add_header(s3,
           "Question 4  ·  The Makeover (Before & After)",
           "From Uninformative Overplotting to Conditional Means",
           "Demonstrating core data visualization principles: identical sample (N = 2,357), radically different insight.")

fig_w = Inches(4.38)
fig_y = Inches(1.58)
add_svg_picture(s3, "figures/default.svg", Inches(0.50), fig_y, fig_w)
add_svg_picture(s3, "figures/literacy.svg", Inches(5.12), fig_y, fig_w)

# Figure Sub-Labels
box_dl = s3.shapes.add_textbox(Inches(0.50), Inches(4.80), fig_w, Inches(0.24))
tfdl = box_dl.text_frame; tfdl.margin_left = tfdl.margin_right = tfdl.margin_top = tfdl.margin_bottom = 0
p = tfdl.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "DEFAULT: R ggplot2 raw scatter (severely overplotted 0/1 points)"; style_run(r, 9.5, TEXT_MUTED, bold=True)

box_dr = s3.shapes.add_textbox(Inches(5.12), Inches(4.80), fig_w, Inches(0.24))
tfdr = box_dr.text_frame; tfdr.margin_left = tfdr.margin_right = tfdr.margin_top = tfdr.margin_bottom = 0
p = tfdr.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "REDESIGN: Weighted conditional means by 5-year age band and sex"; style_run(r, 9.5, COLOR_NAVY, bold=True)

# 3 Principle Cards Below
p_w = Inches(2.88)
p_h = Inches(1.88)
p_y = Inches(5.10)

# Principle 1: Conditional Means
tf_p1 = add_card(s3, xs[0], p_y, p_w, p_h, stripe_color=COLOR_ALERT)
p = tf_p1.paragraphs[0]
r = p.add_run(); r.text = "1. CONDITIONAL MEANS"; style_run(r, 9, COLOR_ALERT, bold=True)
p = tf_p1.add_paragraph(); p.space_before = Pt(2); p.space_after = Pt(4)
r = p.add_run(); r.text = "Overcoming Overplotting"; style_run(r, 12.5, TEXT_MAIN, bold=True, font=FONT_TITLE)
p = tf_p1.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("The default plots 2,357 overlapping dots at 0 and 1, completely hiding density and age patterns. Grouping by 5-year bands and taking weighted conditional means reveals the true cohort trajectory.")
style_run(r, 9.8, TEXT_BODY)

# Principle 2: Channel Ranking
tf_p2 = add_card(s3, xs[1], p_y, p_w, p_h, stripe_color=COLOR_MEN)
p = tf_p2.paragraphs[0]
r = p.add_run(); r.text = "2. CHANNEL RANKING"; style_run(r, 9, COLOR_MEN, bold=True)
p = tf_p2.add_paragraph(); p.space_before = Pt(2); p.space_after = Pt(4)
r = p.add_run(); r.text = "Position Over Color"; style_run(r, 12.5, TEXT_MAIN, bold=True, font=FONT_TITLE)
p = tf_p2.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("Position on a common scale—the highest-ranked visual channel—is dedicated to the continuous variables: age and literacy rate. Color is reserved strictly for the categorical split by sex.")
style_run(r, 9.8, TEXT_BODY)

# Principle 3: Data-Ink & Tradeoff
tf_p3 = add_card(s3, xs[2], p_y, p_w, p_h, stripe_color=COLOR_TEAL)
p = tf_p3.paragraphs[0]
r = p.add_run(); r.text = "3. DATA-INK & TRADEOFF"; style_run(r, 9, COLOR_TEAL, bold=True)
p = tf_p3.add_paragraph(); p.space_before = Pt(2); p.space_after = Pt(4)
r = p.add_run(); r.text = "Direct Labels & Hidden Spread"; style_run(r, 12.5, TEXT_MAIN, bold=True, font=FONT_TITLE)
p = tf_p3.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("Direct end-labels eliminate legend clutter. Trade-off: the default figure shows individual 0/1 spread and sample density, which the redesign hides by aggregating into summary points.")
style_run(r, 9.8, TEXT_BODY)

add_footer(s3, 3)

# ------------------------------------------------------------------------------
# SLIDE 4: Question 3 & 6 - Key Finding: The Gender Literacy Gap
# ------------------------------------------------------------------------------
s4 = prs.slides.add_slide(blank)
fill_slide(s4)
add_header(s4,
           "Question 3 & 6  ·  Primary Demographic Finding",
           "Widest Gender Literacy Gap at Ages 40–44 (47.4 Points)",
           "Weighted literacy rates by age band and sex in urban Khyber Pakhtunkhwa (PSLM 2019–20).")

# Left: Redesign SVG figure
add_svg_picture(s4, "figures/literacy.svg", Inches(0.50), Inches(1.58), Inches(5.30))

# Right: Metric Cards Stack
s4_rx = Inches(5.95)
s4_rw = Inches(3.55)

# Stat 1: Youth Cohort (15-19)
tf_s1 = add_card(s4, s4_rx, Inches(1.58), s4_rw, Inches(1.30), stripe_color=COLOR_TEAL)
p = tf_s1.paragraphs[0]
r = p.add_run(); r.text = "YOUTH COHORT (AGES 15–19)  ·  21.6 POINT GAP"; style_run(r, 8.5, COLOR_TEAL, bold=True)

box_y1 = s4.shapes.add_textbox(s4_rx + Inches(0.12), Inches(1.84), Inches(1.60), Inches(0.50))
tfy1 = box_y1.text_frame; tfy1.margin_left = tfy1.margin_right = tfy1.margin_top = tfy1.margin_bottom = 0
p = tfy1.paragraphs[0]
r = p.add_run(); r.text = "89.7%\n"; style_run(r, 19, COLOR_MEN, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "Men (n = 193)"; style_run(r, 9, TEXT_MUTED)

box_y2 = s4.shapes.add_textbox(s4_rx + Inches(1.80), Inches(1.84), Inches(1.60), Inches(0.50))
tfy2 = box_y2.text_frame; tfy2.margin_left = tfy2.margin_right = tfy2.margin_top = tfy2.margin_bottom = 0
p = tfy2.paragraphs[0]
r = p.add_run(); r.text = "68.1%\n"; style_run(r, 19, COLOR_WOMEN, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "Women (n = 163)"; style_run(r, 9, TEXT_MUTED)

box_ynote = s4.shapes.add_textbox(s4_rx + Inches(0.12), Inches(2.52), s4_rw - Inches(0.24), Inches(0.30))
tfynote = box_ynote.text_frame; tfynote.margin_left = tfynote.margin_right = tfynote.margin_top = tfynote.margin_bottom = 0
p = tfynote.paragraphs[0]
r = p.add_run(); r.text = "Highest female literacy band in adult cohorts; gap narrows among youth."; style_run(r, 9, TEXT_BODY)

# Stat 2: Older Cohort (55-59)
tf_s2 = add_card(s4, s4_rx, Inches(2.98), s4_rw, Inches(1.30), stripe_color=COLOR_NAVY)
p = tf_s2.paragraphs[0]
r = p.add_run(); r.text = "OLDER COHORT (AGES 55–59)  ·  42.6 POINT GAP"; style_run(r, 8.5, COLOR_NAVY, bold=True)

box_o1 = s4.shapes.add_textbox(s4_rx + Inches(0.12), Inches(3.24), Inches(1.60), Inches(0.50))
tfo1 = box_o1.text_frame; tfo1.margin_left = tfo1.margin_right = tfo1.margin_top = tfo1.margin_bottom = 0
p = tfo1.paragraphs[0]
r = p.add_run(); r.text = "58.8%\n"; style_run(r, 19, COLOR_MEN, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "Men (n = 41)"; style_run(r, 9, TEXT_MUTED)

box_o2 = s4.shapes.add_textbox(s4_rx + Inches(1.80), Inches(3.24), Inches(1.60), Inches(0.50))
tfo2 = box_o2.text_frame; tfo2.margin_left = tfo2.margin_right = tfo2.margin_top = tfo2.margin_bottom = 0
p = tfo2.paragraphs[0]
r = p.add_run(); r.text = "16.2%\n"; style_run(r, 19, COLOR_WOMEN, bold=True, font=FONT_TITLE)
r = p.add_run(); r.text = "Women (n = 41)"; style_run(r, 9, TEXT_MUTED)

box_onote = s4.shapes.add_textbox(s4_rx + Inches(0.12), Inches(3.92), s4_rw - Inches(0.24), Inches(0.30))
tfonote = box_onote.text_frame; tfonote.margin_left = tfonote.margin_right = tfonote.margin_top = tfonote.margin_bottom = 0
p = tfonote.paragraphs[0]
r = p.add_run(); r.text = "Severe female disadvantage in older cohorts educated in earlier decades."; style_run(r, 9, TEXT_BODY)

# Stat 3: Peak Gap Callout Box
tf_s3 = add_card(s4, s4_rx, Inches(4.38), s4_rw, Inches(1.42), stripe_color=COLOR_ALERT)
p = tf_s3.paragraphs[0]
r = p.add_run(); r.text = "LARGEST DISPARITY: AGES 40–44"; style_run(r, 9, COLOR_ALERT, bold=True)
p = tf_s3.add_paragraph(); p.space_before = Pt(2); p.space_after = Pt(2)
r = p.add_run(); r.text = "47.4 Points Gap"; style_run(r, 20, COLOR_ALERT, bold=True, font=FONT_TITLE)
p = tf_s3.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("Men: 77.1% (n = 100) vs Women: 29.7% (n = 110).\n"
          "Wider than both youth (21.6 pts) and oldest adults (37.8 pts).")
style_run(r, 9.8, TEXT_BODY)

# Bottom Economic Caveat Banner
y_bot = Inches(5.92)
tf_bot = add_card(s4, Inches(0.50), y_bot, Inches(9.0), Inches(0.95), stripe_color=COLOR_NAVY)
p = tf_bot.paragraphs[0]
r = p.add_run(); r.text = "CRITICAL ECONOMIC INTERPRETATION: WHAT THIS CANNOT TELL US"; style_run(r, 9, COLOR_NAVY, bold=True)
p = tf_bot.add_paragraph(); p.space_before = Pt(2)
r = p.add_run()
r.text = ("This cross-section reflects the historical expansion of schooling across birth cohorts over time, not individual people losing literacy as they age. A single cross-section cannot separate aging from birth cohort effects, nor identify causal drivers.")
style_run(r, 10, TEXT_BODY)

add_footer(s4, 4)

# ==============================================================================
# Save Presentation & Patch SVG Content-Type Extension
# ==============================================================================
os.makedirs("slides", exist_ok=True)
out = "slides/FA24-BEDS-105.pptx"
prs.save(out)

tmp = out + ".tmp"
with zipfile.ZipFile(out, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")
            if 'Extension="svg"' not in text:
                open_end = text.index("<Types")
                close = text.index(">", open_end) + 1
                text = (text[:close]
                        + '<Default Extension="svg" ContentType="image/svg+xml"/>'
                        + text[close:])
            data = text.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, out)

print(f"SUCCESS: Saved {out} with {len(prs.slides._sldIdLst)} slides.")
