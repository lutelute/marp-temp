#!/usr/bin/env python3
"""Generate a Pandoc reference PPTX with academic theme styling."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import copy

# Academic theme colors
PRIMARY   = RGBColor(0x16, 0x21, 0x3e)
SECONDARY = RGBColor(0x0f, 0x34, 0x60)
ACCENT    = RGBColor(0xe9, 0x45, 0x60)
BG        = RGBColor(0xff, 0xff, 0xff)
FG        = RGBColor(0x1a, 0x1a, 0x2e)
MUTED     = RGBColor(0x6c, 0x75, 0x7d)
LIGHT     = RGBColor(0xf0, 0xf2, 0xf5)
WHITE     = RGBColor(0xff, 0xff, 0xff)

FONT_HEAD = "Helvetica Neue"
FONT_BODY = "Helvetica Neue"


def style_title_placeholder(ph, color=PRIMARY, size=Pt(40), bold=True):
    """Style a title placeholder."""
    for para in ph.text_frame.paragraphs:
        para.font.name = FONT_HEAD
        para.font.size = size
        para.font.bold = bold
        para.font.color.rgb = color


def style_body_placeholder(ph, color=FG, size=Pt(20)):
    """Style a body/content placeholder."""
    for para in ph.text_frame.paragraphs:
        para.font.name = FONT_BODY
        para.font.size = size
        para.font.color.rgb = color


def set_slide_bg(slide_layout, color):
    """Set solid background color on a slide layout."""
    bg = slide_layout.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_gradient_bg(slide_layout, color1, color2):
    """Set gradient background on a slide layout."""
    bg = slide_layout.background
    fill = bg.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[0].position = 0.0
    fill.gradient_stops[1].color.rgb = color2
    fill.gradient_stops[1].position = 1.0


def main():
    prs = Presentation("/tmp/reference_default.pptx")

    # --- Layout 0: Title Slide ---
    layout0 = prs.slide_layouts[0]
    set_gradient_bg(layout0, PRIMARY, SECONDARY)
    for ph in layout0.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:  # Title
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_HEAD
                para.font.size = Pt(44)
                para.font.bold = True
                para.font.color.rgb = WHITE
                para.alignment = PP_ALIGN.CENTER
        elif idx == 1:  # Subtitle
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_BODY
                para.font.size = Pt(22)
                para.font.bold = False
                para.font.color.rgb = WHITE
                para.alignment = PP_ALIGN.CENTER

    # --- Layout 1: Title and Content (default body slides) ---
    layout1 = prs.slide_layouts[1]
    set_slide_bg(layout1, BG)
    for ph in layout1.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:  # Title
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_HEAD
                para.font.size = Pt(36)
                para.font.bold = True
                para.font.color.rgb = PRIMARY
        elif idx == 1:  # Content
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_BODY
                para.font.size = Pt(20)
                para.font.color.rgb = FG

    # --- Layout 2: Section Header (divider) ---
    layout2 = prs.slide_layouts[2]
    set_slide_bg(layout2, LIGHT)
    for ph in layout2.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:  # Title
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_HEAD
                para.font.size = Pt(40)
                para.font.bold = True
                para.font.color.rgb = PRIMARY
                para.alignment = PP_ALIGN.CENTER
        elif idx == 1:  # Subtitle
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_BODY
                para.font.size = Pt(22)
                para.font.bold = False
                para.font.color.rgb = MUTED
                para.alignment = PP_ALIGN.CENTER

    # --- Layout 3: Two Content (2-column) ---
    layout3 = prs.slide_layouts[3]
    set_slide_bg(layout3, BG)
    for ph in layout3.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_HEAD
                para.font.size = Pt(36)
                para.font.bold = True
                para.font.color.rgb = PRIMARY
        elif idx in (1, 2):
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_BODY
                para.font.size = Pt(18)
                para.font.color.rgb = FG

    # --- Layout 5: Title Only ---
    layout5 = prs.slide_layouts[5]
    set_slide_bg(layout5, BG)
    for ph in layout5.placeholders:
        if ph.placeholder_format.idx == 0:
            for para in ph.text_frame.paragraphs:
                para.font.name = FONT_HEAD
                para.font.size = Pt(36)
                para.font.bold = True
                para.font.color.rgb = PRIMARY

    # --- Layout 6: Blank ---
    layout6 = prs.slide_layouts[6]
    set_slide_bg(layout6, BG)

    # Save
    out_path = "/Users/shigenoburyuto/Documents/GitHub/marp_temp/pptx/reference.pptx"
    prs.save(out_path)
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
