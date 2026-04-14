"""Slide dimensions and layout constants."""

from pptx.util import Inches, Pt

# Slide dimensions (16:9 standard)
SW = Inches(13.333)
SH = Inches(7.5)

# Common regions
MARGIN_L = Inches(1.0)
MARGIN_R = Inches(1.0)
MARGIN_T = Inches(0.45)
CONTENT_W = SW - MARGIN_L - MARGIN_R
TITLE_H = Inches(0.45)
TITLE_TOP = MARGIN_T
BODY_TOP = MARGIN_T + TITLE_H + Inches(0.12)
BODY_H = SH - BODY_TOP - Inches(0.6)

# Font size scale
SZ_TITLE = Pt(18)
SZ_H2 = Pt(16)
SZ_H3 = Pt(14)
SZ_BODY = Pt(14)
SZ_COL = Pt(13)
SZ_SMALL = Pt(12)
SZ_FOOT = Pt(9)
SZ_EQ = Pt(28)
SZ_EQ_VAR = Pt(14)
SZ_ZONE_L = Pt(15)
SZ_ZONE_B = Pt(13)
