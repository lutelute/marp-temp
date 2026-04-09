#!/usr/bin/env python3
"""
Marp Academic Template → Editable PPTX converter.

Template-driven: each slide class maps to a dedicated builder
with fixed layout positions. No Pandoc dependency.

Usage:
    python pptx/convert.py example.md
    python pptx/convert.py example.md -o output.pptx
"""

import re
import sys
import argparse
import tempfile
from pathlib import Path
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False

# ============================================================
# Theme constants
# ============================================================
PRIMARY   = RGBColor(0x16, 0x21, 0x3e)
SECONDARY = RGBColor(0x0f, 0x34, 0x60)
ACCENT    = RGBColor(0xe9, 0x45, 0x60)
BG_WHITE  = RGBColor(0xff, 0xff, 0xff)
FG        = RGBColor(0x1a, 0x1a, 0x2e)
MUTED     = RGBColor(0x6c, 0x75, 0x7d)
LIGHT     = RGBColor(0xf0, 0xf2, 0xf5)
WHITE     = RGBColor(0xff, 0xff, 0xff)

FONT = "Helvetica Neue"

# Slide dimensions (16:9 standard)
SW = Inches(13.333)
SH = Inches(7.5)

# Common regions
MARGIN_L  = Inches(0.8)
MARGIN_R  = Inches(0.8)
MARGIN_T  = Inches(0.5)
CONTENT_W = SW - MARGIN_L - MARGIN_R
TITLE_H   = Inches(0.9)
TITLE_TOP  = MARGIN_T
BODY_TOP   = MARGIN_T + TITLE_H + Inches(0.15)
BODY_H     = SH - BODY_TOP - Inches(0.5)

# ============================================================
# Data structures
# ============================================================
@dataclass
class SlideData:
    index: int
    slide_class: str | None
    paginate: bool
    raw: str
    # Parsed fields (populated by parser)
    h1: str = ""
    h2: str = ""
    body_lines: list = field(default_factory=list)
    columns: list = field(default_factory=list)  # list of list of lines
    top_text: str = ""      # sandwich lead
    bottom_text: str = ""   # sandwich conclusion
    table_rows: list = field(default_factory=list)
    image_path: str = ""
    caption: str = ""
    footnote: str = ""
    timeline_items: list = field(default_factory=list)
    eq_main: str = ""
    eq_vars: list = field(default_factory=list)  # [(sym, desc), ...]
    ref_items: list = field(default_factory=list)  # [(author, title, venue), ...]


# ============================================================
# Parser
# ============================================================
def strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text).strip()


def extract_div(text: str, cls: str) -> str | None:
    pattern = rf'<div\s+class="[^"]*{re.escape(cls)}[^"]*">'
    m = re.search(pattern, text)
    if not m:
        return None
    start = m.end()
    depth = 1
    pos = start
    while pos < len(text) and depth > 0:
        no = text.find("<div", pos)
        nc = text.find("</div>", pos)
        if nc == -1:
            break
        if no != -1 and no < nc:
            depth += 1
            pos = no + 4
        else:
            depth -= 1
            if depth == 0:
                return text[start:nc].strip()
            pos = nc + 6
    return text[start:].strip()


def extract_child_divs(text: str) -> list[str]:
    children = []
    pos = 0
    while pos < len(text):
        m = re.search(r"<div[^>]*>", text[pos:])
        if not m:
            break
        ds = pos + m.end()
        depth = 1
        scan = ds
        while scan < len(text) and depth > 0:
            no = text.find("<div", scan)
            nc = text.find("</div>", scan)
            if nc == -1:
                break
            if no != -1 and no < nc:
                depth += 1
                scan = no + 4
            else:
                depth -= 1
                if depth == 0:
                    children.append(text[ds:nc].strip())
                    pos = nc + 6
                    break
                scan = nc + 6
        else:
            break
    return children


def parse_markdown_lines(text: str) -> list[str]:
    """Parse markdown text into clean lines, stripping HTML."""
    lines = []
    for line in text.split("\n"):
        s = line.strip()
        if s.startswith("<div") or s.startswith("</div>") or s.startswith("<p ") or s.startswith("<ol") or s.startswith("</ol") or s.startswith("<li") or s.startswith("</li"):
            inner = strip_html(s)
            if inner:
                lines.append(inner)
        elif s.startswith("<span"):
            inner = strip_html(s)
            if inner:
                lines.append(inner)
        else:
            lines.append(strip_html(line.rstrip()))
    # Remove consecutive blank lines
    result = []
    prev_blank = False
    for l in lines:
        if not l.strip():
            if not prev_blank:
                result.append("")
            prev_blank = True
        else:
            result.append(l)
            prev_blank = False
    return result


def parse_slide(index: int, raw: str) -> SlideData:
    """Parse a raw slide chunk into SlideData."""
    # Extract directives
    directives = {}
    def repl(m):
        directives[m.group(1)] = m.group(2)
        return ""
    content = re.sub(r"<!--\s+_(\w+):\s*(.+?)\s*-->", repl, raw).strip()

    sd = SlideData(
        index=index,
        slide_class=directives.get("class"),
        paginate=directives.get("paginate", "true") != "false",
        raw=content,
    )

    # H1, H2
    h1m = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    h2m = re.search(r"^##\s+(.+)$", content, re.MULTILINE)
    if h1m:
        sd.h1 = h1m.group(1).strip()
    if h2m:
        sd.h2 = h2m.group(1).strip()

    cls = sd.slide_class

    # --- Equation ---
    if cls == "equation":
        eq = extract_div(content, "eq-main")
        if eq:
            # Extract $$...$$ from eq-main
            mm = re.search(r"\$\$(.*?)\$\$", eq, re.DOTALL)
            sd.eq_main = mm.group(1).strip() if mm else strip_html(eq)

        desc = extract_div(content, "eq-desc")
        if desc:
            spans = re.findall(r"<span[^>]*>(.*?)</span>", desc, re.DOTALL)
            for i in range(0, len(spans) - 1, 2):
                sym = strip_html(spans[i])
                d = strip_html(spans[i + 1])
                sd.eq_vars.append((sym, d))

        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Columns ---
    elif cls in ("cols-2", "cols-2-wide-l", "cols-2-wide-r", "cols-3"):
        cols = extract_div(content, "columns")
        if cols:
            for child in extract_child_divs(cols):
                sd.columns.append(parse_markdown_lines(child))
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Sandwich ---
    elif cls == "sandwich":
        top = extract_div(content, "top")
        if top:
            lead = extract_div(top, "lead")
            sd.top_text = strip_html(lead) if lead else strip_html(top)

        cols = extract_div(content, "columns")
        if cols:
            for child in extract_child_divs(cols):
                sd.columns.append(parse_markdown_lines(child))

        bottom = extract_div(content, "bottom")
        if bottom:
            conc = extract_div(bottom, "conclusion")
            if conc:
                sd.bottom_text = strip_html(conc)
            else:
                box = extract_div(bottom, "box")
                sd.bottom_text = strip_html(box) if box else strip_html(bottom)

    # --- Figure ---
    elif cls == "figure":
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)
        desc = extract_div(content, "description")
        if desc:
            sd.body_lines = parse_markdown_lines(desc)

    # --- Table ---
    elif cls == "table-slide":
        rows = []
        for line in content.split("\n"):
            s = line.strip()
            if s.startswith("|") and not re.match(r"^\|[-:|]+\|$", s):
                cells = [c.strip() for c in s.strip("|").split("|")]
                rows.append(cells)
        sd.table_rows = rows

        ba = extract_div(content, "box-accent")
        if ba:
            sd.bottom_text = strip_html(ba)
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- References ---
    elif cls == "references":
        lis = re.findall(r"<li>(.*?)</li>", content, re.DOTALL)
        for li in lis:
            am = re.search(r'class="author"[^>]*>(.*?)</span>', li)
            tm = re.search(r'class="title"[^>]*>(.*?)</span>', li)
            vm = re.search(r'class="venue"[^>]*>(.*?)</span>', li)
            sd.ref_items.append((
                am.group(1).strip() if am else "",
                tm.group(1).strip() if tm else "",
                vm.group(1).strip() if vm else "",
            ))

    # --- Timeline ---
    elif cls == "timeline-h":
        container = extract_div(content, "tl-h-container")
        if container:
            items = extract_child_divs(container)
            for item in items:
                block = extract_child_divs(item)
                inner = block[0] if block else item
                ym = re.search(r'class="tl-h-year"[^>]*>(.*?)</span>', inner, re.DOTALL)
                tm = re.search(r'class="tl-h-text"[^>]*>(.*?)</span>', inner, re.DOTALL)
                dm = re.search(r'class="tl-h-detail"[^>]*>(.*?)</div>', inner, re.DOTALL)
                sd.timeline_items.append({
                    "year": strip_html(ym.group(1)) if ym else "",
                    "text": strip_html(tm.group(1)) if tm else "",
                    "detail": re.sub(r"\s+", " ", strip_html(dm.group(1))) if dm else "",
                    "highlight": "highlight" in item,
                })

    elif cls == "timeline":
        container = extract_div(content, "tl-container")
        if container:
            items = extract_child_divs(container)
            for item in items:
                ym = re.search(r'class="tl-year"[^>]*>(.*?)</span>', item, re.DOTALL)
                tm = re.search(r'class="tl-text"[^>]*>(.*?)</span>', item, re.DOTALL)
                dm = re.search(r'class="tl-detail"[^>]*>(.*?)</div>', item, re.DOTALL)
                sd.timeline_items.append({
                    "year": strip_html(ym.group(1)) if ym else "",
                    "text": strip_html(tm.group(1)) if tm else "",
                    "detail": strip_html(dm.group(1)) if dm else "",
                    "highlight": "highlight" in item,
                })

    # --- Default body ---
    else:
        # Remove h1/h2 lines from body
        body = content
        if h1m:
            body = body[:h1m.start()] + body[h1m.end():]
        if h2m:
            body = body[:h2m.start()] + body[h2m.end():]
        # Extract box-accent
        ba = extract_div(body, "box-accent")
        bp = extract_div(body, "box-primary")
        fn = extract_div(body, "footnote")
        # Clean remaining body
        for tag in ("box-accent", "box-primary", "box", "footnote"):
            div = extract_div(body, tag)
            if div:
                # Remove the entire div from body
                pattern = rf'<div\s+class="[^"]*{tag}[^"]*">.*?</div>'
                body = re.sub(pattern, "", body, flags=re.DOTALL)
        sd.body_lines = parse_markdown_lines(body)
        if ba:
            sd.bottom_text = strip_html(ba)
        elif bp:
            sd.bottom_text = strip_html(bp)
        if fn:
            sd.footnote = strip_html(fn)

        # Check for images in default slides
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

    return sd


def parse_marp(path: str) -> list[SlideData]:
    text = Path(path).read_text(encoding="utf-8")
    # Strip frontmatter
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            text = text[end + 3:]
    chunks = re.split(r"\n---\n", text)
    slides = []
    for i, chunk in enumerate(chunks):
        chunk = chunk.strip()
        if chunk:
            slides.append(parse_slide(i, chunk))
    return slides


# ============================================================
# PPTX Builders
# ============================================================
class PptxBuilder:
    def __init__(self, base_path: Path):
        self.prs = Presentation()
        self.prs.slide_width = SW
        self.prs.slide_height = SH
        self.base_path = base_path
        self._img_cache = {}

    def save(self, path: str):
        self.prs.save(path)

    def _blank_slide(self):
        layout = self.prs.slide_layouts[6]  # Blank
        return self.prs.slides.add_slide(layout)

    def _add_textbox(self, slide, left, top, width, height):
        return slide.shapes.add_textbox(left, top, width, height)

    def _set_bg(self, slide, color):
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = color

    def _set_gradient_bg(self, slide, c1, c2):
        bg = slide.background
        bg.fill.gradient()
        bg.fill.gradient_stops[0].color.rgb = c1
        bg.fill.gradient_stops[0].position = 0.0
        bg.fill.gradient_stops[1].color.rgb = c2
        bg.fill.gradient_stops[1].position = 1.0

    def _add_title(self, slide, text, top=None, color=PRIMARY):
        if top is None:
            top = TITLE_TOP
        tb = self._add_textbox(slide, MARGIN_L, top, CONTENT_W, TITLE_H)
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = FONT
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = color
        # Accent underline
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            MARGIN_L, top + TITLE_H - Inches(0.05),
            CONTENT_W, Pt(3)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = ACCENT
        line.line.fill.background()
        return tb

    def _add_para(self, tf, text, size=Pt(20), color=FG, bold=False, italic=False, space_before=Pt(4)):
        p = tf.add_paragraph()
        p.text = text
        p.font.name = FONT
        p.font.size = size
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.italic = italic
        p.space_before = space_before
        return p

    def _add_body_text(self, slide, lines, left=None, top=None, width=None, height=None, size=Pt(20)):
        if left is None: left = MARGIN_L
        if top is None: top = BODY_TOP
        if width is None: width = CONTENT_W
        if height is None: height = BODY_H

        tb = self._add_textbox(slide, left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True

        first = True
        for line in lines:
            s = line.strip()
            if not s:
                continue

            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()

            # Handle markdown formatting
            is_h2 = s.startswith("## ")
            is_h3 = s.startswith("### ")
            is_bullet = s.startswith("- ") or s.startswith("* ")
            is_numbered = re.match(r"^\d+\.\s", s)

            if is_h2:
                p.text = s[3:]
                p.font.name = FONT
                p.font.size = Pt(26)
                p.font.bold = True
                p.font.color.rgb = SECONDARY
                p.space_before = Pt(12)
            elif is_h3:
                p.text = s[4:]
                p.font.name = FONT
                p.font.size = Pt(22)
                p.font.bold = True
                p.font.color.rgb = MUTED
                p.space_before = Pt(8)
            elif is_bullet:
                p.text = s[2:]
                p.font.name = FONT
                p.font.size = size
                p.font.color.rgb = FG
                p.level = 0
                p.space_before = Pt(4)
                # Bullet
                pPr = p._pPr
                if pPr is None:
                    from pptx.oxml.ns import qn
                    pPr = p._p.get_or_add_pPr()
                from pptx.oxml.ns import qn
                buChar = pPr.makeelement(qn("a:buChar"), {"char": "•"})
                # Remove existing bullets
                for existing in pPr.findall(qn("a:buChar")):
                    pPr.remove(existing)
                for existing in pPr.findall(qn("a:buNone")):
                    pPr.remove(existing)
                pPr.append(buChar)
            elif is_numbered:
                num_text = re.sub(r"^\d+\.\s", "", s)
                p.text = s
                p.font.name = FONT
                p.font.size = size
                p.font.color.rgb = FG
                p.space_before = Pt(4)
            else:
                # Apply bold markdown: **text**
                self._set_rich_text(p, s, size, FG)
                p.space_before = Pt(4)

        return tb

    def _set_rich_text(self, para, text, size=Pt(20), color=FG):
        """Parse **bold** and set runs."""
        para.clear()
        parts = re.split(r"(\*\*.*?\*\*)", text)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                run = para.add_run()
                run.text = part[2:-2]
                run.font.name = FONT
                run.font.size = size
                run.font.color.rgb = color
                run.font.bold = True
            else:
                run = para.add_run()
                run.text = part
                run.font.name = FONT
                run.font.size = size
                run.font.color.rgb = color

    def _resolve_image(self, img_path: str) -> str | None:
        """Resolve image path, converting SVG to PNG if needed."""
        p = self.base_path / img_path
        if not p.exists():
            return None
        if p.suffix.lower() == ".svg":
            png_path = p.with_suffix(".png")
            if not png_path.exists() or png_path.stat().st_mtime < p.stat().st_mtime:
                if HAS_CAIROSVG:
                    cairosvg.svg2png(url=str(p), write_to=str(png_path), output_width=1400, dpi=300)
                else:
                    return None
            return str(png_path)
        return str(p)

    def _add_accent_box(self, slide, text, left, top, width, height, border_color=ACCENT):
        """Add a box with thick left border."""
        # Background rect
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = LIGHT
        bg.line.fill.background()
        bg.rotation = 0
        # Left border
        bdr = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, top, Pt(6), height
        )
        bdr.fill.solid()
        bdr.fill.fore_color.rgb = border_color
        bdr.line.fill.background()
        # Text
        tb = self._add_textbox(slide, left + Pt(16), top + Pt(8), width - Pt(32), height - Pt(16))
        tf = tb.text_frame
        tf.word_wrap = True
        self._set_rich_text(tf.paragraphs[0], text, Pt(18), FG)
        return tb

    def _add_conclusion_box(self, slide, text, left, top, width, height):
        """Dark box with white text."""
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = PRIMARY
        bg.line.fill.background()
        tb = self._add_textbox(slide, left + Pt(16), top + Pt(10), width - Pt(32), height - Pt(20))
        tf = tb.text_frame
        tf.word_wrap = True
        self._set_rich_text(tf.paragraphs[0], text, Pt(17), WHITE)
        return tb

    def _add_footnote(self, slide, text):
        left = MARGIN_L
        top = SH - Inches(0.55)
        width = CONTENT_W
        # Line
        ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(1))
        ln.fill.solid()
        ln.fill.fore_color.rgb = RGBColor(0xDE, 0xE2, 0xE6)
        ln.line.fill.background()
        # Text
        tb = self._add_textbox(slide, left, top + Pt(4), width, Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = FONT
        p.font.size = Pt(11)
        p.font.color.rgb = MUTED

    # ---------- Slide type builders ----------

    def build_title(self, sd: SlideData):
        slide = self._blank_slide()
        self._set_gradient_bg(slide, PRIMARY, SECONDARY)

        # Title
        tb = self._add_textbox(slide, Inches(1), Inches(1.5), SW - Inches(2), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sd.h1
        p.font.name = FONT
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

        # Subtitle + author info
        tb2 = self._add_textbox(slide, Inches(1), Inches(3.2), SW - Inches(2), Inches(3.5))
        tf2 = tb2.text_frame
        tf2.word_wrap = True

        # Remaining lines after h1
        remaining = []
        past_h1 = False
        for line in sd.raw.split("\n"):
            s = line.strip()
            if s.startswith("# ") and not past_h1:
                past_h1 = True
                continue
            if s.startswith("## "):
                remaining.append(s[3:])
                continue
            if past_h1 and s:
                remaining.append(strip_html(s))

        first = True
        for line in remaining:
            if first:
                p = tf2.paragraphs[0]
                first = False
            else:
                p = tf2.add_paragraph()
            p.text = line
            p.font.name = FONT
            p.font.size = Pt(20)
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(6)

    def build_divider(self, sd: SlideData):
        slide = self._blank_slide()
        self._set_bg(slide, LIGHT)

        tb = self._add_textbox(slide, Inches(1), Inches(2.5), SW - Inches(2), Inches(1.2))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = sd.h1
        p.font.name = FONT
        p.font.size = Pt(42)
        p.font.bold = True
        p.font.color.rgb = PRIMARY
        p.alignment = PP_ALIGN.CENTER

        if sd.h2:
            p2 = tf.add_paragraph()
            p2.text = sd.h2
            p2.font.name = FONT
            p2.font.size = Pt(22)
            p2.font.color.rgb = MUTED
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(12)

    def build_default(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if sd.body_lines:
            self._add_body_text(slide, sd.body_lines)
        if sd.bottom_text:
            self._add_accent_box(
                slide, sd.bottom_text,
                MARGIN_L, SH - Inches(1.8), CONTENT_W, Inches(1.0)
            )
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_equation(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        # Main equation as large text
        eq_top = BODY_TOP + Inches(0.2)
        tb = self._add_textbox(slide, MARGIN_L, eq_top, CONTENT_W, Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sd.eq_main
        p.font.name = FONT
        p.font.size = Pt(28)
        p.font.color.rgb = FG
        p.alignment = PP_ALIGN.CENTER

        # Variable descriptions
        if sd.eq_vars:
            var_top = eq_top + Inches(1.6)
            var_tb = self._add_textbox(slide, Inches(2.5), var_top, Inches(8), Inches(3))
            tf = var_tb.text_frame
            tf.word_wrap = True
            first = True
            for sym, desc in sd.eq_vars:
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()

                run_sym = p.add_run()
                run_sym.text = sym + "  "
                run_sym.font.name = FONT
                run_sym.font.size = Pt(18)
                run_sym.font.bold = True
                run_sym.font.color.rgb = SECONDARY

                run_desc = p.add_run()
                run_desc.text = desc
                run_desc.font.name = FONT
                run_desc.font.size = Pt(17)
                run_desc.font.color.rgb = FG

                p.space_before = Pt(6)

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def _add_column_content(self, slide, lines, left, top, width, height, size=Pt(18)):
        """Add column content, handling embedded images."""
        # Separate image lines from text lines
        text_lines = []
        images = []
        for line in lines:
            img_m = re.match(r"!\[(?:w:\d+)?\]\(([^)]+)\)", line.strip())
            if img_m:
                images.append(img_m.group(1))
            else:
                text_lines.append(line)

        cur_top = top

        # Insert images first
        for img_path in images:
            img_file = self._resolve_image(img_path)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(width * 0.95)
                max_h = int(Inches(2.5))
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                img_left = left + (width - pw) // 2
                slide.shapes.add_picture(img_file, img_left, cur_top, pw, ph)
                cur_top += ph + Inches(0.1)

        # Add remaining text
        remaining_h = top + height - cur_top
        if text_lines and remaining_h > 0:
            self._add_body_text(slide, text_lines, left=left, top=int(cur_top), width=int(width), height=int(remaining_h), size=size)

    def build_columns(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        n = len(sd.columns)
        if n == 0:
            return

        gap = Inches(0.4)
        total_gap = gap * (n - 1)
        col_w = (CONTENT_W - total_gap) / n

        for i, col_lines in enumerate(sd.columns):
            left = MARGIN_L + i * (col_w + gap)
            self._add_column_content(slide, col_lines, left=int(left), top=BODY_TOP, width=int(col_w), height=BODY_H, size=Pt(18))

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_sandwich(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        cur_top = BODY_TOP

        # Lead text
        if sd.top_text:
            tb = self._add_textbox(slide, MARGIN_L, cur_top, CONTENT_W, Inches(0.8))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = sd.top_text
            p.font.name = FONT
            p.font.size = Pt(19)
            p.font.color.rgb = SECONDARY
            cur_top += Inches(1.0)

        # Columns
        n = len(sd.columns)
        if n > 0:
            gap = Inches(0.4)
            total_gap = gap * (n - 1)
            col_w = (CONTENT_W - total_gap) / n
            col_h = Inches(2.5)

            for i, col_lines in enumerate(sd.columns):
                left = MARGIN_L + i * (col_w + gap)
                self._add_body_text(slide, col_lines, left=left, top=int(cur_top), width=int(col_w), height=int(col_h), size=Pt(17))

            cur_top += col_h + Inches(0.2)

        # Conclusion
        if sd.bottom_text:
            remaining_h = SH - cur_top - Inches(0.4)
            box_h = min(Inches(1.0), int(remaining_h))
            self._add_conclusion_box(slide, sd.bottom_text, MARGIN_L, int(cur_top), CONTENT_W, box_h)

    def build_figure(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        img_file = self._resolve_image(sd.image_path) if sd.image_path else None
        if img_file:
            # Center image
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(CONTENT_W * 0.85)
            max_h = int(Inches(3.5))
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96)
            ph = int(ih * scale * 914400 / 96)
            left = (SW - pw) // 2
            slide.shapes.add_picture(img_file, left, BODY_TOP, pw, ph)
            cap_top = BODY_TOP + ph + Inches(0.15)
        else:
            cap_top = BODY_TOP + Inches(3.5)

        if sd.caption:
            tb = self._add_textbox(slide, MARGIN_L, cap_top, CONTENT_W, Inches(0.8))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = sd.caption
            p.font.name = FONT
            p.font.size = Pt(14)
            p.font.color.rgb = FG
            p.alignment = PP_ALIGN.CENTER

        if sd.body_lines:
            desc_top = cap_top + Inches(0.7)
            self._add_body_text(slide, sd.body_lines, top=int(desc_top), size=Pt(17))

    def build_table(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if sd.h2:
            tb = self._add_textbox(slide, MARGIN_L, BODY_TOP - Inches(0.3), CONTENT_W, Inches(0.5))
            tf = tb.text_frame
            p = tf.paragraphs[0]
            p.text = sd.h2
            p.font.name = FONT
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = SECONDARY

        if not sd.table_rows:
            return

        rows = sd.table_rows
        n_rows = len(rows)
        n_cols = max(len(r) for r in rows) if rows else 0

        tbl_top = BODY_TOP + Inches(0.4)
        tbl_w = CONTENT_W
        tbl_h = Inches(0.4) * n_rows

        tbl_shape = slide.shapes.add_table(n_rows, n_cols, MARGIN_L, tbl_top, tbl_w, tbl_h)
        table = tbl_shape.table

        for ri, row in enumerate(rows):
            for ci, cell_text in enumerate(row):
                if ci >= n_cols:
                    break
                cell = table.cell(ri, ci)
                # Clean bold markers
                clean = cell_text.replace("**", "")
                cell.text = clean
                for para in cell.text_frame.paragraphs:
                    para.font.name = FONT
                    para.font.size = Pt(16)
                    if ri == 0:
                        # Header
                        para.font.color.rgb = WHITE
                        para.font.bold = True
                    elif ri == n_rows - 1 and "**" in cell_text:
                        para.font.bold = True
                        para.font.color.rgb = FG
                    else:
                        para.font.color.rgb = FG

                # Cell fill
                if ri == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = PRIMARY
                elif ri % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = BG_WHITE

        # Bottom text
        if sd.bottom_text:
            bt_top = tbl_top + tbl_h + Inches(0.2)
            self._add_accent_box(slide, sd.bottom_text, MARGIN_L, bt_top, CONTENT_W, Inches(0.8))

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_references(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        tb = self._add_textbox(slide, MARGIN_L, BODY_TOP, CONTENT_W, BODY_H)
        tf = tb.text_frame
        tf.word_wrap = True

        first = True
        for i, (author, title, venue) in enumerate(sd.ref_items, 1):
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()

            run_num = p.add_run()
            run_num.text = f"[{i}] "
            run_num.font.name = FONT
            run_num.font.size = Pt(14)
            run_num.font.color.rgb = FG

            run_auth = p.add_run()
            run_auth.text = author + " "
            run_auth.font.name = FONT
            run_auth.font.size = Pt(14)
            run_auth.font.bold = True
            run_auth.font.color.rgb = FG

            run_title = p.add_run()
            run_title.text = title + " "
            run_title.font.name = FONT
            run_title.font.size = Pt(14)
            run_title.font.italic = True
            run_title.font.color.rgb = FG

            run_venue = p.add_run()
            run_venue.text = venue
            run_venue.font.name = FONT
            run_venue.font.size = Pt(14)
            run_venue.font.color.rgb = MUTED

            p.space_before = Pt(8)

    def build_timeline_h(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.timeline_items
        if not items:
            return

        n = len(items)
        line_y = BODY_TOP + Inches(0.6)
        item_w = (CONTENT_W - Inches(0.5)) / n

        # Horizontal line
        ln = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            MARGIN_L, line_y, CONTENT_W, Pt(3)
        )
        ln.fill.solid()
        ln.fill.fore_color.rgb = RGBColor(0xDE, 0xE2, 0xE6)
        ln.line.fill.background()

        for i, item in enumerate(items):
            cx = MARGIN_L + int(item_w * (i + 0.5))
            color = ACCENT if item.get("highlight") else SECONDARY

            # Circle
            r = Pt(7)
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                cx - r, line_y - r + Pt(1), r * 2, r * 2
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = color
            circle.line.fill.background()

            # Block below
            block_top = line_y + Inches(0.35)
            block_left = MARGIN_L + int(item_w * i) + Inches(0.05)
            block_w = int(item_w) - Inches(0.1)
            block_h = Inches(1.8)

            bg = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                block_left, block_top, block_w, block_h
            )
            if item.get("highlight"):
                bg.fill.solid()
                bg.fill.fore_color.rgb = RGBColor(0xFD, 0xED, 0xEF)
            else:
                bg.fill.solid()
                bg.fill.fore_color.rgb = LIGHT
            bg.line.fill.background()

            # Text in block
            tb = self._add_textbox(slide, block_left + Pt(8), block_top + Pt(8), block_w - Pt(16), block_h - Pt(16))
            tf = tb.text_frame
            tf.word_wrap = True

            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = item["year"]
            run.font.name = FONT
            run.font.size = Pt(18)
            run.font.bold = True
            run.font.color.rgb = PRIMARY

            run2 = p.add_run()
            run2.text = "  " + item["text"]
            run2.font.name = FONT
            run2.font.size = Pt(15)
            run2.font.color.rgb = FG
            if item.get("highlight"):
                run2.font.bold = True

            if item.get("detail"):
                p2 = tf.add_paragraph()
                p2.text = item["detail"]
                p2.font.name = FONT
                p2.font.size = Pt(12)
                p2.font.color.rgb = MUTED
                p2.space_before = Pt(6)

    def build_timeline_v(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.timeline_items
        if not items:
            return

        line_x = MARGIN_L + Inches(0.15)
        n = len(items)
        item_h = min(Inches(1.0), int(BODY_H / n))

        # Vertical line
        ln = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            line_x, BODY_TOP, Pt(3), int(item_h * n)
        )
        ln.fill.solid()
        ln.fill.fore_color.rgb = RGBColor(0xDE, 0xE2, 0xE6)
        ln.line.fill.background()

        for i, item in enumerate(items):
            top = BODY_TOP + int(item_h * i)
            color = ACCENT if item.get("highlight") else SECONDARY

            # Circle
            r = Pt(6)
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                line_x - r + Pt(1), top + Pt(8), r * 2, r * 2
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = color
            circle.line.fill.background()

            # Block
            block_left = MARGIN_L + Inches(0.6)
            block_w = CONTENT_W - Inches(0.8)
            bg = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                block_left, top, block_w, int(item_h) - Inches(0.1)
            )
            if item.get("highlight"):
                bg.fill.solid()
                bg.fill.fore_color.rgb = RGBColor(0xFD, 0xED, 0xEF)
            else:
                bg.fill.solid()
                bg.fill.fore_color.rgb = LIGHT
            bg.line.fill.background()

            tb = self._add_textbox(slide, block_left + Pt(12), top + Pt(6), block_w - Pt(24), int(item_h) - Pt(24))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = item["year"] + "  "
            run.font.name = FONT
            run.font.size = Pt(17)
            run.font.bold = True
            run.font.color.rgb = PRIMARY
            run2 = p.add_run()
            run2.text = item["text"]
            run2.font.name = FONT
            run2.font.size = Pt(16)
            run2.font.color.rgb = FG

            if item.get("detail"):
                p2 = tf.add_paragraph()
                p2.text = item["detail"]
                p2.font.name = FONT
                p2.font.size = Pt(13)
                p2.font.color.rgb = MUTED

    def build_end(self, sd: SlideData):
        slide = self._blank_slide()
        self._set_bg(slide, PRIMARY)

        tb = self._add_textbox(slide, Inches(1), Inches(2), SW - Inches(2), Inches(3))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = sd.h1 or "Thank you"
        p.font.name = FONT
        p.font.size = Pt(50)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

        # Remaining text
        remaining = []
        past_h1 = False
        for line in sd.raw.split("\n"):
            s = line.strip()
            if s.startswith("# ") and not past_h1:
                past_h1 = True
                continue
            if past_h1 and s:
                remaining.append(s)

        for line in remaining:
            p2 = tf.add_paragraph()
            p2.text = line
            p2.font.name = FONT
            p2.font.size = Pt(22)
            p2.font.color.rgb = WHITE
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(8)

    # ---------- Build all ----------

    def build_all(self, slides: list[SlideData]):
        BUILDERS = {
            "title": self.build_title,
            "divider": self.build_divider,
            "cols-2": self.build_columns,
            "cols-2-wide-l": self.build_columns,
            "cols-2-wide-r": self.build_columns,
            "cols-3": self.build_columns,
            "sandwich": self.build_sandwich,
            "equation": self.build_equation,
            "figure": self.build_figure,
            "table-slide": self.build_table,
            "references": self.build_references,
            "timeline-h": self.build_timeline_h,
            "timeline": self.build_timeline_v,
            "end": self.build_end,
        }

        for sd in slides:
            builder = BUILDERS.get(sd.slide_class, self.build_default)
            builder(sd)


# ============================================================
# Main
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Convert Marp academic templates to editable PPTX"
    )
    parser.add_argument("input", help="Input Marp markdown file")
    parser.add_argument("-o", "--output", help="Output .pptx path")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = args.output or str(input_path.with_name(input_path.stem + "_editable.pptx"))

    slides = parse_marp(args.input)
    print(f"Parsed {len(slides)} slides", file=sys.stderr)

    builder = PptxBuilder(base_path=input_path.parent)
    builder.build_all(slides)
    builder.save(output_path)

    print(f"Saved: {output_path}", file=sys.stderr)
    print(f"  {len(slides)} slides, all editable text boxes", file=sys.stderr)


if __name__ == "__main__":
    main()
