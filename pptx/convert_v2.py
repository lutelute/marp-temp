#!/usr/bin/env python3
"""
Marp Academic Template → Editable PPTX converter (v2).

Template-driven: each slide class maps to a dedicated builder
with fixed layout positions. Math is rendered as native OMML
(editable in PowerPoint) via Pandoc, with PNG fallback.

Usage:
    python pptx/convert_v2.py example.md
    python pptx/convert_v2.py example.md -o output.pptx
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
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

import subprocess
import hashlib

try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False

sys.path.insert(0, str(Path(__file__).parent))
from latex_to_omml import latex_to_omml_element, OmmlError  # noqa: E402
from lxml import etree  # noqa: E402

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ============================================================
# Theme loaded from themes/academic.css (single source of truth)
# ============================================================
THEME_CSS_PATH = Path(__file__).resolve().parent.parent / "themes" / "academic.css"

_HEX_RE = re.compile(r"#([0-9a-fA-F]{6})")
_ROOT_RE = re.compile(r":root\s*\{([^}]*)\}", re.DOTALL)
_VAR_RE = re.compile(r"--([\w-]+)\s*:\s*([^;]+);")


def _hex_to_rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _resolve_font(font_stack: str, available: set[str]) -> str:
    """Pick first font in the CSS stack that is actually installed.

    Falls back to the first name in the stack if none are detected (so the
    PPTX still looks right on a machine with the CSS-declared fonts).
    """
    names = [n.strip().strip("'\"") for n in font_stack.split(",")]
    for n in names:
        if n in available:
            return n
    return names[0] if names else "Helvetica Neue"


def _list_installed_fonts() -> set[str]:
    try:
        from matplotlib import font_manager  # type: ignore
        return {f.name for f in font_manager.fontManager.ttflist}
    except Exception:
        return set()


def load_theme(css_path: Path = THEME_CSS_PATH) -> dict:
    """Parse `:root` CSS variables from the Marp theme.

    Returns a dict with keys:
      colors: {name: RGBColor}
      font_body / font_head / font_mono: first-resolvable font name
    """
    text = css_path.read_text(encoding="utf-8")
    m = _ROOT_RE.search(text)
    root = m.group(1) if m else ""
    vars_ = dict(_VAR_RE.findall(root))

    colors: dict[str, RGBColor] = {}
    for k, v in vars_.items():
        if k.startswith("color-"):
            hm = _HEX_RE.search(v)
            if hm:
                colors[k[len("color-"):]] = _hex_to_rgb(hm.group(1))

    installed = _list_installed_fonts()
    fonts = {
        "body": _resolve_font(vars_.get("font-body", ""), installed),
        "head": _resolve_font(vars_.get("font-head", ""), installed),
        "ea":   _resolve_font(vars_.get("font-ea", "Hiragino Sans"), installed),
        "mono": _resolve_font(vars_.get("font-mono", ""), installed),
    }
    return {"colors": colors, "fonts": fonts}


_THEME = load_theme()
_C = _THEME["colors"]

PRIMARY   = _C.get("primary",   RGBColor(0x16, 0x21, 0x3e))
SECONDARY = _C.get("secondary", RGBColor(0x0f, 0x34, 0x60))
ACCENT    = _C.get("accent",    RGBColor(0xe9, 0x45, 0x60))
BG_WHITE  = _C.get("bg",        RGBColor(0xff, 0xff, 0xff))
FG        = _C.get("fg",        RGBColor(0x1a, 0x1a, 0x2e))
MUTED     = _C.get("muted",     RGBColor(0x6c, 0x75, 0x7d))
LIGHT     = _C.get("light",     RGBColor(0xf0, 0xf2, 0xf5))
WHITE     = RGBColor(0xff, 0xff, 0xff)

FONT      = _THEME["fonts"]["body"]
FONT_HEAD = _THEME["fonts"]["head"]
FONT_EA   = _THEME["fonts"]["ea"]
FONT_MONO = _THEME["fonts"]["mono"]

print(f"[theme] latin={FONT}  ea={FONT_EA}  head={FONT_HEAD}", file=sys.stderr)

# Slide dimensions (16:9 standard)
SW = Inches(13.333)
SH = Inches(7.5)

# Common regions
MARGIN_L  = Inches(1.0)
MARGIN_R  = Inches(1.0)
MARGIN_T  = Inches(0.45)
CONTENT_W = SW - MARGIN_L - MARGIN_R
TITLE_H   = Inches(0.45)
TITLE_TOP  = MARGIN_T
BODY_TOP   = MARGIN_T + TITLE_H + Inches(0.12)
BODY_H     = SH - BODY_TOP - Inches(0.6)  # extra bottom for footer

# Font size scale — single place to tune density
SZ_TITLE   = Pt(18)   # slide h1
SZ_H2      = Pt(16)   # sub-heading in body
SZ_H3      = Pt(14)   # h3 in body
SZ_BODY    = Pt(14)   # default body / bullets
SZ_COL     = Pt(13)   # column content
SZ_SMALL   = Pt(12)   # captions, refs, table cells
SZ_FOOT    = Pt(9)    # footnote
SZ_EQ      = Pt(28)   # display equation
SZ_EQ_VAR  = Pt(14)   # equation variable descriptions
SZ_ZONE_L  = Pt(15)   # zone box label
SZ_ZONE_B  = Pt(13)   # zone box body

# ============================================================
# Theme layout config (overridden by --theme flag)
# ============================================================
@dataclass
class ThemeLayout:
    h1_deco: str = "left-bar"           # left-bar | bottom-line | top-line | double-bottom | none
    h1_deco_width: int = 8              # Pt
    h1_deco_color: str = "primary"      # primary | secondary | accent
    title_bg: str = "white"             # white | gradient | dark | light
    title_align: str = "left"           # left | center
    divider_align: str = "left"         # left | center
    end_bg: str = "white"               # white | dark | light
    box_style: str = "border-only"      # border-only | filled | card | accent-border
    box_radius: float = 0.02            # adjustments[0] value
    box_fill: bool = False
    spacing: str = "compact"            # compact | normal | generous

LAYOUT = ThemeLayout()

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
    # For class=="equations": ordered list of (label, latex) pairs, e.g.
    # [("minimize", "f(x) = ..."), ("subject to", "Ax \\le b"), ("", "x \\ge 0")]
    eq_system: list = field(default_factory=list)
    ref_items: list = field(default_factory=list)  # [(author, title, venue), ...]
    # Zone templates
    zone_flow_items: list = field(default_factory=list)     # [{"label", "body"}, ...]
    zone_compare: dict = field(default_factory=dict)        # {"left_label", "left_body", ...}
    zone_matrix: dict = field(default_factory=dict)         # {"x_label", "y_label", "cells": [...]}
    zone_process_items: list = field(default_factory=list)  # [{"step", "title", "body"}, ...]
    # Research presentation templates
    agenda_items: list = field(default_factory=list)         # [str, ...]
    rq_main: str = ""
    rq_sub: str = ""
    summary_points: list = field(default_factory=list)       # [str, ...]
    result_dual_items: list = field(default_factory=list)     # [{"image", "caption"}, ...]
    appendix_label: str = ""
    # Overview / Result / Steps
    overview_text: str = ""
    overview_points: list = field(default_factory=list)  # [str, ...]
    result_text: str = ""
    result_figure: str = ""
    result_caption: str = ""
    result_analysis: list = field(default_factory=list)  # [str, ...]
    steps_items: list = field(default_factory=list)  # [{"num","title","body"}, ...]
    # New slide types (v2)
    quote_text: str = ""
    quote_source: str = ""
    history_items: list = field(default_factory=list)   # [{"year","event"}, ...]
    panorama_text: str = ""
    kpi_items: list = field(default_factory=list)        # [{"value","label"}, ...]
    pros_items: list = field(default_factory=list)       # [str, ...]
    cons_items: list = field(default_factory=list)       # [str, ...]
    def_term: str = ""
    def_body: str = ""
    def_note: str = ""
    gallery_items: list = field(default_factory=list)    # [{"image","caption"}, ...]
    highlight_text: str = ""
    checklist_items: list = field(default_factory=list)  # [{"text","done"}, ...]
    annotation_figure: str = ""
    annotation_notes: list = field(default_factory=list) # [str, ...]
    ba_before: dict = field(default_factory=dict)        # {"label","body"}
    ba_after: dict = field(default_factory=dict)         # {"label","body"}
    funnel_items: list = field(default_factory=list)     # [{"label","value"}, ...]
    stack_items: list = field(default_factory=list)      # [{"name","desc"}, ...]
    card_items: list = field(default_factory=list)       # [{"title","body"}, ...]
    split_left: dict = field(default_factory=dict)       # {"label","body"}
    split_right: dict = field(default_factory=dict)      # {"label","body"}
    code_text: str = ""
    code_desc: str = ""
    multi_result_items: list = field(default_factory=list) # [{"metric","value","desc"}, ...]
    takeaway_main: str = ""
    takeaway_points: list = field(default_factory=list)  # [str, ...]
    profile_name: str = ""
    profile_affiliation: str = ""
    profile_bio: list = field(default_factory=list)      # [str, ...]


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
        sd.h1 = strip_html(h1m.group(1))
    if h2m:
        sd.h2 = strip_html(h2m.group(1))

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

    # --- Equations (multi-line system, e.g. optimization formulation) ---
    elif cls == "equations":
        sys_div = extract_div(content, "eq-system")
        if sys_div:
            # Each row: <div class="row"> <span class="label">minimize</span> $$...$$ </div>
            rows = extract_child_divs(sys_div)
            if rows:
                for row in rows:
                    lm = re.search(r'<span[^>]*class="[^"]*label[^"]*"[^>]*>(.*?)</span>',
                                   row, re.DOTALL)
                    label = strip_html(lm.group(1)) if lm else ""
                    mm = re.search(r"\$\$(.*?)\$\$", row, re.DOTALL)
                    if mm:
                        sd.eq_system.append((label, mm.group(1).strip()))
            else:
                # Fallback: parse bare $$...$$ blocks sequentially, with optional
                # leading <span class="label">…</span>.
                # Split on $$...$$ boundaries while tracking preceding labels.
                pattern = re.compile(
                    r'(?:<span[^>]*class="[^"]*label[^"]*"[^>]*>(.*?)</span>\s*)?\$\$(.*?)\$\$',
                    re.DOTALL,
                )
                for m in pattern.finditer(sys_div):
                    label = strip_html(m.group(1) or "")
                    sd.eq_system.append((label, m.group(2).strip()))

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

    # --- Zone: flow ---
    elif cls == "zone-flow":
        container = extract_div(content, "zf-container")
        if container:
            for box in extract_child_divs(container):
                lbl = re.search(r'class="[^"]*zf-label[^"]*"[^>]*>(.*?)</span>', box, re.DOTALL)
                bod = re.search(r'class="[^"]*zf-body[^"]*"[^>]*>(.*?)</span>', box, re.DOTALL)
                sd.zone_flow_items.append({
                    "label": strip_html(lbl.group(1)) if lbl else "",
                    "body": strip_html(bod.group(1)) if bod else "",
                })
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Zone: compare ---
    elif cls == "zone-compare":
        for side in ("zc-left", "zc-right"):
            div = extract_div(content, side)
            prefix = "left" if "left" in side else "right"
            if div:
                lbl = re.search(r'class="[^"]*zc-label[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                bod = re.search(r'class="[^"]*zc-body[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                sd.zone_compare[f"{prefix}_label"] = strip_html(lbl.group(1)) if lbl else ""
                sd.zone_compare[f"{prefix}_body"] = strip_html(bod.group(1)) if bod else ""
        vs = extract_div(content, "zc-vs")
        sd.zone_compare["vs_text"] = strip_html(vs) if vs else "VS"
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Zone: matrix ---
    elif cls == "zone-matrix":
        container = extract_div(content, "zm-container")
        xl = extract_div(content, "zm-xlabel")
        yl = extract_div(content, "zm-ylabel")
        sd.zone_matrix["x_label"] = strip_html(xl) if xl else ""
        sd.zone_matrix["y_label"] = strip_html(yl) if yl else ""
        cells = []
        for pos in ("zm-tl", "zm-tr", "zm-bl", "zm-br"):
            cell = extract_div(content, pos)
            if cell:
                lbl = re.search(r'class="[^"]*zm-label[^"]*"[^>]*>(.*?)</span>', cell, re.DOTALL)
                bod = re.search(r'class="[^"]*zm-body[^"]*"[^>]*>(.*?)</span>', cell, re.DOTALL)
                cells.append({
                    "label": strip_html(lbl.group(1)) if lbl else "",
                    "body": strip_html(bod.group(1)) if bod else "",
                })
            else:
                cells.append({"label": "", "body": ""})
        sd.zone_matrix["cells"] = cells
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Zone: process ---
    elif cls == "zone-process":
        container = extract_div(content, "zp-container")
        if container:
            for step_div in extract_child_divs(container):
                nm = re.search(r'class="[^"]*zp-num[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                ti = re.search(r'class="[^"]*zp-title[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                bo = re.search(r'class="[^"]*zp-body[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                sd.zone_process_items.append({
                    "step": strip_html(nm.group(1)) if nm else "",
                    "title": strip_html(ti.group(1)) if ti else "",
                    "body": strip_html(bo.group(1)) if bo else "",
                })
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Agenda ---
    elif cls == "agenda":
        agenda = extract_div(content, "agenda-list")
        if agenda:
            for m in re.finditer(r"\d+\.\s*(.+)", agenda):
                sd.agenda_items.append(strip_html(m.group(1).strip()))

    # --- Research Question ---
    elif cls == "rq":
        main = extract_div(content, "rq-main")
        if main:
            sd.rq_main = strip_html(main)
        sub = extract_div(content, "rq-sub")
        if sub:
            sd.rq_sub = strip_html(sub)

    # --- Result Dual ---
    elif cls == "result-dual":
        results = extract_div(content, "results")
        if results:
            items = extract_child_divs(results)
            for item in items:
                img_m = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", item)
                cap = extract_div(item, "caption")
                sd.result_dual_items.append({
                    "image": img_m.group(1) if img_m else "",
                    "caption": strip_html(cap) if cap else "",
                })

    # --- Summary ---
    elif cls == "summary":
        # Extract <li> items from summary-points
        sp = extract_div(content, "summary-points")
        if not sp:
            # Fallback: look for <ol class="summary-points">
            sp_m = re.search(r'<ol[^>]*class="[^"]*summary-points[^"]*"[^>]*>(.*?)</ol>',
                             content, re.DOTALL)
            sp = sp_m.group(1) if sp_m else ""
        if sp:
            for li_m in re.finditer(r"<li>(.*?)</li>", sp, re.DOTALL):
                sd.summary_points.append(strip_html(li_m.group(1)))

    # --- Appendix ---
    elif cls == "appendix":
        lbl = re.search(r'class="[^"]*appendix-label[^"]*"[^>]*>(.*?)</span>', content, re.DOTALL)
        if lbl:
            sd.appendix_label = strip_html(lbl.group(1))
        # Parse body as default (table, text, etc.)
        body = content
        if h1m:
            body = body[:h1m.start()] + body[h1m.end():]
        if h2m:
            body = body[:h2m.start()] + body[h2m.end():]
        for tag in ("appendix-label",):
            pattern = rf'<span\s+class="[^"]*{tag}[^"]*"[^>]*>.*?</span>'
            body = re.sub(pattern, "", body, flags=re.DOTALL)
        # Parse table if present
        rows = []
        for line in body.split("\n"):
            s = line.strip()
            if s.startswith("|") and not re.match(r"^\|[-:|]+\|$", s):
                cells = [c.strip() for c in s.strip("|").split("|")]
                rows.append(cells)
        if rows:
            sd.table_rows = rows
        else:
            sd.body_lines = parse_markdown_lines(body)

    # --- Overview ---
    elif cls == "overview":
        lead = extract_div(content, "ov-lead")
        if lead:
            sd.overview_text = strip_html(lead)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)
        pts = extract_div(content, "ov-points")
        if pts:
            for li in re.finditer(r"<li>(.*?)</li>", pts, re.DOTALL):
                sd.overview_points.append(strip_html(li.group(1)))
            if not sd.overview_points:
                for line in pts.split("\n"):
                    s = line.strip()
                    if s.startswith("- ") or s.startswith("* "):
                        sd.overview_points.append(s[2:].strip())
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Result ---
    elif cls == "result":
        lead = extract_div(content, "rs-lead")
        if lead:
            sd.result_text = strip_html(lead)
        fig = extract_div(content, "rs-figure")
        if fig:
            img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", fig)
            if img:
                sd.result_figure = img.group(1)
            cap = extract_div(fig, "caption")
            if cap:
                sd.result_caption = strip_html(cap)
        analysis = extract_div(content, "rs-analysis")
        if analysis:
            for li in re.finditer(r"<li>(.*?)</li>", analysis, re.DOTALL):
                sd.result_analysis.append(strip_html(li.group(1)))
            if not sd.result_analysis:
                for line in analysis.split("\n"):
                    s = line.strip()
                    if s.startswith("- ") or s.startswith("* "):
                        sd.result_analysis.append(s[2:].strip())
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Steps (horizontal) ---
    elif cls == "steps":
        container = extract_div(content, "st-container")
        if container:
            for step_div in extract_child_divs(container):
                nm = re.search(r'class="[^"]*st-num[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                ti = re.search(r'class="[^"]*st-title[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                bo = re.search(r'class="[^"]*st-body[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                sd.steps_items.append({
                    "num": strip_html(nm.group(1)) if nm else "",
                    "title": strip_html(ti.group(1)) if ti else "",
                    "body": strip_html(bo.group(1)) if bo else "",
                })
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    # --- Quote ---
    elif cls == "quote":
        qt = extract_div(content, "qt-text")
        if qt:
            sd.quote_text = strip_html(qt)
        qs = extract_div(content, "qt-source")
        if qs:
            sd.quote_source = strip_html(qs)

    # --- History ---
    elif cls == "history":
        container = extract_div(content, "hs-container")
        if container:
            for item in extract_child_divs(container):
                ym = re.search(r'class="[^"]*hs-year[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                em = re.search(r'class="[^"]*hs-event[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.history_items.append({
                    "year": strip_html(ym.group(1)) if ym else "",
                    "event": strip_html(em.group(1)) if em else "",
                })

    # --- Panorama ---
    elif cls == "panorama":
        pn = extract_div(content, "pn-text")
        if pn:
            sd.panorama_text = strip_html(pn)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

    # --- KPI ---
    elif cls == "kpi":
        container = extract_div(content, "kpi-container")
        if container:
            for item in extract_child_divs(container):
                vm = re.search(r'class="[^"]*kpi-value[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                lm = re.search(r'class="[^"]*kpi-label[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.kpi_items.append({
                    "value": strip_html(vm.group(1)) if vm else "",
                    "label": strip_html(lm.group(1)) if lm else "",
                })

    # --- Pros-Cons ---
    elif cls == "pros-cons":
        pros = extract_div(content, "pc-pros")
        if pros:
            for li in re.finditer(r"<li>(.*?)</li>", pros, re.DOTALL):
                sd.pros_items.append(strip_html(li.group(1)))
        cons = extract_div(content, "pc-cons")
        if cons:
            for li in re.finditer(r"<li>(.*?)</li>", cons, re.DOTALL):
                sd.cons_items.append(strip_html(li.group(1)))

    # --- Definition ---
    elif cls == "definition":
        dt = extract_div(content, "df-term")
        if dt:
            sd.def_term = strip_html(dt)
        db = extract_div(content, "df-body")
        if db:
            sd.def_body = strip_html(db)
        dn = extract_div(content, "df-note")
        if dn:
            sd.def_note = strip_html(dn)

    # --- Diagram ---
    elif cls == "diagram":
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)

    # --- Gallery-Img ---
    elif cls == "gallery-img":
        container = extract_div(content, "gi-container")
        if container:
            for item in extract_child_divs(container):
                img_m = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", item)
                cap = extract_div(item, "gi-caption")
                sd.gallery_items.append({
                    "image": img_m.group(1) if img_m else "",
                    "caption": strip_html(cap) if cap else "",
                })

    # --- Highlight ---
    elif cls == "highlight":
        hl = extract_div(content, "hl-text")
        if hl:
            sd.highlight_text = strip_html(hl)

    # --- Checklist ---
    elif cls == "checklist":
        container = extract_div(content, "cl-container")
        if container:
            for li in re.finditer(r'<li(\s+class="done")?>(.*?)</li>', container, re.DOTALL):
                sd.checklist_items.append({
                    "text": strip_html(li.group(2)),
                    "done": li.group(1) is not None,
                })

    # --- Annotation ---
    elif cls == "annotation":
        fig = extract_div(content, "an-figure")
        if fig:
            img_m = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", fig)
            if img_m:
                sd.annotation_figure = img_m.group(1)
        notes = extract_div(content, "an-notes")
        if notes:
            for li in re.finditer(r"<li>(.*?)</li>", notes, re.DOTALL):
                sd.annotation_notes.append(strip_html(li.group(1)))

    # --- Before-After ---
    elif cls == "before-after":
        for prefix, div_cls in [("ba_before", "ba-before"), ("ba_after", "ba-after")]:
            div = extract_div(content, div_cls)
            if div:
                lm = re.search(r'class="[^"]*ba-label[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                bm = re.search(r'class="[^"]*ba-body[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                setattr(sd, prefix, {
                    "label": strip_html(lm.group(1)) if lm else "",
                    "body": strip_html(bm.group(1)) if bm else "",
                })

    # --- Funnel ---
    elif cls == "funnel":
        container = extract_div(content, "fn-container")
        if container:
            for item in extract_child_divs(container):
                lm = re.search(r'class="[^"]*fn-label[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                vm = re.search(r'class="[^"]*fn-value[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.funnel_items.append({
                    "label": strip_html(lm.group(1)) if lm else "",
                    "value": strip_html(vm.group(1)) if vm else "",
                })

    # --- Stack ---
    elif cls == "stack":
        container = extract_div(content, "sk-container")
        if container:
            for item in extract_child_divs(container):
                nm = re.search(r'class="[^"]*sk-name[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                dm = re.search(r'class="[^"]*sk-desc[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.stack_items.append({
                    "name": strip_html(nm.group(1)) if nm else "",
                    "desc": strip_html(dm.group(1)) if dm else "",
                })

    # --- Card-Grid ---
    elif cls == "card-grid":
        container = extract_div(content, "cg-container")
        if container:
            for item in extract_child_divs(container):
                tm = re.search(r'class="[^"]*cg-title[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                bm = re.search(r'class="[^"]*cg-body[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.card_items.append({
                    "title": strip_html(tm.group(1)) if tm else "",
                    "body": strip_html(bm.group(1)) if bm else "",
                })

    # --- Split-Text ---
    elif cls == "split-text":
        for prefix, div_cls in [("split_left", "sp-left"), ("split_right", "sp-right")]:
            div = extract_div(content, div_cls)
            if div:
                lm = re.search(r'class="[^"]*sp-label[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                bm = re.search(r'class="[^"]*sp-body[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                setattr(sd, prefix, {
                    "label": strip_html(lm.group(1)) if lm else "",
                    "body": strip_html(bm.group(1)) if bm else "",
                })

    # --- Code ---
    elif cls == "code":
        cd = extract_div(content, "cd-code")
        if cd:
            # Extract text from code block, preserving content
            code_m = re.search(r"```[\w]*\n(.*?)```", cd, re.DOTALL)
            if code_m:
                sd.code_text = code_m.group(1).rstrip()
            else:
                sd.code_text = strip_html(cd)
        desc = extract_div(content, "cd-desc")
        if desc:
            sd.code_desc = strip_html(desc)

    # --- Multi-Result ---
    elif cls == "multi-result":
        container = extract_div(content, "mr-container")
        if container:
            for item in extract_child_divs(container):
                mm = re.search(r'class="[^"]*mr-metric[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                vm = re.search(r'class="[^"]*mr-value[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                dm = re.search(r'class="[^"]*mr-desc[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.multi_result_items.append({
                    "metric": strip_html(mm.group(1)) if mm else "",
                    "value": strip_html(vm.group(1)) if vm else "",
                    "desc": strip_html(dm.group(1)) if dm else "",
                })

    # --- Takeaway ---
    elif cls == "takeaway":
        ta = extract_div(content, "ta-main")
        if ta:
            sd.takeaway_main = strip_html(ta)
        pts = extract_div(content, "ta-points")
        if pts:
            for li in re.finditer(r"<li>(.*?)</li>", pts, re.DOTALL):
                sd.takeaway_points.append(strip_html(li.group(1)))

    # --- Profile ---
    elif cls == "profile":
        container = extract_div(content, "pf-container")
        if container:
            nm = extract_div(container, "pf-name")
            if nm:
                sd.profile_name = strip_html(nm)
            af = extract_div(container, "pf-affiliation")
            if af:
                sd.profile_affiliation = strip_html(af)
            bio = extract_div(container, "pf-bio")
            if bio:
                for li in re.finditer(r"<li>(.*?)</li>", bio, re.DOTALL):
                    sd.profile_bio.append(strip_html(li.group(1)))
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

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
        self._math_cache_dir = Path(tempfile.mkdtemp(prefix="marp_math_"))
        self._render_script = base_path / "pptx" / "render_math.mjs"
        if not self._render_script.exists():
            # Try relative to this script
            self._render_script = Path(__file__).parent / "render_math.mjs"

    def save(self, path: str):
        self._ensure_ea_font()
        self.prs.save(path)

    def _ensure_ea_font(self):
        """Inject <a:ea> / <a:cs> alongside <a:latin> so Japanese text is
        rendered in the same font as Latin text. python-pptx only sets
        <a:latin> via `font.name`, so without this, Japanese falls back to
        the theme default.

        Child order inside <a:rPr>: ... latin, ea, cs, sym, ...
        """
        rpr_tags = (f"{{{NS_A}}}rPr", f"{{{NS_A}}}defRPr", f"{{{NS_A}}}endParaRPr")
        for slide in self.prs.slides:
            root = slide._element
            for tag in rpr_tags:
                for rpr in root.iter(tag):
                    self._patch_rpr(rpr)

    @staticmethod
    def _patch_rpr(rpr):
        latin = rpr.find(f"{{{NS_A}}}latin")
        if latin is None:
            return
        ea = rpr.find(f"{{{NS_A}}}ea")
        if ea is None:
            ea = etree.Element(f"{{{NS_A}}}ea")
            ea.set("typeface", FONT_EA)
            latin.addnext(ea)
        else:
            ea.set("typeface", FONT_EA)
        cs = rpr.find(f"{{{NS_A}}}cs")
        if cs is None:
            cs = etree.Element(f"{{{NS_A}}}cs")
            cs.set("typeface", FONT_EA)
            ea.addnext(cs)
        else:
            cs.set("typeface", FONT_EA)

    def _omml_element(self, latex: str, display: bool):
        """Try to convert LaTeX to OMML element. Returns lxml Element or None."""
        try:
            return latex_to_omml_element(latex, display=display)
        except OmmlError as e:
            print(f"  OMML failed: {latex[:40]}... ({e}) — falling back to PNG", file=sys.stderr)
            return None

    def _add_math_omml_display(self, slide, latex: str, left, top, width, pt_size: int = 28):
        """Insert a display-mode equation as native OMML inside a textbox.

        The box hugs the expected content height (roughly 1.6× font size for
        radicals / fractions). Vertical anchor is TOP so the visual top of the
        equation matches `top`, which lets callers position follow-up elements
        without guessing around hidden padding.
        """
        el = self._omml_element(latex, display=True)
        if el is None:
            return None
        height = Pt(int(pt_size * 1.7))
        tb = self._add_textbox(slide, left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        # Seed run properties so OMML inherits font size/name
        run = p.add_run()
        run.text = ""
        run.font.name = FONT
        run.font.size = Pt(pt_size)
        run.font.color.rgb = FG
        p._p.append(el)
        return tb

    def _append_math_omml_inline(self, para, latex: str, size, color):
        """Append inline (non-display) OMML into an existing paragraph.

        Returns True on success, False on failure (caller should text-fallback).
        """
        el = self._omml_element(latex, display=False)
        if el is None:
            return False
        # Anchor font via an empty run right before the math element
        run = para.add_run()
        run.text = ""
        run.font.name = FONT
        run.font.size = size
        if color is not None:
            run.font.color.rgb = color
        para._p.append(el)
        return True

    def _render_math(self, latex: str, display: bool = False, fontsize: int = 28) -> str | None:
        """Render LaTeX to PNG via KaTeX + Playwright. Returns path to PNG."""
        key = hashlib.md5(f"{latex}:{display}:{fontsize}".encode()).hexdigest()
        png_path = self._math_cache_dir / f"{key}.png"
        if png_path.exists():
            return str(png_path)

        if not self._render_script.exists():
            return None

        cmd = [
            "node", str(self._render_script),
            latex, str(png_path),
            f"--fontsize={fontsize}",
        ]
        if display:
            cmd.append("--display")

        try:
            subprocess.run(cmd, capture_output=True, timeout=15, check=True)
            if png_path.exists():
                return str(png_path)
        except (subprocess.TimeoutExpired, subprocess.CalledProcessError) as e:
            print(f"  Math render failed: {latex[:40]}... ({e})", file=sys.stderr)
        return None

    def _add_math_image(self, slide, latex: str, left, top, max_width, display=True, fontsize=28):
        """Render LaTeX and insert as image. Returns (width, height) in EMU or None."""
        png = self._render_math(latex, display=display, fontsize=fontsize)
        if not png:
            return None
        from PIL import Image
        with Image.open(png) as im:
            iw, ih = im.size
        # Scale: 1 CSS px ≈ 12700 EMU at 96 DPI, but screenshots are 1:1 device px
        # Use DPI-aware scaling
        dpi = 96
        pw = int(iw * 914400 / dpi)
        ph = int(ih * 914400 / dpi)
        # Scale down if too wide
        if pw > max_width:
            scale = max_width / pw
            pw = int(pw * scale)
            ph = int(ph * scale)
        img_left = left + (max_width - pw) // 2  # center
        slide.shapes.add_picture(png, img_left, top, pw, ph)
        return (pw, ph)

    def _blank_slide(self):
        layout = self.prs.slide_layouts[6]  # Blank
        return self.prs.slides.add_slide(layout)

    def _add_textbox(self, slide, left, top, width, height):
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.auto_size = MSO_AUTO_SIZE.NONE
        # Zero out inner padding so the box hugs the text. Extra whitespace
        # around content should come from positioning, never from hidden
        # frame margins. Individual call sites can opt back in if needed.
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        return tb

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

    def _add_title(self, slide, text, top=None, color=None):
        if color is None:
            color = PRIMARY
        if top is None:
            top = TITLE_TOP
        deco_color_map = {"primary": PRIMARY, "secondary": SECONDARY, "accent": ACCENT}
        deco_c = deco_color_map.get(LAYOUT.h1_deco_color, PRIMARY)
        deco_w = Pt(LAYOUT.h1_deco_width)
        text_left = MARGIN_L
        text_w = CONTENT_W

        if LAYOUT.h1_deco == "left-bar":
            bar_h = TITLE_H
            # 3-color bar (6:3:1)
            for frac, fc in [(0.6, PRIMARY), (0.3, SECONDARY), (0.1, ACCENT)]:
                y_off = sum(bar_h * f for f in [0.6, 0.3, 0.1][:[ (0.6,PRIMARY),(0.3,SECONDARY),(0.1,ACCENT)].index((frac,fc))])
                pass
            # Simpler: single solid bar in deco color
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, int(MARGIN_L), int(top), int(deco_w), int(TITLE_H))
            bar.fill.solid()
            bar.fill.fore_color.rgb = deco_c
            bar.line.fill.background()
            text_left = int(MARGIN_L + deco_w + Pt(10))
            text_w = int(CONTENT_W - deco_w - Pt(10))

        elif LAYOUT.h1_deco == "bottom-line":
            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, int(MARGIN_L), int(top + TITLE_H - Pt(2)),
                int(CONTENT_W), int(deco_w))
            line.fill.solid()
            line.fill.fore_color.rgb = deco_c
            line.line.fill.background()

        elif LAYOUT.h1_deco == "top-line":
            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, int(MARGIN_L), int(top), int(CONTENT_W), int(deco_w))
            line.fill.solid()
            line.fill.fore_color.rgb = deco_c
            line.line.fill.background()
            top = int(top + deco_w + Pt(4))

        elif LAYOUT.h1_deco == "double-bottom":
            for offset in [0, Pt(6)]:
                line = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, int(MARGIN_L),
                    int(top + TITLE_H - Pt(2) + offset),
                    int(CONTENT_W), Pt(2))
                line.fill.solid()
                line.fill.fore_color.rgb = deco_c
                line.line.fill.background()

        # else: "none" — no decoration

        tb = self._add_textbox(slide, int(text_left), int(top), int(text_w), TITLE_H)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = FONT_HEAD
        p.font.size = SZ_TITLE
        p.font.bold = True
        p.font.color.rgb = color
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        return tb

    def _add_para(self, tf, text, size=None, color=FG, bold=False, italic=False, space_before=Pt(4)):
        if size is None:
            size = SZ_BODY
        p = tf.add_paragraph()
        p.text = text
        p.font.name = FONT
        p.font.size = size
        p.font.color.rgb = color
        p.font.bold = bold
        # p.font.italic = italic  # disabled by design
        p.space_before = space_before
        return p

    def _add_body_text(self, slide, lines, left=None, top=None, width=None, height=None, size=None):
        if size is None:
            size = SZ_BODY
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
                p.text = strip_html(s[3:])
                p.font.name = FONT_HEAD
                p.font.size = SZ_H2
                p.font.bold = True
                p.font.color.rgb = SECONDARY
                p.space_before = Pt(10)
            elif is_h3:
                p.text = strip_html(s[4:])
                p.font.name = FONT_HEAD
                p.font.size = SZ_H3
                p.font.bold = True
                p.font.color.rgb = MUTED
                p.space_before = Pt(6)
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

        # Hug the content: body textboxes should not carry layout whitespace
        # as hidden padding. They are edit targets for text, not spacers.
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        return tb

    def _set_rich_text(self, para, text, size=None, color=FG):
        if size is None:
            size = SZ_BODY
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
        self._set_rich_text(tf.paragraphs[0], text, SZ_BODY, FG)
        return tb

    def _add_conclusion_box(self, slide, text, left, top, width, height):
        """Bordered box with text."""
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        bg.adjustments[0] = 0.02
        bg.fill.solid()
        bg.fill.fore_color.rgb = LIGHT
        bg.line.fill.background()
        tb = self._add_textbox(slide, left + Pt(16), top + Pt(10), width - Pt(32), height - Pt(20))
        tf = tb.text_frame
        tf.word_wrap = True
        self._set_rich_text(tf.paragraphs[0], text, SZ_BODY, FG)
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
        p.font.size = SZ_FOOT
        p.font.color.rgb = MUTED
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    def _add_zone_box(self, slide, left, top, width, height,
                      label="", body="", fill_color=LIGHT,
                      label_size=None, body_size=None):
        if label_size is None:
            label_size = SZ_ZONE_L
        if body_size is None:
            body_size = SZ_ZONE_B
        """Rectangular zone: shape + overlaid label/body textbox."""
        bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        bg.adjustments[0] = LAYOUT.box_radius

        if LAYOUT.box_style == "filled" or LAYOUT.box_style == "card":
            bg.fill.solid()
            bg.fill.fore_color.rgb = fill_color
            if LAYOUT.box_style == "card":
                bg.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
                bg.line.width = Pt(1)
            else:
                bg.line.fill.background()
        elif LAYOUT.box_style == "accent-border":
            bg.fill.background()
            bg.line.color.rgb = ACCENT
            bg.line.width = Pt(1.5)
        else:  # border-only
            bg.fill.background()
            bg.line.color.rgb = RGBColor(0xE8, 0xE8, 0xE8)
            bg.line.width = Pt(0.75)
        pad = Pt(14)
        tb = self._add_textbox(
            slide, left + pad, top + pad, width - pad * 2, height - pad * 2)
        tf = tb.text_frame
        tf.word_wrap = True
        if label:
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = label
            run.font.name = FONT_HEAD
            run.font.size = label_size
            run.font.bold = True
            run.font.color.rgb = SECONDARY
        if body:
            p2 = tf.add_paragraph() if label else tf.paragraphs[0]
            p2.space_before = Pt(6)
            self._set_text_with_inline_math(p2, body, body_size, FG)
        return bg, tb

    # ---------- Slide type builders ----------

    def build_title(self, sd: SlideData):
        slide = self._blank_slide()

        # Background
        if LAYOUT.title_bg == "gradient":
            self._set_gradient_bg(slide, PRIMARY, SECONDARY)
        elif LAYOUT.title_bg == "dark":
            self._set_bg(slide, PRIMARY)
        elif LAYOUT.title_bg == "light":
            self._set_bg(slide, LIGHT)
        # else: white (default)

        is_dark = LAYOUT.title_bg in ("gradient", "dark")
        align = PP_ALIGN.CENTER if LAYOUT.title_align == "center" else PP_ALIGN.LEFT
        h_color = WHITE if is_dark else PRIMARY
        sub_color = RGBColor(0xCC, 0xCC, 0xCC) if is_dark else MUTED

        # Title
        tb = self._add_textbox(slide, Inches(1), Inches(1.5), SW - Inches(2), Inches(2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sd.h1
        p.font.name = FONT_HEAD
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = h_color
        p.alignment = align

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
            p.font.color.rgb = sub_color
            p.alignment = align
            p.space_before = Pt(6)

    def build_divider(self, sd: SlideData):
        slide = self._blank_slide()

        align = PP_ALIGN.CENTER if LAYOUT.divider_align == "center" else PP_ALIGN.LEFT
        x = Inches(1) if LAYOUT.divider_align == "center" else Inches(1.5)
        w = SW - Inches(2) if LAYOUT.divider_align == "center" else SW - Inches(3)

        tb = self._add_textbox(slide, x, Inches(2.5), w, Inches(1.5))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = sd.h1
        p.font.name = FONT_HEAD
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = PRIMARY
        p.alignment = align

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

        eq_top = BODY_TOP + Inches(0.2)

        # Main equation — prefer native OMML, fall back to KaTeX PNG, then plain text.
        omml_box = self._add_math_omml_display(
            slide, sd.eq_main, MARGIN_L, eq_top, CONTENT_W, pt_size=28
        )
        if omml_box is not None:
            var_top = eq_top + omml_box.height + Inches(0.25)
        else:
            result = self._add_math_image(
                slide, sd.eq_main,
                MARGIN_L, eq_top, CONTENT_W,
                display=True, fontsize=36
            )
            if result:
                _, eq_h = result
                var_top = eq_top + eq_h + Inches(0.3)
            else:
                tb = self._add_textbox(slide, MARGIN_L, eq_top, CONTENT_W, Inches(1.2))
                tf = tb.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = sd.eq_main
                p.font.name = FONT
                p.font.size = Pt(28)
                p.font.color.rgb = FG
                p.alignment = PP_ALIGN.CENTER
                var_top = eq_top + Inches(1.6)

        # Variable descriptions — render each symbol as math image
        if sd.eq_vars:
            desc_left = Inches(2.0)
            desc_w = Inches(9)
            # Row must fit tall constructs (√, fractions, matrices). Keep padding.
            row_h = Inches(0.58)

            for vi, (sym, desc) in enumerate(sd.eq_vars):
                row_top = var_top + int(row_h * vi)

                sym_latex = sym.strip()
                if sym_latex.startswith("$"):
                    sym_latex = sym_latex.strip("$")

                # Render symbol as native OMML inline math in a right-aligned textbox.
                stb = self._add_textbox(slide, desc_left, row_top, Inches(2.0), row_h)
                stf = stb.text_frame
                stf.vertical_anchor = MSO_ANCHOR.MIDDLE
                stf.margin_top = 0
                stf.margin_bottom = 0
                sp = stf.paragraphs[0]
                sp.alignment = PP_ALIGN.RIGHT
                if not self._append_math_omml_inline(sp, sym_latex, Pt(22), SECONDARY):
                    # Fallback: PNG image
                    sym_png = self._render_math(sym_latex, display=False, fontsize=22)
                    if sym_png:
                        from PIL import Image
                        with Image.open(sym_png) as im:
                            sw, sh = im.size
                        pw = int(sw * 914400 / 96)
                        ph = int(sh * 914400 / 96)
                        max_sym_w = Inches(1.8)
                        if pw > max_sym_w:
                            scale = max_sym_w / pw
                            pw = int(pw * scale)
                            ph = int(ph * scale)
                        sym_right = desc_left + Inches(2.0)
                        sym_x = sym_right - pw
                        sym_y = row_top + (row_h - ph) // 2
                        slide.shapes.add_picture(sym_png, sym_x, sym_y, pw, ph)
                        # Remove empty textbox
                        sp_tf = stb
                        sp_tf.element.getparent().remove(sp_tf.element)
                    else:
                        sp.text = sym
                        sp.font.name = FONT
                        sp.font.size = Pt(18)
                        sp.font.bold = True
                        sp.font.color.rgb = SECONDARY
                        sp.alignment = PP_ALIGN.RIGHT

                # Description text
                dtb = self._add_textbox(slide, desc_left + Inches(2.3), row_top, Inches(6.5), row_h)
                dtf = dtb.text_frame
                dtf.word_wrap = True
                dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
                dtf.margin_top = 0
                dtf.margin_bottom = 0
                # Parse inline math in description
                self._set_text_with_inline_math(dtf.paragraphs[0], desc, Pt(17), FG)

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_equations(self, sd: SlideData):
        """Multi-equation slide (e.g. optimization problem formulation).

        Layout: each row has a left label (e.g. "minimize", "subject to")
        and a centered display equation. Optional eq-desc table below.
        """
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        n = len(sd.eq_system)
        if n == 0:
            return

        # Layout geometry
        top = BODY_TOP + Inches(0.1)
        label_left = MARGIN_L + Inches(0.3)
        label_w = Inches(1.9)
        eq_left = label_left + label_w + Inches(0.2)
        eq_w = CONTENT_W - (eq_left - MARGIN_L) - Inches(0.3)

        # Adjust per-row height so the system fits the available space.
        vars_h = Inches(0.58) * max(len(sd.eq_vars), 0)
        footnote_h = Inches(0.4) if sd.footnote else Inches(0)
        avail_h = BODY_H - vars_h - footnote_h - Inches(0.3)
        row_h = min(Inches(1.3), max(Inches(0.75), int(avail_h / n)))
        pt_size = 30 if n <= 3 else (26 if n <= 4 else 22)

        for i, (label, latex) in enumerate(sd.eq_system):
            row_top = top + int(row_h * i)

            # Label (e.g. "minimize", "subject to")
            if label:
                ltb = self._add_textbox(slide, label_left, row_top, label_w, row_h)
                ltf = ltb.text_frame
                ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
                ltf.margin_top = 0
                ltf.margin_bottom = 0
                ltf.word_wrap = True
                lp = ltf.paragraphs[0]
                lp.text = label
                lp.font.name = FONT
                lp.font.size = Pt(max(14, pt_size - 10))
                lp.font.italic = False
                lp.font.color.rgb = SECONDARY
                lp.alignment = PP_ALIGN.RIGHT

            # Equation — prefer native OMML, fall back to KaTeX PNG, then text.
            el = self._omml_element(latex, display=True)
            etb = self._add_textbox(slide, eq_left, row_top, eq_w, row_h)
            etf = etb.text_frame
            etf.word_wrap = True
            etf.vertical_anchor = MSO_ANCHOR.MIDDLE
            etf.margin_top = 0
            etf.margin_bottom = 0
            ep = etf.paragraphs[0]
            ep.alignment = PP_ALIGN.LEFT
            erun = ep.add_run()
            erun.text = ""
            erun.font.name = FONT
            erun.font.size = Pt(pt_size)
            erun.font.color.rgb = FG
            if el is not None:
                ep._p.append(el)
            else:
                # PNG fallback: delete the textbox and place an image.
                etb.element.getparent().remove(etb.element)
                result = self._add_math_image(
                    slide, latex,
                    eq_left, row_top, eq_w,
                    display=True, fontsize=pt_size,
                )
                if not result:
                    # Plain-text last resort
                    tb2 = self._add_textbox(slide, eq_left, row_top, eq_w, row_h)
                    tf2 = tb2.text_frame
                    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
                    p2 = tf2.paragraphs[0]
                    p2.text = latex
                    p2.font.name = FONT
                    p2.font.size = Pt(pt_size - 4)
                    p2.font.color.rgb = FG

        # Variable descriptions (same layout as build_equation)
        if sd.eq_vars:
            var_top = top + int(row_h * n) + Inches(0.25)
            desc_left = Inches(2.0)
            row_h_v = Inches(0.52)

            for vi, (sym, desc) in enumerate(sd.eq_vars):
                row_top = var_top + int(row_h_v * vi)

                sym_latex = sym.strip()
                if sym_latex.startswith("$"):
                    sym_latex = sym_latex.strip("$")

                stb = self._add_textbox(slide, desc_left, row_top, Inches(2.0), row_h_v)
                stf = stb.text_frame
                stf.vertical_anchor = MSO_ANCHOR.MIDDLE
                stf.margin_top = 0
                stf.margin_bottom = 0
                sp = stf.paragraphs[0]
                sp.alignment = PP_ALIGN.RIGHT
                if not self._append_math_omml_inline(sp, sym_latex, Pt(18), SECONDARY):
                    sp.text = sym
                    sp.font.name = FONT
                    sp.font.size = Pt(16)
                    sp.font.bold = True
                    sp.font.color.rgb = SECONDARY

                dtb = self._add_textbox(slide, desc_left + Inches(2.3), row_top, Inches(6.5), row_h_v)
                dtf = dtb.text_frame
                dtf.word_wrap = True
                dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
                dtf.margin_top = 0
                dtf.margin_bottom = 0
                self._set_text_with_inline_math(dtf.paragraphs[0], desc, Pt(15), FG)

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def _set_text_with_inline_math(self, para, text, size, color):
        """Set paragraph with mixed text + inline OMML math on `$...$` segments."""
        para.clear()
        # Split into alternating text / math chunks. Odd indices are math.
        parts = re.split(r'\$([^$]+)\$', text)
        for i, chunk in enumerate(parts):
            if not chunk:
                continue
            if i % 2 == 0:
                run = para.add_run()
                run.text = chunk
                run.font.name = FONT
                run.font.size = size
                if color is not None:
                    run.font.color.rgb = color
            else:
                if not self._append_math_omml_inline(para, chunk, size, color):
                    # Fallback: render the LaTeX as plain text
                    run = para.add_run()
                    run.text = chunk
                    run.font.name = FONT
                    run.font.size = size
                    if color is not None:
                        run.font.color.rgb = color
        run.font.color.rgb = color

    def _add_column_content(self, slide, lines, left, top, width, height, size=None):
        if size is None:
            size = SZ_COL
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

        sub_top = BODY_TOP
        if sd.h2:
            tb = self._add_textbox(slide, MARGIN_L, sub_top, CONTENT_W, Inches(0.35))
            tf = tb.text_frame
            p = tf.paragraphs[0]
            p.text = sd.h2
            p.font.name = FONT
            p.font.size = SZ_H2
            p.font.bold = True
            p.font.color.rgb = SECONDARY
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            sub_top += Inches(0.4)

        if not sd.table_rows:
            return

        rows = sd.table_rows
        n_rows = len(rows)
        n_cols = max(len(r) for r in rows) if rows else 0

        tbl_top = sub_top + Inches(0.1)
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
                    para.font.size = SZ_SMALL
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
            run_title.font.italic = False
            run_title.font.color.rgb = FG

            run_venue = p.add_run()
            run_venue.text = venue
            run_venue.font.name = FONT
            run_venue.font.size = Pt(14)
            run_venue.font.color.rgb = MUTED

            p.space_before = Pt(8)

        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

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
            bg.adjustments[0] = 0.02
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
            bg.adjustments[0] = 0.02
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
        if LAYOUT.end_bg == "dark":
            self._set_bg(slide, PRIMARY)
        elif LAYOUT.end_bg == "light":
            self._set_bg(slide, LIGHT)

        is_dark = LAYOUT.end_bg == "dark"
        h_color = WHITE if is_dark else PRIMARY
        sub_color = RGBColor(0xCC, 0xCC, 0xCC) if is_dark else MUTED

        tb = self._add_textbox(slide, Inches(1), Inches(2), SW - Inches(2), Inches(3))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = sd.h1 or "Thank you"
        p.font.name = FONT_HEAD
        p.font.size = Pt(50)
        p.font.bold = True
        p.font.color.rgb = h_color
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
            p2.font.color.rgb = sub_color
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(8)

    # ---------- Zone builders ----------

    def build_zone_flow(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.zone_flow_items
        n = len(items)
        if n == 0:
            return

        gap = Inches(0.25)
        arrow_w = Inches(0.45)
        total_arrows = (n - 1) * (arrow_w + gap)
        box_w = (CONTENT_W - total_arrows - gap * (n - 1)) / n
        box_h = BODY_H - Inches(0.3)
        y = BODY_TOP + Inches(0.15)

        for i, item in enumerate(items):
            x = MARGIN_L + i * (box_w + gap + arrow_w + gap)
            if i > 0:
                x = MARGIN_L + i * box_w + i * gap + (i - 1) * (arrow_w + gap) + arrow_w + gap

            # Compute x properly
            x = MARGIN_L
            for j in range(i):
                x += box_w + gap
                if j < n - 1:
                    x += arrow_w + gap

            self._add_zone_box(slide, int(x), int(y), int(box_w), int(box_h),
                               label=item["label"], body=item["body"],
                               label_size=Pt(20), body_size=Pt(16))

            # Arrow after box (except last)
            if i < n - 1:
                ax = int(x + box_w + gap)
                ay = int(y + box_h / 2 - Inches(0.2))
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.NOTCHED_RIGHT_ARROW,
                    ax, ay, int(arrow_w), Inches(0.4))
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = MUTED
                arrow.line.fill.background()

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_zone_compare(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        zc = sd.zone_compare
        if not zc:
            return

        gap = Inches(0.8)
        box_w = (CONTENT_W - gap) / 2
        box_h = BODY_H - Inches(0.3)
        y = BODY_TOP + Inches(0.15)
        x_left = MARGIN_L
        x_right = MARGIN_L + box_w + gap

        self._add_zone_box(slide, int(x_left), int(y), int(box_w), int(box_h),
                           label=zc.get("left_label", ""), body=zc.get("left_body", ""),
                           label_size=Pt(22), body_size=Pt(17))

        self._add_zone_box(slide, int(x_right), int(y), int(box_w), int(box_h),
                           label=zc.get("right_label", ""), body=zc.get("right_body", ""),
                           label_size=Pt(22), body_size=Pt(17))

        # VS badge
        vs_text = zc.get("vs_text", "VS")
        badge_d = Inches(0.7)
        bx = int(MARGIN_L + box_w + gap / 2 - badge_d / 2)
        by = int(y + box_h / 2 - badge_d / 2)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, bx, by, int(badge_d), int(badge_d))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT
        badge.line.fill.background()
        vtb = self._add_textbox(slide, bx, by, int(badge_d), int(badge_d))
        vtf = vtb.text_frame
        vtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        vp = vtf.paragraphs[0]
        vp.text = vs_text
        vp.font.name = FONT_HEAD
        vp.font.size = Pt(16)
        vp.font.bold = True
        vp.font.color.rgb = WHITE
        vp.alignment = PP_ALIGN.CENTER

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_zone_matrix(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        zm = sd.zone_matrix
        cells = zm.get("cells", [])
        if len(cells) < 4:
            return

        y_margin = Inches(0.6)  # left strip for Y-axis label
        x_margin = Inches(0.45)  # bottom strip for X-axis label
        gap = Inches(0.15)
        grid_left = MARGIN_L + y_margin
        grid_top = BODY_TOP + Inches(0.1)
        grid_w = CONTENT_W - y_margin
        grid_h = BODY_H - x_margin - Inches(0.2)
        cell_w = (grid_w - gap) / 2
        cell_h = (grid_h - gap) / 2

        positions = [
            (grid_left, grid_top),                        # TL
            (grid_left + cell_w + gap, grid_top),         # TR
            (grid_left, grid_top + cell_h + gap),         # BL
            (grid_left + cell_w + gap, grid_top + cell_h + gap),  # BR
        ]

        for (cx, cy), cell in zip(positions, cells):
            self._add_zone_box(slide, int(cx), int(cy), int(cell_w), int(cell_h),
                               label=cell["label"], body=cell["body"],
                               label_size=Pt(18), body_size=Pt(15))

        # Cross-hair lines
        mid_x = int(grid_left + cell_w + gap / 2 - Pt(1))
        mid_y = int(grid_top + cell_h + gap / 2 - Pt(1))
        vline = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, mid_x, int(grid_top), Pt(2), int(grid_h))
        vline.fill.solid()
        vline.fill.fore_color.rgb = MUTED
        vline.line.fill.background()
        hline = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, int(grid_left), mid_y, int(grid_w), Pt(2))
        hline.fill.solid()
        hline.fill.fore_color.rgb = MUTED
        hline.line.fill.background()

        # Y-axis label (rotated)
        if zm.get("y_label"):
            ytb = self._add_textbox(
                slide, int(MARGIN_L), int(grid_top), int(y_margin - Inches(0.1)), int(grid_h))
            ytf = ytb.text_frame
            ytf.vertical_anchor = MSO_ANCHOR.MIDDLE
            yp = ytf.paragraphs[0]
            yp.text = zm["y_label"]
            yp.font.name = FONT_HEAD
            yp.font.size = Pt(15)
            yp.font.bold = True
            yp.font.color.rgb = SECONDARY
            yp.alignment = PP_ALIGN.CENTER
            ytb.rotation = 270.0

        # X-axis label
        if zm.get("x_label"):
            xtb = self._add_textbox(
                slide, int(grid_left), int(grid_top + grid_h + Inches(0.1)),
                int(grid_w), int(x_margin))
            xp = xtb.text_frame.paragraphs[0]
            xp.text = zm["x_label"]
            xp.font.name = FONT_HEAD
            xp.font.size = Pt(15)
            xp.font.bold = True
            xp.font.color.rgb = SECONDARY
            xp.alignment = PP_ALIGN.CENTER

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_zone_process(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.zone_process_items
        n = len(items)
        if n == 0:
            return

        gap = Inches(0.15)
        circle_d = Inches(0.5)
        item_h = (BODY_H - gap * (n - 1)) / n
        box_left = MARGIN_L + circle_d + Inches(0.25)
        box_w = CONTENT_W - circle_d - Inches(0.25)

        for i, item in enumerate(items):
            y = BODY_TOP + i * (item_h + gap)

            # Number circle
            cx = int(MARGIN_L)
            cy = int(y + item_h / 2 - circle_d / 2)
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, cy, int(circle_d), int(circle_d))
            circle.fill.solid()
            circle.fill.fore_color.rgb = PRIMARY
            circle.line.fill.background()
            ctb = self._add_textbox(slide, cx, cy, int(circle_d), int(circle_d))
            ctf = ctb.text_frame
            ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
            cp = ctf.paragraphs[0]
            cp.text = item["step"]
            cp.font.name = FONT_HEAD
            cp.font.size = Pt(20)
            cp.font.bold = True
            cp.font.color.rgb = WHITE
            cp.alignment = PP_ALIGN.CENTER

            # Connector line (except last)
            if i < n - 1:
                lx = int(MARGIN_L + circle_d / 2 - Pt(1))
                ly = int(y + item_h / 2 + circle_d / 2)
                lh = int(item_h + gap - circle_d)
                conn = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, lx, ly, Pt(2), lh)
                conn.fill.solid()
                conn.fill.fore_color.rgb = MUTED
                conn.line.fill.background()

            # Content box
            self._add_zone_box(slide, int(box_left), int(y), int(box_w), int(item_h),
                               label=item["title"], body=item["body"],
                               label_size=Pt(19), body_size=Pt(16))

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    # ---------- Research presentation builders ----------

    def build_agenda(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.agenda_items
        if not items:
            return

        y = BODY_TOP + Inches(0.1)
        for i, item in enumerate(items, 1):
            row_h = Inches(0.55)
            row_y = y + row_h * (i - 1)

            # Number circle
            cd = Inches(0.42)
            cx = int(MARGIN_L + Inches(0.3))
            cy = int(row_y + (row_h - cd) / 2)
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, cy, int(cd), int(cd))
            circle.fill.solid()
            circle.fill.fore_color.rgb = SECONDARY
            circle.line.fill.background()
            ctb = self._add_textbox(slide, cx, cy, int(cd), int(cd))
            ctf = ctb.text_frame
            ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
            cp = ctf.paragraphs[0]
            cp.text = str(i)
            cp.font.name = FONT_HEAD
            cp.font.size = Pt(18)
            cp.font.bold = True
            cp.font.color.rgb = WHITE
            cp.alignment = PP_ALIGN.CENTER

            # Item text
            tx = int(MARGIN_L + Inches(1.0))
            tb = self._add_textbox(slide, tx, int(row_y), int(CONTENT_W - Inches(1.0)), int(row_h))
            tf = tb.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = item
            p.font.name = FONT
            p.font.size = Pt(22)
            p.font.color.rgb = FG

    def build_rq(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        if sd.rq_main:
            tb = self._add_textbox(
                slide, Inches(1.5), BODY_TOP + Inches(0.8),
                SW - Inches(3), Inches(2.5))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            self._set_text_with_inline_math(p, sd.rq_main, Pt(28), PRIMARY)
            p.font.bold = True

        if sd.rq_sub:
            tb2 = self._add_textbox(
                slide, Inches(2), BODY_TOP + Inches(3.5),
                SW - Inches(4), Inches(1))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            self._set_text_with_inline_math(p2, sd.rq_sub, Pt(18), MUTED)

    def build_result_dual(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.result_dual_items
        n = min(len(items), 2)
        if n == 0:
            return

        gap = Inches(0.4)
        col_w = (CONTENT_W - gap) / n
        img_h = BODY_H - Inches(1.2)

        for i, item in enumerate(items[:2]):
            x = MARGIN_L + i * (col_w + gap)
            img_path = self._resolve_image(item["image"]) if item["image"] else None
            if img_path:
                from PIL import Image
                with Image.open(img_path) as im:
                    iw, ih = im.size
                pw = int(iw * 914400 / 96)
                ph = int(ih * 914400 / 96)
                max_w = int(col_w - Inches(0.2))
                max_h = int(img_h)
                scale = min(max_w / pw, max_h / ph, 1.0)
                pw = int(pw * scale)
                ph = int(ph * scale)
                ix = int(x + (col_w - pw) / 2)
                slide.shapes.add_picture(img_path, ix, int(BODY_TOP), pw, ph)
                cap_y = int(BODY_TOP + ph + Inches(0.15))
            else:
                cap_y = int(BODY_TOP + Inches(3.5))

            if item.get("caption"):
                ctb = self._add_textbox(slide, int(x), cap_y, int(col_w), Inches(0.6))
                cp = ctb.text_frame.paragraphs[0]
                cp.text = item["caption"]
                cp.font.name = FONT
                cp.font.size = Pt(12)
                cp.font.color.rgb = FG
                cp.alignment = PP_ALIGN.CENTER

    def build_summary(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.summary_points
        if not items:
            return

        y = BODY_TOP + Inches(0.1)
        row_h = min(Inches(1.0), (BODY_H - Inches(0.3)) / len(items))

        for i, text in enumerate(items):
            row_y = y + int(row_h * i)

            # Accent left border
            bdr = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                int(MARGIN_L + Inches(0.3)), int(row_y + Pt(4)),
                Pt(4), int(row_h - Pt(8)))
            bdr.fill.solid()
            bdr.fill.fore_color.rgb = ACCENT
            bdr.line.fill.background()

            # Number
            ntb = self._add_textbox(
                slide, int(MARGIN_L + Inches(0.55)), int(row_y), Inches(0.5), int(row_h))
            ntf = ntb.text_frame
            ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
            np = ntf.paragraphs[0]
            np.text = str(i + 1) + "."
            np.font.name = FONT_HEAD
            np.font.size = Pt(22)
            np.font.bold = True
            np.font.color.rgb = PRIMARY

            # Text
            ttb = self._add_textbox(
                slide, int(MARGIN_L + Inches(1.2)), int(row_y),
                int(CONTENT_W - Inches(1.2)), int(row_h))
            ttf = ttb.text_frame
            ttf.word_wrap = True
            ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
            self._set_text_with_inline_math(
                ttf.paragraphs[0], text, Pt(20), FG)

    def build_appendix(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1, color=MUTED)

        # Appendix label in top-right
        if sd.appendix_label:
            ltb = self._add_textbox(
                slide, SW - Inches(2.5), MARGIN_T, Inches(2), Inches(0.4))
            lp = ltb.text_frame.paragraphs[0]
            lp.text = sd.appendix_label
            lp.font.name = FONT
            lp.font.size = Pt(11)
            lp.font.color.rgb = MUTED
            lp.alignment = PP_ALIGN.RIGHT

        # Reuse table builder if table data exists
        if sd.table_rows:
            self._build_table_content(slide, sd)
        elif sd.body_lines:
            self._add_body_text(slide, sd.body_lines, size=Pt(17))

    def _build_table_content(self, slide, sd: SlideData):
        """Shared table rendering used by build_table and build_appendix."""
        rows = sd.table_rows
        if not rows:
            return
        n_cols = max(len(r) for r in rows)
        n_rows = len(rows)
        tbl_w = CONTENT_W
        tbl_h = min(BODY_H - Inches(0.5), Pt(40) * n_rows)
        tbl_top = BODY_TOP + Inches(0.1)
        graphic_frame = slide.shapes.add_table(
            n_rows, n_cols, int(MARGIN_L), int(tbl_top), int(tbl_w), int(tbl_h)
        )
        table = graphic_frame.table
        for ri, row in enumerate(rows):
            for ci, cell_text in enumerate(row):
                if ci >= n_cols:
                    break
                cell = table.cell(ri, ci)
                cell.text = strip_html(cell_text)
                for para in cell.text_frame.paragraphs:
                    para.font.name = FONT
                    para.font.size = Pt(13)
                if ri == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = PRIMARY
                    for para in cell.text_frame.paragraphs:
                        para.font.color.rgb = WHITE
                        para.font.bold = True
                elif ri % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = BG_WHITE

    # ---------- Overview / Result / Steps ----------

    def build_overview(self, sd: SlideData):
        """Overview: lead text → figure → key points."""
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        cur_y = BODY_TOP

        # Lead text
        if sd.overview_text:
            tb = self._add_textbox(slide, MARGIN_L, int(cur_y), CONTENT_W, Inches(0.8))
            tf = tb.text_frame
            tf.word_wrap = True
            self._set_text_with_inline_math(tf.paragraphs[0], sd.overview_text, SZ_BODY, FG)
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            cur_y += Inches(0.9)

        # Figure (centered)
        if sd.image_path:
            img_file = self._resolve_image(sd.image_path)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(CONTENT_W * 0.7)
                max_h = int(Inches(3.0))
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                ix = int(MARGIN_L + (CONTENT_W - pw) / 2)
                slide.shapes.add_picture(img_file, ix, int(cur_y), pw, ph)
                cur_y += ph + Inches(0.1)

        # Caption
        if sd.caption:
            ctb = self._add_textbox(slide, MARGIN_L, int(cur_y), CONTENT_W, Inches(0.3))
            cp = ctb.text_frame.paragraphs[0]
            cp.text = sd.caption
            cp.font.name = FONT
            cp.font.size = SZ_SMALL
            cp.font.color.rgb = MUTED
            cp.alignment = PP_ALIGN.CENTER
            cur_y += Inches(0.4)

        # Key points
        if sd.overview_points:
            pts_left = MARGIN_L + Inches(0.5)
            for i, pt in enumerate(sd.overview_points):
                py = int(cur_y + Inches(0.32) * i)
                # Bullet dot
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                    int(pts_left), int(py + Pt(4)), Pt(6), Pt(6))
                dot.fill.solid()
                dot.fill.fore_color.rgb = SECONDARY
                dot.line.fill.background()
                # Text
                ptb = self._add_textbox(slide, int(pts_left + Pt(14)), py,
                    int(CONTENT_W - Inches(0.8)), Inches(0.3))
                self._set_text_with_inline_math(
                    ptb.text_frame.paragraphs[0], pt, SZ_BODY, FG)

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_result(self, sd: SlideData):
        """Result: lead text → left figure + right analysis."""
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        cur_y = BODY_TOP

        # Lead text (full width)
        if sd.result_text:
            tb = self._add_textbox(slide, MARGIN_L, int(cur_y), CONTENT_W, Inches(0.8))
            tf = tb.text_frame
            tf.word_wrap = True
            self._set_text_with_inline_math(tf.paragraphs[0], sd.result_text, SZ_BODY, FG)
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            cur_y += Inches(0.85)

        # Two-column: figure left, analysis right
        gap = Inches(0.4)
        left_w = CONTENT_W * 0.5
        right_w = CONTENT_W - left_w - gap
        right_x = MARGIN_L + left_w + gap

        # Left: figure
        if sd.result_figure:
            img_file = self._resolve_image(sd.result_figure)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(left_w)
                max_h = int(Inches(3.5))
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                slide.shapes.add_picture(img_file, int(MARGIN_L), int(cur_y), pw, ph)
                # Caption below figure
                if sd.result_caption:
                    ctb = self._add_textbox(slide, int(MARGIN_L), int(cur_y + ph + Pt(4)),
                        int(left_w), Inches(0.3))
                    cp = ctb.text_frame.paragraphs[0]
                    cp.text = sd.result_caption
                    cp.font.name = FONT
                    cp.font.size = Pt(10)
                    cp.font.color.rgb = MUTED
                    cp.alignment = PP_ALIGN.CENTER

        # Right: analysis points
        if sd.result_analysis:
            for i, text in enumerate(sd.result_analysis):
                py = int(cur_y + Inches(0.38) * i)
                # Bullet
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                    int(right_x), int(py + Pt(4)), Pt(5), Pt(5))
                dot.fill.solid()
                dot.fill.fore_color.rgb = ACCENT
                dot.line.fill.background()
                # Text
                atb = self._add_textbox(slide, int(right_x + Pt(12)), py,
                    int(right_w - Pt(12)), Inches(0.35))
                atf = atb.text_frame
                atf.word_wrap = True
                self._set_text_with_inline_math(atf.paragraphs[0], text, SZ_COL, FG)

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_steps(self, sd: SlideData):
        """Horizontal steps: numbered boxes in a row with connectors."""
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.steps_items
        n = len(items)
        if n == 0:
            return

        gap = Inches(0.15)
        connector_w = Inches(0.3)
        total_connectors = (n - 1) * (connector_w + gap * 2)
        box_w = (CONTENT_W - total_connectors) / n
        box_h = BODY_H - Inches(0.2)
        y = BODY_TOP + Inches(0.1)

        x = MARGIN_L
        for i, item in enumerate(items):
            # Number badge at top
            badge_d = Inches(0.4)
            bx = int(x + box_w / 2 - badge_d / 2)
            by = int(y)
            badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, bx, by, int(badge_d), int(badge_d))
            badge.fill.solid()
            badge.fill.fore_color.rgb = PRIMARY
            badge.line.fill.background()
            btb = self._add_textbox(slide, bx, by, int(badge_d), int(badge_d))
            btf = btb.text_frame
            btf.vertical_anchor = MSO_ANCHOR.MIDDLE
            bp = btf.paragraphs[0]
            bp.text = item["num"]
            bp.font.name = FONT_HEAD
            bp.font.size = Pt(16)
            bp.font.bold = True
            bp.font.color.rgb = WHITE
            bp.alignment = PP_ALIGN.CENTER

            # Content box below badge
            box_y = int(y + badge_d + Inches(0.1))
            box_inner_h = int(box_h - badge_d - Inches(0.1))
            self._add_zone_box(slide, int(x), box_y, int(box_w), box_inner_h,
                               label=item["title"], body=item["body"],
                               label_size=SZ_ZONE_L, body_size=SZ_ZONE_B)

            # Connector arrow (except last)
            if i < n - 1:
                ax = int(x + box_w + gap)
                ay = int(y + box_h / 2)
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW,
                    ax, int(ay - Inches(0.1)), int(connector_w), Inches(0.2))
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = MUTED
                arrow.line.fill.background()

            x += box_w + gap + connector_w + gap

        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    # ---------- New slide type builders (v2) ----------

    def build_quote(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        # Large opening quote mark
        qtb = self._add_textbox(slide, MARGIN_L + Inches(0.5), BODY_TOP, Inches(1.5), Inches(1.2))
        qp = qtb.text_frame.paragraphs[0]
        qp.text = "\u275D"
        qp.font.name = FONT_HEAD
        qp.font.size = Pt(72)
        qp.font.color.rgb = LIGHT

        # Quote text with left border
        q_left = MARGIN_L + Inches(1.0)
        q_top = BODY_TOP + Inches(0.8)
        q_w = CONTENT_W - Inches(2.0)
        q_h = Inches(3.0)
        # Left border bar
        bdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
            int(q_left), int(q_top), Pt(4), int(q_h))
        bdr.fill.solid()
        bdr.fill.fore_color.rgb = SECONDARY
        bdr.line.fill.background()
        # Text
        tb = self._add_textbox(slide, int(q_left + Pt(20)), int(q_top + Pt(8)),
            int(q_w), int(q_h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = sd.quote_text
        p.font.name = FONT
        p.font.size = Pt(22)
        p.font.color.rgb = FG
        p.font.italic = False

        # Source attribution (right-aligned)
        if sd.quote_source:
            stb = self._add_textbox(slide, int(q_left), int(q_top + q_h + Inches(0.3)),
                int(q_w + Pt(20)), Inches(0.5))
            sp = stb.text_frame.paragraphs[0]
            sp.text = "-- " + sd.quote_source
            sp.font.name = FONT
            sp.font.size = Pt(16)
            sp.font.color.rgb = MUTED
            sp.alignment = PP_ALIGN.RIGHT

    def build_history(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.history_items
        if not items:
            return

        n = len(items)
        row_h = min(Inches(0.8), int((BODY_H - Inches(0.2)) / n))
        year_w = Inches(1.5)
        event_left = MARGIN_L + year_w + Inches(0.3)
        event_w = CONTENT_W - year_w - Inches(0.3)

        for i, item in enumerate(items):
            y = BODY_TOP + int(row_h * i)

            # Year
            ytb = self._add_textbox(slide, int(MARGIN_L), int(y), int(year_w), int(row_h))
            ytf = ytb.text_frame
            ytf.vertical_anchor = MSO_ANCHOR.MIDDLE
            yp = ytf.paragraphs[0]
            yp.text = item["year"]
            yp.font.name = FONT_HEAD
            yp.font.size = Pt(20)
            yp.font.bold = True
            yp.font.color.rgb = PRIMARY
            yp.alignment = PP_ALIGN.RIGHT

            # Event
            etb = self._add_textbox(slide, int(event_left), int(y), int(event_w), int(row_h))
            etf = etb.text_frame
            etf.word_wrap = True
            etf.vertical_anchor = MSO_ANCHOR.MIDDLE
            ep = etf.paragraphs[0]
            ep.text = item["event"]
            ep.font.name = FONT
            ep.font.size = Pt(16)
            ep.font.color.rgb = FG

            # Horizontal line between items
            if i < n - 1:
                ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                    int(MARGIN_L), int(y + row_h - Pt(1)),
                    int(CONTENT_W), Pt(1))
                ln.fill.solid()
                ln.fill.fore_color.rgb = RGBColor(0xDE, 0xE2, 0xE6)
                ln.line.fill.background()

    def build_panorama(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        text_w = int(CONTENT_W * 0.4)
        img_w = int(CONTENT_W * 0.55)
        gap = int(CONTENT_W * 0.05)

        # Left text
        if sd.panorama_text:
            tb = self._add_textbox(slide, MARGIN_L, BODY_TOP, text_w, BODY_H)
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = sd.panorama_text
            p.font.name = FONT
            p.font.size = Pt(18)
            p.font.color.rgb = FG

        # Right image
        if sd.image_path:
            img_file = self._resolve_image(sd.image_path)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = img_w
                max_h = int(BODY_H)
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                ix = int(MARGIN_L + text_w + gap + (img_w - pw) / 2)
                iy = int(BODY_TOP + (BODY_H - ph) / 2)
                slide.shapes.add_picture(img_file, ix, iy, pw, ph)

    def build_kpi(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.kpi_items
        if not items:
            return

        n = len(items)
        gap = Inches(0.3)
        col_w = (CONTENT_W - gap * (n - 1)) / n
        y = BODY_TOP + Inches(0.5)

        for i, item in enumerate(items):
            x = MARGIN_L + i * (col_w + gap)

            # Large value
            vtb = self._add_textbox(slide, int(x), int(y), int(col_w), Inches(2.0))
            vtf = vtb.text_frame
            vtf.vertical_anchor = MSO_ANCHOR.BOTTOM
            vp = vtf.paragraphs[0]
            vp.text = item["value"]
            vp.font.name = FONT_HEAD
            vp.font.size = Pt(48)
            vp.font.bold = True
            vp.font.color.rgb = PRIMARY
            vp.alignment = PP_ALIGN.CENTER

            # Label below
            ltb = self._add_textbox(slide, int(x), int(y + Inches(2.2)), int(col_w), Inches(0.8))
            ltf = ltb.text_frame
            ltf.word_wrap = True
            lp = ltf.paragraphs[0]
            lp.text = item["label"]
            lp.font.name = FONT
            lp.font.size = Pt(16)
            lp.font.color.rgb = MUTED
            lp.alignment = PP_ALIGN.CENTER

    def build_pros_cons(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        gap = Inches(0.6)
        col_w = (CONTENT_W - gap) / 2
        y = BODY_TOP + Inches(0.1)

        # Pros column (left)
        self._build_pros_cons_col(slide, sd.pros_items, MARGIN_L, y, col_w,
            "Pros", SECONDARY, "\u2713")

        # Cons column (right)
        self._build_pros_cons_col(slide, sd.cons_items, int(MARGIN_L + col_w + gap), y, col_w,
            "Cons", ACCENT, "\u2717")

    def _build_pros_cons_col(self, slide, items, x, y, w, header, color, bullet_char):
        # Header
        htb = self._add_textbox(slide, int(x), int(y), int(w), Inches(0.5))
        hp = htb.text_frame.paragraphs[0]
        hp.text = header
        hp.font.name = FONT_HEAD
        hp.font.size = Pt(20)
        hp.font.bold = True
        hp.font.color.rgb = color
        hp.alignment = PP_ALIGN.CENTER

        # Items
        itb = self._add_textbox(slide, int(x), int(y + Inches(0.6)), int(w), int(BODY_H - Inches(0.8)))
        itf = itb.text_frame
        itf.word_wrap = True
        first = True
        for item in items:
            if first:
                p = itf.paragraphs[0]
                first = False
            else:
                p = itf.add_paragraph()
            run = p.add_run()
            run.text = bullet_char + "  " + item
            run.font.name = FONT
            run.font.size = Pt(16)
            run.font.color.rgb = FG
            p.space_before = Pt(8)

    def build_definition(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        cur_y = BODY_TOP + Inches(0.3)

        # Term
        if sd.def_term:
            tb = self._add_textbox(slide, MARGIN_L, int(cur_y), CONTENT_W, Inches(0.8))
            p = tb.text_frame.paragraphs[0]
            p.text = sd.def_term
            p.font.name = FONT_HEAD
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = PRIMARY
            cur_y += Inches(1.0)

        # Definition body
        if sd.def_body:
            tb = self._add_textbox(slide, MARGIN_L + Inches(0.3), int(cur_y),
                CONTENT_W - Inches(0.3), Inches(2.5))
            tf = tb.text_frame
            tf.word_wrap = True
            self._set_text_with_inline_math(tf.paragraphs[0], sd.def_body, Pt(20), FG)
            cur_y += Inches(2.8)

        # Note
        if sd.def_note:
            tb = self._add_textbox(slide, MARGIN_L + Inches(0.3), int(cur_y),
                CONTENT_W - Inches(0.3), Inches(1.0))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = sd.def_note
            p.font.name = FONT
            p.font.size = Pt(14)
            p.font.color.rgb = MUTED
            p.font.italic = False

    def build_diagram(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        img_file = self._resolve_image(sd.image_path) if sd.image_path else None
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(CONTENT_W * 0.95)
            max_h = int(BODY_H * 0.85)
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
            pw = int(iw * scale * 914400 / 96)
            ph = int(ih * scale * 914400 / 96)
            left = (SW - pw) // 2
            top = int(BODY_TOP + (BODY_H * 0.85 - ph) / 2)
            slide.shapes.add_picture(img_file, left, top, pw, ph)
            cap_top = top + ph + Inches(0.1)
        else:
            cap_top = BODY_TOP + int(BODY_H * 0.85)

        if sd.caption:
            tb = self._add_textbox(slide, MARGIN_L, int(cap_top), CONTENT_W, Inches(0.5))
            p = tb.text_frame.paragraphs[0]
            p.text = sd.caption
            p.font.name = FONT
            p.font.size = SZ_SMALL
            p.font.color.rgb = MUTED
            p.alignment = PP_ALIGN.CENTER

    def build_gallery_img(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.gallery_items
        if not items:
            return

        gap = Inches(0.2)
        cols = 2
        rows = 2
        cell_w = (CONTENT_W - gap) / cols
        cell_h = (BODY_H - gap - Inches(0.2)) / rows

        for i, item in enumerate(items[:4]):
            r = i // cols
            c = i % cols
            x = MARGIN_L + c * (cell_w + gap)
            y = BODY_TOP + r * (cell_h + gap)

            img_file = self._resolve_image(item["image"]) if item.get("image") else None
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(cell_w - Inches(0.1))
                max_h = int(cell_h - Inches(0.5))
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                ix = int(x + (cell_w - pw) / 2)
                slide.shapes.add_picture(img_file, ix, int(y), pw, ph)
                cap_y = int(y + ph + Pt(4))
            else:
                cap_y = int(y + cell_h - Inches(0.4))

            if item.get("caption"):
                ctb = self._add_textbox(slide, int(x), cap_y, int(cell_w), Inches(0.35))
                cp = ctb.text_frame.paragraphs[0]
                cp.text = item["caption"]
                cp.font.name = FONT
                cp.font.size = Pt(11)
                cp.font.color.rgb = MUTED
                cp.alignment = PP_ALIGN.CENTER

    def build_highlight(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        tb = self._add_textbox(slide, Inches(1.5), BODY_TOP + Inches(0.5),
            SW - Inches(3), BODY_H - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        self._set_text_with_inline_math(p, sd.highlight_text, Pt(36), PRIMARY)
        p.font.bold = True

    def build_checklist(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.checklist_items
        if not items:
            return

        tb = self._add_textbox(slide, MARGIN_L + Inches(0.3), BODY_TOP, CONTENT_W - Inches(0.3), BODY_H)
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for item in items:
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()
            check = "\u2611" if item["done"] else "\u2610"
            color = MUTED if item["done"] else FG
            run = p.add_run()
            run.text = check + "  " + item["text"]
            run.font.name = FONT
            run.font.size = Pt(18)
            run.font.color.rgb = color
            if item["done"]:
                run.font.italic = False
            p.space_before = Pt(10)

    def build_annotation(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        left_w = int(CONTENT_W * 0.5)
        gap = Inches(0.4)
        right_x = int(MARGIN_L + left_w + gap)
        right_w = int(CONTENT_W - left_w - gap)

        # Left figure
        if sd.annotation_figure:
            img_file = self._resolve_image(sd.annotation_figure)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = left_w
                max_h = int(BODY_H)
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                ix = int(MARGIN_L + (left_w - pw) / 2)
                iy = int(BODY_TOP + (BODY_H - ph) / 2)
                slide.shapes.add_picture(img_file, ix, iy, pw, ph)

        # Right numbered notes
        if sd.annotation_notes:
            tb = self._add_textbox(slide, right_x, BODY_TOP, right_w, BODY_H)
            tf = tb.text_frame
            tf.word_wrap = True
            first = True
            for i, note in enumerate(sd.annotation_notes, 1):
                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()
                run = p.add_run()
                run.text = f"{i}. {note}"
                run.font.name = FONT
                run.font.size = Pt(15)
                run.font.color.rgb = FG
                p.space_before = Pt(10)

    def build_before_after(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        gap = Inches(0.8)
        col_w = (CONTENT_W - gap) / 2
        y = BODY_TOP + Inches(0.1)
        box_h = BODY_H - Inches(0.3)

        # Before (left)
        ba = sd.ba_before
        self._add_zone_box(slide, int(MARGIN_L), int(y), int(col_w), int(box_h),
            label=ba.get("label", "Before"), body=ba.get("body", ""),
            label_size=Pt(22), body_size=Pt(17))

        # Arrow between
        arrow_w = Inches(0.5)
        ax = int(MARGIN_L + col_w + gap / 2 - arrow_w / 2)
        ay = int(y + box_h / 2 - Inches(0.15))
        arrow = slide.shapes.add_shape(MSO_SHAPE.NOTCHED_RIGHT_ARROW,
            ax, ay, int(arrow_w), Inches(0.3))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = ACCENT
        arrow.line.fill.background()

        # After (right)
        ba2 = sd.ba_after
        self._add_zone_box(slide, int(MARGIN_L + col_w + gap), int(y), int(col_w), int(box_h),
            label=ba2.get("label", "After"), body=ba2.get("body", ""),
            label_size=Pt(22), body_size=Pt(17))

    def build_funnel(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.funnel_items
        if not items:
            return

        n = len(items)
        gap = Inches(0.08)
        total_h = BODY_H - Inches(0.2)
        row_h = (total_h - gap * (n - 1)) / n
        max_w = CONTENT_W
        min_w = CONTENT_W * 0.3

        for i, item in enumerate(items):
            # Width narrows progressively
            frac = i / max(n - 1, 1)
            w = int(max_w - (max_w - min_w) * frac)
            x = int(MARGIN_L + (CONTENT_W - w) / 2)
            y = int(BODY_TOP + i * (row_h + gap))

            # Rectangle
            rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                x, y, w, int(row_h))
            rect.adjustments[0] = 0.04
            # Color gradient from SECONDARY to ACCENT
            rect.fill.solid()
            r1, g1, b1 = SECONDARY
            r2, g2, b2 = ACCENT
            r = int(r1 + (r2 - r1) * frac)
            g = int(g1 + (g2 - g1) * frac)
            b = int(b1 + (b2 - b1) * frac)
            rect.fill.fore_color.rgb = RGBColor(r, g, b)
            rect.line.fill.background()

            # Text overlay
            tb = self._add_textbox(slide, x + Pt(16), y, w - Pt(32), int(row_h))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            run_l = p.add_run()
            run_l.text = item["label"]
            run_l.font.name = FONT_HEAD
            run_l.font.size = Pt(16)
            run_l.font.bold = True
            run_l.font.color.rgb = WHITE
            if item.get("value"):
                run_v = p.add_run()
                run_v.text = "  " + item["value"]
                run_v.font.name = FONT
                run_v.font.size = Pt(14)
                run_v.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

    def build_stack(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.stack_items
        if not items:
            return

        n = len(items)
        gap = Inches(0.08)
        total_h = BODY_H - Inches(0.2)
        row_h = (total_h - gap * (n - 1)) / n
        colors = [PRIMARY, SECONDARY, ACCENT, MUTED]

        for i, item in enumerate(items):
            y = int(BODY_TOP + i * (row_h + gap))
            c = colors[i % len(colors)]

            # Full-width bar
            rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                int(MARGIN_L), y, int(CONTENT_W), int(row_h))
            rect.adjustments[0] = 0.02
            rect.fill.solid()
            rect.fill.fore_color.rgb = c
            rect.line.fill.background()

            # Name (left)
            ntb = self._add_textbox(slide, int(MARGIN_L + Pt(20)), y,
                int(CONTENT_W * 0.3), int(row_h))
            ntf = ntb.text_frame
            ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
            np = ntf.paragraphs[0]
            np.text = item["name"]
            np.font.name = FONT_HEAD
            np.font.size = Pt(18)
            np.font.bold = True
            np.font.color.rgb = WHITE

            # Description (right)
            dtb = self._add_textbox(slide, int(MARGIN_L + CONTENT_W * 0.35), y,
                int(CONTENT_W * 0.6), int(row_h))
            dtf = dtb.text_frame
            dtf.word_wrap = True
            dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
            dp = dtf.paragraphs[0]
            dp.text = item["desc"]
            dp.font.name = FONT
            dp.font.size = Pt(14)
            dp.font.color.rgb = WHITE

    def build_card_grid(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.card_items
        if not items:
            return

        gap = Inches(0.2)
        cols = 2
        rows_count = 2
        cell_w = (CONTENT_W - gap) / cols
        cell_h = (BODY_H - gap - Inches(0.1)) / rows_count

        for i, item in enumerate(items[:4]):
            r = i // cols
            c = i % cols
            x = MARGIN_L + c * (cell_w + gap)
            y = BODY_TOP + r * (cell_h + gap)
            self._add_zone_box(slide, int(x), int(y), int(cell_w), int(cell_h),
                label=item["title"], body=item["body"],
                label_size=Pt(18), body_size=Pt(14))

    def build_split_text(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        gap = Inches(0.5)
        col_w = (CONTENT_W - gap) / 2
        y = BODY_TOP + Inches(0.1)
        box_h = BODY_H - Inches(0.3)

        for i, side in enumerate([sd.split_left, sd.split_right]):
            x = MARGIN_L + i * (col_w + gap)
            label = side.get("label", "")
            body = side.get("body", "")
            self._add_zone_box(slide, int(x), int(y), int(col_w), int(box_h),
                label=label, body=body,
                label_size=Pt(20), body_size=Pt(16))

    def build_code(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        gap = Inches(0.4)
        code_w = int(CONTENT_W * 0.55)
        desc_w = int(CONTENT_W - code_w - gap)

        # Left: dark code box
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
            int(MARGIN_L), int(BODY_TOP), code_w, int(BODY_H - Inches(0.2)))
        bg.adjustments[0] = 0.02
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(0x1E, 0x1E, 0x2E)
        bg.line.fill.background()

        tb = self._add_textbox(slide, int(MARGIN_L + Pt(16)), int(BODY_TOP + Pt(16)),
            int(code_w - Pt(32)), int(BODY_H - Inches(0.2) - Pt(32)))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sd.code_text
        p.font.name = FONT_MONO
        p.font.size = Pt(13)
        p.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

        # Right: explanation
        if sd.code_desc:
            desc_x = int(MARGIN_L + code_w + gap)
            dtb = self._add_textbox(slide, desc_x, BODY_TOP, desc_w, int(BODY_H - Inches(0.2)))
            dtf = dtb.text_frame
            dtf.word_wrap = True
            dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
            dp = dtf.paragraphs[0]
            dp.text = sd.code_desc
            dp.font.name = FONT
            dp.font.size = Pt(16)
            dp.font.color.rgb = FG

    def build_multi_result(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        items = sd.multi_result_items
        if not items:
            return

        n = min(len(items), 4)
        gap = Inches(0.3)
        col_w = (CONTENT_W - gap * (n - 1)) / n
        y = BODY_TOP + Inches(0.3)

        for i, item in enumerate(items[:n]):
            x = MARGIN_L + i * (col_w + gap)

            # Large value
            vtb = self._add_textbox(slide, int(x), int(y), int(col_w), Inches(1.5))
            vtf = vtb.text_frame
            vtf.vertical_anchor = MSO_ANCHOR.BOTTOM
            vp = vtf.paragraphs[0]
            vp.text = item["value"]
            vp.font.name = FONT_HEAD
            vp.font.size = Pt(40)
            vp.font.bold = True
            vp.font.color.rgb = PRIMARY
            vp.alignment = PP_ALIGN.CENTER

            # Metric name
            mtb = self._add_textbox(slide, int(x), int(y + Inches(1.7)), int(col_w), Inches(0.5))
            mp = mtb.text_frame.paragraphs[0]
            mp.text = item["metric"]
            mp.font.name = FONT_HEAD
            mp.font.size = Pt(16)
            mp.font.bold = True
            mp.font.color.rgb = SECONDARY
            mp.alignment = PP_ALIGN.CENTER

            # Description
            if item.get("desc"):
                dtb = self._add_textbox(slide, int(x), int(y + Inches(2.3)), int(col_w), Inches(1.5))
                dtf = dtb.text_frame
                dtf.word_wrap = True
                dp = dtf.paragraphs[0]
                dp.text = item["desc"]
                dp.font.name = FONT
                dp.font.size = Pt(13)
                dp.font.color.rgb = MUTED
                dp.alignment = PP_ALIGN.CENTER

    def build_takeaway(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        # Main message
        if sd.takeaway_main:
            tb = self._add_textbox(slide, Inches(1.5), BODY_TOP + Inches(0.3),
                SW - Inches(3), Inches(2.0))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            self._set_text_with_inline_math(p, sd.takeaway_main, Pt(28), PRIMARY)
            p.font.bold = True

        # Supporting points
        if sd.takeaway_points:
            pts_top = BODY_TOP + Inches(2.8)
            pts_left = MARGIN_L + Inches(1.0)
            pts_w = CONTENT_W - Inches(2.0)
            for i, pt in enumerate(sd.takeaway_points):
                py = int(pts_top + Inches(0.45) * i)
                # Bullet dot
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                    int(pts_left), int(py + Pt(5)), Pt(6), Pt(6))
                dot.fill.solid()
                dot.fill.fore_color.rgb = ACCENT
                dot.line.fill.background()
                # Text
                ptb = self._add_textbox(slide, int(pts_left + Pt(16)), py,
                    int(pts_w), Inches(0.4))
                ptf = ptb.text_frame
                ptf.word_wrap = True
                self._set_text_with_inline_math(ptf.paragraphs[0], pt, Pt(18), FG)

    def build_profile(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)

        left_w = int(CONTENT_W * 0.35)
        gap = Inches(0.4)
        right_x = int(MARGIN_L + left_w + gap)
        right_w = int(CONTENT_W - left_w - gap)

        # Left: image or placeholder
        if sd.image_path:
            img_file = self._resolve_image(sd.image_path)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = left_w
                max_h = int(BODY_H)
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96), 1.0)
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                ix = int(MARGIN_L + (left_w - pw) / 2)
                iy = int(BODY_TOP + (BODY_H - ph) / 2)
                slide.shapes.add_picture(img_file, ix, iy, pw, ph)
        else:
            # Placeholder circle
            cd = min(left_w, int(BODY_H * 0.6))
            cx = int(MARGIN_L + (left_w - cd) / 2)
            cy = int(BODY_TOP + (BODY_H - cd) / 2)
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, cy, cd, cd)
            circle.fill.solid()
            circle.fill.fore_color.rgb = LIGHT
            circle.line.color.rgb = MUTED
            circle.line.width = Pt(1)

        cur_y = BODY_TOP + Inches(0.3)

        # Name
        if sd.profile_name:
            ntb = self._add_textbox(slide, right_x, int(cur_y), right_w, Inches(0.7))
            np = ntb.text_frame.paragraphs[0]
            np.text = sd.profile_name
            np.font.name = FONT_HEAD
            np.font.size = Pt(28)
            np.font.bold = True
            np.font.color.rgb = PRIMARY
            cur_y += Inches(0.8)

        # Affiliation
        if sd.profile_affiliation:
            atb = self._add_textbox(slide, right_x, int(cur_y), right_w, Inches(0.5))
            ap = atb.text_frame.paragraphs[0]
            ap.text = sd.profile_affiliation
            ap.font.name = FONT
            ap.font.size = Pt(18)
            ap.font.color.rgb = SECONDARY
            cur_y += Inches(0.7)

        # Bio list
        if sd.profile_bio:
            btb = self._add_textbox(slide, right_x, int(cur_y), right_w,
                int(BODY_H - (cur_y - BODY_TOP)))
            btf = btb.text_frame
            btf.word_wrap = True
            first = True
            for item in sd.profile_bio:
                if first:
                    p = btf.paragraphs[0]
                    first = False
                else:
                    p = btf.add_paragraph()
                run = p.add_run()
                run.text = "\u2022 " + item
                run.font.name = FONT
                run.font.size = Pt(15)
                run.font.color.rgb = FG
                p.space_before = Pt(6)

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
            "equations": self.build_equations,
            "figure": self.build_figure,
            "table-slide": self.build_table,
            "references": self.build_references,
            "timeline-h": self.build_timeline_h,
            "timeline": self.build_timeline_v,
            "end": self.build_end,
            "zone-flow": self.build_zone_flow,
            "zone-compare": self.build_zone_compare,
            "zone-matrix": self.build_zone_matrix,
            "zone-process": self.build_zone_process,
            "agenda": self.build_agenda,
            "rq": self.build_rq,
            "result-dual": self.build_result_dual,
            "summary": self.build_summary,
            "appendix": self.build_appendix,
            "overview": self.build_overview,
            "result": self.build_result,
            "steps": self.build_steps,
            # New slide types (v2)
            "quote": self.build_quote,
            "history": self.build_history,
            "panorama": self.build_panorama,
            "kpi": self.build_kpi,
            "pros-cons": self.build_pros_cons,
            "definition": self.build_definition,
            "diagram": self.build_diagram,
            "gallery-img": self.build_gallery_img,
            "highlight": self.build_highlight,
            "checklist": self.build_checklist,
            "annotation": self.build_annotation,
            "before-after": self.build_before_after,
            "funnel": self.build_funnel,
            "stack": self.build_stack,
            "card-grid": self.build_card_grid,
            "split-text": self.build_split_text,
            "code": self.build_code,
            "multi-result": self.build_multi_result,
            "takeaway": self.build_takeaway,
            "profile": self.build_profile,
        }

        for sd in slides:
            builder = BUILDERS.get(sd.slide_class, self.build_default)
            builder(sd)

        # Global footer on every slide (except title/end)
        self._add_global_footer()

    def _add_global_footer(self):
        """Add consistent footer to all slides except title and end."""
        for i, slide in enumerate(self.prs.slides):
            if i == 0 or i == len(self.prs.slides) - 1:
                continue  # skip title + end
            tb = self._add_textbox(
                slide, int(MARGIN_L), int(SH - Inches(0.4)),
                int(CONTENT_W), Inches(0.25))
            p = tb.text_frame.paragraphs[0]
            # Slide number on right
            p.text = f"{i + 1}"
            p.font.name = FONT
            p.font.size = Pt(8)
            p.font.color.rgb = MUTED
            p.alignment = PP_ALIGN.RIGHT


# ============================================================
# Main
# ============================================================
def apply_palette(palette_css: Path):
    """Override global color/font constants from a palette CSS file."""
    global PRIMARY, SECONDARY, ACCENT, FG, MUTED, LIGHT, BG_WHITE
    global FONT, FONT_HEAD, FONT_EA, FONT_MONO
    theme = load_theme(palette_css)
    c = theme["colors"]
    PRIMARY   = c.get("primary",   PRIMARY)
    SECONDARY = c.get("secondary", SECONDARY)
    ACCENT    = c.get("accent",    ACCENT)
    FG        = c.get("fg",        FG)
    MUTED     = c.get("muted",     MUTED)
    LIGHT     = c.get("light",     LIGHT)
    BG_WHITE  = c.get("bg",        BG_WHITE)
    # Fonts may also be overridden in palette
    if theme["fonts"]["body"]:
        FONT      = theme["fonts"]["body"]
    if theme["fonts"]["head"]:
        FONT_HEAD = theme["fonts"]["head"]
    if theme["fonts"]["ea"]:
        FONT_EA   = theme["fonts"]["ea"]
    # Load layout config from YAML if it exists alongside the CSS
    global LAYOUT
    name = palette_css.stem.replace("academic-", "")
    yaml_path = palette_css.parent / f"config-{name}.yaml"
    if yaml_path.exists():
        import yaml
        cfg = yaml.safe_load(yaml_path.read_text())
        lo = cfg.get("layout", {})
        LAYOUT = ThemeLayout(
            h1_deco=lo.get("h1_deco", LAYOUT.h1_deco),
            h1_deco_width=lo.get("h1_deco_width", LAYOUT.h1_deco_width),
            h1_deco_color=lo.get("h1_deco_color", LAYOUT.h1_deco_color),
            title_bg=lo.get("title_bg", LAYOUT.title_bg),
            title_align=lo.get("title_align", LAYOUT.title_align),
            divider_align=lo.get("divider_align", LAYOUT.divider_align),
            end_bg=lo.get("end_bg", LAYOUT.end_bg),
            box_style=lo.get("box_style", LAYOUT.box_style),
            box_radius=lo.get("box_radius", LAYOUT.box_radius),
            box_fill=lo.get("box_fill", LAYOUT.box_fill),
            spacing=lo.get("spacing", LAYOUT.spacing),
        )
    print(f"[palette] {name}: primary={PRIMARY} secondary={SECONDARY} accent={ACCENT}",
          file=sys.stderr)
    print(f"[layout]  h1={LAYOUT.h1_deco} title={LAYOUT.title_bg} box={LAYOUT.box_style} spacing={LAYOUT.spacing}",
          file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(
        description="Convert Marp academic templates to editable PPTX"
    )
    parser.add_argument("input", help="Input Marp markdown file")
    parser.add_argument("-o", "--output", help="Output .pptx path")
    parser.add_argument("-t", "--theme", help="Palette CSS (e.g. themes/palettes/academic-navy.css)")
    args = parser.parse_args()

    if args.theme:
        apply_palette(Path(args.theme))

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
