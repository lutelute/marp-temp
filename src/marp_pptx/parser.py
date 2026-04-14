"""Marp markdown parser — extracts structured SlideData from .md files."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


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


@dataclass
class SlideData:
    index: int
    slide_class: str | None
    paginate: bool
    raw: str
    h1: str = ""
    h2: str = ""
    body_lines: list = field(default_factory=list)
    columns: list = field(default_factory=list)
    top_text: str = ""
    bottom_text: str = ""
    table_rows: list = field(default_factory=list)
    image_path: str = ""
    caption: str = ""
    footnote: str = ""
    timeline_items: list = field(default_factory=list)
    eq_main: str = ""
    eq_vars: list = field(default_factory=list)
    eq_system: list = field(default_factory=list)
    ref_items: list = field(default_factory=list)
    zone_flow_items: list = field(default_factory=list)
    zone_compare: dict = field(default_factory=dict)
    zone_matrix: dict = field(default_factory=dict)
    zone_process_items: list = field(default_factory=list)
    agenda_items: list = field(default_factory=list)
    rq_main: str = ""
    rq_sub: str = ""
    summary_points: list = field(default_factory=list)
    result_dual_items: list = field(default_factory=list)
    appendix_label: str = ""
    overview_text: str = ""
    overview_points: list = field(default_factory=list)
    result_text: str = ""
    result_figure: str = ""
    result_caption: str = ""
    result_analysis: list = field(default_factory=list)
    steps_items: list = field(default_factory=list)
    quote_text: str = ""
    quote_source: str = ""
    history_items: list = field(default_factory=list)
    panorama_text: str = ""
    kpi_items: list = field(default_factory=list)
    pros_items: list = field(default_factory=list)
    cons_items: list = field(default_factory=list)
    def_term: str = ""
    def_body: str = ""
    def_note: str = ""
    gallery_items: list = field(default_factory=list)
    highlight_text: str = ""
    checklist_items: list = field(default_factory=list)
    annotation_figure: str = ""
    annotation_notes: list = field(default_factory=list)
    ba_before: dict = field(default_factory=dict)
    ba_after: dict = field(default_factory=dict)
    funnel_items: list = field(default_factory=list)
    stack_items: list = field(default_factory=list)
    card_items: list = field(default_factory=list)
    split_left: dict = field(default_factory=dict)
    split_right: dict = field(default_factory=dict)
    code_text: str = ""
    code_desc: str = ""
    multi_result_items: list = field(default_factory=list)
    takeaway_main: str = ""
    takeaway_points: list = field(default_factory=list)
    profile_name: str = ""
    profile_affiliation: str = ""
    profile_bio: list = field(default_factory=list)


def parse_slide(index: int, raw: str) -> SlideData:
    """Parse a raw slide chunk into SlideData."""
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

    h1m = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    h2m = re.search(r"^##\s+(.+)$", content, re.MULTILINE)
    if h1m:
        sd.h1 = strip_html(h1m.group(1))
    if h2m:
        sd.h2 = strip_html(h2m.group(1))

    cls = sd.slide_class

    if cls == "equation":
        eq = extract_div(content, "eq-main")
        if eq:
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

    elif cls == "equations":
        sys_div = extract_div(content, "eq-system")
        if sys_div:
            rows = extract_child_divs(sys_div)
            if rows:
                for row in rows:
                    lm = re.search(r'<span[^>]*class="[^"]*label[^"]*"[^>]*>(.*?)</span>', row, re.DOTALL)
                    label = strip_html(lm.group(1)) if lm else ""
                    mm = re.search(r"\$\$(.*?)\$\$", row, re.DOTALL)
                    if mm:
                        sd.eq_system.append((label, mm.group(1).strip()))
            else:
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

    elif cls in ("cols-2", "cols-2-wide-l", "cols-2-wide-r", "cols-3"):
        cols = extract_div(content, "columns")
        if cols:
            for child in extract_child_divs(cols):
                sd.columns.append(parse_markdown_lines(child))
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

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

    elif cls == "zone-matrix":
        extract_div(content, "zm-container")
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

    elif cls == "agenda":
        agenda = extract_div(content, "agenda-list")
        if agenda:
            for m in re.finditer(r"\d+\.\s*(.+)", agenda):
                sd.agenda_items.append(strip_html(m.group(1).strip()))

    elif cls == "rq":
        main = extract_div(content, "rq-main")
        if main:
            sd.rq_main = strip_html(main)
        sub = extract_div(content, "rq-sub")
        if sub:
            sd.rq_sub = strip_html(sub)

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

    elif cls == "summary":
        sp = extract_div(content, "summary-points")
        if not sp:
            sp_m = re.search(r'<ol[^>]*class="[^"]*summary-points[^"]*"[^>]*>(.*?)</ol>',
                             content, re.DOTALL)
            sp = sp_m.group(1) if sp_m else ""
        if sp:
            for li_m in re.finditer(r"<li>(.*?)</li>", sp, re.DOTALL):
                sd.summary_points.append(strip_html(li_m.group(1)))

    elif cls == "appendix":
        lbl = re.search(r'class="[^"]*appendix-label[^"]*"[^>]*>(.*?)</span>', content, re.DOTALL)
        if lbl:
            sd.appendix_label = strip_html(lbl.group(1))
        body = content
        if h1m:
            body = body[:h1m.start()] + body[h1m.end():]
        if h2m:
            body = body[:h2m.start()] + body[h2m.end():]
        for tag in ("appendix-label",):
            pattern = rf'<span\s+class="[^"]*{tag}[^"]*"[^>]*>.*?</span>'
            body = re.sub(pattern, "", body, flags=re.DOTALL)
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

    elif cls == "quote":
        qt = extract_div(content, "qt-text")
        if qt:
            sd.quote_text = strip_html(qt)
        qs = extract_div(content, "qt-source")
        if qs:
            sd.quote_source = strip_html(qs)

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

    elif cls == "panorama":
        pn = extract_div(content, "pn-text")
        if pn:
            sd.panorama_text = strip_html(pn)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

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

    elif cls == "pros-cons":
        pros = extract_div(content, "pc-pros")
        if pros:
            for li in re.finditer(r"<li>(.*?)</li>", pros, re.DOTALL):
                sd.pros_items.append(strip_html(li.group(1)))
        cons = extract_div(content, "pc-cons")
        if cons:
            for li in re.finditer(r"<li>(.*?)</li>", cons, re.DOTALL):
                sd.cons_items.append(strip_html(li.group(1)))

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

    elif cls == "diagram":
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)

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

    elif cls == "highlight":
        hl = extract_div(content, "hl-text")
        if hl:
            sd.highlight_text = strip_html(hl)

    elif cls == "checklist":
        container = extract_div(content, "cl-container")
        if container:
            for li in re.finditer(r'<li(\s+class="done")?>(.*?)</li>', container, re.DOTALL):
                sd.checklist_items.append({
                    "text": strip_html(li.group(2)),
                    "done": li.group(1) is not None,
                })

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

    elif cls == "code":
        cd = extract_div(content, "cd-code")
        if cd:
            code_m = re.search(r"```[\w]*\n(.*?)```", cd, re.DOTALL)
            if code_m:
                sd.code_text = code_m.group(1).rstrip()
            else:
                sd.code_text = strip_html(cd)
        desc = extract_div(content, "cd-desc")
        if desc:
            sd.code_desc = strip_html(desc)

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

    elif cls == "takeaway":
        ta = extract_div(content, "ta-main")
        if ta:
            sd.takeaway_main = strip_html(ta)
        pts = extract_div(content, "ta-points")
        if pts:
            for li in re.finditer(r"<li>(.*?)</li>", pts, re.DOTALL):
                sd.takeaway_points.append(strip_html(li.group(1)))

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

    else:
        body = content
        if h1m:
            body = body[:h1m.start()] + body[h1m.end():]
        if h2m:
            body = body[:h2m.start()] + body[h2m.end():]
        ba = extract_div(body, "box-accent")
        bp = extract_div(body, "box-primary")
        fn = extract_div(body, "footnote")
        for tag in ("box-accent", "box-primary", "box", "footnote"):
            div = extract_div(body, tag)
            if div:
                pattern = rf'<div\s+class="[^"]*{tag}[^"]*">.*?</div>'
                body = re.sub(pattern, "", body, flags=re.DOTALL)
        sd.body_lines = parse_markdown_lines(body)
        if ba:
            sd.bottom_text = strip_html(ba)
        elif bp:
            sd.bottom_text = strip_html(bp)
        if fn:
            sd.footnote = strip_html(fn)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

    return sd


def parse_marp(path: str | Path) -> list[SlideData]:
    """Parse a Marp markdown file into a list of SlideData."""
    text = Path(path).read_text(encoding="utf-8")
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
