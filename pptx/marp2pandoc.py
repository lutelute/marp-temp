#!/usr/bin/env python3
"""
Marp Markdown → Pandoc-compatible Markdown preprocessor.

Converts Marp-specific HTML divs and directives into Pandoc's native
slide format so that `pandoc -t pptx` produces editable, structured PPTX.

Usage:
    python marp2pandoc.py example.md > pandoc_ready.md
    python marp2pandoc.py example.md -o pandoc_ready.md
"""

import re
import sys
import argparse
from pathlib import Path


def parse_frontmatter(text: str) -> tuple[dict, str]:
    """Strip YAML frontmatter, return (metadata, body)."""
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            fm = text[3:end].strip()
            body = text[end + 3:].strip()
            meta = {}
            for line in fm.split("\n"):
                if ":" in line:
                    k, v = line.split(":", 1)
                    meta[k.strip()] = v.strip()
            return meta, body
    return {}, text


def split_slides(body: str) -> list[str]:
    """Split on slide separator `---`."""
    slides = re.split(r"\n---\n", body)
    return [s.strip() for s in slides if s.strip()]


def extract_directives(slide: str) -> tuple[dict, str]:
    """Extract <!-- _class: X --> and <!-- _paginate: X --> directives."""
    directives = {}
    def replace_dir(m):
        key, val = m.group(1), m.group(2)
        directives[key] = val
        return ""
    cleaned = re.sub(
        r"<!--\s+_(\w+):\s*(.+?)\s*-->",
        replace_dir,
        slide
    )
    return directives, cleaned.strip()


def strip_html_tags(text: str) -> str:
    """Remove all HTML tags, keeping inner text."""
    return re.sub(r"<[^>]+>", "", text)


def extract_div_content(text: str, class_name: str) -> str | None:
    """Extract content of first <div class="...class_name...">...</div>."""
    pattern = rf'<div\s+class="[^"]*{re.escape(class_name)}[^"]*">'
    m = re.search(pattern, text)
    if not m:
        return None
    start = m.end()
    depth = 1
    pos = start
    while pos < len(text) and depth > 0:
        next_open = text.find("<div", pos)
        next_close = text.find("</div>", pos)
        if next_close == -1:
            break
        if next_open != -1 and next_open < next_close:
            depth += 1
            pos = next_open + 4
        else:
            depth -= 1
            if depth == 0:
                return text[start:next_close].strip()
            pos = next_close + 6
    return text[start:].strip()


def extract_all_child_divs(text: str) -> list[str]:
    """Extract direct child <div> contents from a container."""
    children = []
    pos = 0
    while pos < len(text):
        m = re.search(r"<div[^>]*>", text[pos:])
        if not m:
            break
        div_start = pos + m.end()
        depth = 1
        scan = div_start
        while scan < len(text) and depth > 0:
            next_open = text.find("<div", scan)
            next_close = text.find("</div>", scan)
            if next_close == -1:
                break
            if next_open != -1 and next_open < next_close:
                depth += 1
                scan = next_open + 4
            else:
                depth -= 1
                if depth == 0:
                    children.append(text[div_start:next_close].strip())
                    pos = next_close + 6
                    break
                scan = next_close + 6
        else:
            break
    return children


def clean_markdown(text: str) -> str:
    """Clean HTML from markdown, preserving structure."""
    text = strip_html_tags(text)
    # Clean up blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def convert_image_syntax(text: str) -> str:
    """Convert Marp ![w:NNN](path) to standard ![](path)."""
    return re.sub(r"!\[w:\d+\]", "![]", text)


def process_equation_slide(content: str) -> str:
    """Convert equation slide to Pandoc format."""
    lines = []

    # Extract title
    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    # Extract main equation
    eq_main = extract_div_content(content, "eq-main")
    if eq_main:
        eq_clean = clean_markdown(eq_main).strip()
        lines.append(eq_clean)
        lines.append("")

    # Extract variable descriptions
    eq_desc = extract_div_content(content, "eq-desc")
    if eq_desc:
        # Parse sym/description pairs from alternating spans
        all_spans = re.findall(r"<span[^>]*>(.*?)</span>", eq_desc, re.DOTALL)
        for i in range(0, len(all_spans) - 1, 2):
            sym = strip_html_tags(all_spans[i]).strip()
            desc = strip_html_tags(all_spans[i + 1]).strip() if i + 1 < len(all_spans) else ""
            lines.append(f"- {sym} : {desc}")
        lines.append("")

    # Extract footnote
    footnote = extract_div_content(content, "footnote")
    if footnote:
        lines.append(f"*{clean_markdown(footnote)}*")
        lines.append("")

    return "\n".join(lines)


def process_table_slide(content: str) -> str:
    """Process table slides: keep table + notes on same slide."""
    lines = []
    in_table = False

    for line in content.split("\n"):
        stripped = line.strip()
        # Skip HTML divs, extract text
        if stripped.startswith("<div") or stripped.startswith("</div>"):
            inner = strip_html_tags(stripped).strip()
            if inner:
                # Put box-accent / footnote as plain text after table
                lines.append(f"\n{inner}")
        elif stripped.startswith("<p ") or stripped.startswith("<span"):
            inner = strip_html_tags(stripped).strip()
            if inner:
                lines.append(inner)
        elif stripped.startswith("|"):
            lines.append(line)
        else:
            lines.append(line)

    return "\n".join(lines)


def process_columns(content: str, num_cols: int) -> str:
    """Convert column layout to Pandoc's native columns.

    Pandoc PPTX only supports 2-column layout natively.
    For 3+ columns, output as sequential sections on a single slide.
    """
    lines = []

    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    cols_content = extract_div_content(content, "columns")
    footnote = extract_div_content(content, "footnote")

    if cols_content:
        children = extract_all_child_divs(cols_content)
        n = len(children)

        if n == 2:
            # Native 2-column
            lines.append(":::::::::::::: {.columns}")
            for child in children:
                lines.append("::: {.column width=\"50%\"}")
                lines.append(clean_markdown(convert_image_syntax(child)))
                lines.append(":::")
                lines.append("")
            lines.append("::::::::::::::")
        elif n >= 3:
            # 3+ columns: use 2-col with merged content
            # Split into left half and right half
            mid = (n + 1) // 2
            left_parts = children[:mid]
            right_parts = children[mid:]
            lines.append(":::::::::::::: {.columns}")
            lines.append("::: {.column width=\"50%\"}")
            for part in left_parts:
                lines.append(clean_markdown(convert_image_syntax(part)))
                lines.append("")
            lines.append(":::")
            lines.append("::: {.column width=\"50%\"}")
            for part in right_parts:
                lines.append(clean_markdown(convert_image_syntax(part)))
                lines.append("")
            lines.append(":::")
            lines.append("::::::::::::::")
    else:
        cleaned = content
        if h1_match:
            cleaned = content[h1_match.end():].strip()
        lines.append(clean_markdown(cleaned))

    if footnote:
        lines.append("")
        lines.append(f"*{clean_markdown(footnote)}*")

    return "\n".join(lines)


def process_sandwich(content: str) -> str:
    """Convert sandwich layout to Pandoc format."""
    lines = []

    # Title
    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    # Top / lead
    top = extract_div_content(content, "top")
    if top:
        lead = extract_div_content(top, "lead") or clean_markdown(top)
        if isinstance(lead, str):
            lead = clean_markdown(lead)
        lines.append(lead)
        lines.append("")

    # Collect conclusion text first
    bottom_text = ""
    bottom = extract_div_content(content, "bottom")
    if bottom:
        conclusion = extract_div_content(bottom, "conclusion")
        if conclusion:
            bottom_text = f"\n**{clean_markdown(conclusion)}**"
        else:
            box = extract_div_content(bottom, "box")
            if box:
                bottom_text = f"\n{clean_markdown(box)}"
            else:
                bottom_text = f"\n{clean_markdown(bottom)}"

    # Columns — append conclusion to the last column to keep on same slide
    cols = extract_div_content(content, "columns")
    if cols:
        children = extract_all_child_divs(cols)
        if children:
            lines.append(":::::::::::::: {.columns}")
            for ci, child in enumerate(children):
                w = 100 // max(len(children), 1)
                lines.append(f"::: {{.column width=\"{w}%\"}}")
                lines.append(clean_markdown(child))
                # Append conclusion to last column
                if ci == len(children) - 1 and bottom_text:
                    lines.append(bottom_text)
                lines.append(":::")
                lines.append("")
            lines.append("::::::::::::::")
    elif bottom_text:
        lines.append(bottom_text)

    return "\n".join(lines)


def process_figure(content: str) -> str:
    """Convert figure slide to Pandoc format with caption as image title."""
    lines = []

    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    # Caption text for image title
    caption = extract_div_content(content, "caption")
    cap_text = clean_markdown(caption).strip() if caption else ""

    # Image with caption as title (keeps on same slide)
    img_match = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
    if img_match:
        if cap_text:
            lines.append(f'![{cap_text}]({img_match.group(1)})')
        else:
            lines.append(f"![]({img_match.group(1)})")
        lines.append("")

    # Description
    desc = extract_div_content(content, "description")
    if desc:
        lines.append(clean_markdown(desc))

    return "\n".join(lines)


def process_timeline_h(content: str) -> str:
    """Convert horizontal timeline to a structured list."""
    lines = []

    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    container = extract_div_content(content, "tl-h-container")
    if container:
        items = extract_all_child_divs(container)
        for item in items:
            # Extract from inner block div
            block = extract_all_child_divs(item)
            inner = block[0] if block else item

            year_m = re.search(r'class="tl-h-year"[^>]*>(.*?)</span>', inner, re.DOTALL)
            text_m = re.search(r'class="tl-h-text"[^>]*>(.*?)</span>', inner, re.DOTALL)
            detail_m = re.search(r'class="tl-h-detail"[^>]*>(.*?)</div>', inner, re.DOTALL)

            year = strip_html_tags(year_m.group(1)).strip() if year_m else ""
            text = strip_html_tags(text_m.group(1)).strip() if text_m else ""
            detail = strip_html_tags(detail_m.group(1)).strip() if detail_m else ""
            # Normalize whitespace from <br> etc
            detail = re.sub(r"\s+", " ", detail).strip()

            highlight = "highlight" in item

            entry = f"- **{year}** — {text}"
            if highlight:
                entry = f"- **→ {year}** — **{text}**"
            lines.append(entry)
            if detail:
                lines.append(f"    - {detail}")
        lines.append("")

    return "\n".join(lines)


def process_timeline_v(content: str) -> str:
    """Convert vertical timeline to a structured list."""
    lines = []

    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    container = extract_div_content(content, "tl-container")
    if container:
        items = extract_all_child_divs(container)
        for item in items:
            year_m = re.search(r'class="tl-year"[^>]*>(.*?)</span>', item, re.DOTALL)
            text_m = re.search(r'class="tl-text"[^>]*>(.*?)</span>', item, re.DOTALL)
            detail_m = re.search(r'class="tl-detail"[^>]*>(.*?)</div>', item, re.DOTALL)

            year = strip_html_tags(year_m.group(1)).strip() if year_m else ""
            text = strip_html_tags(text_m.group(1)).strip() if text_m else ""
            detail = strip_html_tags(detail_m.group(1)).strip() if detail_m else ""

            highlight = "highlight" in item
            marker = "**→**" if highlight else "-"

            line = f"{marker} **{year}** — {text}"
            lines.append(line)
            if detail:
                lines.append(f"  - {detail}")
        lines.append("")

    return "\n".join(lines)


def process_references(content: str) -> str:
    """Convert references list."""
    lines = []

    h1_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if h1_match:
        lines.append(f"# {h1_match.group(1)}")
        lines.append("")

    # Parse <li> elements
    lis = re.findall(r"<li>(.*?)</li>", content, re.DOTALL)
    for i, li in enumerate(lis, 1):
        author_m = re.search(r'class="author"[^>]*>(.*?)</span>', li)
        title_m = re.search(r'class="title"[^>]*>(.*?)</span>', li)
        venue_m = re.search(r'class="venue"[^>]*>(.*?)</span>', li)

        author = author_m.group(1).strip() if author_m else ""
        title = title_m.group(1).strip() if title_m else ""
        venue = venue_m.group(1).strip() if venue_m else ""

        lines.append(f"{i}. **{author}** {title} {venue}")

    return "\n".join(lines)


def process_default(content: str) -> str:
    """Process a default (no class) slide."""
    lines = []
    for line in content.split("\n"):
        stripped = line.strip()
        if stripped.startswith("<div") or stripped.startswith("</div>"):
            inner = strip_html_tags(stripped).strip()
            if inner:
                lines.append(inner)
        elif stripped.startswith("<p "):
            inner = strip_html_tags(stripped).strip()
            if inner:
                lines.append(inner)
        elif stripped.startswith("<span"):
            inner = strip_html_tags(stripped).strip()
            if inner:
                lines.append(inner)
        else:
            lines.append(convert_image_syntax(line))
    return "\n".join(lines)


def process_slide(slide_class: str | None, content: str) -> str:
    """Route a slide to the appropriate processor."""
    if slide_class == "title":
        # Title slide: Pandoc uses % title metadata or first h1 + subtitle
        return process_default(content)
    elif slide_class == "divider":
        return process_default(content)
    elif slide_class == "equation":
        return process_equation_slide(content)
    elif slide_class in ("cols-2", "cols-2-wide-l", "cols-2-wide-r"):
        return process_columns(content, 2)
    elif slide_class == "cols-3":
        return process_columns(content, 3)
    elif slide_class == "sandwich":
        return process_sandwich(content)
    elif slide_class == "figure":
        return process_figure(content)
    elif slide_class == "table-slide":
        return process_table_slide(content)
    elif slide_class == "references":
        return process_references(content)
    elif slide_class == "timeline-h":
        return process_timeline_h(content)
    elif slide_class == "timeline":
        return process_timeline_v(content)
    elif slide_class == "end":
        return process_default(content)
    else:
        return process_default(content)


def convert(input_path: str) -> str:
    """Convert a Marp markdown file to Pandoc-compatible markdown."""
    text = Path(input_path).read_text(encoding="utf-8")
    meta, body = parse_frontmatter(text)
    raw_slides = split_slides(body)

    output_parts = []

    # Pandoc YAML header
    output_parts.append("---")
    output_parts.append("title: ''")
    if meta.get("math") == "katex":
        output_parts.append("# math via pandoc native")
    output_parts.append("---")
    output_parts.append("")

    for i, raw in enumerate(raw_slides):
        directives, content = extract_directives(raw)
        slide_class = directives.get("class")

        processed = process_slide(slide_class, content)

        output_parts.append(processed)
        output_parts.append("")

        # Slide separator (except after last)
        if i < len(raw_slides) - 1:
            output_parts.append("---")
            output_parts.append("")

    return "\n".join(output_parts)


def main():
    parser = argparse.ArgumentParser(
        description="Convert Marp markdown to Pandoc-compatible format"
    )
    parser.add_argument("input", help="Input Marp markdown file")
    parser.add_argument("-o", "--output", help="Output file (default: stdout)")
    args = parser.parse_args()

    result = convert(args.input)

    if args.output:
        Path(args.output).write_text(result, encoding="utf-8")
        print(f"Written: {args.output}", file=sys.stderr)
    else:
        print(result)


if __name__ == "__main__":
    main()
