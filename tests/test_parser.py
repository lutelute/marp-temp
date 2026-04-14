"""Tests for the Marp parser."""
from marp_pptx.parser import parse_marp, parse_slide, strip_html, extract_div, SlideData


def test_strip_html():
    assert strip_html("<b>hello</b>") == "hello"
    assert strip_html("plain text") == "plain text"


def test_extract_div():
    html = '<div class="eq-main">$$x^2$$</div>'
    assert extract_div(html, "eq-main") == "$$x^2$$"


def test_extract_div_missing():
    assert extract_div("<div>text</div>", "nonexistent") is None


def test_parse_slide_title():
    raw = '<!-- _class: title -->\n# Hello World\n## Subtitle'
    sd = parse_slide(0, raw)
    assert sd.slide_class == "title"
    assert sd.h1 == "Hello World"
    assert sd.h2 == "Subtitle"


def test_parse_slide_default():
    raw = '# Simple Slide\n- bullet one\n- bullet two'
    sd = parse_slide(0, raw)
    assert sd.slide_class is None
    assert sd.h1 == "Simple Slide"
    assert len(sd.body_lines) > 0


def test_parse_slide_equation():
    raw = '<!-- _class: equation -->\n# Math\n<div class="eq-main">$$E=mc^2$$</div>'
    sd = parse_slide(0, raw)
    assert sd.slide_class == "equation"
    assert sd.eq_main == "E=mc^2"


def test_parse_marp(example_md):
    if not example_md.exists():
        return
    slides = parse_marp(str(example_md))
    assert len(slides) > 0
    assert isinstance(slides[0], SlideData)


def test_parse_slide_kpi():
    raw = '''<!-- _class: kpi -->
# Metrics
<div class="kpi-container">
<div class="kpi-item"><span class="kpi-value">98%</span><span class="kpi-label">Accuracy</span></div>
<div class="kpi-item"><span class="kpi-value">1.2s</span><span class="kpi-label">Latency</span></div>
</div>'''
    sd = parse_slide(0, raw)
    assert sd.slide_class == "kpi"
    assert len(sd.kpi_items) == 2
    assert sd.kpi_items[0]["value"] == "98%"
