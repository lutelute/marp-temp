"""Tests for the PPTX builder."""
import tempfile
from pathlib import Path

from marp_pptx.parser import parse_slide, SlideData
from marp_pptx.theme import ThemeConfig
from marp_pptx.builder import PptxBuilder


def _make_builder(tmp_path=None):
    if tmp_path is None:
        tmp_path = Path(tempfile.mkdtemp())
    tc = ThemeConfig()
    return PptxBuilder(base_path=tmp_path, theme=tc)


def test_builder_creates_presentation():
    b = _make_builder()
    assert b.prs is not None
    assert len(b.prs.slides) == 0


def test_build_title_slide():
    b = _make_builder()
    sd = parse_slide(0, "<!-- _class: title -->\n# Test Title\n## Subtitle")
    b.build_title(sd)
    assert len(b.prs.slides) == 1


def test_build_default_slide():
    b = _make_builder()
    sd = parse_slide(0, "# Simple Content\n- bullet 1\n- bullet 2")
    b.build_default(sd)
    assert len(b.prs.slides) == 1


def test_build_all():
    b = _make_builder()
    slides = [
        parse_slide(0, "<!-- _class: title -->\n# Title"),
        parse_slide(1, "# Content\n- item"),
        parse_slide(2, "<!-- _class: end -->\n# Thank You"),
    ]
    b.build_all(slides)
    assert len(b.prs.slides) == 3


def test_save_pptx():
    b = _make_builder()
    sd = parse_slide(0, "<!-- _class: title -->\n# Save Test")
    b.build_title(sd)
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        b.save(f.name)
        assert Path(f.name).stat().st_size > 0
        Path(f.name).unlink()


def test_build_equation_slide():
    b = _make_builder()
    raw = '<!-- _class: equation -->\n# Math\n<div class="eq-main">$$E=mc^2$$</div>'
    sd = parse_slide(0, raw)
    b.build_equation(sd)
    assert len(b.prs.slides) == 1


def test_build_all_type_builders():
    """Smoke test: each registered builder runs without error on minimal input."""
    b = _make_builder()
    # Build a simple slide for each type
    for css_class in b.BUILDERS:
        sd = parse_slide(0, f"<!-- _class: {css_class} -->\n# Test {css_class}")
        method = getattr(b, b.BUILDERS[css_class])
        method(sd)
    assert len(b.prs.slides) == len(b.BUILDERS)
