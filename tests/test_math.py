"""Tests for math rendering."""
import shutil
from pathlib import Path

from marp_pptx.math.renderer import render_latex_png


def test_render_simple_latex():
    result = render_latex_png(r"x^2 + y^2 = z^2", fontsize=20)
    assert result is not None
    assert Path(result).exists()
    assert Path(result).stat().st_size > 0


def test_render_fraction():
    result = render_latex_png(r"\frac{a}{b}", fontsize=24, display=True)
    assert result is not None


def test_render_cache():
    # Same expression should return cached path
    r1 = render_latex_png(r"e^{i\pi} + 1 = 0", fontsize=20)
    r2 = render_latex_png(r"e^{i\pi} + 1 = 0", fontsize=20)
    assert r1 == r2


def test_omml_with_pandoc():
    if not shutil.which("pandoc"):
        return  # skip if pandoc not installed
    from marp_pptx.math.omml import latex_to_omml_element
    el = latex_to_omml_element(r"\frac{a}{b}", display=True)
    assert el is not None
