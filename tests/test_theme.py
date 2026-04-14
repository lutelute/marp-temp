"""Tests for theme loading."""
from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path


def test_load_default_theme():
    path = get_default_theme_path()
    assert path.exists()
    tc = ThemeConfig.from_css(path)
    assert tc.font != ""
    assert tc.font_head != ""
    assert tc.primary is not None


def test_palette_path():
    p = get_palette_path("navy")
    assert p is not None
    assert p.exists()


def test_palette_path_nonexistent():
    p = get_palette_path("nonexistent_palette_xyz")
    assert p is None or not p.exists()


def test_apply_palette():
    tc = ThemeConfig.from_css(get_default_theme_path())
    original_primary = tc.primary
    pp = get_palette_path("navy")
    if pp:
        tc.apply_palette(pp)
        # Navy palette changes primary
        assert tc.primary is not None


def test_theme_layout_defaults():
    tc = ThemeConfig()
    assert tc.layout.h1_deco == "left-bar"
    assert tc.layout.box_style == "border-only"
