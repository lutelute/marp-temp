"""Theme configuration loaded from CSS variables."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from pptx.dml.color import RGBColor

_HEX_RE = re.compile(r"#([0-9a-fA-F]{6})")
_ROOT_RE = re.compile(r":root\s*\{([^}]*)\}", re.DOTALL)
_VAR_RE = re.compile(r"--([\w-]+)\s*:\s*([^;]+);")


def _hex_to_rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _resolve_font(font_stack: str, available: set[str]) -> str:
    names = [n.strip().strip("'\"") for n in font_stack.split(",")]
    for n in names:
        if n in available:
            return n
    return names[0] if names else "Helvetica Neue"


def _list_installed_fonts() -> set[str]:
    try:
        from matplotlib import font_manager
        return {f.name for f in font_manager.fontManager.ttflist}
    except Exception:
        return set()


@dataclass
class ThemeLayout:
    h1_deco: str = "left-bar"
    h1_deco_width: int = 8
    h1_deco_color: str = "primary"
    title_bg: str = "white"
    title_align: str = "left"
    divider_align: str = "left"
    end_bg: str = "white"
    box_style: str = "border-only"
    box_radius: float = 0.02
    box_fill: bool = False
    spacing: str = "compact"


@dataclass
class ThemeConfig:
    """Holds all color/font/layout state for a presentation."""
    # Colors
    primary: RGBColor = field(default_factory=lambda: RGBColor(0x16, 0x21, 0x3e))
    secondary: RGBColor = field(default_factory=lambda: RGBColor(0x0f, 0x34, 0x60))
    accent: RGBColor = field(default_factory=lambda: RGBColor(0xe9, 0x45, 0x60))
    bg: RGBColor = field(default_factory=lambda: RGBColor(0xff, 0xff, 0xff))
    fg: RGBColor = field(default_factory=lambda: RGBColor(0x1a, 0x1a, 0x2e))
    muted: RGBColor = field(default_factory=lambda: RGBColor(0x6c, 0x75, 0x7d))
    light: RGBColor = field(default_factory=lambda: RGBColor(0xf0, 0xf2, 0xf5))
    white: RGBColor = field(default_factory=lambda: RGBColor(0xff, 0xff, 0xff))
    # Fonts
    font: str = "Helvetica Neue"
    font_head: str = "Helvetica Neue"
    font_ea: str = "Hiragino Sans"
    font_mono: str = "SF Mono"
    # Scale multipliers (applied by builder)
    font_scale: float = 1.0
    margin_scale: float = 1.0
    # Math rendering mode: "omml" (editable in PowerPoint, default) or
    # "png" (matplotlib image — use for live-preview where the viewer
    # has poor OMML support, e.g. LibreOffice).
    math_mode: str = "omml"
    # Layout
    layout: ThemeLayout = field(default_factory=ThemeLayout)

    @classmethod
    def from_css(cls, css_path: Path) -> ThemeConfig:
        """Parse CSS :root variables into a ThemeConfig."""
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

        # Create defaults first, then override with CSS values
        defaults = cls()
        installed = _list_installed_fonts()
        config = cls(
            primary=colors.get("primary", defaults.primary),
            secondary=colors.get("secondary", defaults.secondary),
            accent=colors.get("accent", defaults.accent),
            bg=colors.get("bg", defaults.bg),
            fg=colors.get("fg", defaults.fg),
            muted=colors.get("muted", defaults.muted),
            light=colors.get("light", defaults.light),
            font=_resolve_font(vars_.get("font-body", ""), installed),
            font_head=_resolve_font(vars_.get("font-head", ""), installed),
            font_ea=_resolve_font(vars_.get("font-ea", "Hiragino Sans"), installed),
            font_mono=_resolve_font(vars_.get("font-mono", ""), installed),
        )
        return config

    def apply_palette(self, palette_css: Path) -> None:
        """Override colors/fonts from a palette CSS file."""
        import sys
        text = palette_css.read_text(encoding="utf-8")
        m = _ROOT_RE.search(text)
        root = m.group(1) if m else ""
        vars_ = dict(_VAR_RE.findall(root))

        for k, v in vars_.items():
            if k.startswith("color-"):
                hm = _HEX_RE.search(v)
                if hm:
                    name = k[len("color-"):]
                    color = _hex_to_rgb(hm.group(1))
                    if hasattr(self, name):
                        setattr(self, name, color)

        installed = _list_installed_fonts()
        if vars_.get("font-body"):
            self.font = _resolve_font(vars_["font-body"], installed)
        if vars_.get("font-head"):
            self.font_head = _resolve_font(vars_["font-head"], installed)
        if vars_.get("font-ea"):
            self.font_ea = _resolve_font(vars_["font-ea"], installed)

        # Load layout config from YAML if it exists alongside the CSS
        name = palette_css.stem.replace("academic-", "")
        yaml_path = palette_css.parent / f"config-{name}.yaml"
        if yaml_path.exists():
            import yaml
            cfg = yaml.safe_load(yaml_path.read_text())
            lo = cfg.get("layout", {})
            self.layout = ThemeLayout(
                h1_deco=lo.get("h1_deco", self.layout.h1_deco),
                h1_deco_width=lo.get("h1_deco_width", self.layout.h1_deco_width),
                h1_deco_color=lo.get("h1_deco_color", self.layout.h1_deco_color),
                title_bg=lo.get("title_bg", self.layout.title_bg),
                title_align=lo.get("title_align", self.layout.title_align),
                divider_align=lo.get("divider_align", self.layout.divider_align),
                end_bg=lo.get("end_bg", self.layout.end_bg),
                box_style=lo.get("box_style", self.layout.box_style),
                box_radius=lo.get("box_radius", self.layout.box_radius),
                box_fill=lo.get("box_fill", self.layout.box_fill),
                spacing=lo.get("spacing", self.layout.spacing),
            )

        print(f"[palette] {name}: primary={self.primary} secondary={self.secondary} accent={self.accent}",
              file=sys.stderr)


def get_default_theme_path() -> Path:
    """Return path to bundled academic.css theme."""
    return Path(__file__).parent / "data" / "themes" / "academic.css"


def get_palette_path(name: str) -> Path | None:
    """Return path to a named palette CSS file."""
    palettes_dir = Path(__file__).parent / "data" / "themes" / "palettes"
    css = palettes_dir / f"academic-{name}.css"
    if css.exists():
        return css
    # Try direct name
    css = palettes_dir / f"{name}.css"
    return css if css.exists() else None
