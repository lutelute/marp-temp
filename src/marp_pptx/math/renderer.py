"""Pure-Python LaTeX-to-PNG renderer using matplotlib.

Replaces the Node.js KaTeX+Playwright dependency entirely.
"""
from __future__ import annotations

import hashlib
import io
import tempfile
from pathlib import Path

_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_math_png"
_CACHE_DIR.mkdir(exist_ok=True)


def render_latex_png(
    latex: str,
    fontsize: int = 28,
    display: bool = False,
    color: str = "#1a1a2e",
    dpi: int = 150,
) -> str | None:
    """Render LaTeX to a PNG file. Returns path to the PNG, or None on failure.

    Uses matplotlib's mathtext engine (no TeX installation required).
    """
    key = hashlib.md5(f"{latex}:{fontsize}:{display}:{dpi}".encode()).hexdigest()
    png_path = _CACHE_DIR / f"{key}.png"
    if png_path.exists():
        return str(png_path)

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import mathtext

        # Wrap in display-mode delimiters if needed
        tex = latex.strip()
        if not tex.startswith("$"):
            tex = f"${tex}$" if not display else f"${tex}$"

        fig = plt.figure(figsize=(0.01, 0.01))
        fig.patch.set_alpha(0.0)

        text = fig.text(
            0, 0, tex,
            fontsize=fontsize,
            color=color,
            math_fontfamily="cm",
        )

        # Render to get bounding box
        fig.canvas.draw()
        bbox = text.get_window_extent(fig.canvas.get_renderer())

        # Resize figure to fit text
        fig.set_size_inches(
            (bbox.width / dpi) + 0.1,
            (bbox.height / dpi) + 0.1,
        )

        # Re-position text
        text.set_position((0.05 * dpi / bbox.width if bbox.width > 0 else 0,
                          0.05 * dpi / bbox.height if bbox.height > 0 else 0))

        fig.savefig(
            str(png_path),
            dpi=dpi,
            transparent=True,
            bbox_inches="tight",
            pad_inches=0.05,
        )
        plt.close(fig)

        if png_path.exists() and png_path.stat().st_size > 0:
            return str(png_path)
        return None
    except Exception as e:
        import sys
        print(f"  Math PNG render failed: {latex[:40]}... ({e})", file=sys.stderr)
        return None
