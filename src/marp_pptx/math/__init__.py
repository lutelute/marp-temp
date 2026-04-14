"""Math rendering: OMML (editable) and PNG (fallback)."""

from marp_pptx.math.omml import latex_to_omml_element, OmmlError
from marp_pptx.math.renderer import render_latex_png

__all__ = ["latex_to_omml_element", "OmmlError", "render_latex_png"]
