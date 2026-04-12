"""LaTeX → OMML (Office Math Markup Language) conversion.

Uses Pandoc as the conversion engine (texmath library under the hood).
Output is a lxml <a14:m> element ready to append to a python-pptx <a:p>
paragraph, producing equations that are natively editable in PowerPoint.

Falls back to None on failure so callers can switch to a PNG image path.
"""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_omml_cache"
_CACHE_DIR.mkdir(exist_ok=True)

_PANDOC = shutil.which("pandoc")


class OmmlError(RuntimeError):
    pass


def _cache_key(tex: str, display: bool) -> str:
    return hashlib.md5(f"{int(display)}:{tex}".encode()).hexdigest()


def _run_pandoc(latex: str, display: bool) -> bytes:
    wrapped = f"$${latex}$$" if display else f"${latex}$"
    md = f"---\ntitle: x\n---\n\n# s\n\n{wrapped}\n"
    with tempfile.TemporaryDirectory() as d:
        dp = Path(d)
        (dp / "in.md").write_text(md, encoding="utf-8")
        pp = dp / "out.pptx"
        subprocess.run(
            [_PANDOC, "-t", "pptx", str(dp / "in.md"), "-o", str(pp)],
            check=True,
            capture_output=True,
            timeout=20,
        )
        with zipfile.ZipFile(pp) as z:
            # Pandoc emits a title slide (slide1) and our content (slide2).
            return z.read("ppt/slides/slide2.xml")


def latex_to_omml_element(latex: str, display: bool = False) -> etree._Element:
    """Convert a LaTeX math string into an <a14:m> lxml element.

    The returned element is a standalone subtree with namespace declarations
    inlined; append it directly to a python-pptx paragraph via
    `paragraph._p.append(element)`.

    Raises OmmlError on any failure.
    """
    if _PANDOC is None:
        raise OmmlError("pandoc not found in PATH")

    key = _cache_key(latex, display)
    cached = _CACHE_DIR / f"{key}.xml"
    if cached.exists():
        return etree.fromstring(cached.read_bytes())

    try:
        slide_xml = _run_pandoc(latex, display)
    except subprocess.CalledProcessError as e:
        raise OmmlError(f"pandoc failed: {e.stderr.decode(errors='ignore')[:200]}") from e
    except subprocess.TimeoutExpired as e:
        raise OmmlError("pandoc timeout") from e
    except KeyError as e:
        raise OmmlError(f"unexpected pandoc pptx layout: {e}") from e

    root = etree.fromstring(slide_xml)
    a14_elems = root.findall(f".//{{{NS_A14}}}m")
    if not a14_elems:
        raise OmmlError("no <a14:m> element in pandoc output")

    # Pandoc wraps in oMathPara for display mode and bare oMath for inline.
    # Pick the matching one (fallback to first if shape is unexpected).
    chosen = None
    for el in a14_elems:
        has_para = el.find(f"{{{NS_M}}}oMathPara") is not None
        if display and has_para:
            chosen = el
            break
        if not display and not has_para:
            chosen = el
            break
    if chosen is None:
        chosen = a14_elems[0]

    # Serialize with inline namespace decls so the element survives append.
    serialized = etree.tostring(chosen)
    cached.write_bytes(serialized)
    return etree.fromstring(serialized)


def latex_to_omml_xml(latex: str, display: bool = False) -> str:
    """String variant — returns the serialized <a14:m ...>...</a14:m> XML."""
    el = latex_to_omml_element(latex, display)
    return etree.tostring(el, encoding="unicode")


if __name__ == "__main__":
    # Self-test
    import sys
    samples = [
        (r"\frac{a}{b}", False),
        (r"\sum_{i=1}^n x_i^2 = \hat{\theta}", True),
        (r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}", True),
        (r"\alpha \beta \gamma", False),
    ]
    for tex, disp in samples:
        try:
            el = latex_to_omml_element(tex, disp)
            s = etree.tostring(el).decode()
            mark = "D" if disp else "I"
            print(f"[{mark}] {tex[:40]:40s} → {len(s):4d} bytes OK")
        except OmmlError as e:
            print(f"[!] {tex[:40]:40s} → {e}", file=sys.stderr)
