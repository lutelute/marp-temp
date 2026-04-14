"""Flask web UI for marp-pptx."""
from __future__ import annotations

import tempfile
from pathlib import Path

from flask import Flask, request, send_file, render_template_string, jsonify

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx Web UI</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', sans-serif; background: #f7f7f7; color: #1a1a1a; }
.container { max-width: 800px; margin: 40px auto; padding: 0 20px; }
h1 { font-size: 1.8em; margin-bottom: 8px; }
.subtitle { color: #999; margin-bottom: 32px; }
.card { background: white; border-radius: 8px; padding: 32px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 24px; }
label { display: block; font-weight: 600; margin-bottom: 8px; }
select, input[type="file"] { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; margin-bottom: 16px; }
button { background: #1a1a1a; color: white; border: none; padding: 12px 32px; border-radius: 4px; font-size: 1em; cursor: pointer; }
button:hover { background: #333; }
.types-table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
.types-table th { background: #1a1a1a; color: white; padding: 10px 12px; text-align: left; }
.types-table td { padding: 8px 12px; border-bottom: 1px solid #eee; }
.types-table tr:nth-child(even) { background: #f9f9f9; }
.cat-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; background: #e8e8e8; }
</style>
</head>
<body>
<div class="container">
<h1>marp-pptx</h1>
<p class="subtitle">Marp Markdown → Editable PowerPoint (49 semantic slide types)</p>

<div class="card">
<form action="/convert" method="post" enctype="multipart/form-data">
<label>Markdown File (.md)</label>
<input type="file" name="file" accept=".md" required>
<label>Palette</label>
<select name="palette">
<option value="">Default (monochrome)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>
<button type="submit">Convert to PPTX</button>
</form>
</div>

<div class="card">
<h2 style="margin-bottom:16px">Available Types ({{ types|length }})</h2>
<table class="types-table">
<thead><tr><th>Type</th><th>Category</th><th>Geometry</th><th>Meaning</th></tr></thead>
<tbody>
{% for t in types %}
<tr>
<td><code>{{ t.name }}</code></td>
<td><span class="cat-badge">{{ categories[t.category] }}</span></td>
<td>{{ t.geometry }}</td>
<td>{{ t.meaning }}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
</div>
</body>
</html>"""


def create_app() -> Flask:
    app = Flask(__name__)

    @app.route("/")
    def index():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        palettes_dir = Path(__file__).parent.parent / "data" / "themes" / "palettes"
        palettes = sorted(
            p.stem.replace("academic-", "")
            for p in palettes_dir.glob("academic-*.css")
        )
        return render_template_string(
            HTML_TEMPLATE,
            types=TYPE_REGISTRY,
            categories=CATEGORIES,
            palettes=palettes,
        )

    @app.route("/convert", methods=["POST"])
    def convert():
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        f = request.files.get("file")
        if not f:
            return "No file uploaded", 400

        palette_name = request.form.get("palette", "")

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            md_path = tmp / f.filename
            f.save(str(md_path))

            tc = ThemeConfig.from_css(get_default_theme_path())
            if palette_name:
                pp = get_palette_path(palette_name)
                if pp:
                    tc.apply_palette(pp)

            slides = parse_marp(str(md_path))
            builder = PptxBuilder(base_path=tmp, theme=tc)
            builder.build_all(slides)

            out_path = tmp / (md_path.stem + "_editable.pptx")
            builder.save(str(out_path))

            return send_file(
                str(out_path),
                as_attachment=True,
                download_name=out_path.name,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

    @app.route("/api/types")
    def api_types():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        data = [
            {
                "name": t.name,
                "css_class": t.css_class,
                "category": t.category,
                "category_ja": CATEGORIES.get(t.category, t.category),
                "geometry": t.geometry,
                "meaning": t.meaning,
                "use_when": t.use_when,
            }
            for t in TYPE_REGISTRY
        ]
        return jsonify(data)

    return app
