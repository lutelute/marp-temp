"""Flask web UI for marp-pptx."""
from __future__ import annotations

import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, send_file, render_template_string, jsonify, redirect, url_for


INDEX_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx Web UI</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; line-height: 1.6; }
.container { max-width: 900px; margin: 40px auto; padding: 0 20px; }
h1 { font-size: 1.8em; margin-bottom: 8px; }
.subtitle { color: #999; margin-bottom: 32px; }
.card { background: white; border-radius: 8px; padding: 32px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 24px; }
label { display: block; font-weight: 600; margin-bottom: 8px; margin-top: 12px; }
select, input[type="file"], input[type="text"] { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; margin-bottom: 16px; }
button { background: #1a1a1a; color: white; border: none; padding: 12px 32px; border-radius: 4px; font-size: 1em; cursor: pointer; margin-right: 8px; }
button:hover { background: #333; }
button.secondary { background: white; color: #1a1a1a; border: 1px solid #ddd; }
.types-table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
.types-table th { background: #1a1a1a; color: white; padding: 10px 12px; text-align: left; }
.types-table td { padding: 8px 12px; border-bottom: 1px solid #eee; }
.types-table tr:nth-child(even) { background: #f9f9f9; }
.cat-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; background: #e8e8e8; }
.tabs { display: flex; border-bottom: 2px solid #ddd; margin-bottom: 20px; }
.tabs a { padding: 10px 20px; color: #666; text-decoration: none; border-bottom: 2px solid transparent; margin-bottom: -2px; }
.tabs a.active { color: #1a1a1a; border-color: #1a1a1a; font-weight: 600; }
</style>
</head>
<body>
<div class="container">
<h1>marp-pptx</h1>
<p class="subtitle">Marp Markdown → Editable PowerPoint (49 semantic slide types)</p>

<div class="tabs">
<a href="/" class="active">変換</a>
<a href="/types-page">型一覧</a>
</div>

<div class="card">
<h2 style="margin-bottom:16px">① 簡易変換 (設定なしで即ダウンロード)</h2>
<form action="/convert" method="post" enctype="multipart/form-data">
<label>Markdown File (.md)</label>
<input type="file" name="file" accept=".md" required>
<label>Palette</label>
<select name="palette">
<option value="">Default (monochrome)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>
<button type="submit">→ PPTX に変換してダウンロード</button>
</form>
</div>

<div class="card">
<h2 style="margin-bottom:16px">② 調整画面 (スライド分析 + フォント倍率 + パレット)</h2>
<form action="/preview" method="post" enctype="multipart/form-data">
<label>Markdown File (.md)</label>
<input type="file" name="file" accept=".md" required>
<button type="submit">→ プレビュー画面へ</button>
</form>
</div>
</div>
</body>
</html>"""


TYPES_PAGE_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx: Slide Types</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; line-height: 1.6; }
.container { max-width: 1100px; margin: 40px auto; padding: 0 20px; }
h1 { margin-bottom: 20px; }
.tabs { display: flex; border-bottom: 2px solid #ddd; margin-bottom: 20px; }
.tabs a { padding: 10px 20px; color: #666; text-decoration: none; border-bottom: 2px solid transparent; margin-bottom: -2px; }
.tabs a.active { color: #1a1a1a; border-color: #1a1a1a; font-weight: 600; }
table { width: 100%; border-collapse: collapse; font-size: 0.9em; background: white; }
th { background: #1a1a1a; color: white; padding: 10px 12px; text-align: left; }
td { padding: 8px 12px; border-bottom: 1px solid #eee; }
tr:nth-child(even) { background: #f9f9f9; }
.cat-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; background: #e8e8e8; }
</style>
</head>
<body>
<div class="container">
<h1>型一覧 ({{ types|length }})</h1>
<div class="tabs">
<a href="/">変換</a>
<a href="/types-page" class="active">型一覧</a>
</div>
<table>
<thead><tr><th>型</th><th>カテゴリ</th><th>図形</th><th>意味</th><th>使いどころ</th></tr></thead>
<tbody>
{% for t in types %}
<tr>
<td><code>{{ t.name }}</code></td>
<td><span class="cat-badge">{{ categories[t.category] }}</span></td>
<td>{{ t.geometry }}</td>
<td>{{ t.meaning }}</td>
<td>{{ t.use_when }}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
</body>
</html>"""


PREVIEW_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx: Preview & Adjust</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; line-height: 1.6; }
.layout { display: grid; grid-template-columns: 320px 1fr; min-height: 100vh; }
aside { background: white; border-right: 1px solid #ddd; padding: 24px; position: sticky; top: 0; height: 100vh; overflow-y: auto; }
main { padding: 24px 32px; overflow-y: auto; }
h1 { font-size: 1.4em; margin-bottom: 16px; }
h2 { font-size: 1.1em; margin: 16px 0 8px; color: #555; }
label { display: block; font-weight: 600; margin: 12px 0 6px; font-size: 0.9em; }
select, input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 0.9em; }
input[type="range"] { padding: 0; }
.slider-row { display: flex; align-items: center; gap: 8px; }
.slider-val { min-width: 36px; font-variant-numeric: tabular-nums; font-size: 0.85em; color: #666; }
button { width: 100%; background: #1a1a1a; color: white; border: none; padding: 14px; border-radius: 4px; font-size: 1em; cursor: pointer; margin-top: 20px; font-weight: 600; }
button:hover { background: #333; }
button.secondary { background: white; color: #1a1a1a; border: 1px solid #ddd; }
.slide-card { background: white; border-radius: 6px; padding: 16px 20px; margin-bottom: 12px; border-left: 4px solid #3d5a80; }
.slide-card.warning { border-left-color: #e07a5f; }
.slide-num { color: #999; font-size: 0.85em; }
.slide-type { display: inline-block; background: #e8e8e8; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; font-family: ui-monospace, monospace; margin-left: 6px; }
.slide-h1 { font-size: 1.1em; font-weight: 600; margin: 6px 0; }
.slide-stats { font-size: 0.85em; color: #666; }
.back-link { color: #666; text-decoration: none; font-size: 0.9em; }
.back-link:hover { color: #1a1a1a; }
</style>
</head>
<body>
<div class="layout">
<aside>
<a href="/" class="back-link">← 戻る</a>
<h1 style="margin-top:12px">設定</h1>
<form action="/generate" method="post" id="gen-form">
<input type="hidden" name="session_id" value="{{ session_id }}">

<h2>パレット</h2>
<label>Color Palette</label>
<select name="palette">
<option value="">Default (monochrome)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>

<h2>サイズ</h2>
<label>Font Scale (0.7 - 1.3)</label>
<div class="slider-row">
<input type="range" name="font_scale" min="0.7" max="1.3" step="0.05" value="1.0" id="fs-range">
<span class="slider-val" id="fs-val">1.00</span>
</div>

<h2>出力</h2>
<label>Filename</label>
<input type="text" name="output_name" value="{{ filename_base }}_editable.pptx">

<button type="submit">→ PPTX を生成してダウンロード</button>
</form>

<div style="margin-top:20px; padding-top:16px; border-top:1px solid #eee; font-size:0.8em; color:#888;">
<p>※ margin_scale / per-slide override は v0.2 で対応予定 (ROADMAP参照)</p>
</div>
</aside>

<main>
<h1>{{ filename }} — {{ slides|length }} slides</h1>
<p style="color:#666; margin-bottom:20px">
左のパネルで設定を調整 → 下部の「PPTX を生成」ボタンでダウンロード。
</p>

{% for s in slides %}
<div class="slide-card {% if s.warning %}warning{% endif %}">
<span class="slide-num">Slide {{ loop.index }}</span>
<span class="slide-type">{{ s.type_display }}</span>
{% if s.h1 %}<div class="slide-h1">{{ s.h1 }}</div>{% endif %}
<div class="slide-stats">
{% if s.h2 %}<span>H2: {{ s.h2 }}</span> · {% endif %}
<span>{{ s.char_count }} chars</span>
{% if s.bullet_count %} · <span>{{ s.bullet_count }} bullets</span>{% endif %}
{% if s.table_rows %} · <span>{{ s.table_rows }} table rows</span>{% endif %}
{% if s.has_image %} · <span>🖼 image</span>{% endif %}
{% if s.has_math %} · <span>∑ math</span>{% endif %}
</div>
{% if s.warning %}<div style="color:#e07a5f; font-size:0.85em; margin-top:6px">⚠ {{ s.warning }}</div>{% endif %}
</div>
{% endfor %}
</main>
</div>

<script>
const range = document.getElementById('fs-range');
const val = document.getElementById('fs-val');
range.addEventListener('input', () => { val.textContent = parseFloat(range.value).toFixed(2); });
</script>
</body>
</html>"""


# Session-based storage of uploaded MD files
_SESSIONS: dict[str, Path] = {}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

    def _palettes() -> list[str]:
        palettes_dir = Path(__file__).parent.parent / "data" / "themes" / "palettes"
        return sorted(
            p.stem.replace("academic-", "")
            for p in palettes_dir.glob("academic-*.css")
        )

    @app.route("/")
    def index():
        return render_template_string(INDEX_HTML, palettes=_palettes())

    @app.route("/types-page")
    def types_page():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        return render_template_string(
            TYPES_PAGE_HTML,
            types=TYPE_REGISTRY,
            categories=CATEGORIES,
        )

    @app.route("/convert", methods=["POST"])
    def convert():
        return _do_convert(
            palette_name=request.form.get("palette", ""),
            font_scale=1.0,
            output_name=None,
        )

    @app.route("/preview", methods=["POST"])
    def preview():
        from marp_pptx.parser import parse_marp
        from marp_pptx.types import get_type_info

        f = request.files.get("file")
        if not f:
            return "No file uploaded", 400

        # Save to session
        session_id = uuid.uuid4().hex
        tmpdir = Path(tempfile.mkdtemp(prefix="marp_preview_"))
        md_path = tmpdir / (f.filename or "slides.md")
        f.save(str(md_path))
        _SESSIONS[session_id] = md_path

        slides_data = parse_marp(str(md_path))
        slides = []
        for sd in slides_data:
            info = get_type_info(sd.slide_class) if sd.slide_class else None
            type_display = sd.slide_class or "default"
            char_count = len(sd.raw)
            bullet_count = sum(
                1 for line in sd.body_lines
                if line.strip().startswith(("- ", "* "))
            )
            table_rows = len(sd.table_rows)
            has_image = bool(sd.image_path) or bool(sd.annotation_figure) or bool(sd.result_figure) or bool(sd.gallery_items)
            has_math = bool(sd.eq_main) or bool(sd.eq_system) or "$" in sd.raw
            warning = None
            if sd.slide_class and not info and sd.slide_class not in (
                "cols-2-wide-l", "cols-2-wide-r",
            ):
                warning = f"未知の型: {sd.slide_class}"
            slides.append({
                "type_display": type_display,
                "h1": sd.h1,
                "h2": sd.h2,
                "char_count": char_count,
                "bullet_count": bullet_count,
                "table_rows": table_rows,
                "has_image": has_image,
                "has_math": has_math,
                "warning": warning,
            })

        filename = md_path.name
        filename_base = md_path.stem

        return render_template_string(
            PREVIEW_HTML,
            slides=slides,
            palettes=_palettes(),
            session_id=session_id,
            filename=filename,
            filename_base=filename_base,
        )

    @app.route("/generate", methods=["POST"])
    def generate():
        session_id = request.form.get("session_id", "")
        md_path = _SESSIONS.get(session_id)
        if md_path is None or not md_path.exists():
            return "Session expired. Please re-upload.", 400

        return _do_convert(
            md_path=md_path,
            palette_name=request.form.get("palette", ""),
            font_scale=float(request.form.get("font_scale", 1.0)),
            output_name=request.form.get("output_name") or None,
        )

    def _do_convert(md_path=None, palette_name="", font_scale=1.0, output_name=None):
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        if md_path is None:
            f = request.files.get("file")
            if not f:
                return "No file uploaded", 400
            tmpdir = Path(tempfile.mkdtemp())
            md_path = tmpdir / (f.filename or "slides.md")
            f.save(str(md_path))

        tc = ThemeConfig.from_css(get_default_theme_path())
        tc.font_scale = max(0.5, min(2.0, font_scale))
        if palette_name:
            pp = get_palette_path(palette_name)
            if pp:
                tc.apply_palette(pp)

        slides = parse_marp(str(md_path))
        builder = PptxBuilder(base_path=md_path.parent, theme=tc)
        builder.build_all(slides)

        out_name = output_name or (md_path.stem + "_editable.pptx")
        out_path = md_path.parent / out_name
        builder.save(str(out_path))

        return send_file(
            str(out_path),
            as_attachment=True,
            download_name=out_name,
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
