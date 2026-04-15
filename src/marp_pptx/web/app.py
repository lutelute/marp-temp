"""Flask web UI for marp-pptx."""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, send_file, render_template_string, jsonify, redirect, url_for


# Cache for rendered slide thumbnails (keyed by MD content hash + settings)
_PREVIEW_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_pptx_previews"
_PREVIEW_CACHE_DIR.mkdir(exist_ok=True)

_SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
_PDFTOPPM = shutil.which("pdftoppm")


def _render_pptx_to_pngs(pptx_path: Path, out_dir: Path, dpi: int = 100) -> list[Path]:
    """Convert a PPTX file to per-slide PNG thumbnails using soffice + pdftoppm.

    Returns the sorted list of PNG paths. Returns [] if tools unavailable
    or conversion fails.
    """
    if _SOFFICE is None or _PDFTOPPM is None:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        # Step 1: PPTX → PDF via LibreOffice
        subprocess.run(
            [_SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), str(pptx_path)],
            check=True, capture_output=True, timeout=60,
        )
        pdf = out_dir / (pptx_path.stem + ".pdf")
        if not pdf.exists():
            return []
        # Step 2: PDF → PNG per page via pdftoppm
        subprocess.run(
            [_PDFTOPPM, "-png", "-r", str(dpi), str(pdf), str(out_dir / "slide")],
            check=True, capture_output=True, timeout=60,
        )
        return sorted(out_dir.glob("slide-*.png"))
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return []


EDITOR_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx Editor</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body { height: 100%; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; }
.topbar { background: #1a1a1a; color: white; padding: 10px 20px; display: flex; align-items: center; gap: 16px; }
.topbar h1 { font-size: 1.1em; font-weight: 600; }
.topbar a { color: #bbb; text-decoration: none; font-size: 0.9em; }
.topbar a:hover { color: white; }
.topbar .spacer { flex: 1; }
.layout { display: grid; grid-template-columns: 1fr 1fr 300px; height: calc(100vh - 42px); }
.editor-pane { display: flex; flex-direction: column; border-right: 1px solid #ddd; }
.preview-pane { background: #eee; overflow-y: auto; padding: 16px; border-right: 1px solid #ddd; }
.preview-pane h3 { font-size: 0.85em; text-transform: uppercase; letter-spacing: 0.05em; color: #666; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }
.preview-pane .slide-thumb { background: white; margin-bottom: 12px; border: 1px solid #ddd; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
.preview-pane .slide-thumb img { width: 100%; display: block; border-radius: 4px 4px 0 0; }
.preview-pane .slide-thumb .caption { padding: 6px 10px; font-size: 0.75em; color: #999; border-top: 1px solid #eee; display: flex; justify-content: space-between; align-items: center; }
.preview-pane .slide-thumb .slide-actions { display: flex; gap: 3px; }
.preview-pane .slide-thumb .slide-action-btn { background: #f0f0f0; border: 1px solid #ccc; padding: 2px 6px; border-radius: 2px; font-size: 0.85em; cursor: pointer; min-width: 22px; color: #555; }
.preview-pane .slide-thumb .slide-action-btn:hover { background: #ddd; color: #1a1a1a; }
.preview-pane .slide-thumb .slide-action-btn.delete:hover { background: #fce4e4; color: #c62828; border-color: #c62828; }
.preview-empty { color: #999; font-size: 0.85em; text-align: center; padding: 40px 20px; }
.preview-loading { color: #666; font-size: 0.85em; text-align: center; padding: 40px 20px; }
.preview-btn { background: #f0f0f0; border: 1px solid #ccc; padding: 4px 10px; border-radius: 3px; font-size: 0.75em; cursor: pointer; }
.preview-btn:hover { background: #e0e0e0; }
.editor-toolbar { background: white; padding: 10px 14px; border-bottom: 1px solid #eee; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.editor-toolbar button { background: #f0f0f0; border: 1px solid #ddd; padding: 6px 12px; border-radius: 3px; font-size: 0.85em; cursor: pointer; }
.editor-toolbar button:hover { background: #e0e0e0; }
.editor-toolbar button.primary { background: #1a1a1a; color: white; border-color: #1a1a1a; font-weight: 600; }
.editor-toolbar button.primary:hover { background: #333; }
.editor-toolbar .sep { color: #ccc; margin: 0 4px; }

/* Modal */
.modal-bg { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: none; align-items: center; justify-content: center; z-index: 100; }
.modal-bg.open { display: flex; }
.modal { background: white; border-radius: 8px; max-width: 720px; width: 90%; max-height: 85vh; overflow: hidden; display: flex; flex-direction: column; }
.modal-header { padding: 16px 24px; border-bottom: 1px solid #eee; display: flex; align-items: center; justify-content: space-between; }
.modal-header h2 { font-size: 1.1em; }
.modal-header .close { background: none; border: none; font-size: 1.3em; cursor: pointer; color: #999; }
.modal-body { padding: 20px 24px; overflow-y: auto; flex: 1; }
.modal-footer { padding: 14px 24px; border-top: 1px solid #eee; display: flex; gap: 8px; justify-content: flex-end; background: #f9f9f9; }
.modal-footer button { padding: 8px 18px; border-radius: 3px; border: 1px solid #ddd; background: white; cursor: pointer; font-size: 0.9em; }
.modal-footer button.primary { background: #1a1a1a; color: white; border-color: #1a1a1a; font-weight: 600; }

/* Type picker */
.type-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 8px; }
.type-card { border: 1px solid #ddd; border-radius: 4px; padding: 10px 12px; cursor: pointer; background: white; }
.type-card:hover { background: #eef; border-color: #88c; }
.type-card .type-name { font-family: ui-monospace, 'SF Mono', monospace; font-size: 0.85em; font-weight: 600; color: #1a1a1a; }
.type-card .type-geom { font-size: 0.9em; margin-top: 4px; color: #444; }
.type-card .type-meaning { font-size: 0.78em; color: #777; margin-top: 2px; }
.type-category-header { grid-column: 1/-1; font-size: 0.75em; font-weight: 600; color: #555; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 10px; padding: 4px 0; border-bottom: 1px solid #eee; }

/* Form fields */
.form-row { margin-bottom: 14px; }
.form-row label { display: block; font-weight: 600; font-size: 0.85em; margin-bottom: 6px; color: #333; }
.form-row input[type="text"], .form-row textarea { width: 100%; padding: 7px 10px; border: 1px solid #ccc; border-radius: 3px; font-size: 0.9em; font-family: inherit; }
.form-row textarea { resize: vertical; min-height: 60px; font-family: ui-monospace, 'SF Mono', monospace; font-size: 0.85em; }
.form-row .hint { font-size: 0.75em; color: #999; margin-top: 3px; }
.form-row input[type="checkbox"] { margin-right: 6px; }
.array-items { display: flex; flex-direction: column; gap: 8px; }
.array-item { display: flex; gap: 8px; align-items: flex-start; padding: 8px; background: #f7f7f7; border-radius: 3px; }
.array-item > div { flex: 1; }
.array-item .remove-btn { background: #fff; border: 1px solid #ccc; color: #c62828; padding: 4px 10px; border-radius: 3px; cursor: pointer; font-size: 0.85em; flex-shrink: 0; }
.array-item .remove-btn:hover { background: #fce4e4; }
.add-item-btn { margin-top: 8px; background: #eef; color: #1a1a1a; border: 1px dashed #88c; padding: 8px 14px; border-radius: 3px; cursor: pointer; font-size: 0.85em; }
.add-item-btn:hover { background: #ccd; }
.img-preview { max-width: 200px; max-height: 120px; margin-top: 6px; border: 1px solid #ddd; border-radius: 3px; display: block; }
.img-preview-sm { max-width: 100px; max-height: 60px; margin-top: 4px; border: 1px solid #ddd; border-radius: 3px; display: block; }
textarea#md-editor {
    flex: 1; width: 100%; border: none; padding: 16px 20px;
    font-family: ui-monospace, 'SF Mono', Menlo, monospace;
    font-size: 13px; line-height: 1.6; resize: none; outline: none;
    background: white;
}
aside { background: white; padding: 20px; overflow-y: auto; border-left: 1px solid #ddd; }
aside h2 { font-size: 0.95em; margin: 16px 0 8px; color: #555; text-transform: uppercase; letter-spacing: 0.05em; }
aside h2:first-child { margin-top: 0; }
label { display: block; font-weight: 600; margin: 10px 0 4px; font-size: 0.85em; }
select, input[type="text"] { width: 100%; padding: 6px 8px; border: 1px solid #ddd; border-radius: 3px; font-size: 0.9em; }
input[type="range"] { width: 100%; padding: 0; }
.slider-row { display: flex; align-items: center; gap: 8px; }
.slider-val { min-width: 36px; font-variant-numeric: tabular-nums; font-size: 0.85em; color: #666; }
button.primary { width: 100%; background: #1a1a1a; color: white; border: none; padding: 12px; border-radius: 4px; font-size: 0.95em; cursor: pointer; margin-top: 16px; font-weight: 600; }
button.primary:hover { background: #333; }
button.primary:disabled { background: #999; cursor: wait; }
.stats { font-size: 0.8em; color: #666; margin-top: 8px; padding: 10px; background: #f9f9f9; border-radius: 4px; }
.stats span { font-weight: 600; color: #1a1a1a; }
.sample-btn { display: block; width: 100%; text-align: left; padding: 8px 10px; margin-bottom: 4px; background: #f7f7f7; border: 1px solid #eee; border-radius: 3px; font-size: 0.85em; cursor: pointer; color: #555; }
.sample-btn:hover { background: #eef; color: #1a1a1a; }
.status { margin-top: 10px; padding: 8px 10px; border-radius: 3px; font-size: 0.85em; display: none; }
.status.ok { background: #e6f7e6; color: #2e7d32; display: block; }
.status.err { background: #fce4e4; color: #c62828; display: block; }
</style>
</head>
<body>
<div class="topbar">
<h1>marp-pptx Editor</h1>
<a href="/">変換画面に戻る</a>
<a href="/types-page">型一覧</a>
<div class="spacer"></div>
<span style="font-size:0.8em; color:#999">.md 保存不要・ブラウザ内で編集 → PPTX 生成</span>
</div>

<div class="layout">
<div class="editor-pane">
<div class="editor-toolbar">
<button class="primary" onclick="openTypePicker()">+ スライドを追加（型から選ぶ）</button>
<span class="sep">|</span>
<button onclick="insertSnippet('plain')">プレーン</button>
<button onclick="insertSnippet('bullets')">箇条書き</button>
<button onclick="insertSnippet('divider')">区切り</button>
<span class="sep">|</span>
<button onclick="downloadMd()" title="編集中のMarkdownを.mdファイルで保存">📥 MD保存</button>
<button onclick="document.getElementById('md-upload').click()" title=".mdファイルを読み込んでエディタに展開">📤 MD読込</button>
<input type="file" id="md-upload" accept=".md,.markdown,text/*" style="display:none" onchange="loadMdFile(event)">
<button onclick="document.getElementById('pptx-upload').click()" title="PPTXを読み込んでMD構造に変換">🔄 PPTX→MD</button>
<input type="file" id="pptx-upload" accept=".pptx" style="display:none" onchange="loadPptxFile(event)">
<span class="sep">|</span>
<button onclick="if(confirm('エディタ内容を全削除しますか？')) { document.getElementById('md-editor').value=''; updateStats(); autoSave(); triggerAutoPreview(); }">全削除</button>
</div>

<!-- Type picker modal -->
<div class="modal-bg" id="picker-modal">
<div class="modal" style="max-width:820px">
<div class="modal-header">
<h2>スライド型を選ぶ</h2>
<button class="close" onclick="closeModal('picker-modal')">×</button>
</div>
<div style="padding:10px 24px; border-bottom:1px solid #eee">
<input type="text" id="type-search" placeholder="型名・意味・図形で検索 (例: 比較, 時間, VS, □)"
  style="width:100%; padding:8px 10px; border:1px solid #ddd; border-radius:3px; font-size:0.95em"
  oninput="filterTypes(this.value)">
</div>
<div class="modal-body">
<div class="type-grid" id="type-grid"></div>
</div>
</div>
</div>

<!-- Slide edit modal (edit a single slide's raw MD) -->
<div class="modal-bg" id="slide-edit-modal">
<div class="modal" style="max-width:700px">
<div class="modal-header">
<h2>スライドを編集 <span style="font-weight:normal; color:#999; font-size:0.85em" id="slide-edit-idx"></span></h2>
<button class="close" onclick="closeModal('slide-edit-modal')">×</button>
</div>
<div class="modal-body">
<textarea id="slide-edit-ta" style="width:100%; min-height:320px; padding:12px; font-family:ui-monospace,'SF Mono',monospace; font-size:13px; line-height:1.6; border:1px solid #ccc; border-radius:3px; resize:vertical"></textarea>
<div class="hint" style="margin-top:6px">このスライド部分の Markdown を直接編集します。他のスライドには影響しません。</div>
</div>
<div class="modal-footer">
<button onclick="closeModal('slide-edit-modal')">キャンセル</button>
<button class="primary" onclick="saveSlideEdit()">保存</button>
</div>
</div>
</div>

<!-- Form modal -->
<div class="modal-bg" id="form-modal">
<div class="modal">
<div class="modal-header">
<h2 id="form-title">型の入力</h2>
<button class="close" onclick="closeModal('form-modal')">×</button>
</div>
<div class="modal-body" id="form-body"></div>
<div class="modal-footer">
<button onclick="closeModal('form-modal')">キャンセル</button>
<button class="primary" onclick="submitForm()">スライドを追加</button>
</div>
</div>
</div>
<textarea id="md-editor" spellcheck="false" placeholder="ここにMarkdownを書くか、右のサンプルからロードしてください"></textarea>
</div>

<div class="preview-pane">
<h3>
<span>プレビュー (実レンダリング)</span>
<button class="preview-btn" onclick="refreshPreview()">更新</button>
</h3>
<div id="preview-content">
<div class="preview-empty">エディタに内容を入れて<br>「更新」を押してください</div>
</div>
</div>

<aside>
<h2>サンプル</h2>
<button class="sample-btn" onclick="loadSample('minimal')">📄 最小雛形</button>
<button class="sample-btn" onclick="loadSample('all')">📚 全型カタログ</button>
<button class="sample-btn" onclick="loadSample('academic')">🎓 学術発表サンプル</button>

<h2>出力設定</h2>
<label>Palette</label>
<select id="palette">
<option value="">Default (mono)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>

<label>Font Scale</label>
<div class="slider-row">
<input type="range" id="font-scale" min="0.7" max="1.3" step="0.05" value="1.0">
<span class="slider-val" id="fs-val">1.00</span>
</div>

<label>ファイル名</label>
<input type="text" id="output-name" value="slides_editable.pptx">

<button class="primary" id="gen-btn" onclick="generate()">→ PPTX を生成してダウンロード</button>

<div class="status" id="status"></div>
<div class="stats" id="stats"></div>
</aside>
</div>

<script>
const editor = document.getElementById('md-editor');
const stats = document.getElementById('stats');
const fsRange = document.getElementById('font-scale');
const fsVal = document.getElementById('fs-val');
const statusEl = document.getElementById('status');

fsRange.addEventListener('input', () => fsVal.textContent = parseFloat(fsRange.value).toFixed(2));

// Live stats
function updateStats() {
    const t = editor.value;
    if (!t.trim()) { stats.innerHTML = '未入力'; return; }
    const slides = t.split(/\\n---\\n/).filter(x => x.trim()).length;
    const chars = t.length;
    const types = [...t.matchAll(/<!--\\s+_class:\\s+(\\S+)\\s+-->/g)].map(m => m[1]);
    const typeCounts = {};
    types.forEach(x => typeCounts[x] = (typeCounts[x] || 0) + 1);
    const typeList = Object.entries(typeCounts).map(([k,v]) => v > 1 ? `${k}×${v}` : k).join(', ');
    stats.innerHTML = `<span>${slides}</span> slides · <span>${chars}</span> chars${typeList ? '<br>型: ' + typeList : ''}`;
}
editor.addEventListener('input', updateStats);

// ── Type schemas: fields + MD template per type ──
const TYPE_SCHEMAS = {};
let TYPES_META = [];  // loaded from /api/types

async function loadTypeMeta() {
    try {
        const r = await fetch('/api/types');
        TYPES_META = await r.json();
    } catch(e) { console.error('型一覧の取得に失敗', e); }
}

// Helpers for building MD
function esc(s) { return String(s || ''); }
function joinLines(arr) { return arr.filter(Boolean).join('\\n'); }

// Schema: { fields: [...], toMd: (data) => string }
// field: { name, label, type: 'text'|'textarea'|'array'|'checkbox', default, hint, subfields? }

TYPE_SCHEMAS['plain'] = {
    label: 'プレーン本文',
    fields: [
        { name: 'h1', label: 'タイトル (H1)', type: 'text', default: 'タイトル' },
        { name: 'body', label: '本文 (Markdown)', type: 'textarea', default: '- ポイント1\\n- ポイント2' },
    ],
    toMd: d => `# ${esc(d.h1)}\\n${esc(d.body)}`,
};

TYPE_SCHEMAS['title'] = {
    label: 'title — 表紙',
    fields: [
        { name: 'h1', label: 'メインタイトル', type: 'text', default: 'タイトル' },
        { name: 'h2', label: 'サブタイトル', type: 'text', default: 'サブタイトル' },
        { name: 'author', label: '発表者名', type: 'text', default: '' },
        { name: 'date', label: '日付', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: title -->\\n# ${esc(d.h1)}${d.h2 ? '\\n## '+esc(d.h2) : ''}${d.author ? '\\n'+esc(d.author) : ''}${d.date ? '\\n'+esc(d.date) : ''}`,
};

TYPE_SCHEMAS['divider'] = {
    label: 'divider — 章区切り',
    fields: [
        { name: 'h1', label: '章タイトル', type: 'text', default: '第○章' },
        { name: 'h2', label: '章サブ', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: divider -->\\n# ${esc(d.h1)}${d.h2 ? '\\n## '+esc(d.h2) : ''}`,
};

TYPE_SCHEMAS['end'] = {
    label: 'end — 終了',
    fields: [
        { name: 'h1', label: 'メッセージ', type: 'text', default: 'Thank You' },
        { name: 'sub', label: '補足（任意）', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: end -->\\n# ${esc(d.h1)}${d.sub ? '\\n'+esc(d.sub) : ''}`,
};

TYPE_SCHEMAS['cols-2'] = {
    label: 'cols-2 — 2カラム',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '比較' },
        { name: 'left_title', label: '左カラム見出し (H3)', type: 'text', default: '従来' },
        { name: 'left_body', label: '左カラム本文', type: 'textarea', default: '- 項目A\\n- 項目B' },
        { name: 'right_title', label: '右カラム見出し (H3)', type: 'text', default: '提案' },
        { name: 'right_body', label: '右カラム本文', type: 'textarea', default: '- 項目A\\n- 項目B' },
    ],
    toMd: d => `<!-- _class: cols-2 -->\\n# ${esc(d.h1)}\\n<div>\\n\\n### ${esc(d.left_title)}\\n${esc(d.left_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.right_title)}\\n${esc(d.right_body)}\\n</div>`,
};

TYPE_SCHEMAS['cols-3'] = {
    label: 'cols-3 — 3カラム',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '3観点' },
        { name: 'c1_title', label: '列1 見出し', type: 'text', default: '観点1' },
        { name: 'c1_body', label: '列1 本文', type: 'textarea', default: '- A\\n- B' },
        { name: 'c2_title', label: '列2 見出し', type: 'text', default: '観点2' },
        { name: 'c2_body', label: '列2 本文', type: 'textarea', default: '- A\\n- B' },
        { name: 'c3_title', label: '列3 見出し', type: 'text', default: '観点3' },
        { name: 'c3_body', label: '列3 本文', type: 'textarea', default: '- A\\n- B' },
    ],
    toMd: d => `<!-- _class: cols-3 -->\\n# ${esc(d.h1)}\\n<div>\\n\\n### ${esc(d.c1_title)}\\n${esc(d.c1_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.c2_title)}\\n${esc(d.c2_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.c3_title)}\\n${esc(d.c3_body)}\\n</div>`,
};

TYPE_SCHEMAS['sandwich'] = {
    label: 'sandwich — 上・中・下',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'タイトル' },
        { name: 'lead', label: '上部リード文', type: 'textarea', default: '概要を1行で。' },
        { name: 'left_title', label: '中央左 見出し', type: 'text', default: '従来' },
        { name: 'left_body', label: '中央左 本文', type: 'textarea', default: '- 項目A' },
        { name: 'right_title', label: '中央右 見出し', type: 'text', default: '提案' },
        { name: 'right_body', label: '中央右 本文', type: 'textarea', default: '- 項目A' },
        { name: 'conclusion', label: '下部結論', type: 'textarea', default: '**結論:** ...' },
    ],
    toMd: d => `<!-- _class: sandwich -->\\n# ${esc(d.h1)}\\n<div class="top">\\n<div class="lead">${esc(d.lead)}</div>\\n</div>\\n<div class="columns">\\n<div>\\n\\n### ${esc(d.left_title)}\\n${esc(d.left_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.right_title)}\\n${esc(d.right_body)}\\n</div>\\n</div>\\n<div class="bottom">\\n<div class="conclusion">${esc(d.conclusion)}</div>\\n</div>`,
};

TYPE_SCHEMAS['equation'] = {
    label: 'equation — 数式',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '数式' },
        { name: 'formula', label: 'LaTeX式（$$なしで記入）', type: 'textarea', default: 'E = mc^2', hint: '例: \\\\frac{a}{b}, \\\\sum_{i=1}^n x_i' },
        { name: 'vars', label: '変数説明', type: 'array',
          subfields: [
            { name: 'sym', label: '記号 (LaTeX)', type: 'text', default: 'E' },
            { name: 'desc', label: '説明', type: 'text', default: 'エネルギー' },
          ], default: [{sym:'E',desc:'エネルギー'},{sym:'m',desc:'質量'},{sym:'c',desc:'光速'}] },
    ],
    toMd: d => {
        const vars_ = (d.vars||[]).map(v => `<span>$${esc(v.sym)}$</span><span>${esc(v.desc)}</span>`).join('\\n');
        return `<!-- _class: equation -->\\n# ${esc(d.h1)}\\n<div class="eq-main">\\n$$${esc(d.formula)}$$\\n</div>${vars_?`\\n<div class="eq-desc">\\n${vars_}\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['kpi'] = {
    label: 'kpi — 主要指標',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '主要指標' },
        { name: 'items', label: '指標項目', type: 'array',
          subfields: [
            { name: 'value', label: '値', type: 'text', default: '98%' },
            { name: 'label', label: 'ラベル', type: 'text', default: '精度' },
          ], default: [{value:'98%',label:'精度'},{value:'1.2s',label:'応答'},{value:'10x',label:'高速化'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="kpi-value">${esc(it.value)}</span><span class="kpi-label">${esc(it.label)}</span></div>`).join('\\n');
        return `<!-- _class: kpi -->\\n# ${esc(d.h1)}\\n<div class="kpi-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['funnel'] = {
    label: 'funnel — 絞り込み',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '絞り込み' },
        { name: 'items', label: 'ステージ', type: 'array',
          subfields: [
            { name: 'label', label: 'ラベル', type: 'text', default: 'ステージ' },
            { name: 'value', label: '値', type: 'text', default: '' },
          ], default: [{label:'応募',value:'1,000'},{label:'書類通過',value:'200'},{label:'面接通過',value:'50'},{label:'採用',value:'10'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="fn-label">${esc(it.label)}</span><span class="fn-value">${esc(it.value)}</span></div>`).join('\\n');
        return `<!-- _class: funnel -->\\n# ${esc(d.h1)}\\n<div class="fn-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['pros-cons'] = {
    label: 'pros-cons — 賛否',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '賛否' },
        { name: 'pros', label: 'Pros 項目', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'高速'},{text:'省メモリ'}] },
        { name: 'cons', label: 'Cons 項目', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'実装コスト高'},{text:'依存多い'}] },
    ],
    toMd: d => {
        const pros = (d.pros||[]).map(it => `<li>${esc(it.text)}</li>`).join('');
        const cons = (d.cons||[]).map(it => `<li>${esc(it.text)}</li>`).join('');
        return `<!-- _class: pros-cons -->\\n# ${esc(d.h1)}\\n<div class="pc-pros">\\n<ul>${pros}</ul>\\n</div>\\n<div class="pc-cons">\\n<ul>${cons}</ul>\\n</div>`;
    },
};

TYPE_SCHEMAS['timeline-h'] = {
    label: 'timeline-h — 横タイムライン',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'タイムライン' },
        { name: 'items', label: '時点', type: 'array',
          subfields: [
            { name: 'year', label: '年/時点', type: 'text', default: '2024' },
            { name: 'text', label: '出来事', type: 'text', default: '' },
            { name: 'highlight', label: '強調', type: 'checkbox', default: false },
          ], default: [{year:'2024',text:'企画'},{year:'2025',text:'開発',highlight:true},{year:'2026',text:'リリース'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div class="tl-h-item${it.highlight?' highlight':''}"><div><span class="tl-h-year">${esc(it.year)}</span><span class="tl-h-text">${esc(it.text)}</span></div></div>`).join('\\n');
        return `<!-- _class: timeline-h -->\\n# ${esc(d.h1)}\\n<div class="tl-h-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['zone-matrix'] = {
    label: 'zone-matrix — 2×2 マトリクス',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '2軸評価' },
        { name: 'x_label', label: 'X軸ラベル', type: 'text', default: '重要度' },
        { name: 'y_label', label: 'Y軸ラベル', type: 'text', default: '緊急度' },
        { name: 'tl_label', label: '左上 ラベル', type: 'text', default: 'A' },
        { name: 'tl_body', label: '左上 本文', type: 'text', default: '' },
        { name: 'tr_label', label: '右上 ラベル', type: 'text', default: 'B' },
        { name: 'tr_body', label: '右上 本文', type: 'text', default: '' },
        { name: 'bl_label', label: '左下 ラベル', type: 'text', default: 'C' },
        { name: 'bl_body', label: '左下 本文', type: 'text', default: '' },
        { name: 'br_label', label: '右下 ラベル', type: 'text', default: 'D' },
        { name: 'br_body', label: '右下 本文', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: zone-matrix -->\\n# ${esc(d.h1)}\\n<div class="zm-xlabel">${esc(d.x_label)}</div>\\n<div class="zm-ylabel">${esc(d.y_label)}</div>\\n<div class="zm-tl"><span class="zm-label">${esc(d.tl_label)}</span><span class="zm-body">${esc(d.tl_body)}</span></div>\\n<div class="zm-tr"><span class="zm-label">${esc(d.tr_label)}</span><span class="zm-body">${esc(d.tr_body)}</span></div>\\n<div class="zm-bl"><span class="zm-label">${esc(d.bl_label)}</span><span class="zm-body">${esc(d.bl_body)}</span></div>\\n<div class="zm-br"><span class="zm-label">${esc(d.br_label)}</span><span class="zm-body">${esc(d.br_body)}</span></div>`,
};

TYPE_SCHEMAS['agenda'] = {
    label: 'agenda — 目次',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '本日の内容' },
        { name: 'items', label: '項目', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'背景'},{text:'手法'},{text:'結果'},{text:'まとめ'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map((it,i) => `${i+1}. ${esc(it.text)}`).join('\\n');
        return `<!-- _class: agenda -->\\n# ${esc(d.h1)}\\n<div class="agenda-list">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['summary'] = {
    label: 'summary — まとめ',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'まとめ' },
        { name: 'items', label: 'ポイント', type: 'array',
          subfields: [{ name: 'text', label: 'ポイント', type: 'text', default: '' }],
          default: [{text:'提案手法は従来比10倍高速'},{text:'精度は同等'},{text:'OSSとして公開'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: summary -->\\n# ${esc(d.h1)}\\n<ol class="summary-points">\\n${items}\\n</ol>`;
    },
};

TYPE_SCHEMAS['takeaway'] = {
    label: 'takeaway — キーメッセージ',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'Takeaway' },
        { name: 'main', label: '中央の一文', type: 'textarea', default: '型を選ぶだけで伝わるプレゼンに' },
        { name: 'points', label: '補足ポイント', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'49種の意味的な型'},{text:'完全編集可能'}] },
    ],
    toMd: d => {
        const pts = (d.points||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: takeaway -->\\n# ${esc(d.h1)}\\n<div class="ta-main">${esc(d.main)}</div>${pts?`\\n<div class="ta-points">\\n<ul>\\n${pts}\\n</ul>\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['quote'] = {
    label: 'quote — 引用',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '引用' },
        { name: 'text', label: '引用文', type: 'textarea', default: '' },
        { name: 'source', label: '出典', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: quote -->\\n# ${esc(d.h1)}\\n<div class="qt-text">${esc(d.text)}</div>${d.source?`\\n<div class="qt-source">${esc(d.source)}</div>`:''}`,
};

TYPE_SCHEMAS['definition'] = {
    label: 'definition — 用語定義',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '定義' },
        { name: 'term', label: '用語', type: 'text', default: '' },
        { name: 'body', label: '定義文', type: 'textarea', default: '' },
        { name: 'note', label: '補足', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: definition -->\\n# ${esc(d.h1)}\\n<div class="df-term">${esc(d.term)}</div>\\n<div class="df-body">${esc(d.body)}</div>${d.note?`\\n<div class="df-note">${esc(d.note)}</div>`:''}`,
};

TYPE_SCHEMAS['checklist'] = {
    label: 'checklist — チェックリスト',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'チェックリスト' },
        { name: 'items', label: '項目', type: 'array',
          subfields: [
            { name: 'text', label: '項目', type: 'text', default: '' },
            { name: 'done', label: '完了', type: 'checkbox', default: false },
          ], default: [{text:'要件定義',done:true},{text:'実装',done:false}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<li${it.done?' class="done"':''}>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: checklist -->\\n# ${esc(d.h1)}\\n<div class="cl-container">\\n<ul>\\n${items}\\n</ul>\\n</div>`;
    },
};

TYPE_SCHEMAS['highlight'] = {
    label: 'highlight — 強調メッセージ',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '強調' },
        { name: 'text', label: '強調する文', type: 'textarea', default: '' },
    ],
    toMd: d => `<!-- _class: highlight -->\\n# ${esc(d.h1)}\\n<div class="hl-text">${esc(d.text)}</div>`,
};

TYPE_SCHEMAS['rq'] = {
    label: 'rq — 研究課題',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '研究課題' },
        { name: 'main', label: 'メインの問い', type: 'textarea', default: '' },
        { name: 'sub', label: '補足', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: rq -->\\n# ${esc(d.h1)}\\n<div class="rq-main">${esc(d.main)}</div>${d.sub?`\\n<div class="rq-sub">${esc(d.sub)}</div>`:''}`,
};

TYPE_SCHEMAS['code'] = {
    label: 'code — コード',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'コード例' },
        { name: 'lang', label: '言語', type: 'text', default: 'python' },
        { name: 'code', label: 'コード', type: 'textarea', default: 'def hello():\\n    print("hi")' },
        { name: 'desc', label: '説明', type: 'text', default: '' },
    ],
    toMd: function(d) {
        var bt = String.fromCharCode(96,96,96);
        return '<!-- _class: code -->\\n# ' + esc(d.h1) +
               '\\n<div class="cd-code">\\n\\n' + bt + esc(d.lang) +
               '\\n' + esc(d.code) + '\\n' + bt + '\\n</div>' +
               (d.desc ? '\\n<div class="cd-desc">' + esc(d.desc) + '</div>' : '');
    },
};

TYPE_SCHEMAS['zone-flow'] = {
    label: 'zone-flow — フロー',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'フロー' },
        { name: 'items', label: 'ステップ', type: 'array',
          subfields: [
            { name: 'label', label: 'ラベル', type: 'text', default: '' },
            { name: 'body', label: '本文', type: 'text', default: '' },
          ], default: [{label:'入力',body:''},{label:'処理',body:''},{label:'出力',body:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="zf-label">${esc(it.label)}</span><span class="zf-body">${esc(it.body)}</span></div>`).join('\\n');
        return `<!-- _class: zone-flow -->\\n# ${esc(d.h1)}\\n<div class="zf-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['steps'] = {
    label: 'steps — 手順',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '手順' },
        { name: 'items', label: 'ステップ', type: 'array',
          subfields: [
            { name: 'num', label: 'ステップ番号', type: 'text', default: '1' },
            { name: 'title', label: 'タイトル', type: 'text', default: '' },
            { name: 'body', label: '説明', type: 'text', default: '' },
          ], default: [{num:'1',title:'準備',body:''},{num:'2',title:'実行',body:''},{num:'3',title:'確認',body:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="st-num">${esc(it.num)}</span><span class="st-title">${esc(it.title)}</span><span class="st-body">${esc(it.body)}</span></div>`).join('\\n');
        return `<!-- _class: steps -->\\n# ${esc(d.h1)}\\n<div class="st-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['figure'] = {
    label: 'figure — 画像＋キャプション',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '図' },
        { name: 'image', label: '画像パス (MDファイルからの相対)', type: 'text', default: 'assets/image.png' },
        { name: 'width', label: '幅 (px、省略可)', type: 'text', default: '800' },
        { name: 'caption', label: 'キャプション', type: 'text', default: '' },
        { name: 'desc', label: '補足説明', type: 'textarea', default: '' },
    ],
    toMd: d => {
        const w = d.width ? `w:${esc(d.width)}` : '';
        return `<!-- _class: figure -->\\n# ${esc(d.h1)}\\n![${w}](${esc(d.image)})${d.caption?`\\n<div class="caption">${esc(d.caption)}</div>`:''}${d.desc?`\\n<div class="description">\\n${esc(d.desc)}\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['diagram'] = {
    label: 'diagram — 構造図',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '図解' },
        { name: 'image', label: '画像パス', type: 'text', default: 'assets/diagram.png' },
        { name: 'caption', label: 'キャプション', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: diagram -->\\n# ${esc(d.h1)}\\n![](${esc(d.image)})${d.caption?`\\n<div class="caption">${esc(d.caption)}</div>`:''}`,
};

TYPE_SCHEMAS['panorama'] = {
    label: 'panorama — 横長画像',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '全景' },
        { name: 'image', label: '画像パス', type: 'text', default: 'assets/panorama.png' },
        { name: 'text', label: '下部コメント', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: panorama -->\\n# ${esc(d.h1)}\\n![](${esc(d.image)})${d.text?`\\n<div class="pn-text">${esc(d.text)}</div>`:''}`,
};

TYPE_SCHEMAS['annotation'] = {
    label: 'annotation — 図＋注釈',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '注釈' },
        { name: 'image', label: '画像パス', type: 'text', default: 'assets/figure.png' },
        { name: 'notes', label: '注釈項目', type: 'array',
          subfields: [{ name: 'text', label: '注釈', type: 'text', default: '' }],
          default: [{text:'注釈1'},{text:'注釈2'}] },
    ],
    toMd: d => {
        const notes = (d.notes||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: annotation -->\\n# ${esc(d.h1)}\\n<div class="an-figure">\\n![](${esc(d.image)})\\n</div>\\n<div class="an-notes">\\n<ul>\\n${notes}\\n</ul>\\n</div>`;
    },
};

TYPE_SCHEMAS['gallery-img'] = {
    label: 'gallery-img — 画像ギャラリー',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'ギャラリー' },
        { name: 'items', label: '画像項目', type: 'array',
          subfields: [
            { name: 'image', label: '画像パス', type: 'text', default: '' },
            { name: 'caption', label: 'キャプション', type: 'text', default: '' },
          ], default: [{image:'',caption:''},{image:'',caption:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div>\\n![](${esc(it.image)})\\n<div class="gi-caption">${esc(it.caption)}</div>\\n</div>`).join('\\n');
        return `<!-- _class: gallery-img -->\\n# ${esc(d.h1)}\\n<div class="gi-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['table-slide'] = {
    label: 'table-slide — 表',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '表' },
        { name: 'table', label: 'MD表 (| col1 | col2 |形式)', type: 'textarea', hint: '先頭行がヘッダー、2行目が | --- | --- |',
          default: '| 手法 | 精度 | 速度 |\\n| --- | --- | --- |\\n| A | 85% | 1.0s |\\n| B | 92% | 1.5s |\\n| **Ours** | **97%** | **0.8s** |' },
        { name: 'note', label: '補足（任意）', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: table-slide -->\\n# ${esc(d.h1)}\\n${esc(d.table)}${d.note?`\\n<div class="box-accent">${esc(d.note)}</div>`:''}`,
};

TYPE_SCHEMAS['timeline'] = {
    label: 'timeline — 縦タイムライン',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '経過' },
        { name: 'items', label: '時点', type: 'array',
          subfields: [
            { name: 'year', label: '年', type: 'text', default: '' },
            { name: 'text', label: '出来事', type: 'text', default: '' },
            { name: 'detail', label: '詳細', type: 'text', default: '' },
            { name: 'highlight', label: '強調', type: 'checkbox', default: false },
          ], default: [{year:'2024',text:'企画'},{year:'2025',text:'開発',highlight:true},{year:'2026',text:'リリース'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => {
            const dt = it.detail ? `<div class="tl-detail">${esc(it.detail)}</div>` : '';
            return `<div class="tl-item${it.highlight?' highlight':''}"><span class="tl-year">${esc(it.year)}</span><span class="tl-text">${esc(it.text)}</span>${dt}</div>`;
        }).join('\\n');
        return `<!-- _class: timeline -->\\n# ${esc(d.h1)}\\n<div class="tl-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['history'] = {
    label: 'history — 沿革',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '沿革' },
        { name: 'items', label: '出来事', type: 'array',
          subfields: [
            { name: 'year', label: '年', type: 'text', default: '' },
            { name: 'event', label: '出来事', type: 'text', default: '' },
          ], default: [{year:'2000',event:'設立'},{year:'2010',event:'○○賞受賞'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="hs-year">${esc(it.year)}</span><span class="hs-event">${esc(it.event)}</span></div>`).join('\\n');
        return `<!-- _class: history -->\\n# ${esc(d.h1)}\\n<div class="hs-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['before-after'] = {
    label: 'before-after — 変化',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '改善' },
        { name: 'before_label', label: 'Before ラベル', type: 'text', default: 'Before' },
        { name: 'before_body', label: 'Before 内容', type: 'textarea', default: '' },
        { name: 'after_label', label: 'After ラベル', type: 'text', default: 'After' },
        { name: 'after_body', label: 'After 内容', type: 'textarea', default: '' },
    ],
    toMd: d => `<!-- _class: before-after -->\\n# ${esc(d.h1)}\\n<div class="ba-before">\\n<span class="ba-label">${esc(d.before_label)}</span>\\n<span class="ba-body">${esc(d.before_body)}</span>\\n</div>\\n<div class="ba-after">\\n<span class="ba-label">${esc(d.after_label)}</span>\\n<span class="ba-body">${esc(d.after_body)}</span>\\n</div>`,
};

TYPE_SCHEMAS['stack'] = {
    label: 'stack — 積層',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '構成' },
        { name: 'items', label: 'レイヤー (下から上に積む)', type: 'array',
          subfields: [
            { name: 'name', label: '名前', type: 'text', default: '' },
            { name: 'desc', label: '説明', type: 'text', default: '' },
          ], default: [{name:'OS',desc:'基盤'},{name:'DB',desc:'データ'},{name:'App',desc:'アプリケーション'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="sk-name">${esc(it.name)}</span><span class="sk-desc">${esc(it.desc)}</span></div>`).join('\\n');
        return `<!-- _class: stack -->\\n# ${esc(d.h1)}\\n<div class="sk-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['card-grid'] = {
    label: 'card-grid — カード一覧',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'カード一覧' },
        { name: 'items', label: 'カード', type: 'array',
          subfields: [
            { name: 'title', label: 'タイトル', type: 'text', default: '' },
            { name: 'body', label: '本文', type: 'text', default: '' },
          ], default: [{title:'A',body:''},{title:'B',body:''},{title:'C',body:''},{title:'D',body:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="cg-title">${esc(it.title)}</span><span class="cg-body">${esc(it.body)}</span></div>`).join('\\n');
        return `<!-- _class: card-grid -->\\n# ${esc(d.h1)}\\n<div class="cg-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['split-text'] = {
    label: 'split-text — 左右分割テキスト',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '左右' },
        { name: 'left_label', label: '左ラベル', type: 'text', default: '観点A' },
        { name: 'left_body', label: '左本文', type: 'textarea', default: '' },
        { name: 'right_label', label: '右ラベル', type: 'text', default: '観点B' },
        { name: 'right_body', label: '右本文', type: 'textarea', default: '' },
    ],
    toMd: d => `<!-- _class: split-text -->\\n# ${esc(d.h1)}\\n<div class="sp-left">\\n<span class="sp-label">${esc(d.left_label)}</span>\\n<span class="sp-body">${esc(d.left_body)}</span>\\n</div>\\n<div class="sp-right">\\n<span class="sp-label">${esc(d.right_label)}</span>\\n<span class="sp-body">${esc(d.right_body)}</span>\\n</div>`,
};

TYPE_SCHEMAS['zone-compare'] = {
    label: 'zone-compare — VS比較',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '対比' },
        { name: 'left_label', label: '左 ラベル', type: 'text', default: '案A' },
        { name: 'left_body', label: '左 本文', type: 'textarea', default: '' },
        { name: 'vs', label: '中央記号', type: 'text', default: 'VS' },
        { name: 'right_label', label: '右 ラベル', type: 'text', default: '案B' },
        { name: 'right_body', label: '右 本文', type: 'textarea', default: '' },
    ],
    toMd: d => `<!-- _class: zone-compare -->\\n# ${esc(d.h1)}\\n<div class="zc-left">\\n<span class="zc-label">${esc(d.left_label)}</span>\\n<span class="zc-body">${esc(d.left_body)}</span>\\n</div>\\n<div class="zc-vs">${esc(d.vs)}</div>\\n<div class="zc-right">\\n<span class="zc-label">${esc(d.right_label)}</span>\\n<span class="zc-body">${esc(d.right_body)}</span>\\n</div>`,
};

TYPE_SCHEMAS['zone-process'] = {
    label: 'zone-process — プロセス (番号+詳細)',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'プロセス' },
        { name: 'items', label: 'ステップ', type: 'array',
          subfields: [
            { name: 'step', label: '番号', type: 'text', default: '1' },
            { name: 'title', label: 'タイトル', type: 'text', default: '' },
            { name: 'body', label: '説明', type: 'text', default: '' },
          ], default: [{step:'1',title:'計画'},{step:'2',title:'実行'},{step:'3',title:'評価'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="zp-num">${esc(it.step)}</span><span class="zp-title">${esc(it.title)}</span><span class="zp-body">${esc(it.body)}</span></div>`).join('\\n');
        return `<!-- _class: zone-process -->\\n# ${esc(d.h1)}\\n<div class="zp-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['overview'] = {
    label: 'overview — 全体像 (画像+ポイント)',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '全体像' },
        { name: 'lead', label: 'リード文', type: 'text', default: '' },
        { name: 'image', label: '画像パス', type: 'text', default: '' },
        { name: 'caption', label: 'キャプション', type: 'text', default: '' },
        { name: 'points', label: 'ポイント', type: 'array',
          subfields: [{ name: 'text', label: 'ポイント', type: 'text', default: '' }],
          default: [{text:'ポイント1'},{text:'ポイント2'}] },
    ],
    toMd: d => {
        const pts = (d.points||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: overview -->\\n# ${esc(d.h1)}${d.lead?`\\n<div class="ov-lead">${esc(d.lead)}</div>`:''}${d.image?`\\n![](${esc(d.image)})`:''}${d.caption?`\\n<div class="caption">${esc(d.caption)}</div>`:''}${pts?`\\n<div class="ov-points">\\n<ul>\\n${pts}\\n</ul>\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['result'] = {
    label: 'result — 結果 (図+分析)',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '結果' },
        { name: 'lead', label: 'リード文', type: 'text', default: '' },
        { name: 'figure', label: '図の画像パス', type: 'text', default: '' },
        { name: 'caption', label: '図のキャプション', type: 'text', default: '' },
        { name: 'analysis', label: '分析ポイント', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'従来比10倍高速'},{text:'精度も向上'}] },
    ],
    toMd: d => {
        const analysis = (d.analysis||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: result -->\\n# ${esc(d.h1)}${d.lead?`\\n<div class="rs-lead">${esc(d.lead)}</div>`:''}${d.figure?`\\n<div class="rs-figure">\\n![](${esc(d.figure)})${d.caption?`\\n<div class="caption">${esc(d.caption)}</div>`:''}\\n</div>`:''}${analysis?`\\n<div class="rs-analysis">\\n<ul>\\n${analysis}\\n</ul>\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['result-dual'] = {
    label: 'result-dual — 2つの結果並列',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '結果' },
        { name: 'items', label: '結果', type: 'array',
          subfields: [
            { name: 'image', label: '画像パス', type: 'text', default: '' },
            { name: 'caption', label: 'キャプション', type: 'text', default: '' },
          ], default: [{image:'',caption:'結果A'},{image:'',caption:'結果B'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div>\\n![](${esc(it.image)})\\n<div class="caption">${esc(it.caption)}</div>\\n</div>`).join('\\n');
        return `<!-- _class: result-dual -->\\n# ${esc(d.h1)}\\n<div class="results">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['multi-result'] = {
    label: 'multi-result — 複数結果の一覧',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '結果一覧' },
        { name: 'items', label: '結果カード', type: 'array',
          subfields: [
            { name: 'metric', label: '指標名', type: 'text', default: '' },
            { name: 'value', label: '値', type: 'text', default: '' },
            { name: 'desc', label: '説明', type: 'text', default: '' },
          ], default: [{metric:'精度',value:'97%',desc:''},{metric:'速度',value:'0.8s',desc:''},{metric:'メモリ',value:'-50%',desc:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="mr-metric">${esc(it.metric)}</span><span class="mr-value">${esc(it.value)}</span><span class="mr-desc">${esc(it.desc)}</span></div>`).join('\\n');
        return `<!-- _class: multi-result -->\\n# ${esc(d.h1)}\\n<div class="mr-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['references'] = {
    label: 'references — 参考文献',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '参考文献' },
        { name: 'items', label: '文献', type: 'array',
          subfields: [
            { name: 'author', label: '著者', type: 'text', default: '' },
            { name: 'title', label: 'タイトル', type: 'text', default: '' },
            { name: 'venue', label: '出典', type: 'text', default: '' },
          ], default: [{author:'',title:'',venue:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<li><span class="author">${esc(it.author)}</span> <span class="title">${esc(it.title)}</span> <span class="venue">${esc(it.venue)}</span></li>`).join('\\n');
        return `<!-- _class: references -->\\n# ${esc(d.h1)}\\n<ol>\\n${items}\\n</ol>`;
    },
};

TYPE_SCHEMAS['appendix'] = {
    label: 'appendix — 補足',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '補足' },
        { name: 'label', label: 'ラベル (右上)', type: 'text', default: 'APPENDIX A' },
        { name: 'body', label: '本文 (Markdown可)', type: 'textarea', default: '' },
    ],
    toMd: d => `<!-- _class: appendix -->\\n# ${esc(d.h1)}\\n<span class="appendix-label">${esc(d.label)}</span>\\n\\n${esc(d.body)}`,
};

TYPE_SCHEMAS['profile'] = {
    label: 'profile — 人物紹介',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'プロフィール' },
        { name: 'image', label: '写真パス', type: 'text', default: '' },
        { name: 'name', label: '名前', type: 'text', default: '' },
        { name: 'affiliation', label: '所属', type: 'text', default: '' },
        { name: 'bio', label: '経歴', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'2020 大学卒業'},{text:'2022 入社'}] },
    ],
    toMd: d => {
        const bio = (d.bio||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: profile -->\\n# ${esc(d.h1)}${d.image?`\\n![](${esc(d.image)})`:''}\\n<div class="pf-container">\\n<div class="pf-name">${esc(d.name)}</div>\\n<div class="pf-affiliation">${esc(d.affiliation)}</div>${bio?`\\n<div class="pf-bio">\\n<ul>\\n${bio}\\n</ul>\\n</div>`:''}\\n</div>`;
    },
};

TYPE_SCHEMAS['equations'] = {
    label: 'equations — 連立/最適化問題',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '最適化問題' },
        { name: 'rows', label: '式の行', type: 'array',
          subfields: [
            { name: 'label', label: 'ラベル (例: minimize)', type: 'text', default: '' },
            { name: 'latex', label: 'LaTeX式', type: 'text', default: '' },
          ], default: [{label:'minimize',latex:'f(x) = \\\\|Ax - b\\\\|^2'},{label:'subject to',latex:'Ax \\\\le b'},{label:'',latex:'x \\\\ge 0'}] },
    ],
    toMd: d => {
        const rows = (d.rows||[]).map(r => {
            const lbl = r.label ? `<span class="label">${esc(r.label)}</span> ` : '';
            return `<div class="row">${lbl}$$${esc(r.latex)}$$</div>`;
        }).join('\\n');
        return `<!-- _class: equations -->\\n# ${esc(d.h1)}\\n<div class="eq-system">\\n${rows}\\n</div>`;
    },
};

// ── Reverse parser: MD slide text → form data ──
// Helpers
function _rxH1(t) { const m = /^#\\s+(.+)$/m.exec(t); return m ? m[1].trim() : ''; }
function _rxH2(t) { const m = /^##\\s+(.+)$/m.exec(t); return m ? m[1].trim() : ''; }
function _rxStripDirective(t) { return t.replace(/<!--\\s*_class:\\s*\\S+\\s*-->/, '').replace(/<!--\\s*_paginate:[^-]+-->/, ''); }
function _rxDiv(t, cls) {
    const re = new RegExp('<div\\\\s+class="[^"]*\\\\b' + cls + '\\\\b[^"]*">', 'i');
    const m = re.exec(t);
    if (!m) return null;
    const start = m.index + m[0].length;
    let depth = 1, pos = start;
    while (pos < t.length && depth > 0) {
        const no = t.indexOf('<div', pos);
        const nc = t.indexOf('</div>', pos);
        if (nc === -1) break;
        if (no !== -1 && no < nc) { depth++; pos = no + 4; }
        else { depth--; if (depth === 0) return t.substring(start, nc).trim(); pos = nc + 6; }
    }
    return null;
}
function _rxChildDivs(t) {
    const results = [];
    let pos = 0;
    while (pos < t.length) {
        const m = /<div[^>]*>/.exec(t.substring(pos));
        if (!m) break;
        const ds = pos + m.index + m[0].length;
        let depth = 1, scan = ds;
        while (scan < t.length && depth > 0) {
            const no = t.indexOf('<div', scan);
            const nc = t.indexOf('</div>', scan);
            if (nc === -1) break;
            if (no !== -1 && no < nc) { depth++; scan = no + 4; }
            else {
                depth--;
                if (depth === 0) { results.push(t.substring(ds, nc).trim()); pos = nc + 6; break; }
                scan = nc + 6;
            }
        }
        if (depth !== 0) break;
    }
    return results;
}
function _rxSpan(t, cls) {
    const re = new RegExp('<span[^>]*class="[^"]*\\\\b' + cls + '\\\\b[^"]*"[^>]*>([\\\\s\\\\S]*?)<\\\\/span>', 'i');
    const m = re.exec(t);
    return m ? _rxStripHtml(m[1]) : '';
}
function _rxLis(t) {
    const results = [];
    const re = /<li(\\s+class="([^"]+)")?>([\\s\\S]*?)<\\/li>/gi;
    let m;
    while ((m = re.exec(t)) !== null) {
        results.push({ cls: m[2] || '', text: _rxStripHtml(m[3]) });
    }
    return results;
}
function _rxStripHtml(s) { return (s || '').replace(/<[^>]+>/g, '').trim(); }
function _rxImage(t) { const m = /!\\[(?:w:(\\d+))?\\]\\(([^)]+)\\)/.exec(t); return m ? { path: m[2], width: m[1] || '' } : null; }
function _rxMath(t, display) {
    const re = display ? /\\$\\$([\\s\\S]+?)\\$\\$/ : /\\$([^$]+)\\$/;
    const m = re.exec(t);
    return m ? m[1].trim() : '';
}

// Main parser dispatcher
function parseMdForType(cls, mdText) {
    const t = mdText;
    const h1 = _rxH1(t);
    const h2 = _rxH2(t);
    switch (cls) {
        case 'title': {
            const lines = t.split('\\n').map(l => l.trim()).filter(l => l && !l.startsWith('#') && !l.startsWith('<!--'));
            return { h1, h2, author: lines[0] || '', date: lines[1] || '' };
        }
        case 'divider': return { h1, h2 };
        case 'end': {
            const lines = t.split('\\n').map(l => l.trim()).filter(l => l && !l.startsWith('#') && !l.startsWith('<!--'));
            return { h1: h1 || 'Thank You', sub: lines[0] || '' };
        }
        case 'plain': {
            const body = t.replace(/^#\\s+.+$/m, '').replace(/<!--[^>]*-->/g, '').trim();
            return { h1, body };
        }
        case 'rq':
            return { h1, main: _rxStripHtml(_rxDiv(t, 'rq-main') || ''), sub: _rxStripHtml(_rxDiv(t, 'rq-sub') || '') };
        case 'quote':
            return { h1, text: _rxStripHtml(_rxDiv(t, 'qt-text') || ''), source: _rxStripHtml(_rxDiv(t, 'qt-source') || '') };
        case 'definition':
            return { h1, term: _rxStripHtml(_rxDiv(t, 'df-term') || ''), body: _rxStripHtml(_rxDiv(t, 'df-body') || ''), note: _rxStripHtml(_rxDiv(t, 'df-note') || '') };
        case 'highlight':
            return { h1, text: _rxStripHtml(_rxDiv(t, 'hl-text') || '') };
        case 'takeaway': {
            const main = _rxStripHtml(_rxDiv(t, 'ta-main') || '');
            const pts = _rxDiv(t, 'ta-points') || '';
            return { h1, main, points: _rxLis(pts).map(li => ({ text: li.text })) };
        }
        case 'agenda': {
            const block = _rxDiv(t, 'agenda-list') || '';
            const items = [];
            const re = /^\\s*\\d+\\.\\s+(.+)$/gm;
            let m;
            while ((m = re.exec(block)) !== null) items.push({ text: _rxStripHtml(m[1]) });
            return { h1, items };
        }
        case 'summary': {
            let block = _rxDiv(t, 'summary-points');
            if (!block) {
                const m = /<ol[^>]*class="[^"]*summary-points[^"]*"[^>]*>([\\s\\S]*?)<\\/ol>/i.exec(t);
                block = m ? m[1] : '';
            }
            return { h1, items: _rxLis(block).map(li => ({ text: li.text })) };
        }
        case 'kpi': {
            const container = _rxDiv(t, 'kpi-container') || '';
            const items = _rxChildDivs(container).map(d => ({ value: _rxSpan(d, 'kpi-value'), label: _rxSpan(d, 'kpi-label') }));
            return { h1, items };
        }
        case 'funnel': {
            const container = _rxDiv(t, 'fn-container') || '';
            const items = _rxChildDivs(container).map(d => ({ label: _rxSpan(d, 'fn-label'), value: _rxSpan(d, 'fn-value') }));
            return { h1, items };
        }
        case 'pros-cons': {
            const pros = _rxLis(_rxDiv(t, 'pc-pros') || '').map(li => ({ text: li.text }));
            const cons = _rxLis(_rxDiv(t, 'pc-cons') || '').map(li => ({ text: li.text }));
            return { h1, pros, cons };
        }
        case 'timeline-h': {
            const container = _rxDiv(t, 'tl-h-container') || '';
            const items = _rxChildDivs(container).map(d => {
                const inner = _rxChildDivs(d)[0] || d;
                return {
                    year: _rxSpan(inner, 'tl-h-year'),
                    text: _rxSpan(inner, 'tl-h-text'),
                    highlight: /\\bhighlight\\b/.test(d.split('>')[0] || d),
                };
            });
            return { h1, items };
        }
        case 'timeline': {
            const container = _rxDiv(t, 'tl-container') || '';
            const items = _rxChildDivs(container).map(d => ({
                year: _rxSpan(d, 'tl-year'),
                text: _rxSpan(d, 'tl-text'),
                detail: _rxStripHtml(_rxDiv(d, 'tl-detail') || ''),
                highlight: /\\bhighlight\\b/.test(d.split('>')[0] || d),
            }));
            return { h1, items };
        }
        case 'history': {
            const container = _rxDiv(t, 'hs-container') || '';
            const items = _rxChildDivs(container).map(d => ({ year: _rxSpan(d, 'hs-year'), event: _rxSpan(d, 'hs-event') }));
            return { h1, items };
        }
        case 'checklist': {
            const container = _rxDiv(t, 'cl-container') || '';
            const items = _rxLis(container).map(li => ({ text: li.text, done: /\\bdone\\b/.test(li.cls) }));
            return { h1, items };
        }
        case 'steps': {
            const container = _rxDiv(t, 'st-container') || '';
            const items = _rxChildDivs(container).map((d, i) => ({
                num: _rxSpan(d, 'st-num') || String(i+1),
                title: _rxSpan(d, 'st-title'),
                body: _rxSpan(d, 'st-body'),
            }));
            return { h1, items };
        }
        case 'stack': {
            const container = _rxDiv(t, 'sk-container') || '';
            const items = _rxChildDivs(container).map(d => ({ name: _rxSpan(d, 'sk-name'), desc: _rxSpan(d, 'sk-desc') }));
            return { h1, items };
        }
        case 'card-grid': {
            const container = _rxDiv(t, 'cg-container') || '';
            const items = _rxChildDivs(container).map(d => ({ title: _rxSpan(d, 'cg-title'), body: _rxSpan(d, 'cg-body') }));
            return { h1, items };
        }
        case 'zone-flow': {
            const container = _rxDiv(t, 'zf-container') || '';
            const items = _rxChildDivs(container).map(d => ({ label: _rxSpan(d, 'zf-label'), body: _rxSpan(d, 'zf-body') }));
            return { h1, items };
        }
        case 'zone-process': {
            const container = _rxDiv(t, 'zp-container') || '';
            const items = _rxChildDivs(container).map((d, i) => ({
                step: _rxSpan(d, 'zp-num') || String(i+1),
                title: _rxSpan(d, 'zp-title'),
                body: _rxSpan(d, 'zp-body'),
            }));
            return { h1, items };
        }
        case 'zone-compare': {
            const left = _rxDiv(t, 'zc-left') || '';
            const right = _rxDiv(t, 'zc-right') || '';
            const vs = _rxStripHtml(_rxDiv(t, 'zc-vs') || '') || 'VS';
            return {
                h1,
                left_label: _rxSpan(left, 'zc-label'), left_body: _rxSpan(left, 'zc-body'),
                right_label: _rxSpan(right, 'zc-label'), right_body: _rxSpan(right, 'zc-body'),
                vs
            };
        }
        case 'zone-matrix': {
            const get = c => { const d = _rxDiv(t, c) || ''; return { label: _rxSpan(d, 'zm-label'), body: _rxSpan(d, 'zm-body') }; };
            const tl = get('zm-tl'), tr = get('zm-tr'), bl = get('zm-bl'), br = get('zm-br');
            return {
                h1,
                x_label: _rxStripHtml(_rxDiv(t, 'zm-xlabel') || ''),
                y_label: _rxStripHtml(_rxDiv(t, 'zm-ylabel') || ''),
                tl_label: tl.label, tl_body: tl.body,
                tr_label: tr.label, tr_body: tr.body,
                bl_label: bl.label, bl_body: bl.body,
                br_label: br.label, br_body: br.body,
            };
        }
        case 'before-after': {
            const before = _rxDiv(t, 'ba-before') || '';
            const after = _rxDiv(t, 'ba-after') || '';
            return {
                h1,
                before_label: _rxSpan(before, 'ba-label'), before_body: _rxSpan(before, 'ba-body'),
                after_label: _rxSpan(after, 'ba-label'), after_body: _rxSpan(after, 'ba-body'),
            };
        }
        case 'split-text': {
            const left = _rxDiv(t, 'sp-left') || '';
            const right = _rxDiv(t, 'sp-right') || '';
            return {
                h1,
                left_label: _rxSpan(left, 'sp-label'), left_body: _rxSpan(left, 'sp-body'),
                right_label: _rxSpan(right, 'sp-label'), right_body: _rxSpan(right, 'sp-body'),
            };
        }
        case 'cols-2': case 'cols-3': {
            const cols = _rxDiv(t, 'columns') || '';
            const children = _rxChildDivs(cols);
            if (cls === 'cols-2') {
                const [l, r] = [children[0] || '', children[1] || ''];
                const parseH3 = s => { const m = /^###\\s+(.+)$/m.exec(s); return m ? m[1].trim() : ''; };
                const stripH3 = s => s.replace(/^###\\s+.+$/m, '').trim();
                return { h1, left_title: parseH3(l), left_body: stripH3(l), right_title: parseH3(r), right_body: stripH3(r) };
            } else {
                const parseH3 = s => { const m = /^###\\s+(.+)$/m.exec(s); return m ? m[1].trim() : ''; };
                const stripH3 = s => s.replace(/^###\\s+.+$/m, '').trim();
                return {
                    h1,
                    c1_title: parseH3(children[0] || ''), c1_body: stripH3(children[0] || ''),
                    c2_title: parseH3(children[1] || ''), c2_body: stripH3(children[1] || ''),
                    c3_title: parseH3(children[2] || ''), c3_body: stripH3(children[2] || ''),
                };
            }
        }
        case 'sandwich': {
            const top = _rxDiv(t, 'top') || '';
            const lead = _rxStripHtml(_rxDiv(top, 'lead') || top);
            const cols = _rxDiv(t, 'columns') || '';
            const children = _rxChildDivs(cols);
            const parseH3 = s => { const m = /^###\\s+(.+)$/m.exec(s); return m ? m[1].trim() : ''; };
            const stripH3 = s => s.replace(/^###\\s+.+$/m, '').trim();
            const conclusion = _rxStripHtml(_rxDiv(_rxDiv(t, 'bottom') || '', 'conclusion') || '');
            return {
                h1, lead,
                left_title: parseH3(children[0] || ''), left_body: stripH3(children[0] || ''),
                right_title: parseH3(children[1] || ''), right_body: stripH3(children[1] || ''),
                conclusion
            };
        }
        case 'figure': {
            const img = _rxImage(t);
            return {
                h1,
                image: img ? img.path : '',
                width: img ? img.width : '',
                caption: _rxStripHtml(_rxDiv(t, 'caption') || ''),
                desc: _rxStripHtml(_rxDiv(t, 'description') || ''),
            };
        }
        case 'diagram': case 'panorama': {
            const img = _rxImage(t);
            const result = { h1, image: img ? img.path : '' };
            if (cls === 'diagram') result.caption = _rxStripHtml(_rxDiv(t, 'caption') || '');
            if (cls === 'panorama') result.text = _rxStripHtml(_rxDiv(t, 'pn-text') || '');
            return result;
        }
        case 'annotation': {
            const fig = _rxDiv(t, 'an-figure') || '';
            const img = _rxImage(fig);
            const notes = _rxLis(_rxDiv(t, 'an-notes') || '').map(li => ({ text: li.text }));
            return { h1, image: img ? img.path : '', notes };
        }
        case 'gallery-img': {
            const container = _rxDiv(t, 'gi-container') || '';
            const items = _rxChildDivs(container).map(d => {
                const img = _rxImage(d);
                return { image: img ? img.path : '', caption: _rxStripHtml(_rxDiv(d, 'gi-caption') || '') };
            });
            return { h1, items };
        }
        case 'table-slide': {
            const lines = t.split('\\n').filter(l => l.trim().startsWith('|'));
            const note = _rxStripHtml(_rxDiv(t, 'box-accent') || '');
            return { h1, table: lines.join('\\n'), note };
        }
        case 'overview': {
            const img = _rxImage(t);
            const points = _rxLis(_rxDiv(t, 'ov-points') || '').map(li => ({ text: li.text }));
            return {
                h1,
                lead: _rxStripHtml(_rxDiv(t, 'ov-lead') || ''),
                image: img ? img.path : '',
                caption: _rxStripHtml(_rxDiv(t, 'caption') || ''),
                points,
            };
        }
        case 'result': {
            const fig = _rxDiv(t, 'rs-figure') || '';
            const img = _rxImage(fig);
            const analysis = _rxLis(_rxDiv(t, 'rs-analysis') || '').map(li => ({ text: li.text }));
            return {
                h1,
                lead: _rxStripHtml(_rxDiv(t, 'rs-lead') || ''),
                figure: img ? img.path : '',
                caption: _rxStripHtml(_rxDiv(fig, 'caption') || ''),
                analysis,
            };
        }
        case 'result-dual': {
            const container = _rxDiv(t, 'results') || '';
            const items = _rxChildDivs(container).map(d => {
                const img = _rxImage(d);
                return { image: img ? img.path : '', caption: _rxStripHtml(_rxDiv(d, 'caption') || '') };
            });
            return { h1, items };
        }
        case 'multi-result': {
            const container = _rxDiv(t, 'mr-container') || '';
            const items = _rxChildDivs(container).map(d => ({
                metric: _rxSpan(d, 'mr-metric'),
                value: _rxSpan(d, 'mr-value'),
                desc: _rxSpan(d, 'mr-desc'),
            }));
            return { h1, items };
        }
        case 'references': {
            const re = /<li>([\\s\\S]*?)<\\/li>/gi;
            const items = []; let m;
            while ((m = re.exec(t)) !== null) {
                const inner = m[1];
                const gs = (s, c) => { const mm = new RegExp('<span[^>]*class="' + c + '"[^>]*>([\\\\s\\\\S]*?)<\\\\/span>', 'i').exec(s); return mm ? _rxStripHtml(mm[1]) : ''; };
                items.push({ author: gs(inner, 'author'), title: gs(inner, 'title'), venue: gs(inner, 'venue') });
            }
            return { h1, items };
        }
        case 'appendix': {
            const lbl = /<span[^>]*class="[^"]*appendix-label[^"]*"[^>]*>([\\s\\S]*?)<\\/span>/i.exec(t);
            let body = t.replace(/^#\\s+.+$/m, '').replace(/<!--[^>]*-->/g, '').replace(/<span[^>]*appendix-label[\\s\\S]*?<\\/span>/i, '').trim();
            return { h1, label: lbl ? _rxStripHtml(lbl[1]) : 'APPENDIX', body };
        }
        case 'profile': {
            const img = _rxImage(t);
            const container = _rxDiv(t, 'pf-container') || '';
            const bio = _rxLis(_rxDiv(container, 'pf-bio') || '').map(li => ({ text: li.text }));
            return {
                h1,
                image: img ? img.path : '',
                name: _rxStripHtml(_rxDiv(container, 'pf-name') || ''),
                affiliation: _rxStripHtml(_rxDiv(container, 'pf-affiliation') || ''),
                bio,
            };
        }
        case 'equation': {
            const main = _rxDiv(t, 'eq-main') || '';
            const formula = _rxMath(main, true) || _rxStripHtml(main);
            const desc = _rxDiv(t, 'eq-desc') || '';
            const spans = [...desc.matchAll(/<span[^>]*>([\\s\\S]*?)<\\/span>/g)].map(m => _rxStripHtml(m[1]));
            const vars = [];
            for (let i = 0; i + 1 < spans.length; i += 2) {
                vars.push({ sym: spans[i].replace(/^\\$|\\$$/g, ''), desc: spans[i+1] });
            }
            return { h1, formula, vars };
        }
        case 'equations': {
            const sys = _rxDiv(t, 'eq-system') || '';
            const rows = _rxChildDivs(sys).map(row => {
                const lm = /<span[^>]*class="[^"]*label[^"]*"[^>]*>([\\s\\S]*?)<\\/span>/i.exec(row);
                const eq = /\\$\\$([\\s\\S]+?)\\$\\$/.exec(row);
                return { label: lm ? _rxStripHtml(lm[1]) : '', latex: eq ? eq[1].trim() : '' };
            });
            return { h1, rows };
        }
        case 'code': {
            const cd = _rxDiv(t, 'cd-code') || '';
            const m = /\\x60\\x60\\x60(\\w*)\\s*\\n([\\s\\S]*?)\\x60\\x60\\x60/.exec(cd);
            return {
                h1,
                lang: m ? m[1] : 'python',
                code: m ? m[2].replace(/\\s+$/, '') : _rxStripHtml(cd),
                desc: _rxStripHtml(_rxDiv(t, 'cd-desc') || ''),
            };
        }
    }
    return null;
}

// ── Modal management ──
function openModal(id) { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

const CAT_LABEL = {meta:'メタ',structure:'構造',temporal:'時間',convergence:'収束・拡散',evaluation:'評価・判断',knowledge:'知識・定義',flow:'流れ・構造',narrative:'ナラティブ'};
const CAT_ORDER = ['meta','structure','temporal','convergence','evaluation','knowledge','flow','narrative'];

function renderTypeGrid(filter) {
    const q = (filter || '').trim().toLowerCase();
    const grid = document.getElementById('type-grid');
    const byCategory = {};
    TYPES_META.forEach(t => {
        if (q) {
            const hay = `${t.name} ${t.meaning} ${t.geometry} ${t.use_when} ${CAT_LABEL[t.category] || ''}`.toLowerCase();
            if (!hay.includes(q)) return;
        }
        (byCategory[t.category] = byCategory[t.category] || []).push(t);
    });
    const parts = [];
    CAT_ORDER.forEach(cat => {
        if (!byCategory[cat]) return;
        parts.push(`<div class="type-category-header">${CAT_LABEL[cat] || cat}</div>`);
        byCategory[cat].forEach(t => {
            const hasForm = !!TYPE_SCHEMAS[t.css_class];
            parts.push(`<div class="type-card" onclick="selectType('${t.css_class}')" title="${t.use_when}">
                <div class="type-name">${t.name}${hasForm ? ' ✓' : ''}</div>
                <div class="type-geom">${t.geometry}</div>
                <div class="type-meaning">${t.meaning}</div>
            </div>`);
        });
    });
    if (parts.length === 0) {
        grid.innerHTML = '<div style="grid-column:1/-1; text-align:center; padding:40px; color:#999">一致する型がありません</div>';
    } else {
        grid.innerHTML = parts.join('');
    }
}

function filterTypes(q) { renderTypeGrid(q); }

async function openTypePicker() {
    if (TYPES_META.length === 0) await loadTypeMeta();
    document.getElementById('type-search').value = '';
    renderTypeGrid('');
    openModal('picker-modal');
    setTimeout(() => document.getElementById('type-search').focus(), 50);
}

let currentType = null;
let currentData = {};

function selectType(cssClass) {
    closeModal('picker-modal');
    const schema = TYPE_SCHEMAS[cssClass];
    if (!schema) {
        // No form yet — fall back to inserting a minimal snippet
        const sep = editor.value.trim() ? '\\n\\n---\\n\\n' : '';
        const meta = TYPES_META.find(t => t.css_class === cssClass);
        editor.value += sep + `<!-- _class: ${cssClass} -->\\n# ${meta ? meta.meaning : 'タイトル'}\\n`;
        updateStats();
        triggerAutoPreview();
        editor.focus();
        return;
    }
    currentType = cssClass;
    currentData = {};
    // Init defaults
    schema.fields.forEach(f => {
        currentData[f.name] = (f.default !== undefined)
            ? (f.type === 'array' ? JSON.parse(JSON.stringify(f.default)) : f.default)
            : (f.type === 'array' ? [] : (f.type === 'checkbox' ? false : ''));
    });
    document.getElementById('form-title').textContent = schema.label;
    document.getElementById('form-body').innerHTML = buildFormHtml(schema);
    openModal('form-modal');
}

function buildFormHtml(schema) {
    return schema.fields.map(f => buildFieldHtml(f, currentData[f.name], f.name)).join('');
}

function buildFieldHtml(f, value, path) {
    if (f.type === 'text') {
        // Image-ish fields (by naming convention) get an upload button
        const isImg = /^(image|figure)$/.test(f.name) || f.type === 'image';
        return `<div class="form-row">
            <label>${f.label}</label>
            <div style="display:flex; gap:6px; align-items:center">
                <input type="text" value="${escAttr(value||'')}" oninput="setField('${path}', this.value); previewImgAt('${path}', this.value)" style="flex:1" id="fld-${path.replace(/[\\[\\].]/g,'_')}">
                ${isImg ? `<button type="button" onclick="document.getElementById('imgup-${path.replace(/[\\[\\].]/g,'_')}').click()" style="padding:6px 10px">📎 選択</button>` : ''}
                ${isImg ? `<input type="file" id="imgup-${path.replace(/[\\[\\].]/g,'_')}" accept="image/*" style="display:none" onchange="handleImageUpload('${path}', this)">` : ''}
            </div>
            ${isImg && value ? `<img src="${imgToUrl(value)}" class="img-preview" alt="preview">` : ''}
            ${f.hint ? `<div class="hint">${f.hint}</div>` : ''}
        </div>`;
    } else if (f.type === 'textarea') {
        return `<div class="form-row">
            <label>${f.label}</label>
            <textarea oninput="setField('${path}', this.value)">${esc(value||'')}</textarea>
            ${f.hint ? `<div class="hint">${f.hint}</div>` : ''}
        </div>`;
    } else if (f.type === 'checkbox') {
        return `<div class="form-row">
            <label><input type="checkbox" ${value?'checked':''} onchange="setField('${path}', this.checked)"> ${f.label}</label>
        </div>`;
    } else if (f.type === 'array') {
        const items = (value||[]).map((item, i) => buildArrayItemHtml(f, item, path, i)).join('');
        return `<div class="form-row">
            <label>${f.label}</label>
            <div class="array-items" id="arr-${path}">${items}</div>
            <button type="button" class="add-item-btn" onclick="addArrayItem('${path}')">+ 項目を追加</button>
        </div>`;
    }
    return '';
}

function buildArrayItemHtml(f, item, path, index) {
    const sub = f.subfields.map(sf => {
        const val = item[sf.name] || '';
        const subPath = `${path}[${index}].${sf.name}`;
        const safeId = subPath.replace(/[\\[\\].]/g, '_');
        if (sf.type === 'checkbox') {
            return `<label style="font-size:0.8em"><input type="checkbox" ${item[sf.name]?'checked':''} onchange="setField('${subPath}', this.checked)"> ${sf.label}</label>`;
        }
        const isImg = /^(image|figure)$/.test(sf.name);
        const pickerBtn = isImg ? ` <button type="button" style="padding:3px 8px; font-size:0.75em" onclick="document.getElementById('imgup-${safeId}').click()">📎</button><input type="file" id="imgup-${safeId}" accept="image/*" style="display:none" onchange="handleImageUpload('${subPath}', this)">` : '';
        const imgPreview = isImg && val ? `<img src="${imgToUrl(val)}" class="img-preview-sm" alt="">` : '';
        return `<div>
            <label style="font-size:0.75em; margin-bottom:2px">${sf.label}</label>
            <div style="display:flex; gap:4px; align-items:center">
                <input type="text" value="${escAttr(val)}" oninput="setField('${subPath}', this.value)" style="padding:5px 8px; flex:1">${pickerBtn}
            </div>
            ${imgPreview}
        </div>`;
    }).join('');
    return `<div class="array-item">
        <div style="display:flex; flex-direction:column; gap:4px; flex:1">${sub}</div>
        <button type="button" class="remove-btn" onclick="removeArrayItem('${path}', ${index})">×</button>
    </div>`;
}

function imgToUrl(value) {
    if (!value) return '';
    // If it's our assets/xxx.yyy path, resolve to /editor/asset/xxx.yyy
    const m = /^assets\\/(.+)$/.exec(value);
    if (m) return '/editor/asset/' + m[1];
    return value;  // leave as-is (relative to MD; may 404 in the preview)
}

async function handleImageUpload(path, inputEl) {
    const f = inputEl.files[0];
    if (!f) return;
    try {
        const form = new FormData();
        form.append('file', f);
        const r = await fetch('/editor/upload-image', { method: 'POST', body: form });
        if (!r.ok) throw new Error(await r.text());
        const data = await r.json();
        setField(path, data.path);
        // Re-render the form to update the preview thumb
        document.getElementById('form-body').innerHTML = buildFormHtml(TYPE_SCHEMAS[currentType]);
    } catch(e) {
        alert('画像アップロード失敗: ' + e.message);
    }
    inputEl.value = '';
}

function previewImgAt(path, value) { /* reserved */ }

function escAttr(s) { return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;'); }

function setField(path, value) {
    // path like "h1" or "items[0].label"
    const m = path.match(/^(\\w+)(?:\\[(\\d+)\\]\\.(\\w+))?$/);
    if (!m) return;
    if (m[2] !== undefined) {
        currentData[m[1]][+m[2]][m[3]] = value;
    } else {
        currentData[m[1]] = value;
    }
}

function addArrayItem(path) {
    const schema = TYPE_SCHEMAS[currentType];
    const field = schema.fields.find(f => f.name === path);
    const newItem = {};
    field.subfields.forEach(sf => { newItem[sf.name] = sf.default !== undefined ? sf.default : (sf.type === 'checkbox' ? false : ''); });
    currentData[path].push(newItem);
    rerenderArray(path);
}

function removeArrayItem(path, index) {
    currentData[path].splice(index, 1);
    rerenderArray(path);
}

function rerenderArray(path) {
    const schema = TYPE_SCHEMAS[currentType];
    const field = schema.fields.find(f => f.name === path);
    const container = document.getElementById('arr-' + path);
    container.innerHTML = currentData[path].map((item, i) => buildArrayItemHtml(field, item, path, i)).join('');
}

function submitForm() {
    const schema = TYPE_SCHEMAS[currentType];
    const md = schema.toMd(currentData);
    const mode = document.querySelector('#form-modal .modal-footer .primary').getAttribute('data-mode');
    if (mode === 'edit' && editingSlideIndex >= 0) {
        // Replace the existing slide's MD with the regenerated form output
        const all = editor.value;
        const slides = splitIntoSlides(all);
        const s = slides[editingSlideIndex];
        if (s) {
            editor.value = all.substring(0, s.start) + md + '\\n' + all.substring(s.end);
        }
        editingSlideIndex = -1;
        document.querySelector('#form-modal .modal-footer .primary').textContent = 'スライドを追加';
        document.querySelector('#form-modal .modal-footer .primary').removeAttribute('data-mode');
    } else {
        const sep = editor.value.trim() ? '\\n\\n---\\n\\n' : '';
        editor.value += sep + md + '\\n';
    }
    updateStats();
    autoSave();
    triggerAutoPreview();
    closeModal('form-modal');
    editor.focus();
}

function insertSnippet(type) {
    if (type === 'plain') selectType('plain');
    else if (type === 'bullets') {
        const sep = editor.value.trim() ? '\\n\\n---\\n\\n' : '';
        editor.value += sep + `# 箇条書き\\n- ポイント1\\n- ポイント2\\n- ポイント3\\n`;
        updateStats(); triggerAutoPreview();
    }
    else if (type === 'divider') selectType('divider');
}

async function loadSample(name) {
    try {
        const r = await fetch('/editor/sample/' + name);
        if (!r.ok) throw new Error(await r.text());
        editor.value = await r.text();
        updateStats();
        editor.scrollTop = 0;
    } catch(e) {
        setStatus('サンプル読込失敗: ' + e.message, 'err');
    }
}

function setStatus(msg, kind) {
    statusEl.textContent = msg;
    statusEl.className = 'status ' + kind;
    if (kind === 'ok') setTimeout(() => { statusEl.className = 'status'; }, 3000);
}

async function generate() {
    const btn = document.getElementById('gen-btn');
    const md = editor.value;
    if (!md.trim()) { setStatus('Markdownが空です', 'err'); return; }
    btn.disabled = true; btn.textContent = '生成中...';
    try {
        const form = new FormData();
        form.append('markdown', md);
        form.append('palette', document.getElementById('palette').value);
        form.append('font_scale', fsRange.value);
        form.append('output_name', document.getElementById('output-name').value || 'slides.pptx');
        const r = await fetch('/editor/generate', { method: 'POST', body: form });
        if (!r.ok) throw new Error(await r.text());
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = document.getElementById('output-name').value || 'slides.pptx';
        document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(url);
        setStatus('生成完了 → ダウンロード', 'ok');
    } catch(e) {
        setStatus('生成失敗: ' + e.message, 'err');
    } finally {
        btn.disabled = false;
        btn.textContent = '→ PPTX を生成してダウンロード';
    }
}

async function refreshPreview() {
    const panel = document.getElementById('preview-content');
    const md = editor.value;
    if (!md.trim()) {
        panel.innerHTML = '<div class="preview-empty">エディタに内容を入れて<br>「更新」を押してください</div>';
        return;
    }
    panel.innerHTML = '<div class="preview-loading">レンダリング中...<br><small>(LibreOffice経由・数秒かかります)</small></div>';
    try {
        const form = new FormData();
        form.append('markdown', md);
        form.append('palette', document.getElementById('palette').value);
        form.append('font_scale', fsRange.value);
        const r = await fetch('/editor/preview', { method: 'POST', body: form });
        if (!r.ok) throw new Error(await r.text());
        const data = await r.json();
        if (!data.slides || data.slides.length === 0) {
            panel.innerHTML = '<div class="preview-empty">プレビュー生成失敗</div>';
            return;
        }
        panel.innerHTML = data.slides.map((url, i) => `
            <div class="slide-thumb">
                <img src="${url}" alt="slide ${i+1}" loading="lazy">
                <div class="caption"><span>Slide ${i+1}</span></div>
            </div>
        `).join('');
    } catch(e) {
        panel.innerHTML = '<div class="preview-empty" style="color:#c62828">エラー: ' + e.message + '</div>';
    }
}

// Debounced auto-preview
let previewTimer = null;
let autoPreviewEnabled = true;
function triggerAutoPreview() {
    if (!autoPreviewEnabled) return;
    if (previewTimer) clearTimeout(previewTimer);
    previewTimer = setTimeout(() => refreshPreview(), 1500);
}
editor.addEventListener('input', triggerAutoPreview);

// ── MD save / load ──
function downloadMd() {
    const md = editor.value;
    if (!md.trim()) { setStatus('Markdownが空です', 'err'); return; }
    const name = (document.getElementById('output-name').value || 'slides.pptx').replace(/\\.pptx$/, '.md');
    const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    setStatus('保存: ' + name, 'ok');
}

async function loadPptxFile(ev) {
    const f = ev.target.files[0];
    if (!f) return;
    if (editor.value.trim() && !confirm('現在の編集内容を破棄してPPTXを読み込みますか？')) {
        ev.target.value = ''; return;
    }
    setStatus('PPTX→MD変換中...', 'ok');
    try {
        const form = new FormData();
        form.append('file', f);
        const r = await fetch('/editor/pptx-to-md', { method: 'POST', body: form });
        if (!r.ok) throw new Error(await r.text());
        const data = await r.json();
        editor.value = data.markdown;
        const base = f.name.replace(/\\.pptx$/i, '');
        document.getElementById('output-name').value = base + '_editable.pptx';
        updateStats();
        autoSave();
        triggerAutoPreview();
        // Show inference report
        const inferred = data.slides.map((s, i) => `${i+1}: ${s.inferred_class || '(default)'}`).join(', ');
        setStatus(`読込: ${f.name} — 推論型: ${inferred}`, 'ok');
        editor.scrollTop = 0;
    } catch(e) {
        setStatus('PPTX読込失敗: ' + e.message, 'err');
    }
    ev.target.value = '';
}

function loadMdFile(ev) {
    const f = ev.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = e => {
        if (editor.value.trim() && !confirm('現在の編集内容を破棄して読み込みますか？')) {
            ev.target.value = ''; return;
        }
        editor.value = e.target.result;
        // Adjust output filename to match the loaded file
        const base = f.name.replace(/\\.(md|markdown)$/i, '');
        document.getElementById('output-name').value = base + '_editable.pptx';
        updateStats();
        autoSave();
        triggerAutoPreview();
        setStatus('読込: ' + f.name, 'ok');
        editor.scrollTop = 0;
    };
    reader.readAsText(f, 'utf-8');
    ev.target.value = '';
}

// ── LocalStorage autosave ──
const STORAGE_KEY = 'marp-pptx-editor-md';
const STORAGE_SETTINGS = 'marp-pptx-editor-settings';

function autoSave() {
    try {
        localStorage.setItem(STORAGE_KEY, editor.value);
        localStorage.setItem(STORAGE_SETTINGS, JSON.stringify({
            palette: document.getElementById('palette').value,
            fontScale: fsRange.value,
            outputName: document.getElementById('output-name').value,
        }));
    } catch (e) { /* ignore quota */ }
}

function restoreFromStorage() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved && saved.trim()) {
            editor.value = saved;
        }
        const s = localStorage.getItem(STORAGE_SETTINGS);
        if (s) {
            const o = JSON.parse(s);
            if (o.palette) document.getElementById('palette').value = o.palette;
            if (o.fontScale) { fsRange.value = o.fontScale; fsVal.textContent = parseFloat(o.fontScale).toFixed(2); }
            if (o.outputName) document.getElementById('output-name').value = o.outputName;
        }
    } catch (e) { /* ignore */ }
}

// Save on every change (debounced)
let saveTimer = null;
editor.addEventListener('input', () => {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(autoSave, 500);
});
['palette', 'output-name'].forEach(id => {
    document.getElementById(id).addEventListener('change', autoSave);
});
fsRange.addEventListener('change', autoSave);

// ── Slide jump / reorder / delete from preview ──
function splitIntoSlides(md) {
    // Returns array of {start, end, text} positions in the md string
    const slides = [];
    let pos = 0;
    // Strip leading frontmatter
    if (md.startsWith('---')) {
        const fmEnd = md.indexOf('---', 3);
        if (fmEnd !== -1) pos = fmEnd + 3;
    }
    const text = md.substring(pos);
    const re = /\\n---\\n/g;
    let last = 0;
    let m;
    while ((m = re.exec(text)) !== null) {
        const body = text.substring(last, m.index);
        if (body.trim()) slides.push({ start: pos + last, end: pos + m.index, text: body });
        last = m.index + m[0].length;
    }
    const body = text.substring(last);
    if (body.trim()) slides.push({ start: pos + last, end: pos + text.length, text: body });
    return slides;
}

function jumpToSlide(index) {
    const md = editor.value;
    const slides = splitIntoSlides(md);
    if (index < 0 || index >= slides.length) return;
    const s = slides[index];
    editor.focus();
    editor.setSelectionRange(s.start, s.end);
    // Scroll to the location
    const before = md.substring(0, s.start);
    const lineCount = (before.match(/\\n/g) || []).length;
    const lineHeight = 13 * 1.6;
    editor.scrollTop = Math.max(0, lineCount * lineHeight - 80);
}

let editingSlideIndex = -1;
function detectSlideType(mdText) {
    const m = /<!--\\s*_class:\\s*(\\S+)\\s*-->/.exec(mdText);
    return m ? m[1] : null;
}

function editSlide(index, forceRaw) {
    const md = editor.value;
    const slides = splitIntoSlides(md);
    if (index < 0 || index >= slides.length) return;
    editingSlideIndex = index;
    const text = slides[index].text.trim();
    const cls = detectSlideType(text) || 'plain';
    const schema = TYPE_SCHEMAS[cls];
    if (!forceRaw && schema) {
        // Try reverse-parse into form data
        try {
            const data = parseMdForType(cls, text);
            if (data) {
                currentType = cls;
                currentData = data;
                // Fill in defaults for any missing schema fields
                schema.fields.forEach(f => {
                    if (currentData[f.name] === undefined) {
                        currentData[f.name] = f.default !== undefined
                            ? (f.type === 'array' ? JSON.parse(JSON.stringify(f.default)) : f.default)
                            : (f.type === 'array' ? [] : (f.type === 'checkbox' ? false : ''));
                    }
                });
                document.getElementById('form-title').textContent = schema.label + ' を編集';
                document.getElementById('form-body').innerHTML = buildFormHtml(schema) +
                    `<div style="margin-top:16px; padding-top:12px; border-top:1px solid #eee; font-size:0.85em">
                       <a href="javascript:void(0)" onclick="closeModal('form-modal'); editSlide(${index}, true)" style="color:#666">生MDで編集する</a>
                     </div>`;
                openModal('form-modal');
                // Mark the form as "editing" mode
                document.querySelector('#form-modal .modal-footer .primary').textContent = '保存';
                document.querySelector('#form-modal .modal-footer .primary').setAttribute('data-mode', 'edit');
                return;
            }
        } catch (e) { console.warn('fromMd failed:', e); }
    }
    // Fall back to raw MD editor
    document.getElementById('slide-edit-ta').value = text;
    document.getElementById('slide-edit-idx').textContent = `Slide ${index + 1} [${cls}]`;
    openModal('slide-edit-modal');
    setTimeout(() => document.getElementById('slide-edit-ta').focus(), 50);
}

function saveSlideEdit() {
    if (editingSlideIndex < 0) return;
    const newText = document.getElementById('slide-edit-ta').value.trim();
    const md = editor.value;
    const slides = splitIntoSlides(md);
    const s = slides[editingSlideIndex];
    if (!s) return;
    // Replace just this slide's content
    editor.value = md.substring(0, s.start) + newText + '\\n' + md.substring(s.end);
    updateStats(); autoSave(); triggerAutoPreview();
    closeModal('slide-edit-modal');
    editingSlideIndex = -1;
}

function deleteSlide(index) {
    if (!confirm(`Slide ${index+1} を削除しますか？`)) return;
    const md = editor.value;
    const slides = splitIntoSlides(md);
    if (index < 0 || index >= slides.length) return;
    const s = slides[index];
    // Remove the slide body plus a neighboring `---` separator
    let left = s.start;
    let right = s.end;
    // If there's a `---` separator right after (between this and next), include it
    if (md.substring(right, right + 5) === '\\n---\\n') right += 5;
    else if (left >= 5 && md.substring(left - 5, left) === '\\n---\\n') left -= 5;
    editor.value = md.substring(0, left) + md.substring(right);
    updateStats(); autoSave(); triggerAutoPreview();
}

function moveSlide(index, direction) {
    const md = editor.value;
    const slides = splitIntoSlides(md);
    const target = index + direction;
    if (target < 0 || target >= slides.length || index < 0 || index >= slides.length) return;
    // Rebuild MD with slides in new order
    const frontmatter = md.startsWith('---') ? md.substring(0, md.indexOf('---', 3) + 3) : '';
    const newOrder = [...slides];
    [newOrder[index], newOrder[target]] = [newOrder[target], newOrder[index]];
    const rebuilt = frontmatter + '\\n\\n' + newOrder.map(s => s.text.trim()).join('\\n\\n---\\n\\n') + '\\n';
    editor.value = rebuilt;
    updateStats(); autoSave(); triggerAutoPreview();
}

// Override refreshPreview to decorate each slide thumb with buttons
const _origRefreshPreview = refreshPreview;
refreshPreview = async function() {
    await _origRefreshPreview();
    // Add buttons if not present
    const panel = document.getElementById('preview-content');
    const thumbs = panel.querySelectorAll('.slide-thumb');
    thumbs.forEach((thumb, i) => {
        const caption = thumb.querySelector('.caption');
        if (!caption || caption.querySelector('.slide-actions')) return;
        const actions = document.createElement('div');
        actions.className = 'slide-actions';
        actions.innerHTML = `
            <button class="slide-action-btn" onclick="editSlide(${i})" title="このスライドを編集">✎</button>
            <button class="slide-action-btn" onclick="jumpToSlide(${i})" title="MDのこのスライドへジャンプ">↵</button>
            <button class="slide-action-btn" onclick="moveSlide(${i}, -1)" title="上に移動">↑</button>
            <button class="slide-action-btn" onclick="moveSlide(${i}, 1)" title="下に移動">↓</button>
            <button class="slide-action-btn delete" onclick="deleteSlide(${i})" title="削除">×</button>
        `;
        caption.appendChild(actions);
        const img = thumb.querySelector('img');
        if (img) { img.style.cursor = 'pointer'; img.onclick = () => jumpToSlide(i); }
    });
};

// ── Keyboard shortcuts ──
document.addEventListener('keydown', (e) => {
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key === 's') {
        e.preventDefault();
        downloadMd();
    } else if (mod && e.key === 'p') {
        e.preventDefault();
        refreshPreview();
    } else if (mod && e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        generate();
    } else if (e.key === 'Escape') {
        document.querySelectorAll('.modal-bg.open').forEach(m => m.classList.remove('open'));
    } else if (mod && e.key === 'k') {
        e.preventDefault();
        openTypePicker();
    }
});

// Close modal on backdrop click
document.querySelectorAll('.modal-bg').forEach(bg => {
    bg.addEventListener('click', (e) => {
        if (e.target === bg) bg.classList.remove('open');
    });
});

// Show shortcuts hint in topbar
const shortcutHint = document.createElement('span');
shortcutHint.style.cssText = 'font-size:0.75em; color:#888; margin-left:auto';
shortcutHint.innerHTML = '⌘S: MD保存 · ⌘P: プレビュー · ⌘K: 型追加 · ⌘↵: PPTX';
document.querySelector('.topbar .spacer')?.replaceWith(shortcutHint);

// Initialize
loadTypeMeta();
restoreFromStorage();
updateStats();
if (editor.value.trim()) triggerAutoPreview();
</script>
</body>
</html>"""


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

<div class="card" style="background:#1a1a1a; color:white;">
<h2 style="margin-bottom:16px">✏️ ブラウザで直接編集 → PPTX 生成</h2>
<p style="margin-bottom:16px; color:#ccc; font-size:0.9em">
.md ファイルを用意せず、その場で Markdown を書いて PPTX にします。型の挿入ボタンあり。
</p>
<a href="/editor"><button type="button" style="background:white; color:#1a1a1a; cursor:pointer">→ エディタを開く</button></a>
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

    @app.route("/editor")
    def editor():
        return render_template_string(EDITOR_HTML, palettes=_palettes())

    @app.route("/editor/sample/<name>")
    def editor_sample(name: str):
        """Return a sample Markdown document to populate the editor."""
        templates_dir = Path(__file__).parent.parent / "data" / "templates"
        if name == "minimal":
            md = """---
marp: true
theme: academic
---

<!-- _class: title -->
# 発表タイトル
## サブタイトル
発表者名 / 2026年

---

# 概要
- 背景
- 手法
- 結果
- 考察

---

<!-- _class: end -->
# Thank You
"""
        elif name == "all":
            # Concatenate all 49 templates
            parts = ["---\nmarp: true\ntheme: academic\n---\n"]
            for tpl in sorted(templates_dir.glob("*.md")):
                text = tpl.read_text(encoding="utf-8")
                if text.startswith("---"):
                    end = text.find("---", 3)
                    if end != -1:
                        text = text[end + 3:]
                parts.append(text.strip())
            md = "\n\n---\n\n".join(parts)
        elif name == "academic":
            md = """---
marp: true
theme: academic
---

<!-- _class: title -->
# 研究タイトル
## サブタイトル
山田 太郎 / 2026年4月

---

<!-- _class: agenda -->
# 本日の内容
<div class="agenda-list">
1. 背景と研究目的
2. 提案手法
3. 実験結果
4. 考察とまとめ
</div>

---

<!-- _class: rq -->
# 研究課題
<div class="rq-main">既存手法は大規模データに対してスケールするか？</div>
<div class="rq-sub">計算量 $O(n^2)$ がボトルネックとなっている。</div>

---

<!-- _class: sandwich -->
# 提案手法
<div class="top">
<div class="lead">従来の $O(n^2)$ を $O(n \\log n)$ に改善。</div>
</div>
<div class="columns">
<div>

### 従来
- 計算量: `O(n^2)`
- メモリ: 多い
</div>
<div>

### 提案
- 計算量: `O(n \\log n)`
- メモリ: 少ない
</div>
</div>
<div class="bottom">
<div class="conclusion"><strong>結論:</strong> 大規模データでも実用的な速度を実現。</div>
</div>

---

<!-- _class: kpi -->
# 実験結果
<div class="kpi-container">
<div><span class="kpi-value">97%</span><span class="kpi-label">精度</span></div>
<div><span class="kpi-value">10x</span><span class="kpi-label">高速化</span></div>
<div><span class="kpi-value">50%</span><span class="kpi-label">省メモリ</span></div>
</div>

---

<!-- _class: takeaway -->
# Takeaway
<div class="ta-main">型を選ぶだけで、伝わるプレゼンに</div>
<div class="ta-points">
<ul>
<li>計算量の改善により大規模データに対応</li>
<li>精度は従来と同等</li>
<li>OSS として公開予定</li>
</ul>
</div>

---

<!-- _class: end -->
# Thank You
"""
        else:
            return "unknown sample", 404
        from flask import Response
        return Response(md, mimetype="text/plain; charset=utf-8")

    @app.route("/editor/preview", methods=["POST"])
    def editor_preview():
        """Render MD → PPTX → per-slide PNGs; return list of URLs."""
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        md_text = request.form.get("markdown", "")
        if not md_text.strip():
            return jsonify({"slides": []})
        palette_name = request.form.get("palette", "")
        try:
            font_scale = float(request.form.get("font_scale", 1.0))
        except ValueError:
            font_scale = 1.0

        # Cache key based on content + settings
        key_src = f"{md_text}|{palette_name}|{font_scale}".encode("utf-8")
        key = hashlib.md5(key_src).hexdigest()
        out_dir = _PREVIEW_CACHE_DIR / key
        if out_dir.exists():
            pngs = sorted(out_dir.glob("slide-*.png"))
            if pngs:
                return jsonify({"slides": [f"/editor/preview-img/{key}/{p.name}" for p in pngs]})
        out_dir.mkdir(parents=True, exist_ok=True)

        # Build PPTX (expose uploaded images via assets symlink)
        _link_shared_assets_to(out_dir)
        md_path = out_dir / "slides.md"
        md_path.write_text(md_text, encoding="utf-8")
        tc = ThemeConfig.from_css(get_default_theme_path())
        tc.font_scale = max(0.5, min(2.0, font_scale))
        if palette_name:
            pp = get_palette_path(palette_name)
            if pp:
                tc.apply_palette(pp)
        slides = parse_marp(str(md_path))
        builder = PptxBuilder(base_path=out_dir, theme=tc)
        builder.build_all(slides)
        pptx_path = out_dir / "slides.pptx"
        builder.save(str(pptx_path))

        # Render to PNGs
        pngs = _render_pptx_to_pngs(pptx_path, out_dir, dpi=90)
        if not pngs:
            return jsonify({"error": "LibreOffice not available or render failed", "slides": []}), 500
        return jsonify({"slides": [f"/editor/preview-img/{key}/{p.name}" for p in pngs]})

    @app.route("/editor/preview-img/<key>/<name>")
    def editor_preview_img(key: str, name: str):
        """Serve a cached preview PNG."""
        if not key.isalnum() or not name.startswith("slide-") or not name.endswith(".png"):
            return "bad path", 400
        png = _PREVIEW_CACHE_DIR / key / name
        if not png.exists():
            return "not found", 404
        return send_file(str(png), mimetype="image/png")

    @app.route("/editor/pptx-to-md", methods=["POST"])
    def editor_pptx_to_md():
        """Upload a PPTX, extract text+structure, return best-effort MD."""
        from marp_pptx.pptx2md import pptx_to_md_with_report

        f = request.files.get("file")
        if not f:
            return jsonify({"error": "no file"}), 400
        tmpdir = Path(tempfile.mkdtemp(prefix="marp_pptx2md_"))
        pptx_path = tmpdir / (f.filename or "input.pptx")
        f.save(str(pptx_path))

        # Extract images into the shared assets dir so the editor can reference them
        assets_dir = _PREVIEW_CACHE_DIR / "shared_assets"
        try:
            report = pptx_to_md_with_report(pptx_path, extract_images_to=assets_dir)
        except Exception as e:
            return jsonify({"error": f"pptx parse failed: {e}"}), 500
        return jsonify(report)

    @app.route("/editor/upload-image", methods=["POST"])
    def editor_upload_image():
        """Receive an image upload, store in shared assets dir, return assets/<name>."""
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "no file"}), 400
        name = (f.filename or "image").lower()
        ext = name.rsplit(".", 1)[-1] if "." in name else ""
        if ext not in ("png", "jpg", "jpeg", "gif", "svg", "webp"):
            return jsonify({"error": "unsupported extension: " + ext}), 400

        upload_dir = _PREVIEW_CACHE_DIR / "shared_assets"
        upload_dir.mkdir(parents=True, exist_ok=True)

        data = f.read()
        digest = hashlib.md5(data).hexdigest()[:10]
        safe_name = f"{digest}_{Path(name).stem[:40]}.{ext}"
        dest = upload_dir / safe_name
        dest.write_bytes(data)
        return jsonify({
            "path": f"assets/{safe_name}",
            "url": f"/editor/asset/{safe_name}",
        })

    @app.route("/editor/asset/<name>")
    def editor_asset_get(name: str):
        if "/" in name or ".." in name:
            return "bad path", 400
        p = _PREVIEW_CACHE_DIR / "shared_assets" / name
        if not p.exists():
            return "not found", 404
        return send_file(str(p))

    def _link_shared_assets_to(out_dir: Path):
        """Symlink uploaded assets into out_dir/assets/ so the builder can resolve
        'assets/foo.png' paths during PPTX generation.
        """
        src = _PREVIEW_CACHE_DIR / "shared_assets"
        if not src.exists():
            return
        dst = out_dir / "assets"
        if dst.exists():
            return
        try:
            dst.symlink_to(src, target_is_directory=True)
        except (OSError, NotImplementedError):
            # fallback: copy (Windows etc.)
            import shutil as _sh
            _sh.copytree(src, dst)

    @app.route("/editor/generate", methods=["POST"])
    def editor_generate():
        """Generate PPTX from raw Markdown text (no file upload)."""
        md_text = request.form.get("markdown", "")
        if not md_text.strip():
            return "empty markdown", 400
        palette_name = request.form.get("palette", "")
        try:
            font_scale = float(request.form.get("font_scale", 1.0))
        except ValueError:
            font_scale = 1.0
        output_name = request.form.get("output_name") or "slides.pptx"

        tmpdir = Path(tempfile.mkdtemp(prefix="marp_editor_"))
        md_path = tmpdir / "slides.md"
        md_path.write_text(md_text, encoding="utf-8")
        return _do_convert(
            md_path=md_path,
            palette_name=palette_name,
            font_scale=font_scale,
            output_name=output_name,
        )

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

        # Expose uploaded images to the builder's base_path
        _link_shared_assets_to(md_path.parent)

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
