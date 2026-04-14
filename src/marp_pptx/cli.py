"""CLI entry point for marp-pptx."""

from __future__ import annotations

import sys
from pathlib import Path

import click

from marp_pptx import __version__


@click.group(invoke_without_command=True)
@click.version_option(__version__)
@click.pass_context
def main(ctx):
    """marp-pptx: Convert Marp markdown to editable PowerPoint with 49 semantic slide types."""
    if ctx.invoked_subcommand is None:
        click.echo(ctx.get_help())


@main.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("-o", "--output", help="Output .pptx path")
@click.option("-p", "--palette", help="Palette name (e.g. navy, copper, earth)")
@click.option("-t", "--theme", help="Custom palette CSS path")
def convert(input_file: str, output: str | None, palette: str | None, theme: str | None):
    """Convert a Marp markdown file to editable PPTX."""
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder

    input_path = Path(input_file)

    # Load theme
    tc = ThemeConfig.from_css(get_default_theme_path())

    # Apply palette
    if palette:
        palette_path = get_palette_path(palette)
        if palette_path:
            tc.apply_palette(palette_path)
        else:
            click.echo(f"Warning: palette '{palette}' not found, using default", err=True)
    elif theme:
        tc.apply_palette(Path(theme))

    print(f"[theme] latin={tc.font}  ea={tc.font_ea}  head={tc.font_head}", file=sys.stderr)

    # Parse
    slides = parse_marp(str(input_path))
    click.echo(f"Parsed {len(slides)} slides", err=True)

    # Build
    builder = PptxBuilder(base_path=input_path.parent, theme=tc)
    builder.build_all(slides)

    # Save
    output_path = output or str(input_path.with_name(input_path.stem + "_editable.pptx"))
    builder.save(output_path)

    click.echo(f"Saved: {output_path}", err=True)
    click.echo(f"  {len(slides)} slides, all editable text boxes", err=True)


@main.command("types")
@click.option("--category", "-c", help="Filter by category")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
def list_types(category: str | None, as_json: bool):
    """List all available slide types with their semantic meanings."""
    from marp_pptx.types import TYPE_REGISTRY, CATEGORIES

    types = TYPE_REGISTRY
    if category:
        types = [t for t in types if t.category == category]

    if as_json:
        import json
        data = [
            {
                "name": t.name,
                "css_class": t.css_class,
                "category": t.category,
                "category_ja": CATEGORIES.get(t.category, t.category),
                "geometry": t.geometry,
                "meaning": t.meaning,
                "use_when": t.use_when,
                "template": t.template_file,
            }
            for t in types
        ]
        click.echo(json.dumps(data, ensure_ascii=False, indent=2))
        return

    # Table output
    click.echo(f"{'Type':<18} {'Category':<12} {'Geometry':<16} {'Meaning':<24} Use When")
    click.echo("-" * 100)

    current_cat = None
    for t in types:
        cat_ja = CATEGORIES.get(t.category, t.category)
        if t.category != current_cat:
            current_cat = t.category
            if t != types[0]:
                click.echo()
        click.echo(f"{t.name:<18} {cat_ja:<12} {t.geometry:<16} {t.meaning:<24} {t.use_when}")

    click.echo(f"\nTotal: {len(types)} types in {len(CATEGORIES)} categories")
    if not category:
        click.echo(f"Categories: {', '.join(f'{v} ({k})' for k, v in CATEGORIES.items())}")


@main.command()
@click.option("-o", "--output", default="type_catalog.pptx", help="Output catalog PPTX")
@click.option("-p", "--palette", help="Palette name")
def preview(output: str, palette: str | None):
    """Generate a visual catalog PPTX showing all 49 slide types."""
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder
    from marp_pptx.types import TYPE_REGISTRY

    tc = ThemeConfig.from_css(get_default_theme_path())
    if palette:
        palette_path = get_palette_path(palette)
        if palette_path:
            tc.apply_palette(palette_path)

    templates_dir = Path(__file__).parent / "data" / "templates"

    # Concatenate all template files
    all_md = "---\nmarp: true\ntheme: academic\n---\n"
    for t in TYPE_REGISTRY:
        template_path = templates_dir / t.template_file
        if template_path.exists():
            text = template_path.read_text(encoding="utf-8")
            # Strip frontmatter
            if text.startswith("---"):
                end = text.find("---", 3)
                if end != -1:
                    text = text[end + 3:]
            all_md += f"\n---\n{text.strip()}\n"

    # Write temp file and parse
    import tempfile
    with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as f:
        f.write(all_md)
        tmp_path = Path(f.name)

    slides = parse_marp(str(tmp_path))
    builder = PptxBuilder(base_path=templates_dir, theme=tc)
    builder.build_all(slides)
    builder.save(output)

    tmp_path.unlink()
    click.echo(f"Generated catalog: {output} ({len(slides)} slides)")


@main.command()
@click.option("--host", default="127.0.0.1")
@click.option("--port", default=8080, type=int)
def serve(host: str, port: int):
    """Start the web UI for shared/lab use (requires marp-pptx[web])."""
    try:
        from marp_pptx.web.app import create_app
    except ImportError:
        click.echo("Web UI requires Flask. Install with: pip install marp-pptx[web]", err=True)
        raise SystemExit(1)

    app = create_app()
    click.echo(f"Starting marp-pptx web UI on http://{host}:{port}")
    app.run(host=host, port=port)
