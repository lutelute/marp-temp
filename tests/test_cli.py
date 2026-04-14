"""Tests for the CLI."""
from click.testing import CliRunner
from marp_pptx.cli import main


def test_cli_version():
    runner = CliRunner()
    result = runner.invoke(main, ["--version"])
    assert result.exit_code == 0
    assert "0.1.0" in result.output


def test_cli_help():
    runner = CliRunner()
    result = runner.invoke(main, ["--help"])
    assert result.exit_code == 0
    assert "marp-pptx" in result.output


def test_cli_types():
    runner = CliRunner()
    result = runner.invoke(main, ["types"])
    assert result.exit_code == 0
    assert "funnel" in result.output
    assert "Total:" in result.output


def test_cli_types_json():
    runner = CliRunner()
    result = runner.invoke(main, ["types", "--json"])
    assert result.exit_code == 0
    import json
    data = json.loads(result.output)
    assert len(data) > 40


def test_cli_types_category():
    runner = CliRunner()
    result = runner.invoke(main, ["types", "-c", "temporal"])
    assert result.exit_code == 0
    assert "timeline" in result.output


def test_cli_convert(example_md):
    if not example_md.exists():
        return
    import tempfile
    runner = CliRunner()
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        result = runner.invoke(main, ["convert", str(example_md), "-o", f.name])
        assert result.exit_code == 0
        from pathlib import Path
        assert Path(f.name).stat().st_size > 0
        Path(f.name).unlink()
