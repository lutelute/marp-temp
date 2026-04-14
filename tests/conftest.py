"""Shared fixtures for tests."""
from pathlib import Path

import pytest


@pytest.fixture
def project_root():
    return Path(__file__).parent.parent


@pytest.fixture
def example_md(project_root):
    return project_root / "example.md"


@pytest.fixture
def theme_css():
    from marp_pptx.theme import get_default_theme_path
    return get_default_theme_path()
