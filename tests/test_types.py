"""Tests for the type registry."""
from marp_pptx.types import TYPE_REGISTRY, CATEGORIES, get_type_info, SlideTypeInfo


def test_registry_not_empty():
    assert len(TYPE_REGISTRY) > 40


def test_all_types_have_fields():
    for t in TYPE_REGISTRY:
        assert t.name
        assert t.css_class
        assert t.category in CATEGORIES
        assert t.geometry
        assert t.meaning
        assert t.use_when
        assert t.template_file


def test_get_type_info():
    t = get_type_info("funnel")
    assert t is not None
    assert t.name == "funnel"
    assert t.category == "convergence"


def test_get_type_info_missing():
    assert get_type_info("nonexistent") is None


def test_categories_cover_all_types():
    for t in TYPE_REGISTRY:
        assert t.category in CATEGORIES


def test_no_duplicate_css_classes():
    classes = [t.css_class for t in TYPE_REGISTRY]
    assert len(classes) == len(set(classes))
