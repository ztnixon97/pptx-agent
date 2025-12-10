"""
Tests for PPTXHandler.
"""

import pytest
from pathlib import Path
from pptx_agent.core.pptx_handler import PPTXHandler


def test_pptx_handler_initialization():
    """Test PPTXHandler can be initialized without template."""
    handler = PPTXHandler()
    assert handler.prs is not None
    assert handler.template_path is None


def test_add_title_slide():
    """Test adding a title slide."""
    handler = PPTXHandler()
    slide = handler.add_title_slide("Test Title", "Test Subtitle")
    assert slide is not None
    assert handler.get_slide_count() == 1


def test_add_text_box():
    """Test adding a text box to a slide."""
    handler = PPTXHandler()
    slide = handler.add_slide(0)
    textbox = handler.add_text_box(
        slide, "Test content",
        left=1, top=2, width=5, height=3
    )
    assert textbox is not None


def test_add_bullet_points():
    """Test adding bullet points."""
    handler = PPTXHandler()
    slide = handler.add_slide(1)
    points = ["Point 1", "Point 2", "Point 3"]
    textbox = handler.add_bullet_points(slide, points)
    assert textbox is not None


def test_slide_count():
    """Test slide counting."""
    handler = PPTXHandler()
    assert handler.get_slide_count() == 0

    handler.add_slide(0)
    assert handler.get_slide_count() == 1

    handler.add_slide(1)
    assert handler.get_slide_count() == 2


def test_list_layouts():
    """Test listing available layouts."""
    handler = PPTXHandler()
    layouts = handler.list_layouts()
    assert isinstance(layouts, list)
    assert len(layouts) > 0


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
