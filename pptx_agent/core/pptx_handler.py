"""
PowerPoint Handler - Core utilities for working with PowerPoint presentations.
"""

from pathlib import Path
from typing import Optional, Dict, Any, List
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


class PPTXHandler:
    """Handles PowerPoint presentation creation and manipulation."""

    def __init__(self, template_path: Optional[Path] = None):
        """
        Initialize the PPTX handler.

        Args:
            template_path: Optional path to a template .pptx file
        """
        if template_path and template_path.exists():
            self.prs = Presentation(str(template_path))
            self.template_path = template_path
        else:
            self.prs = Presentation()
            self.template_path = None

        self.slide_width = self.prs.slide_width
        self.slide_height = self.prs.slide_height

    def get_layout_by_name(self, name: str):
        """Get a slide layout by name."""
        for layout in self.prs.slide_layouts:
            if layout.name.lower() == name.lower():
                return layout
        return None

    def get_layout_by_index(self, index: int):
        """Get a slide layout by index."""
        try:
            return self.prs.slide_layouts[index]
        except IndexError:
            return self.prs.slide_layouts[0]

    def add_slide(self, layout_index: int = 0):
        """
        Add a new slide with the specified layout.

        Args:
            layout_index: Index of the layout to use (default: 0 - title slide)

        Returns:
            The newly created slide
        """
        layout = self.get_layout_by_index(layout_index)
        return self.prs.slides.add_slide(layout)

    def add_title_slide(self, title: str, subtitle: str = ""):
        """
        Add a title slide.

        Args:
            title: Main title text
            subtitle: Subtitle text

        Returns:
            The newly created slide
        """
        slide = self.add_slide(0)

        if slide.shapes.title:
            slide.shapes.title.text = title

        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = subtitle

        return slide

    def add_text_box(self, slide, text: str, left: float, top: float,
                     width: float, height: float, font_size: int = 18,
                     bold: bool = False, color: Optional[tuple] = None):
        """
        Add a text box to a slide.

        Args:
            slide: The slide to add the text box to
            text: Text content
            left, top: Position in inches
            width, height: Size in inches
            font_size: Font size in points
            bold: Whether text should be bold
            color: RGB color tuple (r, g, b)

        Returns:
            The text box shape
        """
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )

        text_frame = textbox.text_frame
        text_frame.text = text

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(font_size)
                run.font.bold = bold
                if color:
                    run.font.color.rgb = RGBColor(*color)

        return textbox

    def add_bullet_points(self, slide, points: List[str], left: float = 1,
                         top: float = 2, width: float = 8, height: float = 4,
                         font_size: int = 18):
        """
        Add bullet points to a slide.

        Args:
            slide: The slide to add bullets to
            points: List of bullet point texts
            left, top: Position in inches
            width, height: Size in inches
            font_size: Font size in points

        Returns:
            The text box shape with bullets
        """
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )

        text_frame = textbox.text_frame
        text_frame.word_wrap = True

        for i, point in enumerate(points):
            if i > 0:
                p = text_frame.add_paragraph()
            else:
                p = text_frame.paragraphs[0]

            p.text = point
            p.level = 0
            p.font.size = Pt(font_size)

        return textbox

    def format_shape(self, shape, fill_color: Optional[tuple] = None,
                    line_color: Optional[tuple] = None, line_width: float = 1):
        """
        Format a shape with colors and line properties.

        Args:
            shape: The shape to format
            fill_color: RGB tuple for fill color
            line_color: RGB tuple for line color
            line_width: Line width in points
        """
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(*fill_color)

        if line_color:
            shape.line.color.rgb = RGBColor(*line_color)
            shape.line.width = Pt(line_width)

    def save(self, output_path: Path):
        """
        Save the presentation to a file.

        Args:
            output_path: Path where the presentation should be saved
        """
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(output_path))

    def get_slide_count(self) -> int:
        """Get the total number of slides in the presentation."""
        return len(self.prs.slides)

    def delete_slide(self, index: int):
        """Delete a slide by index."""
        if 0 <= index < len(self.prs.slides):
            rId = self.prs.slides._sldIdLst[index].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[index]

    def list_layouts(self) -> List[str]:
        """List all available slide layouts."""
        return [layout.name for layout in self.prs.slide_layouts]
