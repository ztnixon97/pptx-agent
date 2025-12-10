"""
Presentation Builder - Orchestrates the creation of presentations.
"""

from pathlib import Path
from typing import Optional, Dict, Any, List
from ..core.pptx_handler import PPTXHandler
from ..core.template_manager import TemplateManager
from ..llm.openai_client import OpenAIClient
from ..llm.content_planner import ContentPlanner
from ..builders.text_builder import TextSlideBuilder
from ..builders.table_builder import TableSlideBuilder
from ..builders.chart_builder import ChartSlideBuilder
from ..builders.image_builder import ImageSlideBuilder
from ..builders.smartart_builder import SmartArtBuilder
from ..builders.shapes_builder import ShapesBuilder


class PresentationBuilder:
    """Main class for building PowerPoint presentations with LLM assistance."""

    def __init__(self, template_path: Optional[Path] = None,
                 openai_api_key: Optional[str] = None):
        """
        Initialize the presentation builder.

        Args:
            template_path: Optional path to a PowerPoint template
            openai_api_key: Optional OpenAI API key
        """
        self.handler = PPTXHandler(template_path)
        self.template_manager = TemplateManager(template_path)
        self.openai_client = OpenAIClient(api_key=openai_api_key)
        self.content_planner = ContentPlanner(self.openai_client)

        self.text_builder = TextSlideBuilder()
        self.table_builder = TableSlideBuilder()
        self.chart_builder = ChartSlideBuilder()
        self.image_builder = ImageSlideBuilder()
        self.smartart_builder = SmartArtBuilder()
        self.shapes_builder = ShapesBuilder()

        self.current_outline: Optional[Dict[str, Any]] = None

    def create_outline(self, topic: str, summary: str,
                      num_slides: Optional[int] = None,
                      reference_docs: Optional[str] = None) -> Dict[str, Any]:
        """
        Create a presentation outline.

        Args:
            topic: Main topic of the presentation
            summary: Content summary or key points
            num_slides: Suggested number of slides
            reference_docs: Optional reference documentation

        Returns:
            Presentation outline dictionary
        """
        self.current_outline = self.content_planner.create_presentation_outline(
            topic, summary, num_slides, reference_docs
        )
        return self.current_outline

    def build_from_outline(self, outline: Optional[Dict[str, Any]] = None,
                          image_dir: Optional[Path] = None) -> None:
        """
        Build a presentation from an outline.

        Args:
            outline: Presentation outline (uses current_outline if None)
            image_dir: Directory containing images for the presentation
        """
        if outline is None:
            outline = self.current_outline

        if not outline:
            raise ValueError("No outline provided")

        # Add title slide
        title = outline.get('title', 'Untitled Presentation')
        self.handler.add_title_slide(title, "")

        # Process each slide
        for slide_spec in outline.get('slides', []):
            self._build_slide(slide_spec, image_dir)

    def _build_slide(self, slide_spec: Dict[str, Any],
                     image_dir: Optional[Path] = None) -> None:
        """Build a single slide from specification."""
        slide_type = slide_spec.get('slide_type', 'content')
        title = slide_spec.get('title', '')
        content = slide_spec.get('content', '')
        elements = slide_spec.get('elements', [])

        # Determine layout
        layout_idx = self.template_manager.suggest_layout(slide_type) or 1

        if slide_type == 'title':
            self.handler.add_title_slide(title, content)

        elif slide_type == 'section':
            self.text_builder.add_section_slide(self.handler, title, content, layout_idx)

        elif slide_type == 'blank':
            self.text_builder.add_blank_slide(self.handler, layout_idx)

        else:
            # Process elements
            if not elements:
                # Default to simple content slide
                self.text_builder.add_content_slide(self.handler, title, content, layout_idx)
            else:
                self._build_slide_with_elements(title, elements, layout_idx, image_dir)

    def _build_slide_with_elements(self, title: str, elements: List[Dict[str, Any]],
                                   layout_idx: int, image_dir: Optional[Path] = None):
        """Build a slide with specific elements."""
        # Identify primary element type
        primary_type = elements[0].get('type', 'text') if elements else 'text'

        if primary_type == 'bullet_points':
            data = elements[0].get('data', [])
            if isinstance(data, str):
                data = [line.strip() for line in data.split('\n') if line.strip()]
            self.text_builder.add_bullet_slide(self.handler, title, data, layout_idx)

        elif primary_type == 'table':
            table_data = elements[0].get('details', {})
            headers = table_data.get('headers', [])
            rows = table_data.get('rows', [])

            if headers and rows:
                self.table_builder.add_table_slide(
                    self.handler, title, headers, rows, layout_idx
                )
            else:
                # Generate table data
                description = elements[0].get('data', 'Sample table')
                table_spec = self.content_planner.generate_table_data(description)
                self.table_builder.add_table_slide(
                    self.handler, title,
                    table_spec['headers'], table_spec['rows'],
                    layout_idx
                )

        elif primary_type == 'chart':
            chart_data = elements[0].get('details', {})
            chart_type = chart_data.get('chart_type', 'bar')
            categories = chart_data.get('categories', [])
            series = chart_data.get('series', [])

            if not categories or not series:
                # Generate chart data
                description = elements[0].get('data', 'Sample chart')
                chart_spec = self.content_planner.generate_chart_data(description, chart_type)
                categories = chart_spec['categories']
                series = chart_spec['series']

            if chart_type == 'bar':
                self.chart_builder.add_bar_chart_slide(
                    self.handler, title, categories, series, layout_idx
                )
            elif chart_type == 'line':
                self.chart_builder.add_line_chart_slide(
                    self.handler, title, categories, series, layout_idx
                )
            elif chart_type == 'pie':
                values = series[0]['values'] if series else []
                self.chart_builder.add_pie_chart_slide(
                    self.handler, title, categories, values, layout_idx
                )

        elif primary_type == 'image' and image_dir:
            image_name = elements[0].get('data', '')
            image_path = image_dir / image_name

            if image_path.exists():
                caption = elements[0].get('details', {}).get('caption', '')
                self.image_builder.add_image_slide(
                    self.handler, title, image_path, layout_idx,
                    caption=caption
                )
            else:
                # Fallback to text slide
                self.text_builder.add_content_slide(
                    self.handler, title,
                    f"Image not found: {image_name}",
                    layout_idx
                )

        elif primary_type == 'process_flow':
            details = elements[0].get('details', {})
            steps = details.get('items', [])
            if isinstance(steps, str):
                steps = [s.strip() for s in steps.split(',') if s.strip()]
            self.smartart_builder.add_process_flow(self.handler, title, steps, layout_idx)

        elif primary_type == 'cycle':
            details = elements[0].get('details', {})
            items = details.get('items', [])
            if isinstance(items, str):
                items = [i.strip() for i in items.split(',') if i.strip()]
            self.smartart_builder.add_cycle_diagram(self.handler, title, items, layout_idx)

        elif primary_type == 'hierarchy':
            details = elements[0].get('details', {})
            root = details.get('root', 'Root')
            children = details.get('children', [])
            if isinstance(children, str):
                children = [c.strip() for c in children.split(',') if c.strip()]
            self.smartart_builder.add_hierarchy_diagram(self.handler, title, root, children, layout_idx)

        elif primary_type == 'comparison':
            details = elements[0].get('details', {})
            left_items = details.get('left_items', [])
            right_items = details.get('right_items', [])
            left_label = details.get('left_label', 'Option A')
            right_label = details.get('right_label', 'Option B')
            self.smartart_builder.add_comparison_diagram(
                self.handler, title, left_items, right_items, left_label, right_label, layout_idx
            )

        elif primary_type == 'venn':
            details = elements[0].get('details', {})
            left_label = details.get('left_label', 'A')
            right_label = details.get('right_label', 'B')
            left_items = details.get('left_items', [])
            right_items = details.get('right_items', [])
            overlap_items = details.get('overlap_items', [])
            self.smartart_builder.add_venn_diagram(
                self.handler, title, left_label, right_label,
                left_items, right_items, overlap_items, layout_idx
            )

        elif primary_type == 'timeline':
            details = elements[0].get('details', {})
            events = details.get('events', [])
            self.smartart_builder.add_timeline(self.handler, title, events, layout_idx)

        elif primary_type == 'flowchart':
            details = elements[0].get('details', {})
            steps = details.get('steps', [])
            self.shapes_builder.add_flowchart_slide(self.handler, title, steps, layout_idx)

        else:
            # Default to text content
            content = elements[0].get('data', '') if elements else ''
            self.text_builder.add_content_slide(self.handler, title, content, layout_idx)

    def add_text_slide(self, title: str, content: str,
                      slide_type: str = 'content') -> None:
        """Manually add a text slide."""
        layout_idx = self.template_manager.suggest_layout(slide_type) or 1

        if slide_type == 'section':
            self.text_builder.add_section_slide(self.handler, title, content, layout_idx)
        else:
            self.text_builder.add_content_slide(self.handler, title, content, layout_idx)

    def add_bullet_slide(self, title: str, points: List[str]) -> None:
        """Manually add a bullet point slide."""
        layout_idx = self.template_manager.suggest_layout('content') or 1
        self.text_builder.add_bullet_slide(self.handler, title, points, layout_idx)

    def add_table_slide(self, title: str, headers: List[str],
                       rows: List[List[str]]) -> None:
        """Manually add a table slide."""
        layout_idx = self.template_manager.suggest_layout('content') or 1
        self.table_builder.add_table_slide(
            self.handler, title, headers, rows, layout_idx
        )

    def add_chart_slide(self, title: str, chart_type: str,
                       categories: List[str], series: List[Dict[str, Any]]) -> None:
        """Manually add a chart slide."""
        layout_idx = self.template_manager.suggest_layout('content') or 1

        if chart_type == 'bar':
            self.chart_builder.add_bar_chart_slide(
                self.handler, title, categories, series, layout_idx
            )
        elif chart_type == 'line':
            self.chart_builder.add_line_chart_slide(
                self.handler, title, categories, series, layout_idx
            )
        elif chart_type == 'pie':
            values = series[0]['values'] if series else []
            self.chart_builder.add_pie_chart_slide(
                self.handler, title, categories, values, layout_idx
            )

    def add_image_slide(self, title: str, image_path: Path,
                       caption: str = "") -> None:
        """Manually add an image slide."""
        layout_idx = self.template_manager.suggest_layout('content') or 1
        self.image_builder.add_image_slide(
            self.handler, title, image_path, layout_idx,
            caption=caption
        )

    def refine_slide(self, slide_index: int, feedback: str) -> None:
        """
        Refine a specific slide based on feedback.

        Args:
            slide_index: Index of the slide to refine
            feedback: User feedback for refinement
        """
        # This would require more complex slide extraction and rebuilding
        # For now, we'll note it as a future enhancement
        pass

    def save(self, output_path: Path) -> None:
        """
        Save the presentation.

        Args:
            output_path: Path where the presentation should be saved
        """
        self.handler.save(output_path)

    def get_slide_count(self) -> int:
        """Get the current number of slides."""
        return self.handler.get_slide_count()

    def get_template_info(self) -> str:
        """Get information about the current template."""
        return self.template_manager.get_template_summary()

    def suggest_improvements(self) -> List[str]:
        """Get AI-powered suggestions for improving the presentation."""
        if not self.current_outline:
            return ["Create an outline first"]

        return self.content_planner.suggest_improvements(self.current_outline)
