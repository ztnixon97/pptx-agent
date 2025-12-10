"""
Autonomous Presentation Builder - Fully autonomous presentation creation with AI.
"""

from pathlib import Path
from typing import Optional, Dict, Any, List
from ..core.pptx_handler import PPTXHandler
from ..core.template_manager import TemplateManager
from ..core.content_validator import ContentValidator
from ..core.layout_optimizer import LayoutOptimizer
from ..llm.openai_client import OpenAIClient
from ..llm.content_planner import ContentPlanner
from ..llm.autonomous_designer import AutonomousDesigner
from ..builders.text_builder import TextSlideBuilder
from ..builders.table_builder import TableSlideBuilder
from ..builders.chart_builder import ChartSlideBuilder
from ..builders.image_builder import ImageSlideBuilder


class AutonomousPresentationBuilder:
    """
    Fully autonomous presentation builder that makes all design decisions.

    This builder uses AI to:
    - Determine optimal layouts
    - Choose color schemes
    - Decide formatting and styling
    - Validate and optimize content fit
    - Automatically reformat when needed
    - Make all design decisions without user input
    """

    def __init__(self, template_path: Optional[Path] = None,
                 openai_api_key: Optional[str] = None,
                 target_audience: str = "professional"):
        """
        Initialize the autonomous builder.

        Args:
            template_path: Optional path to PowerPoint template
            openai_api_key: Optional OpenAI API key
            target_audience: Target audience for automatic adjustments
        """
        self.handler = PPTXHandler(template_path)
        self.template_manager = TemplateManager(template_path)
        self.content_validator = ContentValidator()
        self.layout_optimizer = LayoutOptimizer(
            self.template_manager.list_layouts() or []
        )

        self.openai_client = OpenAIClient(api_key=openai_api_key)
        self.content_planner = ContentPlanner(self.openai_client)
        self.autonomous_designer = AutonomousDesigner(self.openai_client)

        self.text_builder = TextSlideBuilder()
        self.table_builder = TableSlideBuilder()
        self.chart_builder = ChartSlideBuilder()
        self.image_builder = ImageSlideBuilder()

        self.target_audience = target_audience
        self.color_scheme: Optional[Dict[str, Any]] = None
        self.current_outline: Optional[Dict[str, Any]] = None

    def create_presentation_autonomously(self, topic: str, summary: str,
                                        num_slides: Optional[int] = None,
                                        reference_docs: Optional[str] = None,
                                        image_dir: Optional[Path] = None) -> Dict[str, Any]:
        """
        Create a complete presentation with full autonomy.

        The AI makes all decisions about:
        - Content structure and organization
        - Layout selection
        - Color schemes and styling
        - Font sizes and formatting
        - Content optimization and fitting

        Args:
            topic: Presentation topic
            summary: Content summary
            num_slides: Target number of slides
            reference_docs: Reference documentation
            image_dir: Directory with images

        Returns:
            Creation report with details
        """
        report = {
            'topic': topic,
            'decisions_made': [],
            'optimizations_performed': [],
            'slides_created': 0
        }

        # Step 1: Decide color scheme autonomously
        print("Making autonomous design decisions...")
        self.color_scheme = self.autonomous_designer.decide_color_scheme(
            topic, "professional"
        )
        report['decisions_made'].append({
            'decision': 'color_scheme',
            'choice': self.color_scheme['primary_color']['name'],
            'reasoning': self.color_scheme['reasoning']
        })

        # Step 2: Create outline
        print("Generating presentation outline...")
        self.current_outline = self.content_planner.create_presentation_outline(
            topic, summary, num_slides, reference_docs
        )

        # Step 3: Optimize outline structure
        print("Optimizing presentation structure...")
        flow_analysis = self.layout_optimizer.analyze_presentation_flow(
            self.current_outline.get('slides', [])
        )

        if flow_analysis['recommendations']:
            report['optimizations_performed'].extend(flow_analysis['recommendations'])

        # Step 4: Build slides with autonomous decisions
        print("Building slides with autonomous design...")
        slides_built = self._build_slides_autonomously(
            self.current_outline,
            image_dir,
            report
        )

        report['slides_created'] = slides_built

        return report

    def _build_slides_autonomously(self, outline: Dict[str, Any],
                                   image_dir: Optional[Path],
                                   report: Dict[str, Any]) -> int:
        """Build slides with full autonomous decision making."""
        slides_built = 0

        # Add title slide
        title = outline.get('title', 'Untitled Presentation')
        self.handler.add_title_slide(title, "")
        slides_built += 1

        # Process each slide
        for slide_spec in outline.get('slides', []):
            built = self._build_slide_autonomously(slide_spec, image_dir, report)
            slides_built += built

        return slides_built

    def _build_slide_autonomously(self, slide_spec: Dict[str, Any],
                                  image_dir: Optional[Path],
                                  report: Dict[str, Any]) -> int:
        """Build a single slide with autonomous decisions."""
        # Get autonomous layout decision
        available_layouts = [layout['name'] for layout in self.template_manager.list_layouts()]
        if not available_layouts:
            available_layouts = ['Title Slide', 'Title and Content', 'Section Header', 'Blank']

        layout_decision = self.autonomous_designer.analyze_content_and_decide_layout(
            slide_spec, available_layouts
        )

        # Select optimal layout
        layout_idx, layout_reasoning = self.layout_optimizer.select_optimal_layout(slide_spec)

        report['decisions_made'].append({
            'decision': 'layout',
            'slide': slide_spec.get('title', 'Untitled'),
            'choice': layout_decision['layout_choice'],
            'reasoning': layout_decision['reasoning']
        })

        # Check if content needs optimization
        content = slide_spec.get('content', '')
        elements = slide_spec.get('elements', [])

        if content:
            validation = self.content_validator.validate_text_content(content)

            if not validation['fits']:
                # Autonomously optimize
                optimization = self.autonomous_designer.optimize_content_formatting(
                    content,
                    slide_spec.get('slide_type', 'text'),
                    {
                        'width': self.content_validator.content_area_width,
                        'height': self.content_validator.content_area_height
                    }
                )

                if optimization['action'] == 'split':
                    # Split into multiple slides
                    report['optimizations_performed'].append(
                        f"Split '{slide_spec.get('title')}' into {optimization['split_into_slides']} slides"
                    )
                    return self._build_split_slides(slide_spec, optimization, layout_idx, report)

                elif optimization['action'] in ['shorten', 'reformat']:
                    content = optimization['optimized_content']
                    report['optimizations_performed'].append(
                        f"Optimized content for '{slide_spec.get('title')}': {optimization['reasoning']}"
                    )

        # Build the slide with optimized content
        slide_type = slide_spec.get('slide_type', 'content')
        title = slide_spec.get('title', '')

        if slide_type == 'title':
            self.handler.add_title_slide(title, content)
        elif slide_type == 'section':
            self.text_builder.add_section_slide(self.handler, title, content, layout_idx)
        elif not elements:
            self.text_builder.add_content_slide(self.handler, title, content, layout_idx)
        else:
            self._build_slide_with_elements_autonomously(
                title, elements, layout_idx, layout_decision, image_dir, report
            )

        return 1

    def _build_slide_with_elements_autonomously(self, title: str,
                                               elements: List[Dict[str, Any]],
                                               layout_idx: int,
                                               layout_decision: Dict[str, Any],
                                               image_dir: Optional[Path],
                                               report: Dict[str, Any]):
        """Build slide with elements using autonomous decisions."""
        # Decide visual hierarchy
        hierarchy = self.autonomous_designer.decide_visual_hierarchy(elements)

        report['decisions_made'].append({
            'decision': 'visual_hierarchy',
            'slide': title,
            'reasoning': hierarchy['reasoning']
        })

        # Reorder elements based on hierarchy
        ordered_elements = [elements[i] for i in hierarchy['element_order']]

        # Build based on primary element type
        primary_type = ordered_elements[0].get('type', 'text')

        if primary_type == 'bullet_points':
            data = ordered_elements[0].get('data', [])
            if isinstance(data, str):
                data = [line.strip() for line in data.split('\n') if line.strip()]

            # Validate bullets
            validation = self.content_validator.validate_bullet_points(data)

            if not validation['fits']:
                # Auto-optimize bullet points
                max_points = validation['max_recommended']
                data = data[:max_points]  # Trim to recommended max
                report['optimizations_performed'].append(
                    f"Trimmed bullet points in '{title}' to {max_points} items"
                )

            # Apply autonomous styling
            font_size = hierarchy['sizing']['0'].get('font_size', 18)
            self.text_builder.add_bullet_slide(self.handler, title, data, layout_idx, font_size)

        elif primary_type == 'table':
            table_data = ordered_elements[0].get('details', {})
            headers = table_data.get('headers', [])
            rows = table_data.get('rows', [])

            if not headers or not rows:
                # Generate table data autonomously
                description = ordered_elements[0].get('data', 'Sample table')
                table_spec = self.content_planner.generate_table_data(description)
                headers = table_spec['headers']
                rows = table_spec['rows']

            # Validate table
            validation = self.content_validator.validate_table_size(
                len(rows), len(headers)
            )

            if not validation['fits']:
                # Trim or split table
                max_rows = validation['max_recommended_rows']
                rows = rows[:max_rows]
                report['optimizations_performed'].append(
                    f"Trimmed table in '{title}' to {max_rows} rows"
                )

            # Apply color scheme
            if self.color_scheme:
                primary = self.color_scheme['primary_color']
                self.table_builder.add_styled_table_slide(
                    self.handler, title, headers, rows,
                    header_color=(primary['r'], primary['g'], primary['b']),
                    layout_index=layout_idx
                )
            else:
                self.table_builder.add_table_slide(
                    self.handler, title, headers, rows, layout_idx
                )

        elif primary_type == 'chart':
            chart_data = ordered_elements[0].get('details', {})

            if not chart_data.get('categories'):
                # Autonomously generate chart data
                description = ordered_elements[0].get('data', 'Sample chart')
                chart_decision = self.autonomous_designer.decide_chart_visualization(
                    description,
                    {'num_series': 1, 'data_type': 'numeric'}
                )

                chart_spec = self.content_planner.generate_chart_data(
                    description,
                    chart_decision['chart_type']
                )

                chart_type = chart_decision['chart_type']
                categories = chart_spec['categories']
                series = chart_spec['series']

                report['decisions_made'].append({
                    'decision': 'chart_type',
                    'slide': title,
                    'choice': chart_type,
                    'reasoning': chart_decision['reasoning']
                })
            else:
                chart_type = chart_data.get('chart_type', 'bar')
                categories = chart_data['categories']
                series = chart_data['series']

            # Build chart
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
            image_name = ordered_elements[0].get('data', '')
            image_path = image_dir / image_name

            if image_path.exists():
                caption = ordered_elements[0].get('details', {}).get('caption', '')
                self.image_builder.add_image_slide(
                    self.handler, title, image_path, layout_idx,
                    caption=caption
                )
            else:
                # Fallback to text
                self.text_builder.add_content_slide(
                    self.handler, title,
                    f"Image not found: {image_name}",
                    layout_idx
                )

        else:
            # Default to text content
            content = ordered_elements[0].get('data', '')
            self.text_builder.add_content_slide(self.handler, title, content, layout_idx)

    def _build_split_slides(self, slide_spec: Dict[str, Any],
                           optimization: Dict[str, Any],
                           layout_idx: int,
                           report: Dict[str, Any]) -> int:
        """Build multiple slides from split content."""
        num_splits = optimization['split_into_slides']
        title = slide_spec.get('title', 'Content')
        content = slide_spec.get('content', '')

        # Simple split by length
        chunk_size = len(content) // num_splits
        slides_built = 0

        for i in range(num_splits):
            start = i * chunk_size
            end = start + chunk_size if i < num_splits - 1 else len(content)
            chunk = content[start:end]

            slide_title = f"{title} ({i + 1}/{num_splits})" if num_splits > 1 else title

            self.text_builder.add_content_slide(
                self.handler, slide_title, chunk, layout_idx
            )
            slides_built += 1

        return slides_built

    def save(self, output_path: Path) -> None:
        """Save the presentation."""
        self.handler.save(output_path)

    def get_slide_count(self) -> int:
        """Get current slide count."""
        return self.handler.get_slide_count()
