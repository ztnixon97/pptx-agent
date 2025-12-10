"""
Layout Optimizer - Intelligently selects and optimizes slide layouts.
"""

from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
import json


class LayoutOptimizer:
    """
    Optimizes layout selection based on content type, amount, and presentation context.
    """

    # Layout templates categorized by use case
    LAYOUT_PATTERNS = {
        'title': {
            'keywords': ['title', 'cover', 'intro', 'opening'],
            'best_for': 'First slide, section headers',
            'elements': ['title', 'subtitle']
        },
        'content': {
            'keywords': ['content', 'body', 'text', 'default'],
            'best_for': 'General content with title and body',
            'elements': ['title', 'content']
        },
        'section': {
            'keywords': ['section', 'divider', 'separator', 'chapter'],
            'best_for': 'Section breaks, topic transitions',
            'elements': ['title', 'subtitle']
        },
        'two_content': {
            'keywords': ['two', 'comparison', 'double', 'split'],
            'best_for': 'Comparing two items, side-by-side content',
            'elements': ['title', 'left_content', 'right_content']
        },
        'blank': {
            'keywords': ['blank', 'empty', 'custom'],
            'best_for': 'Custom layouts, images, diagrams',
            'elements': []
        },
        'title_only': {
            'keywords': ['title only', 'heading'],
            'best_for': 'Slides with title and custom content',
            'elements': ['title']
        },
        'picture': {
            'keywords': ['picture', 'image', 'photo'],
            'best_for': 'Image-centric slides',
            'elements': ['title', 'picture']
        },
        'caption': {
            'keywords': ['caption', 'quote'],
            'best_for': 'Quotes, testimonials, highlighted text',
            'elements': ['caption']
        }
    }

    def __init__(self, available_layouts: Optional[List[str]] = None):
        """
        Initialize layout optimizer.

        Args:
            available_layouts: List of available layout names from template
        """
        self.available_layouts = available_layouts or []
        self.layout_map = self._map_layouts()

    def _map_layouts(self) -> Dict[str, int]:
        """Map layout types to indices based on available layouts."""
        layout_map = {}

        for idx, layout_name in enumerate(self.available_layouts):
            layout_lower = layout_name.lower()

            for layout_type, pattern in self.LAYOUT_PATTERNS.items():
                for keyword in pattern['keywords']:
                    if keyword in layout_lower:
                        if layout_type not in layout_map:
                            layout_map[layout_type] = idx
                        break

        # Ensure we have fallbacks
        if 'content' not in layout_map and len(self.available_layouts) > 1:
            layout_map['content'] = 1
        if 'title' not in layout_map and len(self.available_layouts) > 0:
            layout_map['title'] = 0

        return layout_map

    def select_optimal_layout(self, slide_spec: Dict[str, Any]) -> Tuple[int, str]:
        """
        Select optimal layout for a slide specification.

        Args:
            slide_spec: Slide specification with content and type

        Returns:
            Tuple of (layout_index, layout_reasoning)
        """
        slide_type = slide_spec.get('slide_type', 'content')
        elements = slide_spec.get('elements', [])
        content = slide_spec.get('content', '')

        # Direct type mapping
        if slide_type in self.layout_map:
            return (
                self.layout_map[slide_type],
                f"Direct match for {slide_type} slide type"
            )

        # Analyze elements to determine best layout
        element_types = [e.get('type') for e in elements]

        # Image-heavy content
        if 'image' in element_types:
            if 'picture' in self.layout_map:
                return (
                    self.layout_map['picture'],
                    "Image content detected, using picture layout"
                )
            elif 'blank' in self.layout_map:
                return (
                    self.layout_map['blank'],
                    "Image content, using blank layout for flexibility"
                )

        # Comparison content (two columns)
        if len(elements) == 2 and all(e.get('type') in ['text', 'bullet_points'] for e in elements):
            if 'two_content' in self.layout_map:
                return (
                    self.layout_map['two_content'],
                    "Two content elements, using comparison layout"
                )

        # Table or chart content
        if any(t in element_types for t in ['table', 'chart']):
            if 'title_only' in self.layout_map:
                return (
                    self.layout_map['title_only'],
                    "Data visualization, using title-only for clean layout"
                )

        # Default to content layout
        default_layout = self.layout_map.get('content', 1)
        return (default_layout, "Default content layout")

    def optimize_slide_structure(self, slide_spec: Dict[str, Any],
                                 max_elements: int = 5) -> Dict[str, Any]:
        """
        Optimize slide structure for better layout.

        Args:
            slide_spec: Original slide specification
            max_elements: Maximum elements per slide

        Returns:
            Optimized slide specification (may split into multiple slides)
        """
        elements = slide_spec.get('elements', [])

        if len(elements) <= max_elements:
            return {
                'slides': [slide_spec],
                'optimization': 'no_changes',
                'reasoning': 'Slide structure is optimal'
            }

        # Need to split
        optimized_slides = []
        current_elements = []
        base_title = slide_spec.get('title', 'Content')

        for i, element in enumerate(elements):
            current_elements.append(element)

            if len(current_elements) >= max_elements or i == len(elements) - 1:
                new_slide = {
                    'slide_type': slide_spec.get('slide_type', 'content'),
                    'title': f"{base_title} ({len(optimized_slides) + 1})" if len(optimized_slides) > 0 else base_title,
                    'content': slide_spec.get('content', ''),
                    'elements': current_elements.copy()
                }
                optimized_slides.append(new_slide)
                current_elements = []

        return {
            'slides': optimized_slides,
            'optimization': 'split',
            'reasoning': f'Split into {len(optimized_slides)} slides to avoid overcrowding'
        }

    def suggest_layout_alternatives(self, slide_spec: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Suggest alternative layouts for a slide.

        Args:
            slide_spec: Slide specification

        Returns:
            List of alternative layout suggestions
        """
        alternatives = []
        primary_layout, primary_reason = self.select_optimal_layout(slide_spec)

        # Add primary suggestion
        alternatives.append({
            'layout_index': primary_layout,
            'layout_name': self.available_layouts[primary_layout] if primary_layout < len(self.available_layouts) else 'Default',
            'score': 1.0,
            'reasoning': primary_reason,
            'is_primary': True
        })

        # Suggest secondary options based on content
        elements = slide_spec.get('elements', [])

        if elements:
            # Blank layout for custom control
            if 'blank' in self.layout_map:
                alternatives.append({
                    'layout_index': self.layout_map['blank'],
                    'layout_name': self.available_layouts[self.layout_map['blank']],
                    'score': 0.7,
                    'reasoning': 'Blank layout for maximum customization',
                    'is_primary': False
                })

            # Title-only for data-heavy slides
            if 'title_only' in self.layout_map and any(e.get('type') in ['table', 'chart'] for e in elements):
                alternatives.append({
                    'layout_index': self.layout_map['title_only'],
                    'layout_name': self.available_layouts[self.layout_map['title_only']],
                    'score': 0.8,
                    'reasoning': 'Title-only for prominent data display',
                    'is_primary': False
                })

        return alternatives

    def analyze_presentation_flow(self, slides: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Analyze overall presentation layout flow.

        Args:
            slides: List of slide specifications

        Returns:
            Flow analysis and recommendations
        """
        analysis = {
            'total_slides': len(slides),
            'layout_distribution': {},
            'flow_issues': [],
            'recommendations': []
        }

        # Track layout usage
        for slide in slides:
            layout_idx, _ = self.select_optimal_layout(slide)
            layout_name = self.available_layouts[layout_idx] if layout_idx < len(self.available_layouts) else 'default'

            if layout_name not in analysis['layout_distribution']:
                analysis['layout_distribution'][layout_name] = 0
            analysis['layout_distribution'][layout_name] += 1

        # Check for issues
        if analysis['total_slides'] > 0:
            content_count = analysis['layout_distribution'].get('content', 0)
            content_ratio = content_count / analysis['total_slides']

            if content_ratio > 0.8:
                analysis['flow_issues'].append({
                    'issue': 'monotonous_layout',
                    'severity': 'medium',
                    'description': 'Too many slides use the same layout'
                })
                analysis['recommendations'].append(
                    'Add section dividers to break up content slides'
                )

            # Check for missing title slide
            if slides[0].get('slide_type') != 'title':
                analysis['flow_issues'].append({
                    'issue': 'missing_title_slide',
                    'severity': 'high',
                    'description': 'Presentation should start with a title slide'
                })
                analysis['recommendations'].append(
                    'Add a title slide at the beginning'
                )

            # Check for long sequences of content
            sequence_length = 0
            max_sequence = 0
            for slide in slides:
                if slide.get('slide_type') in ['content', 'text']:
                    sequence_length += 1
                    max_sequence = max(max_sequence, sequence_length)
                else:
                    sequence_length = 0

            if max_sequence > 5:
                analysis['flow_issues'].append({
                    'issue': 'long_content_sequence',
                    'severity': 'medium',
                    'description': f'Sequence of {max_sequence} similar slides may be fatiguing'
                })
                analysis['recommendations'].append(
                    'Insert visual breaks (images, charts) or section dividers'
                )

        return analysis

    def get_layout_best_practices(self, layout_type: str) -> Dict[str, Any]:
        """
        Get best practices for a specific layout type.

        Args:
            layout_type: Type of layout

        Returns:
            Best practices and guidelines
        """
        practices = {
            'title': {
                'max_title_length': 60,
                'max_subtitle_length': 100,
                'font_size_title': 44,
                'font_size_subtitle': 32,
                'tips': [
                    'Keep title concise and impactful',
                    'Subtitle should clarify or expand on title',
                    'Use title case for main title'
                ]
            },
            'content': {
                'max_bullet_points': 7,
                'max_chars_per_bullet': 80,
                'font_size_title': 32,
                'font_size_body': 18,
                'tips': [
                    'Follow 6x6 rule: max 6 bullets, 6 words each',
                    'Use parallel structure for bullets',
                    'Leave white space for readability'
                ]
            },
            'two_content': {
                'max_items_per_column': 5,
                'font_size': 16,
                'tips': [
                    'Ensure balanced content between columns',
                    'Use for comparisons or parallel concepts',
                    'Keep columns aligned vertically'
                ]
            }
        }

        return practices.get(layout_type, {
            'tips': ['Follow standard presentation design principles']
        })
