"""
Content Validator - Validates and optimizes content to fit on slides.
"""

from typing import Dict, Any, List, Optional, Tuple
from pptx.util import Pt, Inches
from pptx.text.text import TextFrame
import re


class ContentValidator:
    """Validates and optimizes content to ensure it fits on slides properly."""

    def __init__(self, slide_width: int = 10, slide_height: int = 7.5):
        """
        Initialize content validator.

        Args:
            slide_width: Slide width in inches
            slide_height: Slide height in inches
        """
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.content_area_width = slide_width - 2  # 1 inch margins
        self.content_area_height = slide_height - 2.5  # Title + margins

    def estimate_text_height(self, text: str, font_size: int = 18,
                            width: float = 8.0) -> float:
        """
        Estimate height needed for text.

        Args:
            text: The text content
            font_size: Font size in points
            width: Available width in inches

        Returns:
            Estimated height in inches
        """
        # Rough estimation: chars per line based on font size and width
        chars_per_line = int((width * 72) / (font_size * 0.6))
        lines = len(text) / chars_per_line
        line_height = font_size / 72  # Convert pts to inches
        return lines * line_height * 1.2  # 1.2 for line spacing

    def validate_text_content(self, text: str, font_size: int = 18,
                             available_width: Optional[float] = None,
                             available_height: Optional[float] = None) -> Dict[str, Any]:
        """
        Validate if text content fits in available space.

        Args:
            text: Text content
            font_size: Font size
            available_width: Available width in inches
            available_height: Available height in inches

        Returns:
            Dictionary with validation results and suggestions
        """
        width = available_width or self.content_area_width
        height = available_height or self.content_area_height

        estimated_height = self.estimate_text_height(text, font_size, width)

        fits = estimated_height <= height

        result = {
            'fits': fits,
            'estimated_height': estimated_height,
            'available_height': height,
            'overflow': max(0, estimated_height - height),
            'suggestions': []
        }

        if not fits:
            # Generate suggestions
            if font_size > 14:
                suggested_font_size = max(12, font_size - 2)
                result['suggestions'].append({
                    'type': 'reduce_font_size',
                    'current': font_size,
                    'suggested': suggested_font_size,
                    'reason': 'Content too long for current font size'
                })

            result['suggestions'].append({
                'type': 'split_content',
                'reason': 'Content should be split across multiple slides',
                'suggested_splits': 2 if estimated_height < height * 2 else 3
            })

            result['suggestions'].append({
                'type': 'summarize',
                'reason': 'Content could be condensed or summarized'
            })

        return result

    def validate_bullet_points(self, points: List[str], font_size: int = 18,
                              max_points_per_slide: int = 7) -> Dict[str, Any]:
        """
        Validate bullet points.

        Args:
            points: List of bullet point texts
            font_size: Font size
            max_points_per_slide: Maximum recommended points per slide

        Returns:
            Validation results
        """
        result = {
            'fits': True,
            'point_count': len(points),
            'max_recommended': max_points_per_slide,
            'suggestions': []
        }

        if len(points) > max_points_per_slide:
            result['fits'] = False
            result['suggestions'].append({
                'type': 'split_bullets',
                'reason': f'Too many bullet points ({len(points)} > {max_points_per_slide})',
                'suggested_splits': (len(points) + max_points_per_slide - 1) // max_points_per_slide
            })

        # Check individual bullet length
        for i, point in enumerate(points):
            if len(point) > 100:
                result['suggestions'].append({
                    'type': 'shorten_bullet',
                    'index': i,
                    'current_length': len(point),
                    'suggested_length': 80,
                    'reason': 'Bullet point too long, should be concise'
                })

        return result

    def validate_table_size(self, num_rows: int, num_cols: int,
                           has_title: bool = True) -> Dict[str, Any]:
        """
        Validate if table fits on slide.

        Args:
            num_rows: Number of table rows
            num_cols: Number of columns
            has_title: Whether slide has a title

        Returns:
            Validation results
        """
        # Rough estimates
        max_rows = 10 if has_title else 12
        max_cols = 6

        result = {
            'fits': True,
            'num_rows': num_rows,
            'num_cols': num_cols,
            'max_recommended_rows': max_rows,
            'max_recommended_cols': max_cols,
            'suggestions': []
        }

        if num_rows > max_rows:
            result['fits'] = False
            result['suggestions'].append({
                'type': 'split_table',
                'reason': f'Too many rows ({num_rows} > {max_rows})',
                'suggested_approach': 'Split into multiple slides or summarize data'
            })

        if num_cols > max_cols:
            result['fits'] = False
            result['suggestions'].append({
                'type': 'reduce_columns',
                'reason': f'Too many columns ({num_cols} > {max_cols})',
                'suggested_approach': 'Consider transposing or splitting table'
            })

        if num_rows > 7 and num_cols > 4:
            result['suggestions'].append({
                'type': 'consider_chart',
                'reason': 'Large table might be better visualized as a chart'
            })

        return result

    def optimize_content_for_slide(self, content: str,
                                   content_type: str = 'text') -> Dict[str, Any]:
        """
        Optimize content to fit on a slide.

        Args:
            content: Content to optimize
            content_type: Type of content ('text', 'bullets', etc.)

        Returns:
            Optimized content and metadata
        """
        if content_type == 'bullets':
            points = [p.strip() for p in content.split('\n') if p.strip()]
            validation = self.validate_bullet_points(points)

            if not validation['fits']:
                # Auto-optimize
                max_points = validation['max_recommended']
                if len(points) > max_points:
                    return {
                        'optimized': True,
                        'original_count': len(points),
                        'slides': [
                            points[i:i+max_points]
                            for i in range(0, len(points), max_points)
                        ],
                        'reason': 'Split into multiple slides'
                    }

        elif content_type == 'text':
            validation = self.validate_text_content(content)

            if not validation['fits']:
                # Try reducing font size first
                for suggestion in validation['suggestions']:
                    if suggestion['type'] == 'reduce_font_size':
                        return {
                            'optimized': True,
                            'content': content,
                            'font_size': suggestion['suggested'],
                            'reason': 'Reduced font size to fit content'
                        }

                # Otherwise suggest splitting
                return {
                    'optimized': False,
                    'needs_split': True,
                    'suggestions': validation['suggestions']
                }

        return {'optimized': False, 'content': content}

    def suggest_layout_improvements(self, slide_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Suggest layout improvements for a slide.

        Args:
            slide_data: Dictionary with slide content and structure

        Returns:
            List of improvement suggestions
        """
        suggestions = []

        # Check content density
        content_elements = slide_data.get('elements', [])

        if len(content_elements) > 5:
            suggestions.append({
                'type': 'reduce_complexity',
                'reason': 'Slide has too many elements, consider simplifying',
                'current_elements': len(content_elements),
                'recommended': 3
            })

        # Check for mixed content types
        element_types = [e.get('type') for e in content_elements]
        if len(set(element_types)) > 2:
            suggestions.append({
                'type': 'consistency',
                'reason': 'Too many content types on one slide',
                'suggestion': 'Focus on one or two content types per slide'
            })

        return suggestions

    def calculate_optimal_font_size(self, text_length: int,
                                   available_height: float) -> int:
        """
        Calculate optimal font size for text length.

        Args:
            text_length: Length of text in characters
            available_height: Available height in inches

        Returns:
            Recommended font size in points
        """
        # Simple heuristic
        if text_length < 200:
            return 24
        elif text_length < 400:
            return 20
        elif text_length < 600:
            return 18
        elif text_length < 1000:
            return 16
        else:
            return 14

    def validate_image_placement(self, image_width: float, image_height: float,
                                has_title: bool = True,
                                has_text: bool = False) -> Dict[str, Any]:
        """
        Validate image placement on slide.

        Args:
            image_width: Image width in inches
            image_height: Image height in inches
            has_title: Whether slide has title
            has_text: Whether slide has accompanying text

        Returns:
            Validation results
        """
        available_width = self.content_area_width
        available_height = self.content_area_height

        if has_text:
            # Leave room for text
            available_width = available_width / 2

        fits_width = image_width <= available_width
        fits_height = image_height <= available_height

        result = {
            'fits': fits_width and fits_height,
            'suggested_width': min(image_width, available_width),
            'suggested_height': min(image_height, available_height),
            'suggestions': []
        }

        if not fits_width or not fits_height:
            # Calculate aspect ratio preserving resize
            scale = min(
                available_width / image_width if image_width > available_width else 1,
                available_height / image_height if image_height > available_height else 1
            )

            result['suggestions'].append({
                'type': 'resize_image',
                'scale_factor': scale,
                'new_width': image_width * scale,
                'new_height': image_height * scale,
                'reason': 'Image too large for available space'
            })

        return result
