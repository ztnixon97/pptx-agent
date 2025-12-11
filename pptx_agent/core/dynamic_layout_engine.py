"""
Dynamic Layout Engine - Algorithmically position slide elements without template constraints.

This module enables flexible, LLM-driven slide creation by:
- Calculating optimal element positions based on content
- Supporting custom layouts not defined in templates
- Allowing dynamic reshaping and reorganization
- Using templates only for styling (colors, fonts), not structure
"""

from typing import Dict, List, Any, Tuple, Optional
from dataclasses import dataclass
from pptx.util import Inches, Pt
from enum import Enum


class LayoutRegion(Enum):
    """Predefined regions for common layout patterns."""
    FULL = "full"  # Entire slide
    TOP_HALF = "top_half"
    BOTTOM_HALF = "bottom_half"
    LEFT_HALF = "left_half"
    RIGHT_HALF = "right_half"
    TOP_LEFT = "top_left"
    TOP_RIGHT = "top_right"
    BOTTOM_LEFT = "bottom_left"
    BOTTOM_RIGHT = "bottom_right"
    CENTER = "center"
    HEADER = "header"  # Top title area
    CONTENT = "content"  # Main content area below header


@dataclass
class ElementBox:
    """Represents a positioned element on a slide."""
    left: float  # inches
    top: float  # inches
    width: float  # inches
    height: float  # inches
    element_type: str  # 'title', 'text', 'chart', 'table', 'image', etc.
    content: Any = None
    z_index: int = 0  # layering order


class DynamicLayoutEngine:
    """
    Generate slide layouts dynamically based on content requirements.

    Templates are used ONLY for theming (colors, fonts), not layout structure.
    """

    def __init__(self, slide_width: float = 10.0, slide_height: float = 7.5):
        """
        Initialize layout engine.

        Args:
            slide_width: Slide width in inches (default 10.0 for 4:3)
            slide_height: Slide height in inches (default 7.5)
        """
        self.slide_width = slide_width
        self.slide_height = slide_height

        # Standard margins
        self.margin_left = 0.5
        self.margin_right = 0.5
        self.margin_top = 0.5
        self.margin_bottom = 0.5

        # Calculate available content area
        self.content_width = self.slide_width - self.margin_left - self.margin_right
        self.content_height = self.slide_height - self.margin_top - self.margin_bottom

        # Standard element heights
        self.title_height = 1.0
        self.subtitle_height = 0.6
        self.content_gap = 0.3  # Space between elements

    def get_region_bounds(self, region: LayoutRegion) -> Tuple[float, float, float, float]:
        """
        Get bounds for a predefined region.

        Returns:
            Tuple of (left, top, width, height) in inches
        """
        regions = {
            LayoutRegion.FULL: (
                self.margin_left,
                self.margin_top,
                self.content_width,
                self.content_height
            ),
            LayoutRegion.HEADER: (
                self.margin_left,
                self.margin_top,
                self.content_width,
                self.title_height
            ),
            LayoutRegion.CONTENT: (
                self.margin_left,
                self.margin_top + self.title_height + self.content_gap,
                self.content_width,
                self.content_height - self.title_height - self.content_gap
            ),
            LayoutRegion.TOP_HALF: (
                self.margin_left,
                self.margin_top,
                self.content_width,
                self.content_height / 2
            ),
            LayoutRegion.BOTTOM_HALF: (
                self.margin_left,
                self.margin_top + self.content_height / 2,
                self.content_width,
                self.content_height / 2
            ),
            LayoutRegion.LEFT_HALF: (
                self.margin_left,
                self.margin_top,
                self.content_width / 2,
                self.content_height
            ),
            LayoutRegion.RIGHT_HALF: (
                self.margin_left + self.content_width / 2,
                self.margin_top,
                self.content_width / 2,
                self.content_height
            ),
            LayoutRegion.TOP_LEFT: (
                self.margin_left,
                self.margin_top,
                self.content_width / 2,
                self.content_height / 2
            ),
            LayoutRegion.TOP_RIGHT: (
                self.margin_left + self.content_width / 2,
                self.margin_top,
                self.content_width / 2,
                self.content_height / 2
            ),
            LayoutRegion.BOTTOM_LEFT: (
                self.margin_left,
                self.margin_top + self.content_height / 2,
                self.content_width / 2,
                self.content_height / 2
            ),
            LayoutRegion.BOTTOM_RIGHT: (
                self.margin_left + self.content_width / 2,
                self.margin_top + self.content_height / 2,
                self.content_width / 2,
                self.content_height / 2
            ),
            LayoutRegion.CENTER: (
                self.margin_left + self.content_width / 4,
                self.margin_top + self.content_height / 4,
                self.content_width / 2,
                self.content_height / 2
            ),
        }

        return regions.get(region, regions[LayoutRegion.FULL])

    def create_title_content_layout(self, title_height: Optional[float] = None) -> List[ElementBox]:
        """
        Standard title + content layout.

        Returns:
            List of ElementBox for [title, content_area]
        """
        if title_height is None:
            title_height = self.title_height

        return [
            ElementBox(
                left=self.margin_left,
                top=self.margin_top,
                width=self.content_width,
                height=title_height,
                element_type='title'
            ),
            ElementBox(
                left=self.margin_left,
                top=self.margin_top + title_height + self.content_gap,
                width=self.content_width,
                height=self.content_height - title_height - self.content_gap,
                element_type='content'
            )
        ]

    def create_two_column_layout(self, title: bool = True,
                                 column_ratio: float = 0.5) -> List[ElementBox]:
        """
        Two-column layout with optional title.

        Args:
            title: Include title at top
            column_ratio: Ratio of left column width (0.5 = equal columns)

        Returns:
            List of ElementBox
        """
        elements = []

        content_top = self.margin_top
        content_height_available = self.content_height

        if title:
            elements.append(ElementBox(
                left=self.margin_left,
                top=self.margin_top,
                width=self.content_width,
                height=self.title_height,
                element_type='title'
            ))
            content_top = self.margin_top + self.title_height + self.content_gap
            content_height_available = self.content_height - self.title_height - self.content_gap

        # Left column
        left_width = self.content_width * column_ratio - self.content_gap / 2
        elements.append(ElementBox(
            left=self.margin_left,
            top=content_top,
            width=left_width,
            height=content_height_available,
            element_type='column_left'
        ))

        # Right column
        right_width = self.content_width * (1 - column_ratio) - self.content_gap / 2
        elements.append(ElementBox(
            left=self.margin_left + left_width + self.content_gap,
            top=content_top,
            width=right_width,
            height=content_height_available,
            element_type='column_right'
        ))

        return elements

    def create_grid_layout(self, rows: int, cols: int,
                          title: bool = True) -> List[ElementBox]:
        """
        Grid layout for multiple elements (e.g., image grid).

        Args:
            rows: Number of rows
            cols: Number of columns
            title: Include title at top

        Returns:
            List of ElementBox for each grid cell
        """
        elements = []

        content_top = self.margin_top
        content_height_available = self.content_height

        if title:
            elements.append(ElementBox(
                left=self.margin_left,
                top=self.margin_top,
                width=self.content_width,
                height=self.title_height,
                element_type='title'
            ))
            content_top = self.margin_top + self.title_height + self.content_gap
            content_height_available = self.content_height - self.title_height - self.content_gap

        # Calculate cell dimensions
        cell_width = (self.content_width - (cols - 1) * self.content_gap) / cols
        cell_height = (content_height_available - (rows - 1) * self.content_gap) / rows

        # Create grid cells
        for row in range(rows):
            for col in range(cols):
                elements.append(ElementBox(
                    left=self.margin_left + col * (cell_width + self.content_gap),
                    top=content_top + row * (cell_height + self.content_gap),
                    width=cell_width,
                    height=cell_height,
                    element_type=f'grid_cell_{row}_{col}'
                ))

        return elements

    def create_image_text_layout(self, image_position: str = 'left',
                                 image_ratio: float = 0.6) -> List[ElementBox]:
        """
        Layout with image and text side-by-side.

        Args:
            image_position: 'left', 'right', 'top', or 'bottom'
            image_ratio: Proportion of space for image (0.0-1.0)

        Returns:
            List of ElementBox for [title, image, text]
        """
        elements = []

        # Title
        elements.append(ElementBox(
            left=self.margin_left,
            top=self.margin_top,
            width=self.content_width,
            height=self.title_height,
            element_type='title'
        ))

        content_top = self.margin_top + self.title_height + self.content_gap
        content_height_available = self.content_height - self.title_height - self.content_gap

        if image_position in ['left', 'right']:
            # Horizontal split
            image_width = self.content_width * image_ratio - self.content_gap / 2
            text_width = self.content_width * (1 - image_ratio) - self.content_gap / 2

            if image_position == 'left':
                elements.append(ElementBox(
                    left=self.margin_left,
                    top=content_top,
                    width=image_width,
                    height=content_height_available,
                    element_type='image'
                ))
                elements.append(ElementBox(
                    left=self.margin_left + image_width + self.content_gap,
                    top=content_top,
                    width=text_width,
                    height=content_height_available,
                    element_type='text'
                ))
            else:  # right
                elements.append(ElementBox(
                    left=self.margin_left,
                    top=content_top,
                    width=text_width,
                    height=content_height_available,
                    element_type='text'
                ))
                elements.append(ElementBox(
                    left=self.margin_left + text_width + self.content_gap,
                    top=content_top,
                    width=image_width,
                    height=content_height_available,
                    element_type='image'
                ))

        else:  # top or bottom
            # Vertical split
            image_height = content_height_available * image_ratio - self.content_gap / 2
            text_height = content_height_available * (1 - image_ratio) - self.content_gap / 2

            if image_position == 'top':
                elements.append(ElementBox(
                    left=self.margin_left,
                    top=content_top,
                    width=self.content_width,
                    height=image_height,
                    element_type='image'
                ))
                elements.append(ElementBox(
                    left=self.margin_left,
                    top=content_top + image_height + self.content_gap,
                    width=self.content_width,
                    height=text_height,
                    element_type='text'
                ))
            else:  # bottom
                elements.append(ElementBox(
                    left=self.margin_left,
                    top=content_top,
                    width=self.content_width,
                    height=text_height,
                    element_type='text'
                ))
                elements.append(ElementBox(
                    left=self.margin_left,
                    top=content_top + text_height + self.content_gap,
                    width=self.content_width,
                    height=image_height,
                    element_type='image'
                ))

        return elements

    def create_vertical_stack_layout(self, num_elements: int,
                                    title: bool = True,
                                    heights: Optional[List[float]] = None) -> List[ElementBox]:
        """
        Stack elements vertically with custom heights.

        Args:
            num_elements: Number of elements to stack
            title: Include title at top
            heights: Optional list of height ratios (must sum to 1.0)

        Returns:
            List of ElementBox
        """
        elements = []

        content_top = self.margin_top
        content_height_available = self.content_height

        if title:
            elements.append(ElementBox(
                left=self.margin_left,
                top=self.margin_top,
                width=self.content_width,
                height=self.title_height,
                element_type='title'
            ))
            content_top = self.margin_top + self.title_height + self.content_gap
            content_height_available = self.content_height - self.title_height - self.content_gap

        # Calculate heights
        if heights is None:
            heights = [1.0 / num_elements] * num_elements

        # Account for gaps
        total_gap_height = (num_elements - 1) * self.content_gap
        usable_height = content_height_available - total_gap_height

        # Create stacked elements
        current_top = content_top
        for i, height_ratio in enumerate(heights):
            element_height = usable_height * height_ratio
            elements.append(ElementBox(
                left=self.margin_left,
                top=current_top,
                width=self.content_width,
                height=element_height,
                element_type=f'stack_element_{i}'
            ))
            current_top += element_height + self.content_gap

        return elements

    def create_custom_layout(self, specifications: List[Dict[str, Any]]) -> List[ElementBox]:
        """
        Create completely custom layout from LLM specifications.

        Args:
            specifications: List of dicts with layout specs:
                [
                    {
                        'type': 'title',
                        'region': 'header',  # or custom bounds
                        'left': 0.5,  # optional, overrides region
                        'top': 0.5,
                        'width': 9.0,
                        'height': 1.0
                    },
                    ...
                ]

        Returns:
            List of ElementBox
        """
        elements = []

        for spec in specifications:
            # Check if using predefined region
            if 'region' in spec and spec['region'] in [r.value for r in LayoutRegion]:
                region = LayoutRegion(spec['region'])
                left, top, width, height = self.get_region_bounds(region)
            else:
                # Use custom bounds
                left = spec.get('left', self.margin_left)
                top = spec.get('top', self.margin_top)
                width = spec.get('width', self.content_width)
                height = spec.get('height', 1.0)

            elements.append(ElementBox(
                left=left,
                top=top,
                width=width,
                height=height,
                element_type=spec.get('type', 'content'),
                content=spec.get('content'),
                z_index=spec.get('z_index', 0)
            ))

        # Sort by z_index for layering
        elements.sort(key=lambda e: e.z_index)

        return elements

    def suggest_layout(self, content_description: Dict[str, Any]) -> List[ElementBox]:
        """
        Suggest optimal layout based on content description.

        Args:
            content_description: Dict describing content:
                {
                    'has_title': bool,
                    'has_image': bool,
                    'has_chart': bool,
                    'has_table': bool,
                    'text_amount': 'short'|'medium'|'long',
                    'num_images': int,
                    'comparison': bool  # side-by-side content
                }

        Returns:
            Suggested ElementBox layout
        """
        has_title = content_description.get('has_title', True)
        has_image = content_description.get('has_image', False)
        has_chart = content_description.get('has_chart', False)
        comparison = content_description.get('comparison', False)
        num_images = content_description.get('num_images', 0)

        # Grid layout for multiple images
        if num_images > 1:
            if num_images <= 4:
                return self.create_grid_layout(2, 2, title=has_title)
            elif num_images <= 6:
                return self.create_grid_layout(2, 3, title=has_title)
            else:
                return self.create_grid_layout(3, 3, title=has_title)

        # Comparison layout
        if comparison:
            return self.create_two_column_layout(title=has_title)

        # Image + text
        if has_image or has_chart:
            return self.create_image_text_layout(image_position='right', image_ratio=0.55)

        # Default: title + content
        return self.create_title_content_layout()
