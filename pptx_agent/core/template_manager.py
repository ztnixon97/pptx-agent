"""
Template Manager - Handles PowerPoint template analysis and management.
"""

from pathlib import Path
from typing import Dict, List, Any, Optional
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN


class TemplateManager:
    """Manages PowerPoint templates and extracts their properties."""

    def __init__(self, template_path: Optional[Path] = None):
        """
        Initialize the template manager.

        Args:
            template_path: Path to a template .pptx file
        """
        self.template_path = template_path
        self.template_info = {}

        if template_path and template_path.exists():
            self._analyze_template()

    def _analyze_template(self):
        """Analyze the template and extract useful information."""
        prs = Presentation(str(self.template_path))

        self.template_info = {
            'slide_width': prs.slide_width,
            'slide_height': prs.slide_height,
            'layouts': [],
            'theme_colors': self._extract_theme_colors(prs),
            'master_slides': len(prs.slide_master.slide_layouts)
        }

        # Extract layout information
        for idx, layout in enumerate(prs.slide_layouts):
            layout_info = {
                'index': idx,
                'name': layout.name,
                'placeholders': []
            }

            for placeholder in layout.placeholders:
                ph_info = {
                    'index': placeholder.placeholder_format.idx,
                    'type': placeholder.placeholder_format.type,
                    'name': placeholder.name,
                    'left': placeholder.left,
                    'top': placeholder.top,
                    'width': placeholder.width,
                    'height': placeholder.height
                }
                layout_info['placeholders'].append(ph_info)

            self.template_info['layouts'].append(layout_info)

    def _extract_theme_colors(self, prs: Presentation) -> Dict[str, Any]:
        """Extract theme colors from the presentation."""
        try:
            theme = prs.slide_master.slide_layouts[0].slide.part.related_parts
            return {'extracted': True}
        except Exception:
            return {'extracted': False}

    def get_layout_info(self, layout_name: Optional[str] = None,
                       layout_index: Optional[int] = None) -> Optional[Dict]:
        """
        Get information about a specific layout.

        Args:
            layout_name: Name of the layout
            layout_index: Index of the layout

        Returns:
            Dictionary with layout information or None
        """
        if not self.template_info.get('layouts'):
            return None

        if layout_name:
            for layout in self.template_info['layouts']:
                if layout['name'].lower() == layout_name.lower():
                    return layout

        if layout_index is not None:
            for layout in self.template_info['layouts']:
                if layout['index'] == layout_index:
                    return layout

        return None

    def list_layouts(self) -> List[Dict[str, Any]]:
        """List all available layouts in the template."""
        return self.template_info.get('layouts', [])

    def suggest_layout(self, content_type: str) -> Optional[int]:
        """
        Suggest a layout index based on content type.

        Args:
            content_type: Type of content ('title', 'content', 'section', 'blank', etc.)

        Returns:
            Suggested layout index or None
        """
        if not self.template_info.get('layouts'):
            return 0

        content_type_lower = content_type.lower()

        # Common layout name patterns
        patterns = {
            'title': ['title', 'cover'],
            'content': ['content', 'body', 'text'],
            'section': ['section', 'divider'],
            'blank': ['blank', 'empty'],
            'comparison': ['comparison', 'two'],
            'title_only': ['title only', 'heading']
        }

        for layout in self.template_info['layouts']:
            layout_name = layout['name'].lower()
            if content_type_lower in patterns:
                for pattern in patterns[content_type_lower]:
                    if pattern in layout_name:
                        return layout['index']

        # Default fallbacks
        if content_type_lower == 'title':
            return 0
        else:
            return 1 if len(self.template_info['layouts']) > 1 else 0

    def get_template_summary(self) -> str:
        """Get a human-readable summary of the template."""
        if not self.template_info:
            return "No template loaded"

        summary = [
            f"Template: {self.template_path.name if self.template_path else 'None'}",
            f"Layouts: {len(self.template_info.get('layouts', []))}",
            "\nAvailable Layouts:"
        ]

        for layout in self.template_info.get('layouts', []):
            summary.append(f"  [{layout['index']}] {layout['name']} "
                          f"({len(layout['placeholders'])} placeholders)")

        return "\n".join(summary)
