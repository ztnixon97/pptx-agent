"""
Autonomous Designer - LLM-powered autonomous design decision making.
"""

import json
from typing import Dict, List, Any, Optional, Tuple
from .openai_client import OpenAIClient


class AutonomousDesigner:
    """
    Makes autonomous design decisions for presentations using LLM.

    This agent determines styling, layout, colors, fonts, and content
    organization without requiring explicit user input.
    """

    def __init__(self, client: OpenAIClient):
        """
        Initialize the autonomous designer.

        Args:
            client: OpenAI client instance
        """
        self.client = client

    def analyze_content_and_decide_layout(self, slide_content: Dict[str, Any],
                                         available_layouts: List[str]) -> Dict[str, Any]:
        """
        Analyze content and autonomously decide the best layout.

        Args:
            slide_content: Dictionary with slide content
            available_layouts: List of available layout names

        Returns:
            Layout decision with styling recommendations
        """
        system_prompt = """You are an expert presentation designer. Analyze the content
and decide the optimal layout, styling, and formatting. Return JSON with this structure:

{
  "layout_choice": "layout_name",
  "layout_index": 1,
  "reasoning": "Why this layout was chosen",
  "styling": {
    "font_size": 18,
    "use_bold_title": true,
    "color_scheme": "professional|creative|minimal",
    "alignment": "left|center|justified"
  },
  "content_organization": {
    "should_split": false,
    "suggested_structure": "bullets|paragraphs|columns",
    "max_items": 7
  }
}"""

        user_prompt = f"""Analyze this slide content and decide the best layout and styling:

Content: {json.dumps(slide_content, indent=2)}

Available layouts: {', '.join(available_layouts)}

Make autonomous decisions about layout, styling, and content organization."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.3)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            # Fallback to safe defaults
            return {
                "layout_choice": available_layouts[1] if len(available_layouts) > 1 else available_layouts[0],
                "layout_index": 1,
                "reasoning": "Default layout selection",
                "styling": {
                    "font_size": 18,
                    "use_bold_title": True,
                    "color_scheme": "professional",
                    "alignment": "left"
                },
                "content_organization": {
                    "should_split": False,
                    "suggested_structure": "bullets",
                    "max_items": 7
                }
            }

    def decide_color_scheme(self, presentation_topic: str,
                           presentation_tone: str = "professional") -> Dict[str, Any]:
        """
        Autonomously decide color scheme for presentation.

        Args:
            presentation_topic: Topic of the presentation
            presentation_tone: Desired tone (professional, creative, playful, etc.)

        Returns:
            Color scheme recommendations
        """
        system_prompt = """You are a color theory expert for presentations. Choose an
appropriate color scheme based on topic and tone. Return JSON:

{
  "primary_color": {"r": 68, "g": 114, "b": 196, "name": "Professional Blue"},
  "secondary_color": {"r": 237, "g": 125, "b": 49, "name": "Accent Orange"},
  "text_color": {"r": 0, "g": 0, "b": 0, "name": "Black"},
  "background_color": {"r": 255, "g": 255, "b": 255, "name": "White"},
  "accent_colors": [
    {"r": 165, "g": 165, "b": 165, "name": "Gray"}
  ],
  "reasoning": "Why this scheme works for this topic"
}"""

        user_prompt = f"""Choose an optimal color scheme for:

Topic: {presentation_topic}
Tone: {presentation_tone}

Consider color psychology, readability, and professional standards."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.4)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            # Professional default scheme
            return {
                "primary_color": {"r": 68, "g": 114, "b": 196, "name": "Professional Blue"},
                "secondary_color": {"r": 237, "g": 125, "b": 49, "name": "Accent Orange"},
                "text_color": {"r": 0, "g": 0, "b": 0, "name": "Black"},
                "background_color": {"r": 255, "g": 255, "b": 255, "name": "White"},
                "accent_colors": [{"r": 165, "g": 165, "b": 165, "name": "Gray"}],
                "reasoning": "Classic professional color scheme"
            }

    def optimize_content_formatting(self, content: str,
                                   content_type: str,
                                   available_space: Dict[str, float]) -> Dict[str, Any]:
        """
        Autonomously optimize content formatting to fit space.

        Args:
            content: Content to format
            content_type: Type of content
            available_space: Available space dimensions

        Returns:
            Formatting decisions and optimized content
        """
        system_prompt = """You are an expert at optimizing presentation content.
Given content and space constraints, decide how to format it optimally. Return JSON:

{
  "action": "keep|shorten|split|reformat",
  "optimized_content": "reformatted content",
  "formatting": {
    "font_size": 18,
    "line_spacing": 1.2,
    "use_bullets": true,
    "abbreviate": false
  },
  "split_into_slides": 1,
  "reasoning": "Why these decisions were made"
}"""

        user_prompt = f"""Optimize this content to fit the space:

Content ({content_type}):
{content[:500]}...

Available space: {json.dumps(available_space)}

Make formatting decisions to ensure content fits well and is readable."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.3)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "action": "keep",
                "optimized_content": content,
                "formatting": {
                    "font_size": 18,
                    "line_spacing": 1.2,
                    "use_bullets": True,
                    "abbreviate": False
                },
                "split_into_slides": 1,
                "reasoning": "Content fits current space"
            }

    def decide_visual_hierarchy(self, slide_elements: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Decide visual hierarchy for slide elements.

        Args:
            slide_elements: List of elements to arrange

        Returns:
            Visual hierarchy decisions
        """
        system_prompt = """You are an expert in visual design hierarchy. Arrange elements
to create optimal visual flow and emphasis. Return JSON:

{
  "element_order": [0, 2, 1],
  "emphasis": {
    "primary_element": 0,
    "secondary_elements": [1, 2]
  },
  "sizing": {
    "0": {"font_size": 28, "bold": true},
    "1": {"font_size": 20, "bold": false},
    "2": {"font_size": 16, "bold": false}
  },
  "spacing": {
    "vertical_gaps": [0.3, 0.2, 0.2],
    "margins": {"top": 0.5, "left": 1.0}
  },
  "reasoning": "Visual hierarchy explanation"
}"""

        user_prompt = f"""Determine optimal visual hierarchy for these elements:

{json.dumps(slide_elements, indent=2)}

Apply design principles for emphasis, readability, and visual flow."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.3)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "element_order": list(range(len(slide_elements))),
                "emphasis": {
                    "primary_element": 0,
                    "secondary_elements": list(range(1, len(slide_elements)))
                },
                "sizing": {
                    str(i): {"font_size": 18 - i * 2, "bold": i == 0}
                    for i in range(len(slide_elements))
                },
                "spacing": {
                    "vertical_gaps": [0.2] * len(slide_elements),
                    "margins": {"top": 0.5, "left": 1.0}
                },
                "reasoning": "Standard hierarchy applied"
            }

    def suggest_content_improvements(self, content: str,
                                    context: str = "") -> Dict[str, Any]:
        """
        Suggest autonomous improvements to content.

        Args:
            content: Current content
            context: Additional context

        Returns:
            Improvement suggestions
        """
        system_prompt = """You are an expert presentation content editor. Analyze content
and suggest improvements for clarity, impact, and professionalism. Return JSON:

{
  "improved_content": "enhanced version",
  "changes_made": [
    "Made language more concise",
    "Improved parallel structure",
    "Enhanced key message"
  ],
  "additional_suggestions": [
    "Consider adding a supporting statistic",
    "Could benefit from visual element"
  ]
}"""

        user_prompt = f"""Improve this presentation content:

Content: {content}

Context: {context}

Make it more impactful, clear, and professional while preserving key messages."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.5)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "improved_content": content,
                "changes_made": [],
                "additional_suggestions": []
            }

    def decide_chart_visualization(self, data_description: str,
                                   data_characteristics: Dict[str, Any]) -> Dict[str, Any]:
        """
        Autonomously decide best chart type for data.

        Args:
            data_description: Description of the data
            data_characteristics: Data properties (num_series, data_type, etc.)

        Returns:
            Chart type and styling decisions
        """
        system_prompt = """You are a data visualization expert. Choose the optimal
chart type and styling for the data. Return JSON:

{
  "chart_type": "bar|line|pie|scatter|area",
  "orientation": "horizontal|vertical",
  "styling": {
    "colors": ["#4472C4", "#ED7D31", "#A5A5A5"],
    "show_grid": true,
    "show_legend": true,
    "show_values": false
  },
  "formatting": {
    "title_size": 16,
    "axis_label_size": 12,
    "value_format": "percentage|currency|number"
  },
  "reasoning": "Why this visualization is optimal"
}"""

        user_prompt = f"""Choose the best chart type for this data:

Description: {data_description}

Characteristics: {json.dumps(data_characteristics)}

Consider clarity, impact, and standard visualization best practices."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.3)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "chart_type": "bar",
                "orientation": "vertical",
                "styling": {
                    "colors": ["#4472C4", "#ED7D31", "#A5A5A5"],
                    "show_grid": True,
                    "show_legend": True,
                    "show_values": False
                },
                "formatting": {
                    "title_size": 16,
                    "axis_label_size": 12,
                    "value_format": "number"
                },
                "reasoning": "Bar chart is versatile default"
            }

    def auto_adjust_for_audience(self, content: Dict[str, Any],
                                audience: str = "general") -> Dict[str, Any]:
        """
        Automatically adjust presentation for target audience.

        Args:
            content: Presentation content
            audience: Target audience (executives, technical, general, etc.)

        Returns:
            Adjusted content and style recommendations
        """
        system_prompt = """You are an expert at tailoring presentations for different
audiences. Adjust content, complexity, and style appropriately. Return JSON:

{
  "adjusted_content": "content tailored for audience",
  "complexity_level": "basic|intermediate|advanced",
  "recommended_changes": [
    "Use more technical terminology",
    "Add supporting details",
    "Simplify jargon"
  ],
  "style_adjustments": {
    "formality": "formal|casual",
    "detail_level": "high|medium|low",
    "use_examples": true
  }
}"""

        user_prompt = f"""Adjust this content for the target audience:

Content: {json.dumps(content, indent=2)}

Audience: {audience}

Tailor language, complexity, and style appropriately."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.4)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "adjusted_content": content,
                "complexity_level": "intermediate",
                "recommended_changes": [],
                "style_adjustments": {
                    "formality": "formal",
                    "detail_level": "medium",
                    "use_examples": True
                }
            }
