"""
Content Planner - Uses LLM to plan and generate presentation content.
"""

import json
from typing import Dict, List, Any, Optional
from .openai_client import OpenAIClient


class ContentPlanner:
    """Plans and generates content for PowerPoint presentations using LLM."""

    def __init__(self, client: OpenAIClient):
        """
        Initialize the content planner.

        Args:
            client: OpenAI client instance
        """
        self.client = client

    def create_presentation_outline(self, topic: str, summary: str,
                                   num_slides: Optional[int] = None,
                                   reference_docs: Optional[str] = None) -> Dict[str, Any]:
        """
        Create a presentation outline based on topic and summary.

        Args:
            topic: Main topic of the presentation
            summary: Content summary or key points
            num_slides: Suggested number of slides (optional)
            reference_docs: Reference documentation content (optional)

        Returns:
            Dictionary with presentation outline
        """
        system_prompt = """You are a professional presentation designer with access to comprehensive PowerPoint capabilities.

AVAILABLE CONTENT TYPES:
- Text slides: content, bullets (max 7 points), two-column, quotes, sections
- Tables: data tables, comparisons, summaries (max 6 cols, 10 rows recommended)
- Charts: bar, line, pie, scatter, area (2-8 series, 3-12 categories)
- Images: single, grids (up to 6), with text
- SmartArt-like: process_flow, cycle, hierarchy, comparison, venn, timeline
- Shapes: flowcharts, callouts, icons, annotations

BEST PRACTICES:
- Mix content types for variety
- Use bullets for key points (3-7 per slide)
- Use tables for structured data
- Use charts for trends/comparisons
- Use process flows for sequential steps
- Use hierarchy for organizational structure
- Keep slides focused (one main idea per slide)

Create a detailed outline for a PowerPoint presentation. Return JSON with this structure:

{
  "title": "Presentation Title",
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title|content|section|table|chart|image|smartart|shapes|blank",
      "title": "Slide Title",
      "content": "Main content description",
      "elements": [
        {
          "type": "text|bullet_points|table|chart|image|process_flow|cycle|hierarchy|comparison|venn|timeline|flowchart|shapes",
          "data": "Content or description",
          "details": {
            "for_charts": {
              "chart_type": "bar|line|pie|scatter|area",
              "categories": [],
              "series": []
            },
            "for_tables": {
              "headers": [],
              "rows": []
            },
            "for_smartart": {
              "smartart_type": "process_flow|cycle|hierarchy|comparison|venn|timeline",
              "items": []
            }
          }
        }
      ]
    }
  ]
}

SLIDE TYPES & WHEN TO USE:
- title: Opening slide only
- content: General text content
- section: Section dividers
- table: Structured data, comparisons
- chart: Trends, comparisons, distributions
- smartart: Processes, cycles, hierarchies, comparisons
- shapes: Flowcharts, annotations
- image: Visual content
- blank: Custom layouts

ELEMENT TYPES:
- bullet_points: Key points (provide list)
- table: Data (provide headers and rows)
- chart: Visualizations (specify type: bar/line/pie and data)
- process_flow: Sequential steps (provide steps list)
- cycle: Circular processes (provide items list)
- hierarchy: Organizational structure (provide root and children)
- comparison: Side-by-side (provide left_items and right_items)
- venn: Overlapping concepts (provide left/right/overlap items)
- timeline: Chronological events (provide events with dates)
- flowchart: Decision trees (provide steps with 'decision' flag)
"""

        user_prompt = f"""Create a presentation outline for:

Topic: {topic}

Summary: {summary}"""

        if num_slides:
            user_prompt += f"\n\nTarget number of slides: {num_slides}"

        if reference_docs:
            user_prompt += f"\n\nReference materials:\n{reference_docs[:2000]}"

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.7)

        try:
            outline = json.loads(response)
            return outline
        except json.JSONDecodeError:
            return {
                "title": topic,
                "slides": [
                    {
                        "slide_number": 1,
                        "slide_type": "title",
                        "title": topic,
                        "content": summary,
                        "elements": []
                    }
                ]
            }

    def generate_slide_content(self, slide_title: str, slide_type: str,
                              context: str, reference_docs: Optional[str] = None) -> Dict[str, Any]:
        """
        Generate detailed content for a specific slide.

        Args:
            slide_title: Title of the slide
            slide_type: Type of slide (text, table, chart, etc.)
            context: Context about what the slide should contain
            reference_docs: Optional reference materials

        Returns:
            Dictionary with generated slide content
        """
        system_prompt = f"""You are creating content for a PowerPoint slide.
Generate detailed content for a {slide_type} slide. Return JSON with this structure:

{{
  "title": "Slide Title",
  "subtitle": "Optional subtitle",
  "content": {{
    "main_text": "Main content text",
    "bullet_points": ["point 1", "point 2"],
    "notes": "Speaker notes"
  }}
}}
"""

        user_prompt = f"""Generate content for this slide:

Title: {slide_title}
Type: {slide_type}
Context: {context}"""

        if reference_docs:
            user_prompt += f"\n\nReference materials:\n{reference_docs[:1500]}"

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.7)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "title": slide_title,
                "content": {"main_text": context}
            }

    def generate_table_data(self, description: str, num_rows: int = 5,
                           num_cols: int = 3) -> Dict[str, Any]:
        """
        Generate table data based on description.

        Args:
            description: Description of what the table should contain
            num_rows: Number of rows
            num_cols: Number of columns

        Returns:
            Dictionary with table structure and data
        """
        system_prompt = """Generate table data as JSON with this structure:
{
  "headers": ["Column 1", "Column 2", "Column 3"],
  "rows": [
    ["Data 1", "Data 2", "Data 3"],
    ["Data 4", "Data 5", "Data 6"]
  ]
}
"""

        user_prompt = f"""Create a table with {num_rows} rows and {num_cols} columns.

Description: {description}

Generate realistic data that fits the description."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.7)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "headers": [f"Column {i+1}" for i in range(num_cols)],
                "rows": [[f"Data {i},{j}" for j in range(num_cols)]
                        for i in range(num_rows)]
            }

    def generate_chart_data(self, description: str,
                           chart_type: str = "bar") -> Dict[str, Any]:
        """
        Generate chart data based on description.

        Args:
            description: Description of what the chart should show
            chart_type: Type of chart (bar, line, pie)

        Returns:
            Dictionary with chart data and configuration
        """
        system_prompt = """Generate chart data as JSON with this structure:
{
  "chart_type": "bar|line|pie",
  "title": "Chart Title",
  "categories": ["Category 1", "Category 2", "Category 3"],
  "series": [
    {
      "name": "Series 1",
      "values": [10, 20, 30]
    }
  ]
}
"""

        user_prompt = f"""Create {chart_type} chart data.

Description: {description}

Generate realistic data with 3-6 categories."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        response = self.client.generate_json(messages, temperature=0.7)

        try:
            return json.loads(response)
        except json.JSONDecodeError:
            return {
                "chart_type": chart_type,
                "title": description,
                "categories": ["A", "B", "C"],
                "series": [{"name": "Series 1", "values": [10, 20, 30]}]
            }

    def refine_content(self, current_content: str, feedback: str) -> str:
        """
        Refine content based on user feedback.

        Args:
            current_content: Current content to refine
            feedback: User feedback on what to change

        Returns:
            Refined content
        """
        system_prompt = """You are refining presentation content based on user feedback.
Improve the content while maintaining its structure and purpose."""

        user_prompt = f"""Current content:
{current_content}

User feedback:
{feedback}

Please provide the refined version."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        return self.client.chat_completion(messages, temperature=0.7)

    def suggest_improvements(self, presentation_outline: Dict[str, Any]) -> List[str]:
        """
        Suggest improvements for a presentation outline.

        Args:
            presentation_outline: Current presentation outline

        Returns:
            List of improvement suggestions
        """
        system_prompt = """You are a presentation expert. Analyze the outline and suggest
specific improvements. Return suggestions as a JSON array of strings."""

        user_prompt = f"""Analyze this presentation outline and suggest improvements:

{json.dumps(presentation_outline, indent=2)}

Provide 3-5 specific, actionable suggestions."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        try:
            response = self.client.generate_json(messages, temperature=0.7)
            suggestions = json.loads(response)
            if isinstance(suggestions, dict) and 'suggestions' in suggestions:
                return suggestions['suggestions']
            elif isinstance(suggestions, list):
                return suggestions
            return ["Consider the overall flow and narrative of your presentation"]
        except (json.JSONDecodeError, KeyError):
            return ["Consider the overall flow and narrative of your presentation"]
