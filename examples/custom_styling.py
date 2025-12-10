"""
Example of creating a presentation with custom styling.
"""

from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder
from pptx_agent.builders.table_builder import TableSlideBuilder


def main():
    # Initialize builder
    builder = PresentationBuilder()

    # Add title
    builder.handler.add_title_slide(
        "Custom Styled Presentation",
        "Demonstrating Advanced Formatting"
    )

    # Custom styled table with brand colors
    TableSlideBuilder.add_styled_table_slide(
        builder.handler,
        "Company Comparison",
        headers=["Company", "Revenue", "Employees", "Founded"],
        rows=[
            ["Tech Corp", "$50B", "100,000", "1998"],
            ["Innovate Inc", "$35B", "75,000", "2003"],
            ["Digital Solutions", "$28B", "60,000", "2007"]
        ],
        header_color=(0, 112, 192),  # Blue
        alt_row_color=(242, 242, 242),  # Light gray
        layout_index=1
    )

    # Two-column text slide
    from pptx_agent.builders.text_builder import TextSlideBuilder

    TextSlideBuilder.add_two_column_slide(
        builder.handler,
        "Pros and Cons",
        left_content="""Advantages:
• Fast development
• Cost-effective
• Scalable
• Great community support
• Modern technology""",
        right_content="""Challenges:
• Learning curve
• Integration complexity
• Resource requirements
• Security considerations
• Maintenance needs"""
    )

    # Quote slide
    TextSlideBuilder.add_quote_slide(
        builder.handler,
        "The best way to predict the future is to invent it.",
        "Alan Kay",
        layout_index=6
    )

    # Summary table
    TableSlideBuilder.add_summary_table_slide(
        builder.handler,
        "Project Summary",
        labels=["Duration", "Budget", "Team Size", "Technologies", "Status"],
        values=["6 months", "$250,000", "8 people", "Python, React, AWS", "On Track"]
    )

    # Save
    output_path = Path("custom_styled.pptx")
    builder.save(output_path)
    print(f"Presentation saved to: {output_path}")


if __name__ == '__main__':
    main()
