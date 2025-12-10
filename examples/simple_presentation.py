"""
Simple example of creating a presentation programmatically.
"""

from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder


def main():
    # Initialize builder (uses OPENAI_API_KEY from environment)
    builder = PresentationBuilder()

    # Create outline
    print("Creating outline...")
    outline = builder.create_outline(
        topic="Introduction to Python",
        summary="""
        A beginner-friendly introduction to Python programming covering:
        - Basic syntax and data types
        - Control structures
        - Functions and modules
        - Object-oriented programming
        - Common libraries and frameworks
        """,
        num_slides=8
    )

    print(f"Outline created with {len(outline['slides'])} slides")

    # Build from outline
    print("Building presentation...")
    builder.build_from_outline()

    # Add a few custom slides
    builder.add_bullet_slide(
        "Why Python?",
        [
            "Easy to learn and read",
            "Versatile - web, data science, automation, AI",
            "Large ecosystem of libraries",
            "Strong community support",
            "High demand in job market"
        ]
    )

    builder.add_table_slide(
        "Python vs Other Languages",
        headers=["Feature", "Python", "Java", "JavaScript"],
        rows=[
            ["Learning Curve", "Easy", "Moderate", "Easy"],
            ["Performance", "Good", "Excellent", "Good"],
            ["Use Cases", "Versatile", "Enterprise", "Web"],
            ["Community", "Large", "Large", "Large"]
        ]
    )

    # Save presentation
    output_path = Path("python_intro.pptx")
    builder.save(output_path)
    print(f"Presentation saved to: {output_path}")


if __name__ == '__main__':
    main()
