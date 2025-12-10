"""
Full Features Showcase - Demonstrates all PPTX Agent capabilities.

This example shows all available content types:
- Text (bullets, quotes, sections)
- Tables (data, comparisons)
- Charts (bar, line, pie)
- SmartArt (process, cycle, hierarchy, comparison, venn, timeline)
- Shapes (flowcharts, callouts)
- Images (if provided)
"""

from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder


def main():
    print("Creating comprehensive feature showcase presentation...\n")

    # Initialize builder
    builder = PresentationBuilder()

    # ========== TITLE SLIDE ==========
    builder.handler.add_title_slide(
        "PPTX Agent Feature Showcase",
        "Demonstrating All Capabilities"
    )

    # ========== TEXT SLIDES ==========
    print("Adding text slides...")

    # Bullet points
    builder.add_bullet_slide(
        "Text Capabilities",
        [
            "Bullet point lists with custom formatting",
            "Two-column layouts for comparisons",
            "Quote slides with attribution",
            "Section dividers for organization",
            "Fully formatted text with styles"
        ]
    )

    # Section divider
    builder.text_builder.add_section_slide(
        builder.handler,
        "Data Visualization",
        "Tables and Charts"
    )

    # ========== TABLES ==========
    print("Adding tables...")

    builder.add_table_slide(
        "Data Tables",
        headers=["Product", "Q1", "Q2", "Q3", "Q4"],
        rows=[
            ["Product A", "$125K", "$142K", "$158K", "$181K"],
            ["Product B", "$98K", "$105K", "$118K", "$132K"],
            ["Product C", "$76K", "$84K", "$91K", "$103K"]
        ]
    )

    # Comparison table
    from pptx_agent.builders.table_builder import TableSlideBuilder
    TableSlideBuilder.add_comparison_table_slide(
        builder.handler,
        "Feature Comparison",
        {
            "Feature": ["AI-Powered", "Templates", "Charts", "SmartArt", "Vision Check"],
            "PPTX Agent": ["✓", "✓", "✓", "✓", "✓"],
            "Traditional Tools": ["✗", "✓", "✓", "Limited", "✗"]
        }
    )

    # ========== CHARTS ==========
    print("Adding charts...")

    # Bar chart
    builder.add_chart_slide(
        "Revenue by Product Line",
        chart_type="bar",
        categories=["Q1", "Q2", "Q3", "Q4"],
        series=[
            {"name": "Product A", "values": [125, 142, 158, 181]},
            {"name": "Product B", "values": [98, 105, 118, 132]},
            {"name": "Product C", "values": [76, 84, 91, 103]}
        ]
    )

    # Line chart
    builder.add_chart_slide(
        "Growth Trend",
        chart_type="line",
        categories=["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        series=[
            {"name": "2024", "values": [100, 112, 125, 138, 152, 168]},
            {"name": "2023", "values": [85, 92, 98, 105, 112, 120]}
        ]
    )

    # Pie chart
    builder.add_chart_slide(
        "Market Share",
        chart_type="pie",
        categories=["North America", "Europe", "Asia Pacific", "Other"],
        series=[{"name": "Share", "values": [35, 28, 30, 7]}]
    )

    # ========== SMARTART DIAGRAMS ==========
    print("Adding SmartArt diagrams...")

    # Section divider
    builder.text_builder.add_section_slide(
        builder.handler,
        "SmartArt Diagrams",
        "Process Flows, Cycles, Hierarchies, and More"
    )

    # Process flow
    builder.smartart_builder.add_process_flow(
        builder.handler,
        "Project Phases",
        ["Planning", "Design", "Development", "Testing", "Deployment"]
    )

    # Cycle diagram
    builder.smartart_builder.add_cycle_diagram(
        builder.handler,
        "Continuous Improvement Cycle",
        ["Plan", "Do", "Check", "Act"]
    )

    # Hierarchy
    builder.smartart_builder.add_hierarchy_diagram(
        builder.handler,
        "Organization Structure",
        "Executive Team",
        ["Engineering", "Product", "Marketing", "Sales", "Operations"]
    )

    # Comparison
    builder.smartart_builder.add_comparison_diagram(
        builder.handler,
        "Build vs. Buy Analysis",
        left_items=[
            "Full control",
            "Custom features",
            "No licensing costs",
            "Learn and grow"
        ],
        right_items=[
            "Faster deployment",
            "Proven reliability",
            "Support included",
            "Less maintenance"
        ],
        left_label="Build In-House",
        right_label="Buy Solution"
    )

    # Venn diagram
    builder.smartart_builder.add_venn_diagram(
        builder.handler,
        "Skill Set Overlap",
        "Technical Skills",
        "Business Skills",
        left_items=["Coding", "Architecture", "DevOps"],
        right_items=["Strategy", "Finance", "Marketing"],
        overlap_items=["Product", "Analytics", "Leadership"]
    )

    # Timeline
    builder.smartart_builder.add_timeline(
        builder.handler,
        "Product Roadmap 2024",
        [
            {"date": "Q1", "event": "Beta Launch"},
            {"date": "Q2", "event": "Public Release"},
            {"date": "Q3", "event": "Mobile App"},
            {"date": "Q4", "event": "Enterprise Features"}
        ]
    )

    # ========== SHAPES & FLOWCHARTS ==========
    print("Adding shapes and flowcharts...")

    # Section divider
    builder.text_builder.add_section_slide(
        builder.handler,
        "Shapes & Flowcharts",
        "Custom Shapes and Process Diagrams"
    )

    # Flowchart
    builder.shapes_builder.add_flowchart_slide(
        builder.handler,
        "Decision Process",
        [
            {"text": "Start", "decision": False},
            {"text": "Budget Available?", "decision": True},
            {"text": "Evaluate Options", "decision": False},
            {"text": "Best Option?", "decision": True},
            {"text": "Proceed", "decision": False}
        ]
    )

    # Icon grid
    builder.shapes_builder.add_icon_grid_slide(
        builder.handler,
        "Key Features",
        [
            {"shape": "star", "label": "Quality", "color": (255, 215, 0)},
            {"shape": "diamond", "label": "Value", "color": (68, 114, 196)},
            {"shape": "hexagon", "label": "Speed", "color": (112, 173, 71)},
            {"shape": "circle", "label": "Support", "color": (237, 125, 49)},
            {"shape": "pentagon", "label": "Innovation", "color": (165, 165, 165)},
            {"shape": "octagon", "label": "Security", "color": (255, 0, 0)}
        ],
        cols=3
    )

    # ========== SUMMARY ==========
    print("Adding summary...")

    builder.add_bullet_slide(
        "PPTX Agent Capabilities Summary",
        [
            "✓ Rich text formatting and layouts",
            "✓ Data tables with custom styling",
            "✓ Multiple chart types (bar, line, pie, and more)",
            "✓ SmartArt-like diagrams (7 types)",
            "✓ Custom shapes and flowcharts",
            "✓ Vision-based slide validation",
            "✓ Iterative refinement workflow"
        ]
    )

    # Final slide
    builder.text_builder.add_section_slide(
        builder.handler,
        "Thank You!",
        "Questions?"
    )

    # ========== SAVE ==========
    output_path = Path("feature_showcase.pptx")
    builder.save(output_path)

    print(f"\n✅ Showcase presentation created!")
    print(f"📁 Saved to: {output_path}")
    print(f"📊 Total slides: {builder.get_slide_count()}")
    print(f"\nFeatures demonstrated:")
    print(f"  • Text slides (bullets, sections, quotes)")
    print(f"  • Tables (data, comparisons)")
    print(f"  • Charts (bar, line, pie)")
    print(f"  • SmartArt (process, cycle, hierarchy, comparison, venn, timeline)")
    print(f"  • Shapes (flowcharts, icons)")
    print(f"\nOpen the file to see all capabilities in action!")


if __name__ == '__main__':
    main()
