"""
Example of creating a presentation with various charts.
"""

from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder


def main():
    # Initialize builder
    builder = PresentationBuilder()

    # Add title slide
    builder.handler.add_title_slide(
        "Sales Performance Report",
        "Q4 2024 Analysis"
    )

    # Bar chart - Monthly sales
    builder.add_chart_slide(
        "Monthly Sales Performance",
        chart_type="bar",
        categories=["Oct", "Nov", "Dec"],
        series=[
            {"name": "Product A", "values": [45000, 52000, 61000]},
            {"name": "Product B", "values": [38000, 41000, 47000]},
            {"name": "Product C", "values": [29000, 33000, 38000]}
        ]
    )

    # Line chart - Growth trend
    builder.add_chart_slide(
        "Revenue Growth Trend",
        chart_type="line",
        categories=["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        series=[
            {"name": "2023", "values": [100, 105, 110, 108, 115, 120]},
            {"name": "2024", "values": [110, 118, 125, 130, 138, 145]}
        ]
    )

    # Pie chart - Market share
    builder.add_chart_slide(
        "Market Share Distribution",
        chart_type="pie",
        categories=["North America", "Europe", "Asia Pacific", "Other"],
        series=[{"name": "Share", "values": [35, 28, 30, 7]}]
    )

    # Summary table
    builder.add_table_slide(
        "Key Metrics Summary",
        headers=["Metric", "Q3 2024", "Q4 2024", "Change"],
        rows=[
            ["Total Revenue", "$1.2M", "$1.5M", "+25%"],
            ["New Customers", "450", "620", "+38%"],
            ["Avg Deal Size", "$2,667", "$2,419", "-9%"],
            ["Customer Retention", "92%", "94%", "+2%"]
        ]
    )

    # Key takeaways
    builder.add_bullet_slide(
        "Key Takeaways",
        [
            "Strong revenue growth across all product lines",
            "Product A showing exceptional performance",
            "Asia Pacific market growing rapidly",
            "Customer retention at all-time high",
            "Q1 2025 outlook remains positive"
        ]
    )

    # Save
    output_path = Path("sales_report.pptx")
    builder.save(output_path)
    print(f"Presentation saved to: {output_path}")


if __name__ == '__main__':
    main()
