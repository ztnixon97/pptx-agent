"""
Example demonstrating autonomous validation and optimization.

Shows how the system automatically validates content fits and
optimizes when necessary.
"""

from pathlib import Path
from pptx_agent.core.autonomous_builder import AutonomousPresentationBuilder


def main():
    print("Testing autonomous validation and optimization...\n")

    builder = AutonomousPresentationBuilder()

    # Create a presentation with intentionally long content
    # The autonomous system will detect and optimize automatically
    report = builder.create_presentation_autonomously(
        topic="Comprehensive Product Strategy 2025",
        summary="""
        Detailed strategic overview covering all aspects:

        Market Analysis:
        - Comprehensive market research findings from Q3 and Q4 2024
        - Competitive landscape analysis including all major competitors
        - Customer segmentation and detailed persona development
        - Market trends and future predictions for the next 3-5 years
        - SWOT analysis with detailed breakdowns for each quadrant

        Product Development:
        - Feature roadmap for all product lines through 2025 and beyond
        - Technical architecture decisions and rationale for each choice
        - Development timeline with detailed milestones and dependencies
        - Resource allocation across teams and projects
        - Risk assessment and mitigation strategies for each initiative

        Go-to-Market Strategy:
        - Multi-channel marketing campaign strategies and tactics
        - Sales enablement programs and training initiatives
        - Partnership opportunities and strategic alliances
        - Pricing strategy with detailed tier analysis
        - Launch timeline and coordination across all departments

        Financial Projections:
        - Detailed revenue forecasts by product and region
        - Cost structure analysis and optimization opportunities
        - Investment requirements and ROI projections
        - Break-even analysis for each product line
        - Sensitivity analysis for various market scenarios

        Success Metrics:
        - Key performance indicators for each strategic initiative
        - Measurement methodology and reporting cadence
        - Dashboard development for real-time tracking
        - Quarterly review process and adjustment protocols
        """,
        num_slides=15
    )

    print("\n" + "="*70)
    print("AUTONOMOUS VALIDATION & OPTIMIZATION REPORT")
    print("="*70)

    print(f"\nPresentation: {report['topic']}")
    print(f"Slides Created: {report['slides_created']}")

    if report['optimizations_performed']:
        print("\n--- Automatic Optimizations Applied ---")
        print("The AI detected content issues and autonomously fixed them:\n")
        for i, opt in enumerate(report['optimizations_performed'], 1):
            print(f"{i}. {opt}")

    print("\n--- Design Decisions ---")
    for decision in report['decisions_made'][:5]:  # Show first 5
        print(f"• {decision['decision']}: {decision.get('choice', 'Applied')}")

    # Save
    output_path = Path("autonomous_optimized_strategy.pptx")
    builder.save(output_path)

    print(f"\n{'='*70}")
    print(f"✓ Presentation saved: {output_path}")
    print(f"✓ All content validated and optimized automatically")
    print(f"✓ Total slides: {builder.get_slide_count()}")
    print(f"{'='*70}\n")


if __name__ == '__main__':
    main()
