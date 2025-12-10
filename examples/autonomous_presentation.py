"""
Example of using the fully autonomous presentation builder.

The autonomous builder makes all design decisions including:
- Layout selection
- Color schemes
- Font sizes and styling
- Content optimization and fitting
- Visual hierarchy
"""

from pathlib import Path
from pptx_agent.core.autonomous_builder import AutonomousPresentationBuilder


def main():
    print("Creating presentation with full autonomy...\n")

    # Initialize autonomous builder
    builder = AutonomousPresentationBuilder(
        target_audience="professional"
    )

    # Create presentation - the AI makes ALL design decisions
    report = builder.create_presentation_autonomously(
        topic="The Future of Artificial Intelligence",
        summary="""
        An comprehensive exploration of AI's trajectory and impact:
        - Current state of AI technology and capabilities
        - Breakthrough innovations in machine learning and neural networks
        - Real-world applications across industries (healthcare, finance, education)
        - Ethical considerations and responsible AI development
        - Future predictions and potential societal changes
        - Challenges and opportunities ahead
        - How businesses can prepare for AI transformation
        """,
        num_slides=12
    )

    # Display what the AI decided
    print("\n" + "="*60)
    print("AUTONOMOUS DESIGN REPORT")
    print("="*60)

    print(f"\nPresentation: {report['topic']}")
    print(f"Slides Created: {report['slides_created']}")

    print("\n--- Design Decisions Made by AI ---")
    for decision in report['decisions_made']:
        print(f"\n{decision['decision'].upper()}:")
        print(f"  Choice: {decision.get('choice', 'N/A')}")
        if 'slide' in decision:
            print(f"  Slide: {decision['slide']}")
        print(f"  Reasoning: {decision['reasoning']}")

    if report['optimizations_performed']:
        print("\n--- Autonomous Optimizations ---")
        for opt in report['optimizations_performed']:
            print(f"  • {opt}")

    # Save
    output_path = Path("autonomous_ai_future.pptx")
    builder.save(output_path)

    print(f"\n{'='*60}")
    print(f"Presentation saved to: {output_path}")
    print(f"Total slides: {builder.get_slide_count()}")
    print(f"{'='*60}\n")


if __name__ == '__main__':
    main()
