"""
Example of using the collaborative, iterative workflow.

This demonstrates the full collaborative mode where you work
with the AI to iteratively refine your presentation.
"""

from pathlib import Path
from pptx_agent.cli.collaborative import run_collaborative_mode


def main():
    """
    Run collaborative workflow example.

    In this mode:
    1. You provide requirements (topic, summary, references)
    2. AI creates initial outline
    3. You review and provide feedback
    4. AI refines outline based on feedback
    5. Repeat until you approve the outline
    6. AI builds first draft
    7. Optional: AI visually validates slides
    8. You review and refine individual slides
    9. Final presentation is saved
    """

    print("""
    ╔═══════════════════════════════════════════════════════════════╗
    ║                                                               ║
    ║   COLLABORATIVE PRESENTATION BUILDER                          ║
    ║                                                               ║
    ║   This interactive workflow lets you work WITH the AI         ║
    ║   to create and refine your presentation step-by-step.        ║
    ║                                                               ║
    ╚═══════════════════════════════════════════════════════════════╝
    """)

    # Run collaborative mode
    # You can also provide a custom template
    template_path = None  # Or Path("my_template.pptx")
    output_path = Path("collaborative_example.pptx")

    run_collaborative_mode(
        template_path=template_path,
        output_path=output_path
    )


if __name__ == '__main__':
    main()
