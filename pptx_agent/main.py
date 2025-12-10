"""
Main entry point for PPTX Agent.
"""

import argparse
import sys
from pathlib import Path
from dotenv import load_dotenv

from .core.presentation_builder import PresentationBuilder
from .core.autonomous_builder import AutonomousPresentationBuilder
from .cli.interactive import InteractiveCLI

load_dotenv()


def main():
    """Main application entry point."""
    parser = argparse.ArgumentParser(
        description="PPTX Agent - AI-Powered PowerPoint Builder",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode with default template
  python -m pptx_agent

  # Use a custom template
  python -m pptx_agent --template my_template.pptx

  # Quick presentation creation
  python -m pptx_agent --topic "AI in Healthcare" --summary "Overview of AI applications" --output ai_healthcare.pptx

  # With reference documents
  python -m pptx_agent --topic "Q4 Results" --summary "Financial overview" --reference report.txt --output q4.pptx

  # Autonomous mode - AI makes all design decisions
  python -m pptx_agent --autonomous --topic "Future of AI" --summary "Exploring AI trends" --output ai_future.pptx
        """
    )

    parser.add_argument(
        '--template', '-t',
        type=str,
        help='Path to PowerPoint template file'
    )

    parser.add_argument(
        '--topic',
        type=str,
        help='Presentation topic (for quick mode)'
    )

    parser.add_argument(
        '--summary', '-s',
        type=str,
        help='Content summary or key points (for quick mode)'
    )

    parser.add_argument(
        '--num-slides', '-n',
        type=int,
        help='Target number of slides'
    )

    parser.add_argument(
        '--reference', '-r',
        type=str,
        help='Path to reference document'
    )

    parser.add_argument(
        '--images', '-i',
        type=str,
        help='Path to directory containing images'
    )

    parser.add_argument(
        '--output', '-o',
        type=str,
        default='output.pptx',
        help='Output file path (default: output.pptx)'
    )

    parser.add_argument(
        '--api-key',
        type=str,
        help='OpenAI API key (or set OPENAI_API_KEY env var)'
    )

    parser.add_argument(
        '--interactive',
        action='store_true',
        help='Force interactive mode'
    )

    parser.add_argument(
        '--autonomous',
        action='store_true',
        help='Use fully autonomous mode (AI makes all design decisions)'
    )

    parser.add_argument(
        '--audience',
        type=str,
        default='professional',
        help='Target audience for autonomous mode (professional, technical, general, executive)'
    )

    args = parser.parse_args()

    # Prepare template path
    template_path = Path(args.template) if args.template else None
    if template_path and not template_path.exists():
        print(f"Error: Template file not found: {template_path}")
        sys.exit(1)

    # Initialize builder
    try:
        builder = PresentationBuilder(
            template_path=template_path,
            openai_api_key=args.api_key
        )
    except ValueError as e:
        print(f"Error: {e}")
        print("\nPlease set your OpenAI API key:")
        print("  - Set OPENAI_API_KEY environment variable, or")
        print("  - Create a .env file with OPENAI_API_KEY=your-key, or")
        print("  - Use --api-key argument")
        sys.exit(1)

    # Autonomous mode vs Quick mode vs Interactive mode
    if args.autonomous and args.topic and args.summary:
        # Autonomous mode - AI makes ALL decisions
        print("Creating presentation in AUTONOMOUS mode...")
        print("AI will make all design, layout, and styling decisions.\n")

        # Initialize autonomous builder
        try:
            autonomous_builder = AutonomousPresentationBuilder(
                template_path=template_path,
                openai_api_key=args.api_key,
                target_audience=args.audience
            )
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)

        # Load reference docs if provided
        ref_docs = None
        if args.reference:
            ref_path = Path(args.reference)
            if ref_path.exists():
                try:
                    ref_docs = ref_path.read_text()
                except Exception as e:
                    print(f"Warning: Could not read reference file: {e}")

        # Create presentation autonomously
        try:
            image_dir = Path(args.images) if args.images else None
            report = autonomous_builder.create_presentation_autonomously(
                args.topic,
                args.summary,
                args.num_slides,
                ref_docs,
                image_dir
            )

            # Display report
            print("\n" + "="*60)
            print("AUTONOMOUS CREATION REPORT")
            print("="*60)
            print(f"Slides created: {report['slides_created']}")
            print(f"Design decisions made: {len(report['decisions_made'])}")
            print(f"Optimizations performed: {len(report['optimizations_performed'])}")

            if report['optimizations_performed']:
                print("\nKey optimizations:")
                for opt in report['optimizations_performed'][:3]:
                    print(f"  • {opt}")

        except Exception as e:
            print(f"Error creating autonomous presentation: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)

        # Save
        print(f"\nSaving to {args.output}...")
        try:
            output_path = Path(args.output)
            autonomous_builder.save(output_path)
            print(f"\n✓ Presentation saved successfully to: {output_path}")
            print(f"✓ Total slides: {autonomous_builder.get_slide_count()}")
            print("="*60)
        except Exception as e:
            print(f"Error saving presentation: {e}")
            sys.exit(1)

    elif args.topic and args.summary and not args.interactive:
        # Quick mode - create presentation directly
        print("Creating presentation in quick mode...")

        # Load reference docs if provided
        ref_docs = None
        if args.reference:
            ref_path = Path(args.reference)
            if ref_path.exists():
                try:
                    ref_docs = ref_path.read_text()
                except Exception as e:
                    print(f"Warning: Could not read reference file: {e}")

        # Create outline
        print("Generating outline...")
        try:
            outline = builder.create_outline(
                args.topic,
                args.summary,
                args.num_slides,
                ref_docs
            )
            print(f"Created outline with {len(outline.get('slides', []))} slides")
        except Exception as e:
            print(f"Error creating outline: {e}")
            sys.exit(1)

        # Build presentation
        print("Building presentation...")
        try:
            image_dir = Path(args.images) if args.images else None
            builder.build_from_outline(image_dir=image_dir)
            print(f"Built presentation with {builder.get_slide_count()} slides")
        except Exception as e:
            print(f"Error building presentation: {e}")
            sys.exit(1)

        # Save
        print(f"Saving to {args.output}...")
        try:
            output_path = Path(args.output)
            builder.save(output_path)
            print(f"Presentation saved successfully to: {output_path}")
        except Exception as e:
            print(f"Error saving presentation: {e}")
            sys.exit(1)

    else:
        # Interactive mode
        cli = InteractiveCLI(builder)
        cli.run()


if __name__ == '__main__':
    main()
