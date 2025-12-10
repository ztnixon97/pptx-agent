"""
Collaborative CLI - Interactive interface for iterative presentation building.

This provides a conversational, iterative workflow where the user and AI
collaborate to refine the presentation through multiple feedback cycles.
"""

from pathlib import Path
from typing import Optional
from colorama import Fore, Style, init

from ..core.presentation_builder import PresentationBuilder
from ..core.iterative_workflow import IterativeWorkflow
from ..llm.openai_client import OpenAIClient
from ..llm.vision_validator import VisionValidator

init(autoreset=True)


class CollaborativeCLI:
    """
    Interactive collaborative interface for building presentations.

    Workflow:
    1. Gather requirements from user
    2. Create and refine outline iteratively
    3. Build first draft
    4. Optional: Visual validation with AI
    5. Iterative refinement of individual slides
    6. Final output
    """

    def __init__(self, template_path: Optional[Path] = None,
                 api_key: Optional[str] = None):
        """
        Initialize collaborative CLI.

        Args:
            template_path: Optional PowerPoint template
            api_key: Optional OpenAI API key
        """
        self.client = OpenAIClient(api_key=api_key)
        self.builder = PresentationBuilder(
            template_path=template_path,
            openai_api_key=api_key
        )
        self.workflow = IterativeWorkflow(self.builder, self.client)
        self.vision_validator = VisionValidator(self.client)

        self.topic: Optional[str] = None
        self.summary: Optional[str] = None
        self.num_slides: Optional[int] = None
        self.reference_docs: Optional[str] = None
        self.image_dir: Optional[Path] = None

    def print_header(self):
        """Print welcome header."""
        print(f"\n{Fore.CYAN}{'='*70}")
        print(f"{Fore.CYAN}  PPTX Agent - Collaborative Presentation Builder")
        print(f"{Fore.CYAN}  AI-Powered Iterative Workflow")
        print(f"{Fore.CYAN}{'='*70}{Style.RESET_ALL}\n")

    def print_section(self, title: str):
        """Print section header."""
        print(f"\n{Fore.YELLOW}{'─'*70}")
        print(f"{Fore.YELLOW}{title}")
        print(f"{Fore.YELLOW}{'─'*70}{Style.RESET_ALL}\n")

    def gather_requirements(self):
        """Phase 0: Gather initial requirements from user."""
        self.print_section("📋 Step 1: Gathering Requirements")

        print("Let's start by understanding what you want to create.\n")

        # Topic
        self.topic = input(f"{Fore.GREEN}What is your presentation about?{Style.RESET_ALL}\n→ ").strip()

        # Summary/Content
        print(f"\n{Fore.GREEN}Describe the key points or content to cover:{Style.RESET_ALL}")
        print("(You can enter multiple lines. Press Enter twice when done)\n")

        lines = []
        empty_count = 0
        while empty_count < 2:
            line = input()
            if not line.strip():
                empty_count += 1
            else:
                empty_count = 0
                lines.append(line)

        self.summary = "\n".join(lines).strip()

        # Number of slides
        num_input = input(f"\n{Fore.GREEN}Target number of slides (press Enter to let AI decide):{Style.RESET_ALL}\n→ ").strip()
        if num_input.isdigit():
            self.num_slides = int(num_input)

        # Reference documents
        ref_path = input(f"\n{Fore.GREEN}Path to reference document (optional, press Enter to skip):{Style.RESET_ALL}\n→ ").strip()
        if ref_path:
            path = Path(ref_path)
            if path.exists():
                try:
                    self.reference_docs = path.read_text()
                    print(f"{Fore.GREEN}✓ Reference document loaded{Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.YELLOW}⚠️  Could not read file: {e}{Style.RESET_ALL}")

        # Image directory
        img_path = input(f"\n{Fore.GREEN}Path to images directory (optional, press Enter to skip):{Style.RESET_ALL}\n→ ").strip()
        if img_path:
            path = Path(img_path)
            if path.is_dir():
                self.image_dir = path
                print(f"{Fore.GREEN}✓ Image directory set{Style.RESET_ALL}")

        print(f"\n{Fore.CYAN}✓ Requirements gathered!{Style.RESET_ALL}")

    def outline_refinement_phase(self):
        """Phase 1: Create and refine outline with user feedback."""
        self.print_section("📝 Step 2: Planning & Outline Refinement")

        print("The AI will now create an initial outline based on your requirements.")
        print("You'll have the chance to review and refine it.\n")

        input(f"{Fore.CYAN}Press Enter to continue...{Style.RESET_ALL}")

        # Create initial outline
        outline = self.workflow.phase1_initial_planning(
            self.topic,
            self.summary,
            self.num_slides,
            self.reference_docs
        )

        # Iterative refinement loop
        iteration = 0
        max_iterations = 5

        while iteration < max_iterations:
            # Show outline
            print(self.workflow.get_outline_summary(outline))

            # Get feedback
            print(f"\n{Fore.YELLOW}Review the outline above.{Style.RESET_ALL}")
            print(f"\nOptions:")
            print(f"  1. Type '{Fore.GREEN}approve{Style.RESET_ALL}' to proceed with this outline")
            print(f"  2. Provide feedback for refinement")
            print(f"  3. Type '{Fore.RED}restart{Style.RESET_ALL}' to start over\n")

            feedback = input(f"{Fore.CYAN}Your input:{Style.RESET_ALL} ").strip()

            if feedback.lower() in ['approve', 'approved', 'looks good', 'ok', 'proceed', 'yes']:
                print(f"\n{Fore.GREEN}✓ Outline approved!{Style.RESET_ALL}")
                break

            elif feedback.lower() == 'restart':
                print(f"\n{Fore.YELLOW}Restarting outline creation...{Style.RESET_ALL}")
                outline = self.workflow.phase1_initial_planning(
                    self.topic,
                    self.summary,
                    self.num_slides,
                    self.reference_docs
                )
                iteration = 0
                continue

            elif feedback:
                # Refine based on feedback
                print(f"\n{Fore.CYAN}Refining outline based on your feedback...{Style.RESET_ALL}")
                outline = self.workflow.refine_outline_with_feedback(feedback)
                iteration += 1
                print(f"\n{Fore.GREEN}✓ Outline updated (iteration {iteration})!{Style.RESET_ALL}")

            else:
                print(f"{Fore.YELLOW}Please provide feedback or type 'approve'{Style.RESET_ALL}")

        if iteration >= max_iterations:
            print(f"\n{Fore.YELLOW}⚠️  Maximum refinement iterations reached.{Style.RESET_ALL}")
            confirm = input(f"Proceed with current outline? (yes/no): ").strip()
            if confirm.lower() != 'yes':
                print(f"{Fore.RED}Exiting...{Style.RESET_ALL}")
                return False

        return True

    def build_draft_phase(self):
        """Phase 2: Build first draft from approved outline."""
        self.print_section("🔨 Step 3: Building First Draft")

        print("The AI will now create the first draft of your presentation.")
        print("This may take a moment...\n")

        slide_count = self.workflow.phase2_build_draft(self.image_dir)

        print(f"\n{Fore.GREEN}✓ First draft complete!{Style.RESET_ALL}")
        print(f"   {slide_count} slides created")

        return True

    def visual_validation_phase(self, temp_save_path: Path) -> bool:
        """
        Phase 2.5: Optional visual validation using AI vision.

        Args:
            temp_save_path: Temporary path to save presentation for validation

        Returns:
            True if validation completed or skipped
        """
        self.print_section("👁️  Step 4: Visual Validation (Optional)")

        print("The AI can visually analyze your slides to check:")
        print("  • Content fits properly")
        print("  • Text is readable")
        print("  • Layout is balanced")
        print("  • Colors and formatting look professional\n")

        print(f"{Fore.YELLOW}Note: Visual validation requires slide-to-image conversion{Style.RESET_ALL}")
        print("      and GPT-4 Vision API access.\n")

        choice = input(f"Run visual validation? (yes/no): ").strip().lower()

        if choice != 'yes':
            print(f"{Fore.CYAN}Skipping visual validation{Style.RESET_ALL}")
            return True

        # Check if vision API is available
        if not self.vision_validator.check_vision_api_available():
            print(f"{Fore.YELLOW}⚠️  Vision API not available. Skipping validation.{Style.RESET_ALL}")
            return True

        # Save presentation temporarily
        self.builder.save(temp_save_path)

        # Validate slides
        print(f"\n{Fore.CYAN}Analyzing slides visually...{Style.RESET_ALL}")

        # Create slide contexts
        slide_contexts = [
            {
                'title': slide.get('title', 'Untitled'),
                'intended_content': slide.get('content', '')
            }
            for slide in self.workflow.current_outline.get('slides', [])
        ]

        results = self.vision_validator.validate_all_slides(
            temp_save_path,
            slide_contexts
        )

        # Show report
        report = self.vision_validator.generate_validation_report(results)
        print(report)

        # Collect high-priority issues
        high_priority_issues = []
        for i, result in enumerate(results):
            for issue in result.get('issues', []):
                if issue.get('severity') == 'high':
                    high_priority_issues.append((i, issue))

        if high_priority_issues:
            print(f"\n{Fore.RED}⚠️  Found {len(high_priority_issues)} high-priority issues{Style.RESET_ALL}")
            choice = input(f"\nView suggested fixes? (yes/no): ").strip().lower()

            if choice == 'yes':
                all_issues = [issue for _, issue in high_priority_issues]
                fixes = self.vision_validator.suggest_fixes_for_issues(all_issues)

                print(f"\n{Fore.CYAN}Suggested Fixes:{Style.RESET_ALL}\n")
                for i, fix in enumerate(fixes[:3], 1):  # Show top 3
                    print(f"{i}. {Fore.YELLOW}{fix['issue']}{Style.RESET_ALL}")
                    print(f"   Actions:")
                    for action in fix['actions'][:3]:
                        print(f"   • {action}")
                    print()

        return True

    def refinement_phase(self, output_path: Path):
        """Phase 3: Iterative slide-by-slide refinement."""
        self.print_section("🎨 Step 5: Refinement")

        print("Now you can review and refine individual slides.")
        print("The AI will work with you to make adjustments.\n")

        self.workflow.phase3_iterative_refinement(output_path)

    def run(self, output_path: Path):
        """
        Run the complete collaborative workflow.

        Args:
            output_path: Where to save the final presentation
        """
        self.print_header()

        print("This collaborative workflow guides you through:")
        print("  1. Gathering requirements")
        print("  2. Planning and refining the outline")
        print("  3. Building the first draft")
        print("  4. Visual validation (optional)")
        print("  5. Iterative refinement")
        print("\nLet's get started!\n")

        input(f"{Fore.CYAN}Press Enter to begin...{Style.RESET_ALL}")

        # Phase 0: Requirements
        self.gather_requirements()

        # Phase 1: Outline refinement
        if not self.outline_refinement_phase():
            return

        # Phase 2: Build draft
        if not self.build_draft_phase():
            return

        # Phase 2.5: Visual validation (optional)
        temp_path = output_path.parent / f"temp_{output_path.name}"
        self.visual_validation_phase(temp_path)

        # Phase 3: Refinement
        self.refinement_phase(output_path)

        # Final summary
        self.print_section("✅ Completion")
        print(f"Your presentation is ready!")
        print(f"\n📁 Saved to: {Fore.GREEN}{output_path}{Style.RESET_ALL}")
        print(f"📊 Total slides: {Fore.GREEN}{self.builder.get_slide_count()}{Style.RESET_ALL}")
        print(f"🔄 Refinement iterations: {Fore.GREEN}{self.workflow.refinement_iteration}{Style.RESET_ALL}")
        print(f"\n{Fore.CYAN}Thank you for using PPTX Agent!{Style.RESET_ALL}\n")


def run_collaborative_mode(template_path: Optional[Path] = None,
                          api_key: Optional[str] = None,
                          output_path: Optional[Path] = None):
    """
    Run collaborative mode.

    Args:
        template_path: Optional template file
        api_key: Optional API key
        output_path: Optional output path
    """
    if output_path is None:
        output_path = Path("collaborative_presentation.pptx")

    cli = CollaborativeCLI(template_path, api_key)
    cli.run(output_path)
