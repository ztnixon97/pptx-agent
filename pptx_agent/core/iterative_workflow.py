"""
Iterative Workflow Manager - Manages collaborative presentation refinement.

This module orchestrates the iterative process of:
1. Planning and outlining with user feedback
2. Building first draft
3. Visual validation using vision models
4. Iterative refinement based on user feedback
"""

from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
import json
from ..llm.openai_client import OpenAIClient
from ..llm.content_planner import ContentPlanner
from .presentation_builder import PresentationBuilder


class IterativeWorkflow:
    """Manages the iterative, collaborative presentation creation workflow."""

    def __init__(self, builder: PresentationBuilder, client: OpenAIClient):
        """
        Initialize the iterative workflow manager.

        Args:
            builder: PresentationBuilder instance
            client: OpenAI client for LLM interactions
        """
        self.builder = builder
        self.client = client
        self.content_planner = ContentPlanner(client)

        self.current_outline: Optional[Dict[str, Any]] = None
        self.outline_history: List[Dict[str, Any]] = []
        self.feedback_history: List[Dict[str, str]] = []
        self.refinement_iteration = 0

    def phase1_initial_planning(self, topic: str, summary: str,
                                num_slides: Optional[int] = None,
                                reference_docs: Optional[str] = None) -> Dict[str, Any]:
        """
        Phase 1: Create initial outline based on user inputs.

        Args:
            topic: Presentation topic
            summary: Content summary
            num_slides: Target number of slides
            reference_docs: Optional reference materials

        Returns:
            Initial outline with high-level structure
        """
        print("\n📝 Phase 1: Planning Presentation Structure\n")
        print("Analyzing your inputs and creating initial outline...")

        self.current_outline = self.content_planner.create_presentation_outline(
            topic, summary, num_slides, reference_docs
        )

        self.outline_history.append({
            'iteration': 0,
            'outline': self.current_outline.copy(),
            'feedback': 'Initial creation'
        })

        return self.current_outline

    def get_outline_summary(self, outline: Optional[Dict[str, Any]] = None) -> str:
        """
        Get a human-readable summary of the outline.

        Args:
            outline: Outline to summarize (uses current if None)

        Returns:
            Formatted outline summary
        """
        if outline is None:
            outline = self.current_outline

        if not outline:
            return "No outline available"

        summary = []
        summary.append(f"\n{'='*70}")
        summary.append(f"PRESENTATION OUTLINE")
        summary.append(f"{'='*70}")
        summary.append(f"\nTitle: {outline.get('title', 'Untitled')}")
        summary.append(f"Total Slides: {len(outline.get('slides', []))}\n")
        summary.append("Slide Structure:")
        summary.append("-" * 70)

        for i, slide in enumerate(outline.get('slides', []), 1):
            slide_type = slide.get('slide_type', 'content').upper()
            title = slide.get('title', 'Untitled')
            content_preview = slide.get('content', '')[:60]
            if len(slide.get('content', '')) > 60:
                content_preview += "..."

            summary.append(f"\n{i}. [{slide_type}] {title}")
            if content_preview:
                summary.append(f"   Preview: {content_preview}")

            elements = slide.get('elements', [])
            if elements:
                summary.append(f"   Elements: {len(elements)} ({', '.join(e.get('type', 'unknown') for e in elements[:3])})")

        summary.append("\n" + "="*70)
        return "\n".join(summary)

    def collect_outline_feedback(self) -> str:
        """
        Collect user feedback on the current outline.

        Returns:
            User feedback as string
        """
        print("\n💭 Please review the outline above.")
        print("\nProvide feedback on:")
        print("  - Slide order and flow")
        print("  - Missing topics or slides")
        print("  - Slides that should be combined or split")
        print("  - Content emphasis or focus areas")
        print("  - Any other structural changes")
        print("\nType your feedback (or 'approve' to proceed to building):")
        print("(Press Enter twice to finish)\n")

        lines = []
        empty_count = 0
        while empty_count < 2:
            line = input()
            if not line.strip():
                empty_count += 1
            else:
                empty_count = 0
                lines.append(line)

        feedback = "\n".join(lines).strip()
        return feedback if feedback else "approve"

    def refine_outline_with_feedback(self, feedback: str) -> Dict[str, Any]:
        """
        Refine the outline based on user feedback.

        Args:
            feedback: User feedback on current outline

        Returns:
            Refined outline
        """
        if feedback.lower() in ['approve', 'approved', 'looks good', 'ok', 'proceed']:
            return self.current_outline

        print("\n🔄 Refining outline based on your feedback...\n")

        self.refinement_iteration += 1
        self.feedback_history.append({
            'iteration': self.refinement_iteration,
            'feedback': feedback
        })

        # Use LLM to refine outline
        system_prompt = """You are an expert presentation designer. The user has provided
feedback on the current presentation outline. Refine the outline to address their feedback
while maintaining a logical flow. Return the complete refined outline in the same JSON format."""

        user_prompt = f"""Current outline:
{json.dumps(self.current_outline, indent=2)}

User feedback:
{feedback}

Please refine the outline to address this feedback. Maintain the same JSON structure."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        try:
            response = self.client.generate_json(messages, temperature=0.7)
            refined_outline = json.loads(response)

            self.current_outline = refined_outline
            self.outline_history.append({
                'iteration': self.refinement_iteration,
                'outline': refined_outline.copy(),
                'feedback': feedback
            })

            return refined_outline

        except json.JSONDecodeError:
            print("⚠️  Error refining outline. Keeping current version.")
            return self.current_outline

    def phase2_build_draft(self, image_dir: Optional[Path] = None) -> int:
        """
        Phase 2: Build first draft of presentation.

        Args:
            image_dir: Optional directory with images

        Returns:
            Number of slides created
        """
        print("\n🔨 Phase 2: Building First Draft\n")
        print("Creating presentation from approved outline...")

        self.builder.build_from_outline(self.current_outline, image_dir)
        slide_count = self.builder.get_slide_count()

        print(f"\n✓ First draft complete: {slide_count} slides created")
        return slide_count

    def get_slide_details(self, slide_index: int) -> Dict[str, Any]:
        """
        Get details about a specific slide.

        Args:
            slide_index: Index of slide to inspect

        Returns:
            Slide details and metadata
        """
        if slide_index >= len(self.builder.handler.prs.slides):
            return {'error': 'Slide index out of range'}

        slide = self.builder.handler.prs.slides[slide_index]

        details = {
            'index': slide_index,
            'has_title': slide.shapes.title is not None,
            'title': slide.shapes.title.text if slide.shapes.title else "No title",
            'shape_count': len(slide.shapes),
            'shapes': []
        }

        for shape in slide.shapes:
            shape_info = {
                'type': shape.shape_type,
                'has_text': hasattr(shape, 'text_frame') and shape.has_text_frame,
            }

            if shape_info['has_text']:
                text = shape.text[:100] + "..." if len(shape.text) > 100 else shape.text
                shape_info['text_preview'] = text

            details['shapes'].append(shape_info)

        return details

    def collect_slide_feedback(self, slide_index: int) -> Dict[str, Any]:
        """
        Collect user feedback on a specific slide.

        Args:
            slide_index: Index of slide to get feedback on

        Returns:
            Structured feedback
        """
        slide_details = self.get_slide_details(slide_index)

        print(f"\n📄 Slide {slide_index + 1}: {slide_details.get('title', 'Untitled')}")
        print(f"   Shapes: {slide_details.get('shape_count', 0)}")

        print("\nWhat would you like to change? (or 'skip' to keep as-is)")
        print("  1. Change content/text")
        print("  2. Change layout")
        print("  3. Add elements (image, chart, etc.)")
        print("  4. Remove elements")
        print("  5. Styling/formatting changes")
        print("  6. Delete this slide")
        print("  7. Other (describe)")

        choice = input("\nYour choice (1-7 or 'skip'): ").strip()

        if choice.lower() == 'skip':
            return {'action': 'skip', 'slide_index': slide_index}

        feedback = {
            'slide_index': slide_index,
            'action': choice,
            'details': input("Describe the changes you want: ").strip()
        }

        return feedback

    def apply_slide_feedback(self, feedback: Dict[str, Any]) -> bool:
        """
        Apply user feedback to a specific slide.

        Args:
            feedback: Structured feedback to apply

        Returns:
            True if successfully applied
        """
        if feedback.get('action') == 'skip':
            return True

        slide_index = feedback.get('slide_index', 0)
        action = feedback.get('action')
        details = feedback.get('details', '')

        print(f"\n🔧 Applying changes to slide {slide_index + 1}...")

        # Use LLM to interpret feedback and suggest changes
        system_prompt = """You are helping refine a PowerPoint slide based on user feedback.
Analyze the feedback and provide specific, actionable instructions for modifying the slide.
Return JSON with the modifications needed."""

        slide_details = self.get_slide_details(slide_index)

        user_prompt = f"""Current slide details:
{json.dumps(slide_details, indent=2)}

User wants to: {action}
Details: {details}

Provide specific instructions for modifying this slide."""

        messages = [
            self.client.create_system_message(system_prompt),
            self.client.create_user_message(user_prompt)
        ]

        try:
            response = self.client.chat_completion(messages, temperature=0.5)
            print(f"AI suggestion: {response[:200]}...")
            print("\n⚠️  Manual slide editing not yet implemented.")
            print("   You can open the .pptx file and make changes manually,")
            print("   or adjust the outline and rebuild.")
            return True

        except Exception as e:
            print(f"Error processing feedback: {e}")
            return False

    def phase3_iterative_refinement(self, output_path: Path) -> None:
        """
        Phase 3: Iteratively refine presentation with user feedback.

        Args:
            output_path: Where to save the presentation
        """
        print("\n🎨 Phase 3: Iterative Refinement\n")
        print("Now you can review and refine individual slides.")

        while True:
            # Save current version
            self.builder.save(output_path)
            print(f"\n💾 Current version saved to: {output_path}")
            print(f"   Total slides: {self.builder.get_slide_count()}")

            print("\nOptions:")
            print("  1. Review and refine specific slide")
            print("  2. Review all slides")
            print("  3. Add new slide")
            print("  4. Regenerate entire presentation from outline")
            print("  5. Finish and save")

            choice = input("\nYour choice (1-5): ").strip()

            if choice == '1':
                slide_num = input(f"Which slide? (1-{self.builder.get_slide_count()}): ")
                try:
                    slide_index = int(slide_num) - 1
                    feedback = self.collect_slide_feedback(slide_index)
                    self.apply_slide_feedback(feedback)
                except (ValueError, IndexError):
                    print("Invalid slide number")

            elif choice == '2':
                for i in range(self.builder.get_slide_count()):
                    feedback = self.collect_slide_feedback(i)
                    if feedback.get('action') != 'skip':
                        self.apply_slide_feedback(feedback)

            elif choice == '3':
                print("\n➕ Adding new slide")
                title = input("Slide title: ")
                content = input("Slide content: ")
                self.builder.add_text_slide(title, content)
                print("✓ Slide added")

            elif choice == '4':
                confirm = input("⚠️  This will rebuild the entire presentation. Continue? (yes/no): ")
                if confirm.lower() == 'yes':
                    # Clear current slides
                    while self.builder.get_slide_count() > 0:
                        self.builder.handler.delete_slide(0)
                    # Rebuild from outline
                    self.phase2_build_draft()

            elif choice == '5':
                self.builder.save(output_path)
                print(f"\n✅ Final presentation saved to: {output_path}")
                break

            else:
                print("Invalid choice")

    def run_complete_workflow(self, topic: str, summary: str,
                             output_path: Path,
                             num_slides: Optional[int] = None,
                             reference_docs: Optional[str] = None,
                             image_dir: Optional[Path] = None) -> None:
        """
        Run the complete iterative workflow from start to finish.

        Args:
            topic: Presentation topic
            summary: Content summary
            output_path: Where to save final presentation
            num_slides: Target number of slides
            reference_docs: Optional reference materials
            image_dir: Optional image directory
        """
        print("\n" + "="*70)
        print("ITERATIVE PRESENTATION BUILDER")
        print("="*70)
        print("\nThis workflow has 3 phases:")
        print("  1. Planning & Outline Refinement")
        print("  2. First Draft Creation")
        print("  3. Iterative Slide-by-Slide Refinement")
        print("\nLet's begin!\n")

        # Phase 1: Planning with iterations
        outline = self.phase1_initial_planning(topic, summary, num_slides, reference_docs)

        while True:
            print(self.get_outline_summary(outline))
            feedback = self.collect_outline_feedback()

            if feedback.lower() in ['approve', 'approved', 'looks good', 'ok', 'proceed']:
                print("\n✓ Outline approved! Moving to draft creation...")
                break

            outline = self.refine_outline_with_feedback(feedback)
            print("\n✓ Outline refined. Please review the updated version:")

        # Phase 2: Build draft
        self.phase2_build_draft(image_dir)

        # Phase 3: Iterative refinement
        self.phase3_iterative_refinement(output_path)

        print("\n" + "="*70)
        print("WORKFLOW COMPLETE")
        print("="*70)
        print(f"\n✅ Your presentation is ready: {output_path}")
        print(f"📊 Final slide count: {self.builder.get_slide_count()}")
        print(f"🔄 Refinement iterations: {self.refinement_iteration}")
        print("\nThank you for using PPTX Agent!")
