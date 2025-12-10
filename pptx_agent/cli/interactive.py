"""
Interactive CLI - Provides an interactive interface for building presentations.
"""

import json
from pathlib import Path
from typing import Optional
from colorama import Fore, Style, init

from ..core.presentation_builder import PresentationBuilder

init(autoreset=True)


class InteractiveCLI:
    """Interactive command-line interface for building presentations."""

    def __init__(self, builder: PresentationBuilder):
        """
        Initialize the interactive CLI.

        Args:
            builder: PresentationBuilder instance
        """
        self.builder = builder
        self.running = True

    def print_header(self):
        """Print the application header."""
        print(f"\n{Fore.CYAN}{'='*60}")
        print(f"{Fore.CYAN}  PPTX Agent - AI-Powered PowerPoint Builder")
        print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}\n")

    def print_menu(self):
        """Print the main menu."""
        print(f"\n{Fore.YELLOW}Main Menu:{Style.RESET_ALL}")
        print("  1. Create new presentation outline")
        print("  2. View current outline")
        print("  3. Build presentation from outline")
        print("  4. Add manual slide")
        print("  5. Get improvement suggestions")
        print("  6. Save presentation")
        print("  7. Template information")
        print("  8. Exit")
        print()

    def get_input(self, prompt: str, default: str = "") -> str:
        """Get user input with optional default."""
        if default:
            user_input = input(f"{prompt} [{default}]: ").strip()
            return user_input if user_input else default
        return input(f"{prompt}: ").strip()

    def create_outline(self):
        """Interactive outline creation."""
        print(f"\n{Fore.GREEN}Create Presentation Outline{Style.RESET_ALL}")
        print("-" * 40)

        topic = self.get_input("Presentation topic")
        if not topic:
            print(f"{Fore.RED}Topic is required{Style.RESET_ALL}")
            return

        summary = self.get_input("Content summary/key points")
        if not summary:
            print(f"{Fore.RED}Summary is required{Style.RESET_ALL}")
            return

        num_slides_str = self.get_input("Target number of slides (optional)")
        num_slides = int(num_slides_str) if num_slides_str.isdigit() else None

        ref_docs_path = self.get_input("Path to reference documents (optional)")
        ref_docs = None
        if ref_docs_path:
            path = Path(ref_docs_path)
            if path.exists():
                try:
                    ref_docs = path.read_text()
                except Exception as e:
                    print(f"{Fore.YELLOW}Warning: Could not read reference docs: {e}{Style.RESET_ALL}")

        print(f"\n{Fore.CYAN}Generating outline...{Style.RESET_ALL}")

        try:
            outline = self.builder.create_outline(topic, summary, num_slides, ref_docs)
            print(f"{Fore.GREEN}Outline created successfully!{Style.RESET_ALL}")
            self._display_outline(outline)
        except Exception as e:
            print(f"{Fore.RED}Error creating outline: {e}{Style.RESET_ALL}")

    def _display_outline(self, outline: dict):
        """Display an outline in a readable format."""
        print(f"\n{Fore.CYAN}Presentation Outline:{Style.RESET_ALL}")
        print(f"Title: {outline.get('title', 'N/A')}")
        print(f"Number of slides: {len(outline.get('slides', []))}")
        print("\nSlides:")

        for slide in outline.get('slides', []):
            slide_num = slide.get('slide_number', '?')
            slide_type = slide.get('slide_type', 'unknown')
            title = slide.get('title', 'Untitled')
            print(f"  {slide_num}. [{slide_type.upper()}] {title}")

    def view_outline(self):
        """View the current outline."""
        if not self.builder.current_outline:
            print(f"{Fore.YELLOW}No outline created yet{Style.RESET_ALL}")
            return

        self._display_outline(self.builder.current_outline)

        # Option to see detailed view
        if self.get_input("\nView detailed outline? (y/n)", "n").lower() == 'y':
            print(json.dumps(self.builder.current_outline, indent=2))

    def build_presentation(self):
        """Build presentation from outline."""
        if not self.builder.current_outline:
            print(f"{Fore.YELLOW}No outline created yet. Create one first.{Style.RESET_ALL}")
            return

        image_dir_str = self.get_input("Path to image directory (optional)")
        image_dir = Path(image_dir_str) if image_dir_str else None

        if image_dir and not image_dir.is_dir():
            print(f"{Fore.YELLOW}Warning: Image directory not found{Style.RESET_ALL}")
            image_dir = None

        print(f"\n{Fore.CYAN}Building presentation...{Style.RESET_ALL}")

        try:
            self.builder.build_from_outline(image_dir=image_dir)
            slide_count = self.builder.get_slide_count()
            print(f"{Fore.GREEN}Presentation built successfully!{Style.RESET_ALL}")
            print(f"Total slides: {slide_count}")
        except Exception as e:
            print(f"{Fore.RED}Error building presentation: {e}{Style.RESET_ALL}")

    def add_manual_slide(self):
        """Add a manual slide."""
        print(f"\n{Fore.GREEN}Add Manual Slide{Style.RESET_ALL}")
        print("-" * 40)
        print("Slide types:")
        print("  1. Text slide")
        print("  2. Bullet points slide")
        print("  3. Table slide")
        print("  4. Chart slide")
        print("  5. Image slide")

        choice = self.get_input("\nSelect slide type (1-5)")

        title = self.get_input("Slide title")

        try:
            if choice == '1':
                content = self.get_input("Content")
                self.builder.add_text_slide(title, content)

            elif choice == '2':
                print("Enter bullet points (one per line, empty line to finish):")
                points = []
                while True:
                    point = input("  - ").strip()
                    if not point:
                        break
                    points.append(point)
                if points:
                    self.builder.add_bullet_slide(title, points)

            elif choice == '3':
                headers_str = self.get_input("Column headers (comma-separated)")
                headers = [h.strip() for h in headers_str.split(',')]

                rows = []
                print("Enter rows (comma-separated values, empty line to finish):")
                while True:
                    row_str = input("  ").strip()
                    if not row_str:
                        break
                    row = [cell.strip() for cell in row_str.split(',')]
                    rows.append(row)

                if headers and rows:
                    self.builder.add_table_slide(title, headers, rows)

            elif choice == '4':
                chart_type = self.get_input("Chart type (bar/line/pie)", "bar")
                categories_str = self.get_input("Categories (comma-separated)")
                categories = [c.strip() for c in categories_str.split(',')]

                series_name = self.get_input("Series name", "Data")
                values_str = self.get_input("Values (comma-separated numbers)")
                values = [float(v.strip()) for v in values_str.split(',') if v.strip()]

                series = [{'name': series_name, 'values': values}]
                self.builder.add_chart_slide(title, chart_type, categories, series)

            elif choice == '5':
                image_path_str = self.get_input("Path to image")
                image_path = Path(image_path_str)
                caption = self.get_input("Caption (optional)")

                if image_path.exists():
                    self.builder.add_image_slide(title, image_path, caption)
                else:
                    print(f"{Fore.RED}Image not found{Style.RESET_ALL}")
                    return

            print(f"{Fore.GREEN}Slide added successfully!{Style.RESET_ALL}")

        except Exception as e:
            print(f"{Fore.RED}Error adding slide: {e}{Style.RESET_ALL}")

    def get_suggestions(self):
        """Get AI-powered improvement suggestions."""
        if not self.builder.current_outline:
            print(f"{Fore.YELLOW}No outline created yet{Style.RESET_ALL}")
            return

        print(f"\n{Fore.CYAN}Getting improvement suggestions...{Style.RESET_ALL}")

        try:
            suggestions = self.builder.suggest_improvements()
            print(f"\n{Fore.GREEN}Suggestions:{Style.RESET_ALL}")
            for i, suggestion in enumerate(suggestions, 1):
                print(f"  {i}. {suggestion}")
        except Exception as e:
            print(f"{Fore.RED}Error getting suggestions: {e}{Style.RESET_ALL}")

    def save_presentation(self):
        """Save the presentation."""
        if self.builder.get_slide_count() == 0:
            print(f"{Fore.YELLOW}No slides to save{Style.RESET_ALL}")
            return

        output_path_str = self.get_input("Output file path", "output.pptx")
        output_path = Path(output_path_str)

        if output_path.suffix.lower() != '.pptx':
            output_path = output_path.with_suffix('.pptx')

        try:
            self.builder.save(output_path)
            print(f"{Fore.GREEN}Presentation saved to: {output_path}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}Error saving presentation: {e}{Style.RESET_ALL}")

    def show_template_info(self):
        """Display template information."""
        info = self.builder.get_template_info()
        print(f"\n{Fore.CYAN}{info}{Style.RESET_ALL}")

    def run(self):
        """Run the interactive CLI."""
        self.print_header()

        while self.running:
            self.print_menu()
            choice = self.get_input("Select an option (1-8)")

            if choice == '1':
                self.create_outline()
            elif choice == '2':
                self.view_outline()
            elif choice == '3':
                self.build_presentation()
            elif choice == '4':
                self.add_manual_slide()
            elif choice == '5':
                self.get_suggestions()
            elif choice == '6':
                self.save_presentation()
            elif choice == '7':
                self.show_template_info()
            elif choice == '8':
                print(f"\n{Fore.CYAN}Thank you for using PPTX Agent!{Style.RESET_ALL}\n")
                self.running = False
            else:
                print(f"{Fore.RED}Invalid option{Style.RESET_ALL}")
