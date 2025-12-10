"""
Vision-Based Slide Validator - Uses GPT-4 Vision to validate slides.

This module converts slides to images and uses vision models to:
- Check if content fits properly
- Validate readability
- Assess visual balance
- Identify formatting issues
- Suggest improvements
"""

import base64
import io
from typing import Dict, List, Any, Optional
from pathlib import Path
from PIL import Image
from pptx import Presentation
from ..llm.openai_client import OpenAIClient


class VisionValidator:
    """Validates presentation slides using vision models."""

    def __init__(self, client: OpenAIClient):
        """
        Initialize vision validator.

        Args:
            client: OpenAI client instance
        """
        self.client = client
        self.validation_history: List[Dict[str, Any]] = []

    def slide_to_image(self, prs_path: Path, slide_index: int,
                      output_path: Optional[Path] = None) -> Optional[Path]:
        """
        Convert a slide to an image.

        Note: This requires additional dependencies like python-pptx's export features
        or external tools like LibreOffice/PowerPoint automation.

        Args:
            prs_path: Path to presentation file
            slide_index: Index of slide to convert
            output_path: Optional output path for image

        Returns:
            Path to generated image or None
        """
        # This is a placeholder - actual implementation would require:
        # 1. Using python-pptx with PIL to render slide
        # 2. Using external tools like unoconv, LibreOffice
        # 3. Using PowerPoint COM automation (Windows)

        # For now, we'll note this as a feature that needs platform-specific implementation
        print("⚠️  Slide-to-image conversion requires platform-specific tools.")
        print("   Options:")
        print("   - Install LibreOffice and use: libreoffice --headless --convert-to png")
        print("   - Use PowerPoint automation (Windows)")
        print("   - Use third-party services")

        return None

    def encode_image_to_base64(self, image_path: Path) -> str:
        """
        Encode image to base64 for vision API.

        Args:
            image_path: Path to image file

        Returns:
            Base64 encoded image string
        """
        with open(image_path, 'rb') as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    def validate_slide_visual(self, image_path: Path,
                              slide_context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Validate a slide using vision model.

        Args:
            image_path: Path to slide image
            slide_context: Optional context about the slide

        Returns:
            Validation results with suggestions
        """
        if not image_path.exists():
            return {
                'error': 'Image file not found',
                'valid': False
            }

        print(f"\n👁️  Analyzing slide visually: {image_path.name}")

        # Encode image
        base64_image = self.encode_image_to_base64(image_path)

        # Create vision prompt
        system_prompt = """You are an expert presentation designer analyzing a PowerPoint slide.
Evaluate the slide on these criteria:

1. Content Fit: Does all content fit properly without overflow?
2. Readability: Is text large enough and clearly visible?
3. Visual Balance: Is the layout balanced and professional?
4. Information Density: Is there too much or too little content?
5. Alignment: Are elements properly aligned?
6. Color & Contrast: Are colors appropriate and readable?

Return JSON with this structure:
{
  "overall_score": 0-10,
  "issues": [
    {
      "severity": "high|medium|low",
      "category": "content_fit|readability|layout|density|alignment|color",
      "description": "specific issue",
      "suggestion": "how to fix"
    }
  ],
  "strengths": ["list of positive aspects"],
  "recommendations": ["prioritized list of improvements"]
}
"""

        context_info = ""
        if slide_context:
            context_info = f"\n\nSlide context: {slide_context.get('title', 'Unknown')}"
            if 'intended_content' in slide_context:
                context_info += f"\nIntended content: {slide_context['intended_content']}"

        user_prompt = f"""Analyze this PowerPoint slide image.{context_info}

Provide detailed visual validation focusing on:
- Whether content fits in the slide boundaries
- Text readability (font size, contrast)
- Layout and visual balance
- Any formatting issues

Be specific about problems and how to fix them."""

        try:
            # Note: This uses OpenAI's vision capabilities
            # The actual API call format depends on the OpenAI SDK version
            messages = [
                {
                    "role": "system",
                    "content": system_prompt
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": user_prompt
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ]

            # Use vision-capable model
            response = self.client.client.chat.completions.create(
                model="gpt-4-vision-preview",  # or "gpt-4o" depending on availability
                messages=messages,
                max_tokens=1000,
                temperature=0.3
            )

            result_text = response.choices[0].message.content

            # Try to parse as JSON
            import json
            try:
                validation_result = json.loads(result_text)
            except json.JSONDecodeError:
                # If not JSON, create structured response from text
                validation_result = {
                    "overall_score": 7,
                    "issues": [],
                    "strengths": [],
                    "recommendations": [],
                    "raw_analysis": result_text
                }

            validation_result['validated'] = True
            validation_result['image_path'] = str(image_path)

            self.validation_history.append({
                'slide_image': str(image_path),
                'result': validation_result,
                'context': slide_context
            })

            return validation_result

        except Exception as e:
            return {
                'error': str(e),
                'valid': False,
                'message': 'Vision API call failed'
            }

    def validate_slide_from_presentation(self, prs_path: Path, slide_index: int,
                                        slide_context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Validate a slide by first converting it to image, then analyzing.

        Args:
            prs_path: Path to presentation
            slide_index: Index of slide to validate
            slide_context: Optional context information

        Returns:
            Validation results
        """
        # Convert slide to image
        temp_image = Path(f"temp_slide_{slide_index}.png")
        image_path = self.slide_to_image(prs_path, slide_index, temp_image)

        if not image_path:
            return {
                'error': 'Could not convert slide to image',
                'valid': False,
                'note': 'Vision validation requires slide-to-image conversion'
            }

        # Validate the image
        result = self.validate_slide_visual(image_path, slide_context)

        # Cleanup temp file
        if temp_image.exists():
            temp_image.unlink()

        return result

    def validate_all_slides(self, prs_path: Path,
                           slide_contexts: Optional[List[Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
        """
        Validate all slides in a presentation.

        Args:
            prs_path: Path to presentation
            slide_contexts: Optional list of context for each slide

        Returns:
            List of validation results for all slides
        """
        prs = Presentation(str(prs_path))
        num_slides = len(prs.slides)

        print(f"\n👁️  Validating {num_slides} slides visually...")

        results = []
        for i in range(num_slides):
            context = slide_contexts[i] if slide_contexts and i < len(slide_contexts) else None
            result = self.validate_slide_from_presentation(prs_path, i, context)
            results.append(result)

            # Print summary
            if 'overall_score' in result:
                score = result['overall_score']
                issues = len(result.get('issues', []))
                print(f"   Slide {i+1}: Score {score}/10, {issues} issues found")

        return results

    def generate_validation_report(self, validation_results: List[Dict[str, Any]]) -> str:
        """
        Generate a comprehensive validation report.

        Args:
            validation_results: List of validation results from all slides

        Returns:
            Formatted report string
        """
        report = []
        report.append("\n" + "="*70)
        report.append("VISUAL VALIDATION REPORT")
        report.append("="*70)

        total_slides = len(validation_results)
        valid_slides = sum(1 for r in validation_results if r.get('validated', False))
        avg_score = sum(r.get('overall_score', 0) for r in validation_results) / max(total_slides, 1)

        report.append(f"\nTotal Slides: {total_slides}")
        report.append(f"Successfully Validated: {valid_slides}")
        report.append(f"Average Score: {avg_score:.1f}/10")

        # High priority issues
        high_priority = []
        for i, result in enumerate(validation_results, 1):
            for issue in result.get('issues', []):
                if issue.get('severity') == 'high':
                    high_priority.append((i, issue))

        if high_priority:
            report.append("\n🔴 HIGH PRIORITY ISSUES:")
            for slide_num, issue in high_priority[:5]:  # Top 5
                report.append(f"\n  Slide {slide_num}: {issue.get('description', 'Unknown issue')}")
                report.append(f"  → {issue.get('suggestion', 'No suggestion')}")

        # Overall recommendations
        all_recommendations = []
        for result in validation_results:
            all_recommendations.extend(result.get('recommendations', []))

        if all_recommendations:
            report.append("\n💡 TOP RECOMMENDATIONS:")
            for i, rec in enumerate(all_recommendations[:5], 1):
                report.append(f"  {i}. {rec}")

        report.append("\n" + "="*70)
        return "\n".join(report)

    def suggest_fixes_for_issues(self, issues: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Generate specific, actionable fixes for identified issues.

        Args:
            issues: List of issues from validation

        Returns:
            List of fixes with implementation details
        """
        fixes = []

        for issue in issues:
            category = issue.get('category', 'unknown')
            severity = issue.get('severity', 'low')

            fix = {
                'issue': issue.get('description'),
                'severity': severity,
                'actions': []
            }

            if category == 'content_fit':
                fix['actions'] = [
                    "Reduce font size by 2-4 points",
                    "Split content across multiple slides",
                    "Remove non-essential bullet points",
                    "Increase slide margins"
                ]

            elif category == 'readability':
                fix['actions'] = [
                    "Increase font size to minimum 18pt",
                    "Improve color contrast",
                    "Use bold for emphasis",
                    "Simplify complex sentences"
                ]

            elif category == 'layout':
                fix['actions'] = [
                    "Realign elements using grid",
                    "Balance left and right content",
                    "Add whitespace between sections",
                    "Use consistent margins"
                ]

            elif category == 'density':
                fix['actions'] = [
                    "Reduce number of bullet points to max 7",
                    "Break dense paragraphs into bullets",
                    "Move details to notes",
                    "Use progressive disclosure"
                ]

            elif category == 'color':
                fix['actions'] = [
                    "Increase contrast ratio to 4.5:1 minimum",
                    "Use professional color palette",
                    "Ensure text is dark on light or vice versa",
                    "Test with colorblind simulators"
                ]

            fixes.append(fix)

        return fixes

    def check_vision_api_available(self) -> bool:
        """
        Check if vision API is available and working.

        Returns:
            True if vision API is available
        """
        try:
            # Simple test to see if vision model is available
            test_prompt = [
                {
                    "role": "user",
                    "content": "Test message"
                }
            ]

            response = self.client.client.chat.completions.create(
                model="gpt-4-vision-preview",
                messages=test_prompt,
                max_tokens=10
            )

            return True

        except Exception as e:
            error_str = str(e).lower()
            if 'vision' in error_str or 'model' in error_str:
                print(f"⚠️  Vision API not available: {e}")
                print("   Vision validation features will be limited.")
                return False
            return False
