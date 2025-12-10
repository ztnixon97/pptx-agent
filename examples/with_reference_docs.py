"""
Example of creating a presentation with reference documents.
"""

from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder


def main():
    # Prepare reference documents
    # In a real scenario, this would be loaded from actual files
    reference_docs = """
    Project Overview:
    The new mobile app project aims to deliver a seamless shopping experience
    for our customers. Key features include:
    - Personalized product recommendations using ML
    - One-click checkout
    - Real-time order tracking
    - AR product visualization
    - Social sharing capabilities

    Technical Stack:
    - Frontend: React Native
    - Backend: Node.js with Express
    - Database: PostgreSQL with Redis caching
    - ML Pipeline: Python with TensorFlow
    - Cloud: AWS (ECS, RDS, S3)

    Timeline:
    - Phase 1 (Months 1-3): Core shopping features
    - Phase 2 (Months 4-6): ML recommendations
    - Phase 3 (Months 7-9): AR features and social integration

    Team:
    - 3 Mobile developers
    - 2 Backend developers
    - 1 ML engineer
    - 1 UX designer
    - 1 Product manager

    Budget: $500K for development phase
    Expected ROI: 35% increase in mobile conversions
    """

    # Initialize builder
    builder = PresentationBuilder()

    # Create outline with reference docs
    print("Creating outline with reference documents...")
    outline = builder.create_outline(
        topic="Mobile App Development Project",
        summary="Comprehensive project overview including features, technology, timeline, and resources",
        num_slides=10,
        reference_docs=reference_docs
    )

    print(f"Outline created with {len(outline['slides'])} slides")

    # Build presentation
    print("Building presentation...")
    builder.build_from_outline()

    # Save
    output_path = Path("mobile_app_project.pptx")
    builder.save(output_path)
    print(f"Presentation saved to: {output_path}")
    print(f"Total slides: {builder.get_slide_count()}")


if __name__ == '__main__':
    main()
