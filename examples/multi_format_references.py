"""
Multi-Format Reference Documents Example

This example demonstrates how to use various document formats (.docx, .pptx, .xlsx, .txt)
as reference materials when creating presentations.

The DocumentParser can extract text from:
- Microsoft Word documents (.docx)
- Microsoft PowerPoint presentations (.pptx)
- Microsoft Excel spreadsheets (.xlsx, .xls)
- Plain text files (.txt, .md)
"""

from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder
from pptx_agent.core.document_parser import DocumentParser


def demonstrate_document_parsing():
    """Demonstrate parsing different document formats."""

    print("="*60)
    print("Document Parser Demo - Supported Formats")
    print("="*60)

    # Example 1: Parse a Word document
    print("\n1. Word Document (.docx)")
    print("-" * 40)
    print("Format: Microsoft Word")
    print("Extracts: Paragraphs, tables, headers/footers")
    print("Use case: Policy documents, reports, specifications")
    print("Example: project_requirements.docx")

    # Example 2: Parse a PowerPoint presentation
    print("\n2. PowerPoint Presentation (.pptx)")
    print("-" * 40)
    print("Format: Microsoft PowerPoint")
    print("Extracts: Slide titles, text, tables, speaker notes")
    print("Use case: Previous presentations, templates, content libraries")
    print("Example: quarterly_review_template.pptx")

    # Example 3: Parse an Excel spreadsheet
    print("\n3. Excel Spreadsheet (.xlsx)")
    print("-" * 40)
    print("Format: Microsoft Excel")
    print("Extracts: All sheets, cell values, formatted as tables")
    print("Use case: Financial data, metrics, datasets")
    print("Example: sales_data_q4.xlsx")

    # Example 4: Parse a text file
    print("\n4. Text File (.txt, .md)")
    print("-" * 40)
    print("Format: Plain text / Markdown")
    print("Extracts: Raw text content")
    print("Use case: Notes, documentation, scripts")
    print("Example: meeting_notes.txt")


def create_presentation_from_word_doc():
    """
    Example: Create a presentation using a Word document as reference.

    Typical use case: Convert a project specification document into a presentation.
    """
    print("\n" + "="*60)
    print("Example 1: Creating Presentation from Word Document")
    print("="*60 + "\n")

    # Note: This is a demonstration of the API
    # In practice, you would have actual .docx files

    # Simulate having a Word document with project specifications
    docx_path = Path("project_spec.docx")

    if docx_path.exists():
        # Parse the Word document
        print(f"Parsing: {docx_path.name}")
        content = DocumentParser.parse_file(docx_path)
        summary = DocumentParser.get_document_summary(docx_path)

        print(f"  Format: {summary['extension']}")
        if 'paragraphs' in summary:
            print(f"  Paragraphs: {summary['paragraphs']}")
        if 'tables' in summary:
            print(f"  Tables: {summary['tables']}")

        # Create presentation using the document content
        builder = PresentationBuilder()
        outline = builder.create_outline(
            topic="Project XYZ Specification",
            summary="Overview of project requirements, architecture, and timeline",
            num_slides=8,
            reference_docs=content  # Pass extracted content
        )

        builder.build_from_outline()
        builder.save(Path("project_spec_presentation.pptx"))

        print("\n✓ Presentation created: project_spec_presentation.pptx")
    else:
        print(f"Note: {docx_path} not found. This is a code demonstration.")
        print("Usage: Place a .docx file in the directory and run again.")


def create_presentation_from_excel_data():
    """
    Example: Create a presentation using Excel data as reference.

    Typical use case: Create a financial review presentation from spreadsheet data.
    """
    print("\n" + "="*60)
    print("Example 2: Creating Presentation from Excel Data")
    print("="*60 + "\n")

    xlsx_path = Path("financial_data.xlsx")

    if xlsx_path.exists():
        # Parse the Excel file
        print(f"Parsing: {xlsx_path.name}")
        content = DocumentParser.parse_file(xlsx_path)
        summary = DocumentParser.get_document_summary(xlsx_path)

        print(f"  Format: {summary['extension']}")
        if 'sheets' in summary:
            print(f"  Sheets: {summary['sheets']}")
            print(f"  Sheet names: {', '.join(summary['sheet_names'])}")

        # Create presentation using the data
        builder = PresentationBuilder()
        outline = builder.create_outline(
            topic="Q4 Financial Review",
            summary="Revenue analysis, expense breakdown, and growth metrics",
            num_slides=10,
            reference_docs=content  # AI will extract relevant data
        )

        builder.build_from_outline()
        builder.save(Path("financial_review.pptx"))

        print("\n✓ Presentation created: financial_review.pptx")
    else:
        print(f"Note: {xlsx_path} not found. This is a code demonstration.")
        print("Usage: Place an .xlsx file in the directory and run again.")


def create_presentation_from_multiple_sources():
    """
    Example: Create a presentation using multiple document types.

    Typical use case: Combine data from different sources for a comprehensive report.
    """
    print("\n" + "="*60)
    print("Example 3: Creating Presentation from Multiple Sources")
    print("="*60 + "\n")

    # List of reference documents
    reference_files = [
        Path("executive_summary.docx"),  # Word doc with overview
        Path("metrics.xlsx"),             # Excel with data
        Path("previous_presentation.pptx"), # PowerPoint with history
        Path("notes.txt")                 # Text file with additional notes
    ]

    # Find which files actually exist
    existing_files = [f for f in reference_files if f.exists()]

    if existing_files:
        print(f"Found {len(existing_files)} reference document(s):")
        for file in existing_files:
            print(f"  - {file.name}")

        # Parse all documents
        print("\nParsing documents...")
        combined_content = DocumentParser.parse_multiple_files(existing_files)

        # Create comprehensive presentation
        builder = PresentationBuilder()
        outline = builder.create_outline(
            topic="Comprehensive Business Review",
            summary="Analysis combining executive summary, financial metrics, and historical context",
            num_slides=15,
            reference_docs=combined_content  # Combined content from all sources
        )

        builder.build_from_outline()
        builder.save(Path("comprehensive_review.pptx"))

        print("\n✓ Presentation created: comprehensive_review.pptx")
    else:
        print("Note: No reference documents found. This is a code demonstration.")
        print("\nTo use this feature:")
        print("1. Place reference documents (.docx, .xlsx, .pptx, .txt) in the directory")
        print("2. Update the reference_files list with actual file names")
        print("3. Run the script again")


def parsing_api_examples():
    """Show API examples for document parsing."""
    print("\n" + "="*60)
    print("Document Parser API Examples")
    print("="*60 + "\n")

    print("# Parse a single file")
    print("from pptx_agent.core.document_parser import DocumentParser")
    print("content = DocumentParser.parse_file(Path('document.docx'))\n")

    print("# Get document summary/metadata")
    print("summary = DocumentParser.get_document_summary(Path('document.docx'))")
    print("print(f\"Paragraphs: {summary['paragraphs']}\")")
    print("print(f\"Tables: {summary['tables']}\")\n")

    print("# Parse multiple files")
    print("files = [Path('doc1.docx'), Path('data.xlsx'), Path('notes.txt')]")
    print("combined = DocumentParser.parse_multiple_files(files)\n")

    print("# Use with PresentationBuilder")
    print("builder = PresentationBuilder()")
    print("content = DocumentParser.parse_file(Path('reference.docx'))")
    print("outline = builder.create_outline(")
    print("    topic='My Presentation',")
    print("    summary='Overview',")
    print("    reference_docs=content  # Parsed content")
    print(")")


def main():
    """Run all demonstrations."""

    print("\n" + "="*70)
    print(" PPTX Agent - Multi-Format Reference Documents")
    print("="*70)

    # Show what formats are supported
    demonstrate_document_parsing()

    # Show practical examples
    create_presentation_from_word_doc()
    create_presentation_from_excel_data()
    create_presentation_from_multiple_sources()

    # Show API usage
    parsing_api_examples()

    print("\n" + "="*70)
    print("Summary")
    print("="*70)
    print("""
The DocumentParser supports the following formats:

✓ .docx - Microsoft Word documents
  - Extracts paragraphs, tables, headers/footers
  - Ideal for: Requirements, reports, specifications

✓ .pptx - Microsoft PowerPoint presentations
  - Extracts titles, text, tables, speaker notes
  - Ideal for: Templates, previous presentations, content libraries

✓ .xlsx - Microsoft Excel spreadsheets
  - Extracts all sheets and cell values
  - Ideal for: Financial data, metrics, datasets

✓ .txt/.md - Plain text files
  - Extracts raw text
  - Ideal for: Notes, documentation, scripts

Command-line usage:
  python -m pptx_agent --topic "My Topic" --summary "Overview" \\
      --reference document.docx --output presentation.pptx

Programmatic usage:
  from pptx_agent.core.document_parser import DocumentParser
  content = DocumentParser.parse_file(Path('document.docx'))
  builder.create_outline(topic, summary, reference_docs=content)
    """)


if __name__ == "__main__":
    main()
