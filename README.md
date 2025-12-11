# PPTX Agent

An AI-powered PowerPoint presentation builder that uses Large Language Models (LLMs) to create professional presentations from simple prompts and reference materials.

## Features

### 🤝 Collaborative Mode (RECOMMENDED!)
- **Iterative Workflow**: Work WITH the AI to refine your presentation step-by-step
- **Outline Refinement**: Review and refine the structure before building
- **Vision-Based Validation**: AI visually analyzes slides to check formatting and fit
- **Interactive Feedback**: Provide feedback at every stage of creation
- **User-Controlled**: You maintain creative control throughout the process
- **Perfect for Important Presentations**: Ideal when you need precise control

### 🚀 Autonomous Mode
- **Fully AI-Driven Design**: LLM makes ALL styling, layout, and formatting decisions
- **Intelligent Content Validation**: Automatically checks if content fits and optimizes when needed
- **Smart Layout Selection**: AI chooses optimal layouts for each slide type
- **Automatic Color Schemes**: AI-selected color palettes based on topic and audience
- **Content Optimization**: Auto-splits, reformats, or condenses content to fit perfectly
- **Zero Manual Styling**: Just provide topic and content - AI handles everything else

### Core Capabilities
- **AI-Powered Content Generation**: Uses OpenAI's GPT models to generate presentation outlines and content
- **Template Support**: Work with custom PowerPoint templates or create from scratch
- **Rich Content Types**:
  - Text slides with various formatting options
  - Bullet point slides with two-column layouts
  - Tables with custom styling (data tables, comparisons)
  - Charts (bar, line, pie, scatter, area) using matplotlib
  - Image slides with user-provided images (single or grid layouts)
  - **SmartArt-like Diagrams**: Process flows, cycles, hierarchies, comparisons, Venn diagrams, timelines
  - **Custom Shapes & Flowcharts**: Decision diagrams, icon grids, callouts, annotations
  - Section dividers and title slides
- **Interactive Mode**: Refine and customize presentations interactively
- **Quick Mode**: Generate presentations from command-line arguments
- **Multi-Format Reference Documents**: Incorporate information from .docx, .pptx, .xlsx, .txt files
- **Programmatic API**: Use as a Python library for custom workflows

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd pptx-agent
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up your OpenAI API key:
```bash
# Option 1: Environment variable
export OPENAI_API_KEY='your-api-key-here'

# Option 2: Create .env file
cp .env.example .env
# Edit .env and add your API key
```

## Usage

### Collaborative Mode (Recommended for Important Presentations)

Work WITH the AI through an interactive, iterative process:

```bash
# Start collaborative mode
python -m pptx_agent --collaborative --output my_presentation.pptx

# With custom template
python -m pptx_agent --collaborative \
  --template company_template.pptx \
  --output q4_results.pptx
```

**Collaborative Workflow:**

1. **Gather Requirements**: Provide topic, summary, and optional reference documents
2. **Outline Refinement**: AI creates outline → You review → Provide feedback → AI refines → Repeat until approved
3. **Build First Draft**: AI creates presentation from approved outline
4. **Visual Validation** (Optional): AI visually analyzes slides using GPT-4 Vision
5. **Iterative Refinement**: Review and refine individual slides until satisfied

**Example Interaction:**
```
What is your presentation about?
→ Quarterly Business Review

[AI creates initial outline with 12 slides]

Your feedback: Add a slide about team growth and split
              the financial section into revenue and expenses

[AI refines outline based on your feedback]

Your feedback: approve

[AI builds presentation]

Review slide 3?
→ Make the chart larger and use brand colors

[AI helps refine slide]
```

See [COLLABORATIVE_MODE.md](COLLABORATIVE_MODE.md) for complete documentation.

### Autonomous Mode (Recommended for Quick Drafts)

Let the AI make all design decisions - fastest way to create presentations:

```bash
# Basic autonomous mode
python -m pptx_agent --autonomous \
  --topic "Future of Quantum Computing" \
  --summary "Exploring quantum computing developments, applications, and industry impact" \
  --output quantum.pptx

# With audience targeting
python -m pptx_agent --autonomous \
  --topic "Q4 Financial Results" \
  --summary "Revenue, growth metrics, and strategic initiatives" \
  --audience executive \
  --num-slides 10 \
  --output q4_results.pptx

# With reference documents for factual accuracy
python -m pptx_agent --autonomous \
  --topic "Product Launch Plan" \
  --summary "Go-to-market strategy and timeline" \
  --reference product_spec.txt \
  --output launch_plan.pptx
```

**What Autonomous Mode Does:**
- ✅ Selects optimal layouts for each slide
- ✅ Chooses professional color schemes
- ✅ Determines font sizes and styling
- ✅ Validates content fits on slides
- ✅ Auto-reformats or splits content when needed
- ✅ Organizes visual hierarchy
- ✅ Makes all design decisions autonomously

See [AUTONOMOUS_MODE.md](AUTONOMOUS_MODE.md) for detailed documentation.

### Interactive Mode

Launch the interactive CLI:

```bash
python -m pptx_agent
```

Or with a custom template:

```bash
python -m pptx_agent --template my_template.pptx
```

### Quick Mode

Create a presentation directly from the command line:

```bash
python -m pptx_agent \
  --topic "AI in Healthcare" \
  --summary "Exploring how artificial intelligence is transforming healthcare delivery, diagnostics, and patient outcomes" \
  --num-slides 10 \
  --output ai_healthcare.pptx
```

With reference documents:

```bash
python -m pptx_agent \
  --topic "Q4 Financial Results" \
  --summary "Quarterly financial performance and key metrics" \
  --reference reports/q4_data.txt \
  --images assets/charts/ \
  --output q4_results.pptx
```

### As a Python Library

```python
from pathlib import Path
from pptx_agent.core.presentation_builder import PresentationBuilder

# Initialize builder
builder = PresentationBuilder(
    template_path=Path("template.pptx"),  # Optional
    openai_api_key="your-api-key"  # Optional if set in env
)

# Create outline
outline = builder.create_outline(
    topic="Machine Learning Basics",
    summary="Introduction to ML concepts, algorithms, and applications",
    num_slides=8
)

# Build presentation
builder.build_from_outline()

# Add custom slides
builder.add_bullet_slide(
    "Key Takeaways",
    [
        "Machine learning is transforming industries",
        "Start with simple algorithms",
        "Data quality is crucial"
    ]
)

# Add a chart
builder.add_chart_slide(
    "Model Performance",
    chart_type="bar",
    categories=["Accuracy", "Precision", "Recall"],
    series=[{"name": "Scores", "values": [0.95, 0.92, 0.89]}]
)

# Save
builder.save(Path("ml_basics.pptx"))
```

## Project Structure

```
pptx-agent/
├── pptx_agent/
│   ├── core/              # Core PowerPoint manipulation
│   │   ├── pptx_handler.py
│   │   ├── template_manager.py
│   │   ├── presentation_builder.py
│   │   ├── iterative_workflow.py      # Collaborative workflow manager
│   │   ├── autonomous_builder.py      # Fully autonomous builder
│   │   ├── content_validator.py       # Content fitting validation
│   │   └── layout_optimizer.py        # Intelligent layout selection
│   ├── llm/               # LLM integration
│   │   ├── openai_client.py
│   │   ├── content_planner.py
│   │   ├── vision_validator.py        # Vision-based slide validation
│   │   └── autonomous_designer.py     # AI design decisions
│   ├── builders/          # Slide builders
│   │   ├── text_builder.py
│   │   ├── table_builder.py
│   │   ├── chart_builder.py
│   │   ├── image_builder.py
│   │   ├── smartart_builder.py        # NEW: SmartArt diagrams
│   │   └── shapes_builder.py          # NEW: Custom shapes & flowcharts
│   ├── cli/               # Command-line interface
│   │   ├── interactive.py
│   │   └── collaborative.py           # Collaborative mode CLI
│   └── main.py            # Main entry point
├── examples/              # Example scripts and templates
├── tests/                 # Unit tests
├── requirements.txt       # Dependencies
├── README.md
├── CAPABILITIES.md        # NEW: AI agent capabilities reference
├── COLLABORATIVE_MODE.md  # Collaborative mode documentation
└── AUTONOMOUS_MODE.md     # Autonomous mode documentation
```

## Advanced Features

### Autonomous AI Designer

The autonomous designer makes intelligent decisions about every aspect of your presentation:

```python
from pptx_agent.core.autonomous_builder import AutonomousPresentationBuilder

builder = AutonomousPresentationBuilder(target_audience="professional")

# AI makes ALL decisions - just provide content
report = builder.create_presentation_autonomously(
    topic="Machine Learning in Healthcare",
    summary="Applications, benefits, and challenges of ML in medical field",
    num_slides=10
)

# Review AI decisions
print(f"Color scheme chosen: {report['decisions_made'][0]['choice']}")
print(f"Optimizations performed: {len(report['optimizations_performed'])}")

builder.save(Path("ml_healthcare.pptx"))
```

**AI Decision Examples:**
- Detects data-heavy content → Chooses chart-optimized layout
- Sees comparison content → Selects two-column layout
- Finds too many bullets → Automatically splits across slides
- Content too long → Reduces font size or summarizes
- Images present → Picks picture-friendly layout

### Working with Templates

PPTX Agent can analyze and use PowerPoint templates:

```python
from pptx_agent.core.template_manager import TemplateManager

# Analyze template
manager = TemplateManager(Path("corporate_template.pptx"))

# List available layouts
print(manager.get_template_summary())

# Get layout suggestions
layout_idx = manager.suggest_layout("content")
```

### Custom Chart Creation

Create custom charts with matplotlib:

```python
import matplotlib.pyplot as plt
from pptx_agent.builders.chart_builder import ChartSlideBuilder

# Create custom matplotlib figure
fig, ax = plt.subplots()
ax.plot([1, 2, 3, 4], [1, 4, 2, 3])
ax.set_title("Custom Chart")

# Add to presentation
ChartSlideBuilder.add_custom_chart_slide(
    handler,
    "My Custom Chart",
    fig
)
```

### Programmatic Table Creation

```python
# Simple table
builder.add_table_slide(
    "Quarterly Revenue",
    headers=["Quarter", "Revenue", "Growth"],
    rows=[
        ["Q1", "$1.2M", "15%"],
        ["Q2", "$1.5M", "25%"],
        ["Q3", "$1.8M", "20%"],
    ]
)
```

### Image Management

```python
from pathlib import Path

# Single image
builder.add_image_slide(
    "Product Screenshot",
    Path("images/product.png"),
    caption="Our flagship product"
)

# Multiple images in grid
from pptx_agent.builders.image_builder import ImageSlideBuilder

ImageSlideBuilder.add_multiple_images_slide(
    handler,
    "Gallery",
    [
        Path("img1.png"),
        Path("img2.png"),
        Path("img3.png")
    ]
)
```

### SmartArt-like Diagrams

Create professional diagrams to visualize processes, relationships, and hierarchies:

```python
from pptx_agent.builders.smartart_builder import SmartArtBuilder

smartart = SmartArtBuilder()

# Process flow (sequential steps)
smartart.add_process_flow(
    builder.handler,
    "Development Process",
    ["Plan", "Design", "Build", "Test", "Deploy"]
)

# Cycle diagram (continuous process)
smartart.add_cycle_diagram(
    builder.handler,
    "Agile Sprint Cycle",
    ["Sprint Planning", "Development", "Testing", "Review", "Retrospective"]
)

# Hierarchy (organizational structure)
smartart.add_hierarchy_diagram(
    builder.handler,
    "Organization Chart",
    "CEO",
    ["Engineering", "Product", "Marketing", "Sales"]
)

# Comparison (side-by-side)
smartart.add_comparison_diagram(
    builder.handler,
    "Build vs Buy",
    left_items=["Full control", "Custom features", "No licensing"],
    right_items=["Faster deployment", "Proven solution", "Support included"],
    left_label="Build",
    right_label="Buy"
)

# Venn diagram (overlap analysis)
smartart.add_venn_diagram(
    builder.handler,
    "Skill Set Analysis",
    left_label="Technical Skills",
    right_label="Business Skills",
    left_items=["Coding", "Architecture", "DevOps"],
    right_items=["Strategy", "Finance", "Marketing"],
    overlap_items=["Product", "Analytics", "Leadership"]
)

# Timeline (chronological events)
smartart.add_timeline(
    builder.handler,
    "Product Roadmap 2024",
    [
        {"date": "Q1", "event": "Beta Launch"},
        {"date": "Q2", "event": "Public Release"},
        {"date": "Q3", "event": "Mobile App"},
        {"date": "Q4", "event": "Enterprise Features"}
    ]
)
```

### Custom Shapes & Flowcharts

Create custom shapes and flowchart diagrams:

```python
from pptx_agent.builders.shapes_builder import ShapesBuilder

shapes = ShapesBuilder()

# Flowchart with decision points
shapes.add_flowchart_slide(
    builder.handler,
    "Decision Process",
    [
        {"text": "Start", "decision": False},
        {"text": "Budget Available?", "decision": True},
        {"text": "Evaluate Options", "decision": False},
        {"text": "Best Option?", "decision": True},
        {"text": "Proceed", "decision": False}
    ]
)

# Icon grid with labels
shapes.add_icon_grid_slide(
    builder.handler,
    "Core Values",
    [
        {"shape": "star", "label": "Quality", "color": (255, 215, 0)},
        {"shape": "diamond", "label": "Value", "color": (68, 114, 196)},
        {"shape": "hexagon", "label": "Speed", "color": (112, 173, 71)},
        {"shape": "circle", "label": "Support", "color": (237, 125, 49)}
    ],
    cols=2
)

# Available shapes: rectangle, rounded_rectangle, oval, circle, triangle,
# diamond, pentagon, hexagon, octagon, star, arrow_right, arrow_left,
# arrow_up, arrow_down, callout_rectangle, callout_rounded, cloud
```

### Speaker Notes

Add presenter notes to slides for better presentation delivery:

```python
# Add speaker notes to a slide
handler.add_speaker_notes(
    slide,
    "This slide shows Q4 results. Emphasize the 25% growth in APAC region. "
    "Be prepared to discuss the factors behind this growth if asked."
)

# Get speaker notes from a slide
notes = handler.get_speaker_notes(slide)
```

### Hyperlinks

Create interactive presentations with clickable links:

```python
# External hyperlink on a shape
textbox = handler.add_text_box(slide, "Visit our website", 2, 3, 4, 1)
handler.add_hyperlink_to_shape(textbox, "https://example.com")

# Internal hyperlink to jump to another slide
nav_button = handler.add_text_box(slide, "Jump to Conclusions →", 2, 5, 4, 0.8)
handler.add_internal_hyperlink(nav_button, 15)  # Jump to slide 15
```

### Rich Text Formatting

Apply multiple styles within a single text box:

```python
# Create text with mixed formatting
handler.add_formatted_text_box(
    slide,
    text_runs=[
        {"text": "Important: ", "bold": True, "font_size": 24, "color": (255, 0, 0)},
        {"text": "This feature ", "font_size": 18},
        {"text": "emphasizes", "italic": True, "underline": True, "font_size": 18},
        {"text": " key points", "font_size": 18}
    ],
    left=2, top=2.5, width=6, height=2
)
```

### Advanced Tables

Create sophisticated tables with cell merging and custom styling:

```python
from pptx_agent.builders.table_builder import TableSlideBuilder

TableSlideBuilder.add_advanced_table_slide(
    handler,
    "Q4 Financial Results",
    headers=["Quarter", "Revenue", "Expenses", "Profit"],
    rows=[
        ["Q1", "$500K", "$300K", "$200K"],
        ["Q2", "$600K", "$350K", "$250K"],
        ["Q3", "$700K", "$400K", "$300K"],
        ["Q4", "$800K", "$450K", "$350K"]
    ],
    cell_styles={
        # Highlight the totals row
        (3, 3): {"fill": (255, 215, 0), "bold": True, "font_size": 14}
    }
)
```

### Slide Management

Control your presentation structure:

```python
# Duplicate a slide
handler.duplicate_slide(3)

# Hide a backup slide
handler.hide_slide(10)  # Hidden but available if needed

# Reorder slides
handler.reorder_slides([0, 2, 1, 3, 4])  # New order

# Delete a slide
handler.delete_slide(5)
```

### Footer and Slide Numbers

Add professional touches:

```python
# Add footer to all slides
handler.set_presentation_footer(
    footer_text="Company Confidential | Q4 2024",
    show_slide_number=True,
    show_date=False
)
```

### Slide Dimensions

Set presentation size:

```python
# Widescreen format (16:9) - modern displays
handler.set_slide_size("widescreen")

# Standard format (4:3) - traditional projectors
handler.set_slide_size("standard")

# Custom dimensions
handler.set_custom_slide_size(12, 9)
```

### Image Enhancements

Add accessible, professional images:

```python
# Add image with alt text for accessibility
handler.add_image_with_alt_text(
    slide,
    Path("chart.png"),
    left=2, top=2, width=6, height=4,
    alt_text="Bar chart showing 25% revenue growth in Q4"
)

# Set image transparency for watermarks
pic = slide.shapes.add_picture(str(logo_path), Inches(0.5), Inches(0.5))
handler.set_image_transparency(pic, 0.3)  # 30% transparent
```

### Shape Layering

Control visual stacking order:

```python
# Send background shape to back
handler.send_to_back(slide, 0)

# Bring important shape to front
handler.bring_to_front(slide, 2)
```

### Multi-Format Reference Documents

Extract content from various document formats to use as reference material:

```python
from pptx_agent.core.document_parser import DocumentParser

# Parse a Word document
content = DocumentParser.parse_file(Path("project_spec.docx"))

# Parse an Excel spreadsheet
data = DocumentParser.parse_file(Path("financials.xlsx"))

# Parse a PowerPoint presentation
slides = DocumentParser.parse_file(Path("previous_deck.pptx"))

# Parse multiple documents
files = [
    Path("requirements.docx"),
    Path("data.xlsx"),
    Path("notes.txt")
]
combined = DocumentParser.parse_multiple_files(files)

# Use with PresentationBuilder
builder = PresentationBuilder()
outline = builder.create_outline(
    topic="Product Launch",
    summary="Comprehensive launch strategy",
    reference_docs=content  # AI extracts relevant information
)
```

**Supported Formats:**
- **.docx** - Microsoft Word (paragraphs, tables, headers/footers)
- **.pptx** - Microsoft PowerPoint (titles, text, tables, speaker notes)
- **.xlsx/.xls** - Microsoft Excel (all sheets, cell values)
- **.txt/.md** - Plain text/Markdown files

**Command-line Usage:**
```bash
# Use a Word document as reference
python -m pptx_agent --topic "Q4 Review" --summary "Financial overview" \
    --reference report.docx --output presentation.pptx

# Use an Excel file as reference
python -m pptx_agent --topic "Sales Analysis" --summary "Regional performance" \
    --reference sales_data.xlsx --output sales.pptx
```

## Environment Variables

- `OPENAI_API_KEY`: Your OpenAI API key (required)
- `OPENAI_MODEL`: Model to use (default: gpt-4-turbo-preview)
- `OPENAI_BASE_URL`: Custom API base URL for OpenAI-compatible APIs

## Command-Line Options

```
usage: python -m pptx_agent [options]

options:
  -h, --help            Show help message
  -t, --template PATH   Path to PowerPoint template file
  --topic TOPIC         Presentation topic
  -s, --summary TEXT    Content summary or key points
  -n, --num-slides N    Target number of slides
  -r, --reference PATH  Path to reference document
  -i, --images PATH     Path to directory containing images
  -o, --output PATH     Output file path (default: output.pptx)
  --api-key KEY         OpenAI API key
  --interactive         Force interactive mode
  --collaborative       Use collaborative mode (iterative workflow with user feedback)
  --autonomous          Use fully autonomous mode (AI makes all design decisions)
  --audience TYPE       Target audience (professional/technical/general/executive)
```

## Mode Comparison

| Feature | Collaborative | Autonomous | Quick | Interactive |
|---------|--------------|------------|-------|-------------|
| User Control | High | Low | Medium | High |
| Speed | Moderate | Fast | Fast | Slow |
| Feedback Loops | Multiple | None | None | Manual |
| Vision Validation | Yes | No | No | No |
| Best For | Important presentations | Quick drafts | Simple needs | Exploration |
| Outline Refinement | Iterative | Auto | Auto | Manual |
| Slide-by-slide Review | Yes | No | No | Yes |

## Examples

See the `examples/` directory for:
- `collaborative_workflow.py` - **Interactive collaborative mode** (RECOMMENDED)
- `advanced_features_demo.py` - **Advanced features showcase** (notes, hyperlinks, rich text, advanced tables)
- `multi_format_references.py` - **NEW: Multi-format reference documents** (.docx, .pptx, .xlsx, .txt)
- `full_features_showcase.py` - **Comprehensive content types** (SmartArt, shapes, charts, tables)
- `autonomous_presentation.py` - Full autonomous mode example
- `autonomous_with_validation.py` - Auto-optimization demo
- `simple_presentation.py` - Basic presentation creation
- `chart_presentation.py` - Charts and data visualization
- `with_reference_docs.py` - Using reference documents
- `custom_styling.py` - Custom formatting
- More examples for common use cases

## Requirements

- Python 3.8+
- python-pptx
- openai
- matplotlib
- pandas
- Pillow
- colorama
- python-dotenv
- pydantic
- python-docx (for Word document parsing)
- openpyxl (for Excel parsing)

## Development

### Running Tests

```bash
pytest tests/
```

### Adding Custom Slide Builders

Create a new builder in `pptx_agent/builders/`:

```python
class CustomSlideBuilder:
    @staticmethod
    def add_custom_slide(handler, **kwargs):
        slide = handler.add_slide(1)
        # Add your custom logic
        return slide
```

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Add tests for new features
4. Submit a pull request

## License

MIT License - see LICENSE file for details

## Troubleshooting

### API Key Issues

If you see "OpenAI API key must be provided":
- Ensure `OPENAI_API_KEY` is set in environment or `.env` file
- Use `--api-key` argument
- Check that the API key is valid

### Template Issues

If template layouts aren't working:
- Verify the template file exists and is a valid .pptx
- Use `--template` to specify the correct path
- Check template layouts with the interactive mode (option 7)

### Image Issues

If images aren't appearing:
- Verify image file paths are correct
- Ensure images are in supported formats (PNG, JPG, GIF)
- Check file permissions
- Use absolute paths when possible

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

## Roadmap

### ✅ Completed Features

- [x] ✅ Fully autonomous presentation generation
- [x] ✅ Intelligent content validation and optimization
- [x] ✅ AI-powered layout selection
- [x] ✅ Automatic color scheme generation
- [x] ✅ Collaborative mode with iterative refinement
- [x] ✅ Vision-based slide validation (GPT-4 Vision)
- [x] ✅ Interactive feedback loops
- [x] ✅ SmartArt-like diagrams (process, cycle, hierarchy, comparison, venn, timeline)
- [x] ✅ Custom shapes and flowcharts
- [x] ✅ Comprehensive AI capabilities reference
- [x] ✅ **Speaker notes** - Add presenter notes to slides
- [x] ✅ **Hyperlinks** - External and internal slide links
- [x] ✅ **Rich text formatting** - Bold, italic, colors, multiple fonts
- [x] ✅ **Advanced tables** - Cell merging and individual cell styling
- [x] ✅ **Slide management** - Duplicate, hide, reorder, delete slides
- [x] ✅ **Footer and slide numbers** - Professional slide numbering
- [x] ✅ **Slide dimensions** - Widescreen, standard, custom sizes
- [x] ✅ **Image enhancements** - Alt text for accessibility, transparency
- [x] ✅ **Shape layering** - Control z-order (bring to front/send to back)
- [x] ✅ **Multi-format reference documents** - Parse .docx, .pptx, .xlsx, .txt files

### 🚧 Planned Features

- [ ] Advanced animation options (entrance, emphasis, exit)
- [ ] Slide transitions (fade, push, wipe, etc.)
- [ ] Advanced chart features (combo charts, secondary axis, data labels)
- [ ] Comments and review tracking
- [ ] Full shape grouping support
- [ ] Real-time slide preview in terminal
- [ ] Slide-to-image conversion (platform-independent)
- [ ] Support for video embedding
- [ ] Audio narration
- [ ] Batch processing of multiple presentations
- [ ] Export to PDF
- [ ] Brand guideline integration
- [ ] Multi-user collaborative sessions
- [ ] Web interface
- [ ] Equations and mathematical notation

## Acknowledgments

Built with:
- [python-pptx](https://python-pptx.readthedocs.io/)
- [OpenAI API](https://platform.openai.com/)
- [matplotlib](https://matplotlib.org/)
