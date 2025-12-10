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
- **Reference Documents**: Incorporate information from text files
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
- `full_features_showcase.py` - **Comprehensive feature demonstration** (SmartArt, shapes, charts, tables)
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
- [ ] Real-time slide preview in terminal
- [ ] Slide-to-image conversion (platform-independent)
- [ ] Support for video embedding
- [ ] Advanced animation options
- [ ] Batch processing of multiple presentations
- [ ] Export to PDF
- [ ] Brand guideline integration
- [ ] Speaker notes generation
- [ ] Multi-user collaborative sessions
- [ ] Web interface

## Acknowledgments

Built with:
- [python-pptx](https://python-pptx.readthedocs.io/)
- [OpenAI API](https://platform.openai.com/)
- [matplotlib](https://matplotlib.org/)
