# PPTX Agent

An AI-powered PowerPoint presentation builder that uses Large Language Models (LLMs) to create professional presentations from simple prompts and reference materials.

## Features

- **AI-Powered Content Generation**: Uses OpenAI's GPT models to generate presentation outlines and content
- **Template Support**: Work with custom PowerPoint templates or create from scratch
- **Rich Content Types**:
  - Text slides with various formatting options
  - Bullet point slides
  - Tables with custom styling
  - Charts (bar, line, pie) using matplotlib
  - Image slides with user-provided images
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
│   │   └── presentation_builder.py
│   ├── llm/               # LLM integration
│   │   ├── openai_client.py
│   │   └── content_planner.py
│   ├── builders/          # Slide builders
│   │   ├── text_builder.py
│   │   ├── table_builder.py
│   │   ├── chart_builder.py
│   │   └── image_builder.py
│   ├── cli/               # Command-line interface
│   │   └── interactive.py
│   └── main.py            # Main entry point
├── examples/              # Example scripts and templates
├── tests/                 # Unit tests
├── requirements.txt       # Dependencies
└── README.md
```

## Advanced Features

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
  --topic TOPIC         Presentation topic (for quick mode)
  -s, --summary TEXT    Content summary or key points
  -n, --num-slides N    Target number of slides
  -r, --reference PATH  Path to reference document
  -i, --images PATH     Path to directory containing images
  -o, --output PATH     Output file path (default: output.pptx)
  --api-key KEY         OpenAI API key
  --interactive         Force interactive mode
```

## Examples

See the `examples/` directory for:
- Sample presentation scripts
- Template examples
- Common use cases
- Integration examples

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

- [ ] Support for video embedding
- [ ] Advanced animation options
- [ ] Batch processing of multiple presentations
- [ ] Export to PDF
- [ ] Presentation theme customization
- [ ] Speaker notes generation
- [ ] Collaborative editing features
- [ ] Web interface

## Acknowledgments

Built with:
- [python-pptx](https://python-pptx.readthedocs.io/)
- [OpenAI API](https://platform.openai.com/)
- [matplotlib](https://matplotlib.org/)
