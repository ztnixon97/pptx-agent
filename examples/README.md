# PPTX Agent Examples

This directory contains example scripts demonstrating various features of PPTX Agent.

## Examples Overview

### simple_presentation.py
Basic example showing how to create a simple presentation with AI-generated content.

**Usage:**
```bash
python examples/simple_presentation.py
```

**Features demonstrated:**
- AI-powered outline generation
- Building from outline
- Adding custom bullet slides
- Adding table slides

### chart_presentation.py
Demonstrates creating presentations with various chart types.

**Usage:**
```bash
python examples/chart_presentation.py
```

**Features demonstrated:**
- Bar charts with multiple series
- Line charts for trends
- Pie charts for distributions
- Combining charts with tables

### with_reference_docs.py
Shows how to incorporate reference documents into presentation generation.

**Usage:**
```bash
python examples/with_reference_docs.py
```

**Features demonstrated:**
- Using reference documents
- AI extracting key information
- Structured content generation

### custom_styling.py
Advanced example with custom formatting and styling.

**Usage:**
```bash
python examples/custom_styling.py
```

**Features demonstrated:**
- Custom table colors
- Two-column layouts
- Quote slides
- Summary tables

## Running Examples

1. Make sure you have installed all requirements:
```bash
pip install -r requirements.txt
```

2. Set your OpenAI API key:
```bash
export OPENAI_API_KEY='your-key-here'
# or create a .env file
```

3. Run any example:
```bash
python examples/simple_presentation.py
```

## Creating Your Own Examples

You can use these examples as templates for your own presentations. Key patterns:

### Basic Setup
```python
from pptx_agent.core.presentation_builder import PresentationBuilder

builder = PresentationBuilder()
```

### With Template
```python
from pathlib import Path

builder = PresentationBuilder(
    template_path=Path("my_template.pptx")
)
```

### AI-Generated Content
```python
outline = builder.create_outline(
    topic="Your Topic",
    summary="Your summary",
    num_slides=10
)
builder.build_from_outline()
```

### Manual Slides
```python
builder.add_bullet_slide("Title", ["Point 1", "Point 2"])
builder.add_table_slide("Title", headers, rows)
builder.add_chart_slide("Title", "bar", categories, series)
```

### Save
```python
builder.save(Path("output.pptx"))
```

## Tips

- Start with AI-generated outlines and refine manually
- Use reference documents for factual presentations
- Combine AI generation with manual slides for best results
- Experiment with different chart types for data visualization
- Use templates for consistent branding

## Output

All examples create `.pptx` files in the current directory. These can be opened with:
- Microsoft PowerPoint
- Google Slides
- LibreOffice Impress
- Any PowerPoint-compatible software
