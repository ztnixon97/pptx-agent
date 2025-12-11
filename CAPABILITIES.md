# PPTX Agent Capabilities Reference

This document describes all capabilities available to the AI agent when creating presentations. Use this as a reference when planning and generating presentations.

## Content Types Supported

### 1. Text Content

**Capabilities:**
- Title slides with main title and subtitle
- Content slides with body text
- Bullet point lists (recommended max 7 points)
- Two-column text layouts
- Section divider slides
- Quote slides with attribution
- Formatted text with custom fonts, sizes, colors, and styles

**Builder:** `TextSlideBuilder`

**Methods:**
- `add_content_slide()` - Standard text content
- `add_bullet_slide()` - Bullet point lists
- `add_two_column_slide()` - Side-by-side content
- `add_section_slide()` - Section dividers
- `add_quote_slide()` - Quotations
- `add_formatted_text_slide()` - Custom formatting

**Best Practices:**
- Keep bullets concise (max 80 characters)
- Use parallel structure in lists
- Limit to 7 bullets per slide (6x6 rule)
- Font sizes: Title 32-44pt, Body 18-24pt

### 2. Tables

**Capabilities:**
- Data tables with headers
- Comparison tables (side-by-side)
- Summary tables (key-value pairs)
- Custom colored headers and alternating rows
- Support for any number of columns/rows (recommended max 6 cols, 10 rows)

**Builder:** `TableSlideBuilder`

**Methods:**
- `add_table_slide()` - Standard tables
- `add_comparison_table_slide()` - Comparison format
- `add_styled_table_slide()` - Custom colors
- `add_summary_table_slide()` - Key-value format

**Best Practices:**
- Keep to 6 columns or fewer for readability
- Limit to 10 rows per slide
- Use alternating row colors
- Bold headers with contrasting background
- Consider charts for large datasets

### 3. Charts & Graphs

**Capabilities:**
- Bar charts (vertical, horizontal, stacked)
- Line charts (single or multiple series)
- Pie charts
- Scatter plots (via matplotlib)
- Area charts (via matplotlib)
- Custom matplotlib figures

**Builder:** `ChartSlideBuilder`

**Methods:**
- `add_bar_chart_slide()` - Bar/column charts
- `add_line_chart_slide()` - Line/trend charts
- `add_pie_chart_slide()` - Pie/donut charts
- `add_custom_chart_slide()` - Any matplotlib figure

**Supported Data:**
- Multiple data series
- Custom colors per series
- Axis labels and titles
- Legends
- Grid lines

**Best Practices:**
- Bar charts for comparisons
- Line charts for trends over time
- Pie charts for parts of whole (max 6 slices)
- Include clear axis labels
- Use contrasting colors for multiple series

### 4. Images

**Capabilities:**
- Single images with captions
- Full-screen images
- Multiple images in grid layout (up to 6)
- Image with text side-by-side
- Support for PNG, JPG, GIF formats

**Builder:** `ImageSlideBuilder`

**Methods:**
- `add_image_slide()` - Single image with caption
- `add_full_image_slide()` - Full-screen image
- `add_multiple_images_slide()` - Grid of images
- `add_image_with_text_slide()` - Image + text layout

**Best Practices:**
- Ensure images are high resolution (150+ DPI)
- Use aspect ratios that fit slides (16:9 or 4:3)
- Add captions for context
- Consider image placement in layout

### 5. SmartArt-Like Diagrams

**Capabilities:**
- Process flows (horizontal sequential steps)
- Cycle diagrams (circular processes)
- Hierarchy diagrams (org charts, top-down)
- Comparison diagrams (two-column comparisons)
- Venn diagrams (overlapping concepts)
- Timelines (chronological events)

**Builder:** `SmartArtBuilder`

**Methods:**
- `add_process_flow()` - Step-by-step processes
- `add_cycle_diagram()` - Circular/cyclical processes
- `add_hierarchy_diagram()` - Organizational structure
- `add_comparison_diagram()` - Side-by-side comparisons
- `add_venn_diagram()` - Overlapping concepts
- `add_timeline()` - Chronological events

**Use Cases:**
- Process flows: Project phases, workflows, methodologies
- Cycles: Product lifecycle, continuous improvement, feedback loops
- Hierarchies: Organization charts, category breakdowns
- Comparisons: Pros/cons, before/after, option analysis
- Venn diagrams: Shared characteristics, overlapping features
- Timelines: Project milestones, historical events, roadmaps

### 6. Custom Shapes

**Capabilities:**
- Basic shapes: rectangles, circles, triangles, diamonds
- Arrows: up, down, left, right
- Callouts and annotations
- Flowcharts (decision diamonds, process boxes)
- Icon grids
- Connectors between shapes

**Builder:** `ShapesBuilder`

**Available Shapes:**
```
rectangle, rounded_rectangle, oval, circle, triangle, diamond,
pentagon, hexagon, octagon, star, arrow_right, arrow_left,
arrow_up, arrow_down, callout, cloud
```

**Methods:**
- `add_shape_slide()` - Single custom shape
- `add_callout_slide()` - Multiple callouts/annotations
- `add_flowchart_slide()` - Flowchart diagrams
- `add_icon_grid_slide()` - Grid of icons with labels
- `add_annotation_slide()` - Annotated content
- `add_connector()` - Lines connecting elements

**Use Cases:**
- Flowcharts for decision trees
- Callouts for highlighting key points
- Icons for feature lists
- Annotations for detailed explanations

### 7. Advanced Features

#### Speaker Notes

**Capabilities:**
- Add presenter notes to any slide
- Include talking points, context, and details not shown on slide
- Accessible in presenter view during presentation

**Methods:**
- `handler.add_speaker_notes(slide, notes_text)`
- `handler.get_speaker_notes(slide)`

**Best Practices:**
- Add notes to complex slides for clarity
- Include context and background information
- Add delivery tips and emphasis points
- Keep notes concise but informative (2-3 sentences per slide)
- Include data sources or references

**When to Use:**
- Data-heavy slides (explain what numbers mean)
- Technical content (provide additional context)
- Executive presentations (include supporting details)
- Training materials (add instructor guidance)

#### Hyperlinks

**Capabilities:**
- External hyperlinks to websites or documents
- Internal links to jump to specific slides
- Clickable shapes and text

**Methods:**
- `handler.add_hyperlink_to_shape(shape, url)` - Link shape to URL
- `handler.add_hyperlink_to_text(run, url)` - Link text to URL
- `handler.add_internal_hyperlink(shape, slide_index)` - Jump to slide

**Best Practices:**
- Use for references, resources, or additional information
- Create navigation buttons for non-linear presentations
- Link to agenda slide from section headers
- Add "back to top" buttons on detail slides
- Make links visually distinct (blue, underlined)

**Use Cases:**
- Interactive presentations with branching paths
- Resource slides with external links
- Training decks with reference materials
- Navigation between sections

#### Rich Text Formatting

**Capabilities:**
- Multiple formatted runs in single text box
- Bold, italic, underline
- Custom font sizes and colors within same text
- Different fonts within same text box

**Methods:**
- `handler.add_formatted_text_box(slide, text_runs, left, top, width, height)`

**Text Runs Format:**
```python
text_runs = [
    {"text": "Important: ", "bold": True, "color": (255, 0, 0), "font_size": 24},
    {"text": "Regular text", "font_size": 18},
    {"text": " emphasized", "italic": True, "underline": True}
]
```

**Best Practices:**
- Use bold for emphasis on key terms
- Use color sparingly for impact
- Use italics for definitions or quotes
- Keep formatting consistent across slides
- Don't over-format (reduces readability)

**Use Cases:**
- Highlighting key terms or concepts
- Color-coding different information types
- Emphasizing critical warnings or notes
- Mixed-format instructions or explanations

#### Advanced Table Features

**Capabilities:**
- Cell merging (span multiple rows/columns)
- Individual cell styling (colors, fonts, alignment)
- Custom header colors
- Cell-specific formatting

**Methods:**
- `TableSlideBuilder.add_advanced_table_slide()`

**Options:**
- `merge_cells`: List of (row1, col1, row2, col2) tuples
- `cell_styles`: Dict mapping (row, col) to style dict

**Style Options:**
```python
cell_styles = {
    (0, 0): {
        "fill": (68, 114, 196),      # Background color
        "bold": True,                 # Bold text
        "italic": False,              # Italic text
        "font_size": 14,              # Font size in points
        "font_color": (255, 255, 255), # Text color
        "align": "center"             # left, center, right
    }
}
```

**Best Practices:**
- Merge cells for section headers or totals
- Use contrasting colors for headers
- Highlight important cells (totals, outliers)
- Maintain alignment consistency
- Use cell styling to guide reader's eye

**Use Cases:**
- Financial reports with subtotals
- Comparison matrices
- Scorecards and dashboards
- Complex data presentations

#### Slide Management

**Capabilities:**
- Duplicate slides
- Delete slides
- Reorder slides
- Hide slides (backup content)
- Show hidden slides

**Methods:**
- `handler.duplicate_slide(slide_index)` - Copy entire slide
- `handler.delete_slide(index)` - Remove slide
- `handler.reorder_slides(new_order)` - Rearrange slide order
- `handler.hide_slide(index)` - Hide from presentation
- `handler.show_slide(index)` - Unhide slide

**Best Practices:**
- Hide technical backup slides (show only if asked)
- Duplicate template slides for consistency
- Reorder for optimal flow after initial draft
- Keep backup data slides hidden by default
- Use hidden slides for optional deep-dives

**Use Cases:**
- Backup slides for anticipated questions
- Alternative versions for different audiences
- Technical appendices
- Detailed data not needed in main flow

#### Footer and Slide Numbers

**Capabilities:**
- Add footer text to all slides
- Show slide numbers
- Show date (auto-updated)
- Per-slide or presentation-wide settings

**Methods:**
- `handler.set_slide_footer(slide, footer_text, show_slide_number, show_date)`
- `handler.set_presentation_footer(footer_text, show_slide_number, show_date)`

**Best Practices:**
- Add footer for professional presentations
- Include company name or confidentiality notice
- Show slide numbers for easy reference
- Exclude slide numbers from title and section slides
- Keep footer text concise

**Use Cases:**
- Corporate presentations (company name, confidentiality)
- Conference talks (slide numbers for Q&A)
- Training materials (version numbers)
- Long presentations (orientation for audience)

#### Slide Dimensions

**Capabilities:**
- Widescreen format (16:9) - 13.333" x 7.5"
- Standard format (4:3) - 10" x 7.5"
- Custom dimensions

**Methods:**
- `handler.set_slide_size("widescreen")` - 16:9 format
- `handler.set_slide_size("standard")` - 4:3 format
- `handler.set_custom_slide_size(width, height)` - Custom size

**Best Practices:**
- Use widescreen (16:9) for modern displays
- Use standard (4:3) for older projectors
- Check venue requirements before creating
- Set size before adding content (prevents scaling issues)

**Use Cases:**
- Modern conference presentations (widescreen)
- Traditional projectors (standard)
- Digital displays with specific dimensions
- Printed handouts (custom sizing)

#### Image Enhancements

**Capabilities:**
- Alt text for accessibility (screen readers)
- Image transparency/opacity control
- Proper image positioning and sizing

**Methods:**
- `handler.add_image_with_alt_text(slide, image_path, left, top, width, height, alt_text)`
- `handler.set_image_transparency(image_shape, transparency)`

**Alt Text Guidelines:**
- Describe image content concisely
- Include relevant details for context
- Don't start with "image of" or "picture of"
- Keep to 1-2 sentences

**Transparency Values:**
- 0.0 = Fully opaque (solid)
- 0.5 = Semi-transparent
- 1.0 = Fully transparent (invisible)

**Best Practices:**
- Always add alt text for accessibility
- Use transparency for watermarks or backgrounds
- Ensure images are high resolution
- Position images thoughtfully in layout

**Use Cases:**
- Accessibility compliance
- Watermarked backgrounds
- Layered visual compositions
- Background images with text overlay

#### Shape Layering (Z-Order)

**Capabilities:**
- Send shapes to back (bottom of stack)
- Bring shapes to front (top of stack)
- Control visual stacking order

**Methods:**
- `handler.send_to_back(slide, shape_index)` - Move to back
- `handler.bring_to_front(slide, shape_index)` - Move to front

**Best Practices:**
- Send backgrounds/watermarks to back
- Bring important elements to front
- Layer shapes intentionally for visual effect
- Test layering before finalizing

**Use Cases:**
- Complex diagrams with overlapping shapes
- Background images with text overlays
- Callouts and annotations
- Visual hierarchies

## Layout & Design Capabilities

### Available Layouts

Templates typically include:
1. **Title Slide** - Main title and subtitle
2. **Title and Content** - Title with content area
3. **Section Header** - Section divider
4. **Two Content** - Side-by-side content
5. **Comparison** - Two-column comparison
6. **Title Only** - Title with blank content area
7. **Blank** - No placeholders, full customization
8. **Picture with Caption** - Image-focused layout

### Color & Styling

**Capabilities:**
- Custom RGB colors for any element
- Color schemes (professional, creative, minimal)
- Font customization (size, weight, color)
- Transparency and gradients
- Line styles and widths
- Fill patterns

**Default Professional Palette:**
- Primary Blue: RGB(68, 114, 196)
- Accent Orange: RGB(237, 125, 49)
- Success Green: RGB(112, 173, 71)
- Neutral Gray: RGB(165, 165, 165)
- Background White: RGB(255, 255, 255)
- Text Black: RGB(0, 0, 0)

### Content Validation

**Capabilities:**
- Text overflow detection
- Font size optimization
- Bullet point count validation
- Table size validation
- Image dimension checking
- Content fit analysis

**Validation Criteria:**
- Text: Fits within boundaries, readable font size (14pt+)
- Bullets: Max 7 points, max 80 chars each
- Tables: Max 6 columns, 10 rows
- Images: Proper aspect ratio, fits slide bounds

## Advanced Features

### Vision-Based Validation

**Capability:** If GPT-4 Vision is available, can visually analyze slides

**Can Check:**
- Content overflow (text extending beyond slide)
- Readability (font sizes too small)
- Color contrast (text readable on background)
- Visual balance (layout symmetry)
- Alignment issues
- Information density (too crowded/sparse)

**Output:** Detailed validation report with:
- Overall score (0-10)
- Specific issues by severity (high/medium/low)
- Actionable suggestions for fixes
- Strengths and recommendations

### Reference Document Processing

**Capabilities:**
- Parse multiple document formats:
  * **Word documents (.docx)**: Paragraphs, tables, headers/footers
  * **PowerPoint presentations (.pptx)**: Titles, text, tables, speaker notes
  * **Excel spreadsheets (.xlsx, .xls)**: All sheets, cell values, formatted data
  * **Text files (.txt, .md)**: Plain text and Markdown
- Extract key information from all formats
- Combine content from multiple document types
- Generate content based on factual data
- Maintain context from reference materials

**Methods:**
- `DocumentParser.parse_file(file_path)` - Parse a single document
- `DocumentParser.parse_multiple_files(file_paths)` - Combine multiple documents
- `DocumentParser.get_document_summary(file_path)` - Get document metadata

**Best Practices:**
- Use references for factual presentations
- Combine multiple sources (e.g., Word spec + Excel data)
- Extract specific data points from Excel for charts/tables
- Reuse content from previous PowerPoint presentations
- Parse Word documents for detailed requirements/specifications
- Cite sources when appropriate

**Use Cases:**
- Convert project specifications (Word) into presentations
- Create financial reviews from Excel data
- Reuse slides/content from previous PowerPoint decks
- Combine data from multiple sources (specs + data + notes)
- Extract structured data from Excel for visualization

### Template Support (Basic)

**Capabilities:**
- Load and analyze PowerPoint templates (.pptx)
- Extract available layouts
- Use template color schemes
- Identify placeholder positions
- Suggest optimal layouts for content types

## Flexible Layout System (NEW)

### Template Intelligence & Branding Preservation

**Revolutionary Capability:** PPTX Agent can parse template slides to understand their structure, preserve branding elements, and allow flexible content positioning - combining the best of templates and dynamic layouts.

#### Core Concept

**Traditional Approach:**
- Templates define BOTH styling AND structure
- Locked into predefined layouts
- Can't customize positioning
- Either use template OR create from scratch

**PPTX Agent Approach:**
- Templates define STYLING (colors, fonts, logos) NOT structure
- Parse templates to identify branding vs content
- Preserve branding, flexible content positioning
- Best of both worlds: consistency + flexibility

#### Template Slide Parser

**Capabilities:**
- Analyze any template slide to identify all elements
- Classify elements automatically:
  * **Branding**: Logos, watermarks, company names, persistent headers/footers
  * **Content**: Placeholders, main content areas
  * **Decoration**: Decorative shapes, lines, backgrounds
- Calculate "safe content area" (avoiding branding zones)
- Extract element properties (position, size, type, text)
- Provide layout constraints for intelligent positioning

**Module:** `TemplateSlideParser`

**Methods:**
```python
# Parse a template layout
parser = TemplateSlideParser(template_path)
layout_info = parser.parse_layout(layout_index=1)

# Get safe area for content (avoiding branding)
safe_area = parser.get_content_safe_area(layout_index=1)
# Returns: {'left': 0.5, 'top': 1.5, 'width': 12.3, 'height': 5.0}

# Get human-readable summary
summary = parser.get_layout_summary(layout_index=1)
# Shows: branding elements, content areas, safe zones
```

**Classification Logic:**
- **Branding Elements**:
  - Images in corners (likely logos)
  - Small images (<2" x 2")
  - Text with company keywords (Inc, LLC, ©, etc.)
  - Text at slide edges (headers/footers)
  - Groups in corners (logo + text combinations)
  - Footer/date/number placeholders

- **Content Elements**:
  - Title and body placeholders
  - Large images (>2" in any dimension)
  - Tables and charts
  - Main text areas

- **Decorative Elements**:
  - Thin lines (decorative borders)
  - Very small shapes (<0.5")
  - Background elements

**Output Example:**
```
Layout: Title and Content (Index 1)
Dimensions: 13.3" x 7.5"

Total Elements: 7
  Branding: 2
  Content: 2
  Placeholders: 2
  Decoration: 1

Branding Elements (will be preserved):
  - image: CompanyLogo at (0.5", 0.3")
    [Identified as branding: top-left corner, small size]
  - text: Footer at (0.5", 7.0")
    Text: "© 2024 ACME Corp - Confidential"

Safe Content Area:
  Position: (0.5", 1.8")
  Size: 12.3" x 4.9"
  [Content will be positioned here, automatically avoiding branding]
```

#### Flexible Slide Builder

**Capabilities:**
- Start from actual template slides (not blank)
- Automatically preserve branding elements
- Position content dynamically within safe areas
- Support multiple layout types
- LLM can make layout decisions
- Content never overlaps branding

**Module:** `FlexibleSlideBuilder`

**Basic Usage:**
```python
# Initialize with template
builder = FlexibleSlideBuilder(template_path=Path("corporate_template.pptx"))

# Create slide from template layout
slide = builder.create_slide_from_template(layout_index=1)

# Add content - automatically positioned in safe area
builder.add_content_to_slide(slide, {
    'title': 'Q4 Results',
    'content_type': 'bullets',
    'content': ['Revenue up 25%', 'Expanded to 12 markets', ...]
})

# Save presentation
builder.save(Path("output.pptx"))
```

**Supported Content Types:**
- `'text'`: Paragraph content
- `'bullets'`: Bullet point lists
- `'two_column'`: Side-by-side content with custom ratios
- `'grid'`: Image/content grids with specified rows/columns
- `'image'`: Images with positioning and sizing
- `'custom'`: Exact positioning specified by LLM

**Advanced: Two-Column with Custom Ratio:**
```python
builder.add_content_to_slide(slide, {
    'title': 'Comparison',
    'content_type': 'two_column',
    'content': {
        'ratio': 0.6,  # 60% left, 40% right
        'left': 'Main content gets more space...',
        'right': 'Supporting information...'
    }
})
# Result: Columns automatically sized 60/40, positioned in safe area
```

**Advanced: Image Grid:**
```python
builder.add_content_to_slide(slide, {
    'title': 'Product Gallery',
    'content_type': 'grid',
    'content': {
        'rows': 2,
        'cols': 3,
        'items': [
            {'image': 'product1.png'},
            {'image': 'product2.png'},
            # ... up to 6 items for 2x3 grid
        ]
    }
})
# Result: 2x3 grid automatically positioned in safe area
```

**Advanced: Custom Positioning:**
```python
# LLM can specify exact positions (relative 0-1 within safe area)
builder.add_content_to_slide(slide, {
    'title': 'Custom Layout',
    'content_type': 'custom',
    'content': [
        {'type': 'text', 'left': 0, 'top': 0, 'width': 0.6, 'height': 0.8,
         'content': 'Main content area...', 'font_size': 16},
        {'type': 'text', 'left': 0.65, 'top': 0, 'width': 0.35, 'height': 0.4,
         'content': 'Sidebar info...', 'font_size': 14},
    ]
})
# Result: Elements positioned as specified within safe area
```

**Key Methods:**
```python
# Get safe area for a slide
safe_area = builder.get_content_safe_area(slide)
# Returns: {'left': 0.5, 'top': 1.8, 'width': 12.3, 'height': 4.9}

# Clear placeholder content (keep branding)
builder.clear_placeholder_content(slide)

# Get available layouts
layouts = builder.list_available_layouts()

# Get layout summary
summary = builder.get_layout_summary(layout_index=1)
```

#### Dynamic Layout Engine

**Capabilities:**
- Calculate element positions programmatically (no template required)
- Create common layouts on demand
- Suggest optimal layouts based on content
- Support custom LLM-specified positions
- Work with any slide dimensions

**Module:** `DynamicLayoutEngine`

**When to Use:**
- No template available (creates blank slides with calculated positions)
- Completely custom layouts needed
- LLM wants full creative control
- Non-standard positioning requirements

**Common Layouts:**
```python
engine = DynamicLayoutEngine(slide_width=13.333, slide_height=7.5)

# Title + content
elements = engine.create_title_content_layout()

# Two columns with custom ratio
elements = engine.create_two_column_layout(title=True, column_ratio=0.6)

# Grid layout
elements = engine.create_grid_layout(rows=2, cols=3, title=True)

# Image + text
elements = engine.create_image_text_layout(image_position='left')

# Custom from LLM specification
elements = engine.create_custom_layout([
    {'type': 'title', 'left': 0.5, 'top': 0.5, 'width': 9.0, 'height': 0.8},
    {'type': 'chart', 'left': 0.5, 'top': 1.5, 'width': 5.5, 'height': 4.0},
    # ...
])
```

**Content-Aware Suggestions:**
```python
# Suggest optimal layout based on content
suggested = engine.suggest_layout({
    'has_title': True,
    'has_image': True,
    'has_chart': False,
    'comparison': False
})
# Returns: Appropriate ElementBox layout

suggested = engine.suggest_layout({
    'has_title': True,
    'num_images': 4
})
# Returns: 2x2 grid layout

suggested = engine.suggest_layout({
    'has_title': True,
    'comparison': True
})
# Returns: Two-column layout
```

**Element Box Structure:**
```python
@dataclass
class ElementBox:
    left: float       # Position in inches
    top: float
    width: float      # Size in inches
    height: float
    element_type: str # 'title', 'content', 'image', etc.
    content: Any      # Optional content data
    z_index: int      # Layering order
```

#### LLM Layout Decision Making

**NEW CAPABILITY:** The LLM can now make intelligent layout decisions when creating presentations.

**How It Works:**
1. LLM analyzes content requirements
2. Chooses optimal layout type
3. Specifies layout parameters (ratios, grid size, etc.)
4. Content is positioned automatically within safe areas
5. Branding is always preserved

**Layout Specifications in Outline:**
```json
{
  "slides": [
    {
      "title": "Comparison Slide",
      "layout_config": {
        "type": "two_column",
        "column_ratio": 0.6  // 60/40 split
      },
      "elements": [...]
    },
    {
      "title": "Product Gallery",
      "layout_config": {
        "type": "grid",
        "grid_rows": 2,
        "grid_cols": 3
      },
      "elements": [...]
    },
    {
      "title": "Hero Image Slide",
      "layout_config": {
        "type": "image_text",
        "image_position": "left",
        "image_size": "large"
      },
      "elements": [...]
    }
  ]
}
```

**Layout Decision Guidelines for AI:**

| Content Type | Recommended Layout | Parameters |
|--------------|-------------------|------------|
| **Before/After comparison** | `two_column` | `column_ratio: 0.5` (equal) |
| **Pros/Cons comparison** | `two_column` | `column_ratio: 0.5` (equal) |
| **Main content + sidebar** | `two_column` | `column_ratio: 0.65` (main gets more) |
| **Multiple images (2-6)** | `grid` | `grid_rows: 2, grid_cols: 2/3` |
| **Hero image with text** | `image_text` | `image_position: left/right, image_size: large` |
| **Standard content** | `standard` | Default layout |
| **Complex requirements** | `custom` | Specify exact positions |

**Example AI Decision Process:**
```
User request: "Create slide comparing our approach vs competitors"

AI Analysis:
- Content type: Comparison
- Two distinct groups: us vs them
- Equal emphasis needed

Layout Decision:
{
  "layout_config": {
    "type": "two_column",
    "column_ratio": 0.5  # Equal split for fair comparison
  }
}

Result:
- FlexibleSlideBuilder creates slide from template
- Preserves company logo and footer (branding)
- Creates 50/50 columns within safe content area
- No overlap with branding elements
```

#### Benefits of Flexible Layout System

**For Users:**
✓ Templates provide brand consistency (logos, colors, fonts)
✓ Content layouts not locked to template structure
✓ Professional appearance with creative freedom
✓ No manual positioning needed

**For LLM:**
✓ Can make intelligent layout decisions
✓ Not constrained by predefined layouts
✓ Adapts layout to content requirements
✓ Full control when needed, automation when helpful

**For Quality:**
✓ Branding never accidentally removed or covered
✓ Content always positioned in safe areas
✓ Optimal layouts for each content type
✓ Consistent spacing and alignment

#### Best Practices

**When Using Flexible Layout System:**

1. **Always Start from Template** (if available)
   - Use `FlexibleSlideBuilder` instead of creating blank slides
   - Template provides branding, colors, fonts
   - System automatically preserves branding

2. **Make Layout Decisions Based on Content**
   - Comparison → two-column
   - Multiple images → grid
   - Hero image → image_text
   - Default → standard

3. **Specify Layout Parameters**
   - Column ratios: Consider content importance (60/40 vs 50/50)
   - Grid dimensions: Match number of items (4 items = 2x2, 6 items = 2x3)
   - Image size/position: Match content hierarchy

4. **Trust the Safe Area**
   - System automatically avoids branding
   - Content positioned optimally
   - No need to manually calculate positions

5. **Use Custom Layout for Complex Needs**
   - When standard layouts don't fit
   - Specify relative positions (0-1 within safe area)
   - System converts to absolute positions

**Examples Directory:**
- `examples/dynamic_layout_demo.py` - Dynamic layout engine demonstration
- `examples/flexible_template_demo.py` - Flexible slide builder with template preservation

## Content Generation Guidelines

### When to Use Each Type

| Content Type | Best For | Avoid For |
|--------------|----------|-----------|
| **Text** | Concepts, explanations | Large amounts of data |
| **Bullets** | Key points, lists | Detailed paragraphs |
| **Tables** | Structured data, comparisons | Trends over time |
| **Bar Charts** | Comparing values | Showing parts of whole |
| **Line Charts** | Trends, time series | Static comparisons |
| **Pie Charts** | Parts of whole (2-6 items) | Many categories |
| **Process Flow** | Sequential steps | Non-linear processes |
| **Hierarchy** | Organizational structure | Equal-level comparisons |
| **Venn Diagram** | Overlapping concepts | Unrelated items |
| **Timeline** | Chronological events | Non-temporal data |

### Slide Organization Patterns

**Opening:**
1. Title slide
2. Agenda/Overview
3. Executive summary (optional)

**Body:**
4. Section header
5-N. Content slides (mix of types)

**Closing:**
N+1. Summary/Key takeaways
N+2. Next steps/Call to action
N+3. Q&A or Contact info

### Content Density Rules

**Per Slide:**
- Title: 1 line, max 60 characters
- Bullets: 3-7 points, 40-80 characters each
- Body text: 50-150 words
- Tables: 2-6 columns, 3-10 rows
- Charts: 2-8 data series, 3-12 categories
- Images: 1 large or 2-6 small

## AI Agent Instructions

### Content Planning Phase

When creating outlines:

1. **Analyze Requirements**
   - Topic and summary
   - Target audience
   - Number of slides (if specified)
   - Reference documents

2. **Choose Content Types**
   - Use table above for guidance
   - Mix content types for variety
   - Match type to information

3. **Structure Presentation**
   - Start with title
   - Add section dividers for long presentations
   - Group related content
   - End with summary/action

4. **Specify Elements**
   - For each slide, specify: type, title, content, elements
   - For charts: specify type, data structure
   - For tables: specify headers, rows
   - For diagrams: specify style, items

### Content Generation Phase

When generating specific content:

1. **Follow Best Practices**
   - Apply content density rules
   - Use appropriate fonts/sizes
   - Choose colors from palette
   - Ensure readability

2. **Validate Automatically**
   - Check text fits in boundaries
   - Verify bullet count
   - Validate table dimensions
   - Confirm image sizes

3. **Optimize When Needed**
   - Reduce font size if overflow
   - Split slides if too dense
   - Condense text if too long
   - Suggest chart instead of large table

### Feedback Integration

When user provides feedback:

1. **Understand Intent**
   - What specifically to change
   - Which slide(s) affected
   - Desired outcome

2. **Apply Changes**
   - Modify outline or content
   - Regenerate affected slides
   - Maintain consistency

3. **Validate Results**
   - Ensure changes achieved goal
   - Check for new issues
   - Confirm user satisfaction

## System Prompt Template

When acting as PPTX Agent, use this knowledge:

```
You are PPTX Agent, an AI assistant specialized in creating PowerPoint
presentations. You have access to these capabilities:

CONTENT TYPES:
- Text slides (bullets, paragraphs, quotes, rich formatting)
- Tables (data, comparisons, summaries, advanced cell styling)
- Charts (bar, line, pie, scatter, area)
- Images (single, grids, with text, alt text for accessibility)
- SmartArt-like diagrams (process, cycle, hierarchy, comparison, venn, timeline)
- Custom shapes (flowcharts, icons, callouts, annotations)

ADVANCED FEATURES:
- Speaker notes for presenter guidance
- Hyperlinks (external URLs and internal slide jumps)
- Rich text formatting (bold, italic, colors, multiple fonts)
- Advanced tables (cell merging, individual cell styling)
- Slide management (duplicate, hide, reorder slides)
- Footer and slide numbers
- Slide dimensions (widescreen/standard/custom)
- Image enhancements (alt text, transparency)
- Shape layering (z-order control)

DESIGN FEATURES:
- Multiple layout types
- Custom colors and styling
- Content validation and optimization
- Vision-based slide checking
- Template support

BEST PRACTICES:
- Follow 6x6 rule for bullets (max 6 bullets, 6 words each - adapt as needed)
- Use appropriate content type for data
- Add speaker notes to important/complex slides
- Use hyperlinks for navigation and resources
- Use rich formatting to emphasize key points
- Maintain visual consistency
- Ensure readability (min 18pt for body text)
- Balance content density
- Validate content fits on slides
- Add alt text to images for accessibility

When creating presentations:
1. Analyze user requirements
2. Choose appropriate content types
3. Structure logically (intro → body → conclusion)
4. Validate all content fits
5. Apply professional styling
6. Iterate based on feedback
```

## Troubleshooting

### Common Issues

**Text Overflow:**
- Reduce font size (16pt minimum)
- Split into multiple slides
- Condense content

**Too Many Bullets:**
- Split into 2+ slides
- Convert to table or chart
- Summarize into fewer points

**Chart Too Complex:**
- Reduce data series
- Limit categories
- Use simpler chart type

**Image Quality:**
- Use high-resolution images (150+ DPI)
- Verify dimensions
- Check file format

## Future Capabilities (Planned)

- Video embedding
- Animations and transitions
- Audio narration
- Real-time collaboration
- Export to PDF
- 3D charts
- Advanced chart features (combo charts, secondary axis, data labels)
- Full shape grouping support
- Comments and review tracking
- Equations and mathematical notation

---

**Last Updated:** Based on current implementation
**Version:** 3.0 - Now includes flexible layout system with template intelligence, branding preservation, dynamic positioning, LLM layout decisions, plus speaker notes, hyperlinks, rich text formatting, advanced tables, slide management, footer/slide numbers, slide dimensions, image enhancements, shape layering, and multi-format reference document support
**For AI Agent Use:** Reference this document when planning and creating presentations
