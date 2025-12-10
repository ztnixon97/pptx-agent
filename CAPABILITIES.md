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
- Parse text files (.txt, .md)
- Extract key information
- Generate content based on factual data
- Maintain context from reference materials

**Best Practices:**
- Use references for factual presentations
- Extract specific data points
- Cite sources when appropriate

### Template Support

**Capabilities:**
- Load and analyze PowerPoint templates (.pptx)
- Extract available layouts
- Use template color schemes
- Identify placeholder positions
- Suggest optimal layouts for content types

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
- Text slides (bullets, paragraphs, quotes)
- Tables (data, comparisons, summaries)
- Charts (bar, line, pie, scatter, area)
- Images (single, grids, with text)
- SmartArt-like diagrams (process, cycle, hierarchy, comparison, venn, timeline)
- Custom shapes (flowcharts, icons, callouts, annotations)

DESIGN FEATURES:
- Multiple layout types
- Custom colors and styling
- Content validation and optimization
- Vision-based slide checking
- Template support

BEST PRACTICES:
- Follow 6x6 rule for bullets (max 6 bullets, 6 words each - adapt as needed)
- Use appropriate content type for data
- Maintain visual consistency
- Ensure readability (min 18pt for body text)
- Balance content density
- Validate content fits on slides

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
- Advanced SmartArt conversion
- 3D charts
- Interactive elements

---

**Last Updated:** Based on current implementation
**Version:** 1.0
**For AI Agent Use:** Reference this document when planning and creating presentations
