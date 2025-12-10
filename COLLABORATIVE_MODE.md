## Collaborative Mode - Iterative Presentation Building

### Overview

Collaborative Mode transforms PPTX Agent into your AI presentation partner. Instead of generating everything at once, the AI works **with you** through an iterative refinement process, incorporating your feedback at every stage.

### The Collaborative Workflow

```
┌─────────────────────────────────────────────────────────┐
│  📋 Step 1: Gather Requirements                         │
│  → Topic, summary, target slides, references            │
└───────────────┬─────────────────────────────────────────┘
                ↓
┌─────────────────────────────────────────────────────────┐
│  📝 Step 2: Outline Planning & Refinement               │
│  → AI creates outline                                   │
│  → You review and provide feedback                      │
│  → AI refines based on your input                       │
│  → Iterate until you approve                            │
└───────────────┬─────────────────────────────────────────┘
                ↓
┌─────────────────────────────────────────────────────────┐
│  🔨 Step 3: Build First Draft                           │
│  → AI builds presentation from approved outline         │
└───────────────┬─────────────────────────────────────────┘
                ↓
┌─────────────────────────────────────────────────────────┐
│  👁️  Step 4: Visual Validation (Optional)               │
│  → AI visually analyzes slides using GPT-4 Vision       │
│  → Checks: content fit, readability, balance            │
│  → Provides specific improvement suggestions            │
└───────────────┬─────────────────────────────────────────┘
                ↓
┌─────────────────────────────────────────────────────────┐
│  🎨 Step 5: Iterative Refinement                        │
│  → Review slides one by one                             │
│  → Provide feedback on specific slides                  │
│  → AI helps implement changes                           │
│  → Iterate until satisfied                              │
└───────────────┬─────────────────────────────────────────┘
                ↓
                ✅ Final Presentation
```

## Usage

### Command Line

```bash
# Basic collaborative mode
python -m pptx_agent --collaborative --output my_presentation.pptx

# With custom template
python -m pptx_agent --collaborative \
  --template company_template.pptx \
  --output q4_results.pptx
```

### Python API

```python
from pathlib import Path
from pptx_agent.cli.collaborative import run_collaborative_mode

# Run collaborative workflow
run_collaborative_mode(
    template_path=Path("template.pptx"),  # Optional
    output_path=Path("output.pptx")
)
```

## Step-by-Step Walkthrough

### Step 1: Gathering Requirements

The AI asks you for:

```
What is your presentation about?
→ Quarterly Business Review

Describe the key points or content to cover:
→ Financial performance in Q4
→ Key achievements and milestones
→ Challenges faced and solutions
→ Goals for Q1 next year

Target number of slides (or press Enter):
→ 12

Path to reference document (optional):
→ q4_financial_data.txt

Path to images directory (optional):
→ assets/q4_images/
```

### Step 2: Outline Refinement

**Initial Outline (AI generates):**

```
PRESENTATION OUTLINE
══════════════════════════════════════════════════════════
Title: Q4 2024 Business Review
Total Slides: 12

Slide Structure:
──────────────────────────────────────────────────────────
1. [TITLE] Q4 2024 Business Review
   Preview: Quarterly performance and achievements

2. [CONTENT] Executive Summary
   Preview: Key highlights and overall performance metrics
   Elements: 1 (bullet_points)

3. [CONTENT] Financial Performance
   Preview: Revenue, expenses, and profitability analysis
   Elements: 2 (table, chart)

4. [SECTION] Achievements
   Preview: Major milestones reached in Q4

5. [CONTENT] Product Launches
   Preview: New products and features released
   Elements: 2 (bullet_points, image)

...
```

**You provide feedback:**

```
Your input: I'd like to split the financial performance into
two slides - one for revenue analysis and another for expenses.
Also, add a slide about team growth and hiring.
```

**AI refines outline:**

```
✓ Outline updated (iteration 1)!

PRESENTATION OUTLINE [UPDATED]
══════════════════════════════════════════════════════════
...
3. [CONTENT] Revenue Analysis
   Preview: Revenue breakdown by product line and region

4. [CONTENT] Expense Analysis
   Preview: Operating expenses and cost optimization

5. [CONTENT] Team Growth
   Preview: Hiring achievements and organizational changes
...
```

**You can iterate multiple times until satisfied:**

```
Your input: approve
✓ Outline approved!
```

### Step 3: Building First Draft

```
🔨 Step 3: Building First Draft

Creating presentation from approved outline...

✓ First draft complete!
   12 slides created
```

### Step 4: Visual Validation (Optional)

If you enable visual validation:

```
👁️  Step 4: Visual Validation (Optional)

The AI can visually analyze your slides to check:
  • Content fits properly
  • Text is readable
  • Layout is balanced
  • Colors and formatting look professional

Run visual validation? (yes/no): yes

Analyzing slides visually...
   Slide 1: Score 9/10, 0 issues found
   Slide 2: Score 7/10, 2 issues found
   Slide 3: Score 8/10, 1 issues found
   ...

VISUAL VALIDATION REPORT
══════════════════════════════════════════════════════════
Total Slides: 12
Successfully Validated: 12
Average Score: 8.2/10

🔴 HIGH PRIORITY ISSUES:

  Slide 3: Text content exceeds slide boundaries
  → Reduce font size from 18pt to 16pt or split content

  Slide 7: Low contrast between text and background
  → Increase contrast ratio to meet 4.5:1 minimum

💡 TOP RECOMMENDATIONS:
  1. Reduce bullet points on slide 5 from 9 to 7 maximum
  2. Align chart legends consistently across slides
  3. Increase white space between sections
  4. Use professional color palette throughout
  5. Ensure all images have consistent sizing
```

### Step 5: Iterative Refinement

```
🎨 Step 5: Refinement

Now you can review and refine individual slides.

Options:
  1. Review and refine specific slide
  2. Review all slides
  3. Add new slide
  4. Regenerate entire presentation from outline
  5. Finish and save

Your choice (1-5): 1

Which slide? (1-12): 3

📄 Slide 3: Revenue Analysis
   Shapes: 5

What would you like to change? (or 'skip' to keep as-is)
  1. Change content/text
  2. Change layout
  3. Add elements (image, chart, etc.)
  4. Remove elements
  5. Styling/formatting changes
  6. Delete this slide
  7. Other (describe)

Your choice (1-7): 5

Describe the changes you want: Make the chart larger and
change the color scheme to match our brand colors

🔧 Applying changes to slide 3...
```

## Vision-Based Validation Details

### What the AI Sees

When visual validation is enabled, the AI:

1. **Converts slides to images** (requires platform tools)
2. **Analyzes each slide visually** using GPT-4 Vision
3. **Evaluates** based on professional design criteria
4. **Provides specific feedback** with actionable fixes

### Validation Criteria

The AI checks for:

- **Content Fit**: Does text/content overflow slide boundaries?
- **Readability**: Are fonts large enough? Is contrast sufficient?
- **Visual Balance**: Is layout balanced and professional?
- **Information Density**: Too much or too little content?
- **Alignment**: Are elements properly aligned?
- **Color & Contrast**: Are colors readable and appropriate?

### Example Vision Feedback

```json
{
  "overall_score": 7,
  "issues": [
    {
      "severity": "high",
      "category": "content_fit",
      "description": "Bullet point list extends beyond slide boundary",
      "suggestion": "Reduce to 6 bullets or split across 2 slides"
    },
    {
      "severity": "medium",
      "category": "readability",
      "description": "Font size on footnote is 10pt (too small)",
      "suggestion": "Increase to minimum 14pt for readability"
    }
  ],
  "strengths": [
    "Professional color scheme",
    "Good use of white space",
    "Clear visual hierarchy"
  ],
  "recommendations": [
    "Consider using larger chart for better visibility",
    "Align all text elements to grid",
    "Add subtle background to make text pop"
  ]
}
```

## Benefits of Collaborative Mode

### vs. Autonomous Mode

| Aspect | Autonomous | Collaborative |
|--------|-----------|---------------|
| User Control | Low | High |
| Speed | Fast | Moderate |
| Customization | Limited | Extensive |
| Iteration | None | Multiple |
| Feedback Integration | No | Yes |
| Best For | Quick drafts | Important presentations |

### When to Use Collaborative Mode

✅ **Use Collaborative Mode When:**
- Creating high-stakes presentations
- You have specific vision for structure/content
- You want to maintain creative control
- Presentation represents your brand
- You need to incorporate detailed requirements
- Working on client deliverables

❌ **Use Other Modes When:**
- Need quick draft/prototype
- Content is straightforward
- You're comfortable with AI decisions
- Time is constrained

## Advanced Features

### Outline Refinement Strategies

**Adding Specific Slides:**
```
feedback: Add a slide about customer testimonials after
the product features section
```

**Reordering:**
```
feedback: Move the pricing slide to come before the
features comparison
```

**Splitting Content:**
```
feedback: The roadmap slide has too much - split into
Q1 roadmap and Q2-Q4 roadmap
```

**Changing Emphasis:**
```
feedback: Put more emphasis on the ROI analysis - expand
that section to 2-3 slides with detailed charts
```

### Multi-Round Refinement

You can go through multiple feedback cycles:

```
Iteration 1: Initial outline
  → Feedback: Add competitive analysis section

Iteration 2: Added competitive analysis
  → Feedback: Actually, make it a comparison table
              instead of separate slides

Iteration 3: Updated to comparison table
  → Feedback: Looks good, but add one slide on market
              trends before the competitive analysis

Iteration 4: Added market trends
  → Feedback: approve
```

## Tips for Effective Collaboration

### 1. Be Specific in Feedback

❌ **Vague:** "This doesn't look right"
✅ **Specific:** "The financial chart is too small - make it the focus of the slide"

### 2. Use Reference Materials

Providing reference documents helps the AI understand context:
- Financial reports
- Product specifications
- Strategy documents
- Previous presentations
- Market research

### 3. Iterate on Outline First

It's easier to fix structure at the outline stage than after slides are built.

### 4. Leverage Visual Validation

If available, use vision validation to catch issues the AI might miss when generating slides.

### 5. Save Intermediate Versions

The system saves after each major step, but you can request saves anytime.

## Troubleshooting

### Vision Validation Not Available

Vision validation requires:
- GPT-4 Vision API access (may require specific OpenAI plan)
- Slide-to-image conversion tools (platform-specific)

Alternatives:
- Manual review in PowerPoint
- Export slides as images manually and ask AI to analyze
- Use autonomous validation (text-based only)

### Outline Refinement Not Converging

If you're stuck in refinement loop:
- Be more specific in feedback
- Break down complex requests
- Consider restarting with clearer initial requirements
- Approve current outline and refine slides individually later

### Slide Changes Not Working

Some changes require manual editing in PowerPoint:
- Complex formatting adjustments
- Precise positioning
- Custom animations
- Embedded objects

The AI will guide you on what needs manual editing.

## Example Session

Complete example of collaborative session:

```bash
$ python -m pptx_agent --collaborative

PPTX Agent - Collaborative Presentation Builder
══════════════════════════════════════════════════════════

This collaborative workflow guides you through:
  1. Gathering requirements
  2. Planning and refining the outline
  3. Building the first draft
  4. Visual validation (optional)
  5. Iterative refinement

Let's get started!

Press Enter to begin...

📋 Step 1: Gathering Requirements
──────────────────────────────────────────────────────────

What is your presentation about?
→ Introduction to Machine Learning for Business Leaders

Describe the key points or content to cover:
→ What is machine learning and why it matters
→ Real-world business applications and ROI
→ Common challenges and how to overcome them
→ Getting started with ML in your organization

Target number of slides:
→ 10

✓ Requirements gathered!

📝 Step 2: Planning & Outline Refinement
──────────────────────────────────────────────────────────

[... outline generation and refinement ...]

🔨 Step 3: Building First Draft
──────────────────────────────────────────────────────────

✓ First draft complete!
   10 slides created

[... optional visual validation ...]

🎨 Step 5: Refinement
──────────────────────────────────────────────────────────

[... iterative refinement ...]

✅ Completion
──────────────────────────────────────────────────────────

Your presentation is ready!

📁 Saved to: ml_for_business.pptx
📊 Total slides: 10
🔄 Refinement iterations: 3

Thank you for using PPTX Agent!
```

## Future Enhancements

Planned improvements to collaborative mode:
- Real-time slide preview in terminal
- More granular control over individual elements
- Undo/redo functionality for changes
- Collaborative sessions (multi-user)
- Version history and branching
- Integration with presentation software
- Advanced animation suggestions
- Speaker notes generation during refinement

## See Also

- [AUTONOMOUS_MODE.md](AUTONOMOUS_MODE.md) - For fully AI-driven generation
- [README.md](README.md) - For general usage and features
- [examples/collaborative_workflow.py](examples/collaborative_workflow.py) - Example code
