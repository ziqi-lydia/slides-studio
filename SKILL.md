---
name: slides-studio
description: "Create high-quality, fully editable PPTX presentations with professional design. Every element is a native PPTX object — directly editable in PowerPoint, Google Slides, and Keynote. No AI-template look. Supports template-based design extraction and AI image generation."
---

# Slides Studio — Professional Editable PPTX Builder

Generate **high-quality, fully editable PPTX** files where every element (text, shapes, images) is a native PowerPoint object. Users can import into Google Slides and directly edit any element — change titles, modify text, replace images, adjust shapes.

**Key differentiator:** Unlike most AI slide tools that produce image-based or locked slides, every element here is individually editable after import.

## Architecture

```
SKILL.md                          ← You are here
scripts/create_slides.py          ← Python: JSON spec → editable PPTX
```

**Script location:** `{SKILL_DIR}/scripts/create_slides.py`

## When to Use

- User asks to create a presentation / slide deck / pitch deck
- User wants editable PPTX (not HTML or PDF)
- User wants slides importable into Google Slides
- User provides a template to replicate the design style
- User needs professional, non-AI-looking slides

## Workflow

### Phase 1: Gather Requirements

Ask (if not provided):
1. **Topic / content** — What is the presentation about? Do they have an outline or raw content?
2. **Audience** — Investors, team, clients, students?
3. **Slide count** — Default: 8-12
4. **Design reference** — Do they have a template PPTX or screenshot to match?
5. **Tone** — Professional, creative, bold, minimal?

**Two entry paths:**
- **Path A:** User provides outline → review with user → generate
- **Path B:** User provides topic → you create outline → user approves → generate

### Phase 2: Design System

If user provides a template/reference:
1. Analyze the template (screenshot or PPTX)
2. Extract: color palette, fonts, spacing patterns, layout styles
3. Document as a design system object

If no template, select a bold design system. **NEVER default to generic blue/white.**

**Design system structure (JSON):**
```json
{
  "design": {
    "palette": {
      "primary": "#1E2761",
      "secondary": "#CADCFC",
      "accent": "#F96167",
      "background": "#FFFFFF",
      "surface": "#F5F5F5",
      "text_dark": "#1E1E1E",
      "text_light": "#FFFFFF",
      "text_muted": "#888888"
    },
    "fonts": {
      "heading": "Georgia",
      "body": "Calibri",
      "mono": "Consolas"
    },
    "default_background": { "color": "#FFFFFF" }
  }
}
```

**Google Slides compatible fonts (USE THESE):**
- Headings: Georgia, Arial Black, Trebuchet MS, Impact, Palatino, Cambria, Montserrat, Poppins, Playfair Display, Oswald, Lora
- Body: Calibri, Arial, Calibri Light, Garamond, Open Sans, Roboto, Inter, Noto Sans SC (中文)
- ⚠️ Avoid fonts not in Google Fonts — they'll be substituted

### Phase 3: Plan Slide Sequence

Map content to layout archetypes. Professional rhythm:

```
Slide 1:  HERO COVER          — Full-bleed image or bold color + title
Slide 2:  BIG STATEMENT        — One sentence thesis, huge font
Slide 3:  SPLIT (text+image)   — Context with visual
Slide 4:  GRID CARDS           — 3-4 key points in cards
Slide 5:  BIG NUMBER           — Striking statistic, 72pt+
Slide 6:  SECTION DIVIDER      — Dark bg, section title
Slide 7:  SPLIT (image+text)   — Deep dive (mirror of slide 3)
Slide 8:  PROCESS/TIMELINE     — Numbered steps
Slide 9:  COMPARISON           — Before/after or options
Slide 10: DATA/CHART           — Evidence with large callout numbers
Slide 11: QUOTE                — Testimonial or expert quote
Slide 12: CLOSING/CTA          — Next steps + contact
```

**Critical rules:**
- NEVER use the same layout on consecutive slides
- Alternate text-heavy and visual-heavy
- Section divider every 4-5 slides
- Vary background colors (don't make every slide white)

### Phase 4: Generate Images

For slides that need visuals, use `generateImage()`:
- Style prompts to match the deck's design mood
- Save to artifacts/ folder
- Reference absolute paths in the JSON spec

### Phase 5: Build the JSON Spec

The JSON spec drives `create_slides.py`. Full structure:

```json
{
  "dimensions": { "width": 13.333, "height": 7.5 },
  "design": { /* design system from Phase 2 */ },
  "slides": [
    {
      "title": "Slide Name (internal label)",
      "background": { "color": "#1E2761" },
      "elements": [ /* element objects */ ],
      "notes": "Speaker notes here"
    }
  ]
}
```

### Element Types

#### Textbox
```json
{
  "type": "textbox",
  "x": 0.8, "y": 1.5, "w": 5, "h": 2,
  "font_family": "Georgia",
  "font_size": 44,
  "color": "#FFFFFF",
  "bold": true,
  "align": "left",
  "content": "Your Title Here"
}
```

**Rich text with multiple styles in one textbox:**
```json
{
  "type": "textbox",
  "x": 0.8, "y": 1.5, "w": 8, "h": 3,
  "content": {
    "paragraphs": [
      {
        "align": "left",
        "space_after": 12,
        "runs": [
          { "text": "Bold Title", "font_size": 36, "bold": true, "color": "#1E2761", "font_family": "Georgia" }
        ]
      },
      {
        "runs": [
          { "text": "Regular body text with ", "font_size": 16, "color": "#333333", "font_family": "Calibri" },
          { "text": "highlighted words", "font_size": 16, "color": "#F96167", "bold": true, "font_family": "Calibri" },
          { "text": " in accent color.", "font_size": 16, "color": "#333333", "font_family": "Calibri" }
        ]
      }
    ]
  }
}
```

#### Shape
```json
{
  "type": "shape",
  "shape": "rounded_rectangle",
  "x": 1, "y": 2, "w": 3.5, "h": 4,
  "fill": "#F5F5F5",
  "corner_radius": 200,
  "line": { "type": "none" },
  "shadow": { "color": "000000", "opacity": 0.15 }
}
```

Available shapes: rectangle, rounded_rectangle, circle, oval, triangle, diamond, pentagon, hexagon, arrow_right, arrow_left, chevron, star, heart, callout, cloud, cross, donut

#### Image
```json
{
  "type": "image",
  "path": "/absolute/path/to/image.png",
  "x": 6.5, "y": 0, "w": 6.833, "h": 7.5
}
```

#### Line
```json
{
  "type": "line",
  "x": 1, "y": 3, "x2": 12, "y2": 3,
  "color": "#CADCFC",
  "width": 2
}
```

#### Table
```json
{
  "type": "table",
  "x": 0.8, "y": 2, "w": 11.5, "h": 4.5,
  "data": [
    ["Header 1", "Header 2", "Header 3"],
    ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"]
  ],
  "style": {
    "header_fill": "#1E2761",
    "header_text_color": "#FFFFFF",
    "alt_row_fill": "#F5F5F5",
    "font_size": 12,
    "font_family": "Calibri"
  }
}
```

### Background Options

```json
// Solid color
{ "color": "#1E2761" }

// Gradient
{ "gradient": { "stops": [{"color": "#1E2761", "position": 0}, {"color": "#065A82", "position": 1}], "angle": 45 } }

// Background image (full bleed)
{ "image": "/path/to/bg.png" }
```

### Phase 6: Generate PPTX

```bash
python "{SKILL_DIR}/scripts/create_slides.py" spec.json output.pptx
```

Save output to: `artifacts/{deck-name}.pptx`

### Phase 7: QA (REQUIRED)

**Never skip QA.** Convert to images and visually verify:

```bash
# Convert PPTX to images via LibreOffice
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

Or use the PowerPoint skill's thumbnail.py if available.

Check for:
- Overlapping elements
- Text overflow / cut off
- Color contrast issues
- Alignment problems
- Missing images (placeholder boxes)
- Font rendering issues

Fix issues and re-verify until clean.

## Anti-AI Design Rules (CRITICAL)

These rules prevent the "AI-generated template" look:

### DO NOT:
1. **No gradient blue-purple backgrounds** — the #1 AI tell
2. **No centered-everything layouts** — left-align body text, vary alignment
3. **No accent line under every title** — use whitespace instead
4. **No 3-icon-row on every slide** — vary layouts dramatically
5. **No bullet-point walls** — max 4 bullets per slide, use cards/grids instead
6. **No equal-weight color distribution** — one color dominates (60-70%)
7. **No rounded-corner-card-with-icon as default** — it's the AI cliché
8. **No generic stock photo feel** — use AI-generated images with specific prompts

### DO:
1. **Use asymmetric layouts** — text heavy on left, image right (or vice versa)
2. **Vary backgrounds** — alternate between dark/light/colored/image slides
3. **Use typography as design** — giant numbers (72pt+), bold statements, display fonts
4. **Break the grid** — some elements should extend to edges or overlap
5. **Use whitespace generously** — crowded slides look amateur
6. **Create visual rhythm** — loud slide → quiet slide → loud slide
7. **Match fonts to mood** — serif for premium, sans-serif for modern, mixed for editorial

## Layout Templates (Copy-Paste Ready)

### HERO COVER (Dark)
```json
{
  "title": "Cover",
  "background": { "color": "#1E2761" },
  "elements": [
    { "type": "textbox", "x": 0.8, "y": 1.0, "w": 7, "h": 0.5, "content": "PRESENTATION TITLE", "font_family": "Calibri", "font_size": 12, "color": "#CADCFC", "bold": true },
    { "type": "textbox", "x": 0.8, "y": 1.8, "w": 7, "h": 3, "font_family": "Georgia", "font_size": 54, "color": "#FFFFFF", "bold": true, "content": "The Main\nTitle Goes Here" },
    { "type": "textbox", "x": 0.8, "y": 5.5, "w": 5, "h": 0.8, "font_family": "Calibri", "font_size": 16, "color": "#CADCFC", "content": "Subtitle or author name • Date" },
    { "type": "image", "path": "/path/to/hero.png", "x": 8.5, "y": 0.5, "w": 4.3, "h": 6.5 }
  ]
}
```

### BIG STATEMENT
```json
{
  "title": "Statement",
  "background": { "color": "#FFFFFF" },
  "elements": [
    { "type": "shape", "shape": "rectangle", "x": 0, "y": 0, "w": 0.4, "h": 7.5, "fill": "#F96167" },
    { "type": "textbox", "x": 1.2, "y": 1.5, "w": 10, "h": 4.5, "font_family": "Georgia", "font_size": 42, "color": "#1E2761", "bold": true, "align": "left",
      "content": {
        "paragraphs": [
          { "runs": [
            { "text": "The key insight that\nchanges ", "font_size": 42, "color": "#1E2761", "font_family": "Georgia" },
            { "text": "everything.", "font_size": 42, "color": "#F96167", "font_family": "Georgia", "italic": true }
          ]}
        ]
      }
    }
  ]
}
```

### SPLIT (Text + Image)
```json
{
  "title": "Split Layout",
  "background": { "color": "#F5F5F5" },
  "elements": [
    { "type": "textbox", "x": 0.8, "y": 0.6, "w": 5.5, "h": 0.5, "content": "SECTION 01", "font_family": "Calibri", "font_size": 11, "color": "#888888", "bold": true },
    { "type": "textbox", "x": 0.8, "y": 1.2, "w": 5.5, "h": 1.5, "content": "Section Title Here", "font_family": "Georgia", "font_size": 36, "color": "#1E2761", "bold": true },
    { "type": "textbox", "x": 0.8, "y": 3.0, "w": 5.5, "h": 3.5, "content": "Body text goes here. Keep it concise — max 40 words.", "font_family": "Calibri", "font_size": 16, "color": "#333333" },
    { "type": "image", "path": "/path/to/img.png", "x": 7, "y": 0.3, "w": 5.8, "h": 6.9 }
  ]
}
```

### GRID CARDS (2x2)
```json
{
  "title": "Grid Cards",
  "background": { "color": "#FFFFFF" },
  "elements": [
    { "type": "textbox", "x": 0.8, "y": 0.5, "w": 8, "h": 1, "content": "Key Points", "font_family": "Georgia", "font_size": 36, "color": "#1E2761", "bold": true },
    { "type": "shape", "shape": "rounded_rectangle", "x": 0.8, "y": 2, "w": 5.5, "h": 2.3, "fill": "#F5F5F5", "corner_radius": 150, "text": "Card 1 Title\nDescription text here", "font_family": "Calibri", "font_size": 14, "color": "#333333", "vertical_align": "top" },
    { "type": "shape", "shape": "rounded_rectangle", "x": 6.8, "y": 2, "w": 5.5, "h": 2.3, "fill": "#F5F5F5", "corner_radius": 150, "text": "Card 2 Title\nDescription text here", "font_family": "Calibri", "font_size": 14, "color": "#333333", "vertical_align": "top" },
    { "type": "shape", "shape": "rounded_rectangle", "x": 0.8, "y": 4.7, "w": 5.5, "h": 2.3, "fill": "#F5F5F5", "corner_radius": 150, "text": "Card 3 Title\nDescription text here", "font_family": "Calibri", "font_size": 14, "color": "#333333", "vertical_align": "top" },
    { "type": "shape", "shape": "rounded_rectangle", "x": 6.8, "y": 4.7, "w": 5.5, "h": 2.3, "fill": "#F5F5F5", "corner_radius": 150, "text": "Card 4 Title\nDescription text here", "font_family": "Calibri", "font_size": 14, "color": "#333333", "vertical_align": "top" }
  ]
}
```

### BIG NUMBER
```json
{
  "title": "Big Number",
  "background": { "color": "#1E2761" },
  "elements": [
    { "type": "textbox", "x": 0.8, "y": 0.8, "w": 5, "h": 0.5, "content": "BY THE NUMBERS", "font_family": "Calibri", "font_size": 12, "color": "#CADCFC", "bold": true },
    { "type": "textbox", "x": 0.8, "y": 1.5, "w": 6, "h": 3, "content": "95%", "font_family": "Georgia", "font_size": 120, "color": "#F96167", "bold": true },
    { "type": "textbox", "x": 0.8, "y": 4.8, "w": 6, "h": 1.5, "content": "of users reported improved\nproductivity within 30 days", "font_family": "Calibri", "font_size": 22, "color": "#FFFFFF" },
    { "type": "line", "x": 7.5, "y": 1, "x2": 7.5, "y2": 6.5, "color": "#CADCFC", "width": 1 },
    { "type": "textbox", "x": 8, "y": 2, "w": 4.5, "h": 3.5, "content": "Additional context or supporting data points go here.", "font_family": "Calibri", "font_size": 16, "color": "#CADCFC" }
  ]
}
```

### SECTION DIVIDER
```json
{
  "title": "Section Divider",
  "background": { "color": "#F96167" },
  "elements": [
    { "type": "textbox", "x": 0.8, "y": 2.5, "w": 8, "h": 0.5, "content": "PART TWO", "font_family": "Calibri", "font_size": 14, "color": "#FFFFFF", "bold": true },
    { "type": "textbox", "x": 0.8, "y": 3.2, "w": 8, "h": 2, "content": "Section Title", "font_family": "Georgia", "font_size": 52, "color": "#FFFFFF", "bold": true }
  ]
}
```

## Google Slides Compatibility Notes

| Feature | Supported | Notes |
|---------|:---------:|-------|
| Text editing | ✅ | All text boxes fully editable |
| Font/size/color | ✅ | Preserved on import |
| Shapes | ✅ | Resizable, color changeable |
| Images | ✅ | Replaceable, croppable |
| Tables | ✅ | Cell content editable |
| Charts | ⚠️ | Converted to image — use shapes+text for editable data |
| Gradients | ⚠️ | Simple linear OK, complex may degrade |
| Shadows | ⚠️ | Basic shadows preserved |
| Custom fonts | ⚠️ | Must be in Google Fonts catalog |
| Speaker notes | ✅ | Fully preserved |

## Dependencies

```bash
pip install python-pptx lxml
```

The script auto-installs if missing.