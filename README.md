# 🎨 Slides Studio

> Create high-quality, fully editable PPTX presentations with AI. Every element is a native PowerPoint object — directly editable in PowerPoint, Google Slides, and Keynote. No "AI-generated template" aesthetic.

## ✨ What Makes This Different

| Feature | Slides Studio | Typical AI Slides |
|---|---|---|
| **Output format** | Native PPTX (Open XML) | Images / PDF / locked templates |
| **Editability** | Every text box, shape, image, table is editable | Screenshot-like, can't modify |
| **Design quality** | Extracted from real designer templates | Generic AI template look |
| **Google Slides** | Full import & edit support | Limited or broken |
| **AI images** | Custom generated per slide | Stock photos or clip art |

## 🚀 Quick Start

### 1. Download the Skill

```bash
git clone https://github.com/ziqi-lydia/slides-studio.git
```

Or download as ZIP: **Code → Download ZIP**

### 2. Import into Sai

1. Open the **Sai** app
2. Go to **Settings → Skills**
3. Click **Import Skill**
4. Select the downloaded `slides-studio` folder (or ZIP file)
5. The skill will appear in your skill library

### 3. Start Creating

Just chat with Sai! Example prompts:

```
"Create a 10-slide pitch deck for a fintech startup called PayFlow"
```

```
"Make a presentation about our Q3 marketing results, use a dark modern theme"
```

```
"Build a product launch deck for a gaming headset, reference this template style: [paste URL]"
```

Sai will:
1. Research your topic and generate content
2. Extract a design system from a reference template (or use built-in design rules)
3. Generate custom AI images for decorative elements
4. Build a fully editable PPTX file
5. Upload to Google Drive (optional) for instant editing in Google Slides

## 📁 Repo Structure

```
slides-studio/
├── SKILL.md              # Skill definition (Sai reads this)
├── scripts/
│   └── create_slides.py  # Core PPTX generation engine
├── requirements.txt      # Python dependencies
├── LICENSE
└── README.md
```

## 🛠️ How It Works

### Architecture

```
User Prompt
    ↓
Content Generation (AI research + outline)
    ↓
Design System Extraction (from reference template or built-in rules)
    ↓
AI Image Generation (decorative 3D shapes, product mockups, etc.)
    ↓
JSON Slide Specification (precise positioning, fonts, colors)
    ↓
create_slides.py → Native PPTX
    ↓
Google Drive Upload (optional) → Edit in Google Slides
```

### The JSON Spec Format

Each slide is defined as a JSON object with pixel-precise positioning:

```json
{
  "slides": [
    {
      "slide_number": 1,
      "background": { "color": "#0A1628" },
      "elements": [
        {
          "type": "text",
          "content": "PayFlow",
          "position": { "left": 1.2, "top": 2.5, "width": 7.5, "height": 1.5 },
          "style": {
            "font_family": "Helvetica Neue",
            "font_size": 54,
            "font_color": "#FFFFFF",
            "bold": true,
            "alignment": "left"
          }
        },
        {
          "type": "image",
          "source": "path/to/generated-image.png",
          "position": { "left": 6.0, "top": 1.0, "width": 3.5, "height": 3.5 }
        }
      ]
    }
  ]
}
```

### Supported Element Types

| Type | Description | Key Properties |
|---|---|---|
| `text` | Text boxes with full formatting | font, size, color, bold, italic, alignment |
| `shape` | Rectangles, circles, rounded rects | fill color, border, opacity, rotation |
| `image` | Local files or URLs | auto-scaling, positioning |
| `table` | Data tables with styled headers | row/col colors, borders, font styles |
| `line` | Decorative lines and dividers | color, width, dash style |

## 🎨 Design Philosophy

### Anti-AI-Slop Rules

This skill follows strict design rules to avoid the typical "AI-generated" look:

- **No gradient overload** — solid colors or subtle single-direction gradients only
- **No centered-everything layouts** — asymmetric, editorial-style positioning
- **No clip art** — custom AI-generated images (3D shapes, abstract art, product mockups)
- **No bullet-point walls** — visual hierarchy with varied text sizes and whitespace
- **Real typography** — designer font pairings, not system defaults
- **Consistent color system** — extracted from reference templates, max 4-5 colors

### Template-Driven Design

Instead of generating designs from scratch, the skill extracts design systems from real designer templates:

1. **Color palette** — primary, secondary, accent, background, text colors
2. **Typography** — heading/body font families, size scale, weight usage
3. **Layout patterns** — margin ratios, element spacing, grid alignment
4. **Visual style** — shape usage, image treatment, decorative elements

## 🔧 Running Standalone

You can also use `create_slides.py` independently:

```bash
# Install dependencies
pip install -r requirements.txt

# Generate a PPTX from a JSON spec
python scripts/create_slides.py input_spec.json output.pptx
```

### Python API

```python
from scripts.create_slides import create_presentation

spec = {
    "slides": [
        {
            "slide_number": 1,
            "background": {"color": "#FFFFFF"},
            "elements": [
                {
                    "type": "text",
                    "content": "Hello World",
                    "position": {"left": 1, "top": 3, "width": 8, "height": 1.5},
                    "style": {"font_size": 48, "bold": True, "alignment": "center"}
                }
            ]
        }
    ]
}

create_presentation(spec, "output.pptx")
```

## 📋 Requirements

- Python 3.8+
- `python-pptx` — PPTX generation
- `Pillow` — Image processing
- `requests` — Downloading images from URLs

## 🤝 Compatibility

| Platform | Status |
|---|---|
| Microsoft PowerPoint | ✅ Full support |
| Google Slides (import) | ✅ Full support, all elements editable |
| Apple Keynote | ✅ Supported via PPTX import |
| LibreOffice Impress | ✅ Supported |

## 📄 License

MIT License — see [LICENSE](LICENSE)

---

**Built with [Sai](https://simular.ai)** — the AI coworker that does real work on your computer.

