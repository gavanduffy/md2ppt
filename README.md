# md2ppt

A powerful toolkit for creating PowerPoint presentations programmatically, designed for LLM integration with multiple interaction patterns.

## Overview

This project provides three complementary approaches for LLMs to create PowerPoint presentations, each optimized for different use cases and interaction styles.

## LLM Interaction Methods

### 1. **Markdown-Based Conversion** (`md2ppt.py`)

The most natural and human-readable approach - LLMs can generate structured Markdown with special tags, which is then converted to PowerPoint.

**Best for:**
- Natural language generation
- Document-style presentations
- Quick prototyping
- Human-readable intermediate format

**How it works:**
```markdown
---
title: My Presentation
author: John Doe
theme: material_blue
---

# Title Slide
## Subtitle

---
<!-- slide: content -->

### Main Points

- Point 1
- Point 2
- Point 3
```

**Supported Features:**
- YAML front matter for global config (title, author, theme, colors, fonts)
- Slide type markers: `<!-- slide: title|section|content|two_column|chart|table|code|quote|timeline|image -->`
- Visual tags: backgrounds, transitions, animations, layouts
- Content elements: images, charts, tables, code blocks, math, timelines, mermaid diagrams
- Advanced features: columns, boxes, inline styling, variables, speaker notes

**LLM Usage Pattern:**
1. Generate Markdown content with special syntax
2. Call `convert_markdown_to_pptx` tool with the Markdown string
3. Receive completed presentation file

**Example:**
```python
from md2ppt import MarkdownToPowerPoint

converter = MarkdownToPowerPoint()
result = await converter.convert(markdown_content, "output.pptx")
```

---

### 2. **MCP Server Interface** (`ppt-mcp.py`)

A comprehensive Model Context Protocol server exposing 35+ granular tools for programmatic presentation creation.

**Best for:**
- Fine-grained control
- Interactive multi-step workflows
- Complex custom layouts
- Integration with other MCP tools

**How it works:**
The LLM calls individual tools in sequence to build presentations piece by piece:

```python
# Tool sequence example
1. create_presentation(id="demo", template="corporate", aspect_ratio="16:9")
2. add_title_slide(id="demo", title="Q4 Results", subtitle="Executive Summary")
3. add_chart_slide(id="demo", title="Revenue Growth", chart_type="column", ...)
4. add_smart_art(id="demo", title="Process Flow", smart_art_type="process", ...)
5. save_presentation(id="demo", file_path="q4-results.pptx")
```

**Available Tool Categories:**

**Core Slides:**
- `create_presentation` - Initialize with templates and aspect ratios
- `add_title_slide` - Title and subtitle slides
- `add_content_slide` - Bullet points and text content
- `add_chart_slide` - Bar, column, line, pie, scatter, bubble, radar charts

**Advanced Slides:**
- `add_smart_art` - Process, cycle, hierarchy, pyramid diagrams
- `add_timeline_slide` - Horizontal/vertical timelines
- `add_comparison_slide` - Multi-column comparisons
- `add_quote_slide` - Quotes and testimonials
- `add_agenda_slide` - Table of contents with timing
- `add_team_slide` - Organization and team layouts
- `add_process_flow` - Linear, circular, branching flows
- `add_infographic_slide` - Data visualizations with icons
- `add_gantt_chart` - Project timelines
- `add_swot_analysis` - SWOT matrix layouts

**Enhancements:**
- `add_animation` - Shape and text animations
- `add_slide_notes` - Speaker notes
- `add_hyperlink` - Interactive links
- `add_video_slide` - Embedded video
- `add_audio` - Background audio
- `add_footer` - Consistent footers
- `add_watermark` - Text watermarks
- `add_qr_code` - QR code generation
- `add_math_equation` - LaTeX equations
- `add_code_block` - Syntax-highlighted code

**Slide Management:**
- `apply_slide_master` - Global themes and branding
- `add_custom_shape` - Rectangles, circles, arrows, stars, hearts
- `duplicate_slide` - Copy existing slides
- `reorder_slides` - Change slide sequence
- `delete_slide` - Remove slides

**Export:**
- `save_presentation` - Save as PPTX
- `export_as_pdf` - PDF export
- `export_slides_as_images` - PNG/JPG/SVG export
- `merge_presentations` - Combine multiple decks
- `generate_handouts` - Print-ready handouts

**LLM Usage Pattern:**
1. Plan presentation structure
2. Call tools sequentially to build slides
3. Iterate and refine based on requirements
4. Export in desired format

**Example Workflow:**
```
User: "Create a business presentation about our Q4 results"

LLM:
1. Analyzes requirements
2. Calls create_presentation with corporate template
3. Adds title slide
4. Adds agenda slide with sections
5. For each section:
   - Adds appropriate slide type (charts, content, comparisons)
   - Adds speaker notes
6. Adds closing/thank you slide
7. Applies footer and branding
8. Saves presentation
```

---

### 3. **Material Design Extension** (`material-design.py`)

A specialized extension applying Material Design principles and aesthetics to presentations.

**Best for:**
- Design-conscious presentations
- Brand consistency
- Accessibility compliance
- Professional polish

**How it works:**
Extends the MCP server with design-specific tools:

**Material Design Tools:**
- `apply_material_theme` - Pre-built themes (Material Baseline, Dark, Google Blue, Spotify Green, Notion Minimal)
- `get_material_color_palette` - Generate complementary, analogous, triadic, monochromatic palettes
- `get_design_advice` - Context-aware guidance for color, typography, layout, spacing, animation, accessibility
- `apply_material_layout` - Hero, cards, list, dashboard patterns
- `add_material_components` - FABs, chips, cards, badges
- `check_accessibility` - WCAG AA/AAA compliance checking
- `optimize_for_device` - Responsive layouts for desktop/tablet/mobile
- `generate_style_guide` - Comprehensive style guide slides
- `apply_material_animations` - Standard, emphasized, expressive motion
- `create_mood_board` - Visual style references

**Material Design Principles:**
- **Color:** Semantic color systems with accessibility
- **Typography:** Clear hierarchy with Material type scale
- **Layout:** 8dp/12-column grid systems
- **Spacing:** Consistent spacing scale
- **Motion:** Responsive, natural animations
- **Elevation:** Depth through shadows
- **Accessibility:** WCAG compliance built-in

**LLM Usage Pattern:**
1. Choose or generate Material theme
2. Get design advice for context (corporate, educational, creative)
3. Build slides using MCP tools
4. Apply Material Design components and layouts
5. Validate accessibility
6. Generate style guide for consistency

**Example:**
```python
# Material You dynamic theme
1. apply_material_theme(id="deck", theme="custom", seed_color="2196F3")

# Get color advice
2. get_material_color_palette(base_color="2196F3", type="complementary")

# Get layout guidance
3. get_design_advice(advice_type="layout", context="corporate")

# Apply professional polish
4. apply_material_layout(id="deck", slide=1, pattern="hero")
5. check_accessibility(id="deck", standard="WCAG_AA")
```

---

## Comparison: When to Use Each Method

| Method | Complexity | Control | Best Use Case |
|--------|-----------|---------|---------------|
| **Markdown** | Low | Medium | Quick generation, content-focused |
| **MCP Server** | Medium | High | Complex layouts, precise control |
| **Material Design** | High | Very High | Professional design, brand consistency |

### Hybrid Approach

LLMs can combine methods for optimal results:

1. **Start with Markdown** for rapid content generation
2. **Use MCP tools** to add charts, diagrams, and special slides
3. **Apply Material Design** for final polish and accessibility

## Example: Complete Workflow

```
User: "Create a professional presentation about our product launch"

LLM Strategy:
1. Generate Markdown outline with main content
   - convert_markdown_to_pptx()
   
2. Enhance with MCP tools:
   - add_gantt_chart() for launch timeline
   - add_swot_analysis() for market analysis
   - add_team_slide() for responsible parties
   
3. Apply Material Design polish:
   - apply_material_theme(theme="google_blue")
   - check_accessibility(standard="WCAG_AA")
   - generate_style_guide()
   
4. Export:
   - save_presentation("launch.pptx")
   - export_as_pdf("launch.pdf")
   - export_slides_as_images() for social media
```

## Installation

```bash
pip install python-pptx pillow pyyaml mcp numpy
```

## Usage Examples

### Direct Python API
```python
# Markdown conversion
from md2ppt import MarkdownToPowerPoint
converter = MarkdownToPowerPoint()
result = await converter.convert(markdown_content, "output.pptx")

# MCP Server tools
from ppt_mcp import ExtendedPowerPointServer
server = ExtendedPowerPointServer()
# Use as MCP server or call methods directly

# Material Design
from material_design import MaterialDesignThemes
theme = MaterialDesignThemes.get_material_you_theme("2196F3")
```

### As MCP Server
```bash
# Run the MCP server
python ppt-mcp.py
```

## Features Summary

- **35+ specialized tools** for presentation creation
- **Material Design 3** theming and components
- **Full Markdown syntax** with PowerPoint extensions
- **Accessibility checking** (WCAG AA/AAA)
- **Multiple export formats** (PPTX, PDF, images)
- **Smart Art & diagrams** (process, cycle, hierarchy, timeline, Gantt, SWOT)
- **Rich media support** (images, video, audio, QR codes)
- **Professional templates** (Corporate, Creative, Academic, Minimalist)
- **Responsive layouts** optimized for different devices

## Architecture

```
┌─────────────────────────────────────────────────┐
│                   LLM Agent                      │
└───────────┬─────────────────────────────────────┘
            │
            ├─── Natural Language
            │    └─> Markdown Generator (md2ppt.py)
            │
            ├─── Tool Calls
            │    └─> MCP Server (ppt-mcp.py)
            │
            └─── Design System
                 └─> Material Design (material-design.py)
                          │
                          ▼
                    ┌──────────────┐
                    │  PowerPoint  │
                    │    (.pptx)   │
                    └──────────────┘
```

## Contributing

Each module is designed to work independently or together. Contributions welcome for:
- New slide templates
- Additional chart types
- Material Design components
- Markdown syntax extensions
- Export format options

## License

MIT License - See LICENSE file for details
