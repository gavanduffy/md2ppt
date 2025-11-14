# Unified PowerPoint MCP Server

A single, comprehensive MCP server that integrates **Markdown conversion**, **granular MCP tools**, and **Material Design themes** for PowerPoint creation.

## 🎯 Overview

This unified server simplifies PowerPoint creation by combining three powerful approaches in one:

1. **Markdown Conversion** - Write presentations in markdown and convert instantly
2. **Granular MCP Tools** - Build presentations programmatically with 20+ tools
3. **Material Design** - Apply professional themes with accessibility checking

## 🚀 Quick Start

### Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Test the server
python3 test_unified.py
```

### Basic Usage

The server provides a single interface for all PowerPoint operations. Here are the three main approaches:

#### 1. Markdown Conversion (Fastest)

```python
# Convert markdown directly to PowerPoint
await server.convert_markdown_to_pptx({
    "markdown_content": """---
title: My Presentation
author: Your Name
theme: corporate
---

# Welcome
## Subtitle here

---

# Key Points
- Point 1
- Point 2
- Point 3
""",
    "output_path": "output.pptx"
})
```

#### 2. Granular Tools (Most Control)

```python
# Build presentations step-by-step
await server.create_presentation({"presentation_id": "my_pres"})
await server.add_title_slide({"presentation_id": "my_pres", "title": "My Title"})
await server.add_content_slide({
    "presentation_id": "my_pres",
    "title": "Content",
    "content": ["Point 1", "Point 2"]
})
await server.add_chart_slide({
    "presentation_id": "my_pres",
    "title": "Data",
    "chart_type": "column",
    "categories": ["Q1", "Q2", "Q3"],
    "series": [{"name": "Sales", "values": [100, 150, 200]}]
})
await server.save_presentation({"presentation_id": "my_pres", "file_path": "output.pptx"})
```

#### 3. Material Design (Professional Themes)

```python
# Apply Material Design themes
await server.apply_material_theme({
    "presentation_id": "my_pres",
    "theme_name": "google_blue"
})

# Check accessibility
await server.check_accessibility({
    "background_color": "FFFFFF",
    "text_color": "000000"
})

# Generate color palettes
await server.get_material_color_palette({"seed_color": "4CAF50"})
```

## 📋 Available Tools

### Presentation Management
- `create_presentation` - Create new presentation with templates
- `save_presentation` - Save to file

### Markdown Conversion
- `convert_markdown_to_pptx` - Convert markdown content
- `convert_markdown_file_to_pptx` - Convert markdown file

### Slide Creation
- `add_title_slide` - Title slide
- `add_content_slide` - Bullet points
- `add_two_column_slide` - Two-column layout
- `add_smartart_slide` - SmartArt diagrams
- `add_timeline_slide` - Timeline visualization
- `add_comparison_slide` - Side-by-side comparison
- `add_quote_slide` - Quote with attribution

### Charts & Data
- `add_chart_slide` - Column, bar, line, or pie charts

### Enhancements
- `add_image_to_slide` - Add images
- `add_qr_code` - Generate and add QR codes
- `add_watermark` - Add watermarks to all slides
- `add_slide_notes` - Add speaker notes

### Material Design
- `apply_material_theme` - Apply Material Design themes
- `get_material_color_palette` - Generate color palettes
- `check_accessibility` - WCAG accessibility checking

## 🎨 Available Templates

- **Corporate** - Professional business presentations
- **Creative** - Bold, colorful designs
- **Academic** - Scholarly presentations
- **Minimalist** - Clean, simple layouts

## 🎨 Material Design Themes

- **material_baseline** - Standard Material Design
- **google_blue** - Google brand colors

## 📊 Example: Complete Workflow

```python
from unified_pptx_server import UnifiedPowerPointServer

server = UnifiedPowerPointServer()

# 1. Start with markdown for quick structure
await server.convert_markdown_to_pptx({
    "markdown_content": "...",
    "output_path": "draft.pptx",
    "presentation_id": "my_pres"
})

# 2. Enhance with granular tools
await server.add_chart_slide({
    "presentation_id": "my_pres",
    "title": "Sales Data",
    "chart_type": "column",
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{"name": "Revenue", "values": [100, 150, 200, 250]}]
})

await server.add_qr_code({
    "presentation_id": "my_pres",
    "slide_index": 0,
    "url": "https://example.com"
})

# 3. Apply Material Design theme
await server.apply_material_theme({
    "presentation_id": "my_pres",
    "theme_name": "google_blue"
})

# 4. Save final version
await server.save_presentation({
    "presentation_id": "my_pres",
    "file_path": "final.pptx"
})
```

## 🧪 Testing

Run the comprehensive test suite:

```bash
python3 test_unified.py
```

This creates sample presentations demonstrating all features:
- `test_unified_all_features.pptx` - All slide types and features
- `test_unified_markdown.pptx` - Markdown conversion

## 🏗️ Architecture

The unified server combines three previously separate components:

```
unified_pptx_server.py (1,400 lines)
├── Markdown Parser & Converter (from md2ppt.py)
├── PowerPoint Tools (from ppt-mcp.py)
└── Material Design Themes (from material-design.py)
```

**Benefits of Integration:**
- ✅ Single entry point for all operations
- ✅ Shared presentation state across methods
- ✅ Simpler configuration and deployment
- ✅ Easier maintenance and testing
- ✅ Consistent API across all features

## 📝 Markdown Syntax

Supports standard markdown plus custom syntax:

### Frontmatter
```yaml
---
title: My Presentation
author: Your Name
theme: corporate
aspect_ratio: 16:9
---
```

### Slides
```markdown
# Slide Title
## Subtitle

Content goes here

---

# Next Slide

- Bullet point 1
- Bullet point 2
```

## 🔧 Configuration

Edit `config.json` to customize:
- Server settings
- Default templates
- Material Design themes
- Color schemes

## 📦 Dependencies

- `python-pptx` - PowerPoint generation
- `PyYAML` - YAML parsing
- `Pillow` - Image processing
- `markdown` - Markdown parsing
- `qrcode` - QR code generation (optional)
- `mcp` - Model Context Protocol

## 🤝 MCP Integration

This server implements the Model Context Protocol (MCP), making it easy to integrate with AI assistants and LLMs:

```json
{
  "mcpServers": {
    "unified-powerpoint": {
      "command": "python3",
      "args": ["unified_pptx_server.py"],
      "cwd": "/path/to/md2ppt"
    }
  }
}
```

## 📄 License

MIT License - See LICENSE file for details

## 🙏 Credits

Integrates and simplifies:
- Original `md2ppt.py` markdown converter
- `ppt-mcp.py` MCP tools server
- `material-design.py` theming system
