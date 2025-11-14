# Unified PowerPoint MCP Server Architecture

## System Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                         LLM / AI Application                         │
│                    (Claude, GPT-4, Copilot, etc.)                   │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
                               │ MCP Protocol (stdio)
                               │
┌──────────────────────────────▼──────────────────────────────────────┐
│                    UnifiedPowerPointMCPServer                        │
│                           (server.py)                                │
│                                                                      │
│  ┌────────────────────────────────────────────────────────────┐   │
│  │                    Tool Router                              │   │
│  │  • Routes calls to appropriate module                       │   │
│  │  • Maintains shared presentation state                      │   │
│  │  • Handles errors and responses                             │   │
│  └────────────────────────────────────────────────────────────┘   │
│                                                                      │
│  ┌─────────────┬───────────────────────┬────────────────────┐      │
│  │             │                       │                    │      │
│  ▼             ▼                       ▼                    ▼      │
│  ┌─────────────────┐  ┌──────────────────┐  ┌──────────────────┐ │
│  │   Markdown      │  │   MCP Tools      │  │ Material Design  │ │
│  │   Converter     │  │   (35+ tools)    │  │   & Theming      │ │
│  │                 │  │                  │  │                  │ │
│  │ • Parse MD      │  │ • Create pres    │  │ • Themes         │ │
│  │ • Generate PPTX │  │ • Add slides     │  │ • Color advice   │ │
│  │ • Validate      │  │ • Charts/tables  │  │ • Accessibility  │ │
│  │   syntax        │  │ • SmartArt       │  │ • Layout advice  │ │
│  │                 │  │ • Export/merge   │  │                  │ │
│  │ md2ppt.py       │  │ ppt-mcp.py       │  │ material-        │ │
│  │                 │  │                  │  │ design.py        │ │
│  └─────────────────┘  └──────────────────┘  └──────────────────┘ │
│                                                                      │
│  ┌────────────────────────────────────────────────────────────┐   │
│  │              Shared Presentation Dictionary                 │   │
│  │  presentations = {                                          │   │
│  │    "id1": Presentation(...),                                │   │
│  │    "id2": Presentation(...),                                │   │
│  │  }                                                          │   │
│  └────────────────────────────────────────────────────────────┘   │
└──────────────────────────────┬──────────────────────────────────────┘
                               │
                               ▼
                    ┌──────────────────────┐
                    │   python-pptx        │
                    │   Library            │
                    └──────────────────────┘
                               │
                               ▼
                    ┌──────────────────────┐
                    │  PowerPoint Files    │
                    │     (.pptx)          │
                    └──────────────────────┘
```

## Tool Flow Patterns

### Pattern 1: Markdown → PowerPoint

```
LLM generates markdown
        ↓
convert_markdown_to_pptx
        ↓
MarkdownPresentationParser.parse()
        ↓
PowerPointGenerator.generate()
        ↓
python-pptx creates slides
        ↓
Save to file
```

### Pattern 2: Sequential MCP Tools

```
create_presentation(id="demo")
        ↓
add_title_slide(id="demo", ...)
        ↓
add_chart_slide(id="demo", ...)
        ↓
add_smart_art(id="demo", ...)
        ↓
save_presentation(id="demo", path)
```

### Pattern 3: Hybrid with Material Design

```
convert_markdown_to_pptx(...)          ← Generate content
        ↓
Presentation stored in memory
        ↓
add_timeline_slide(id, ...)            ← Add specialized slides
        ↓
add_swot_analysis(id, ...)
        ↓
apply_material_theme(id, theme)        ← Apply design
        ↓
check_accessibility(id, standard)      ← Validate
        ↓
save_presentation(id, path)            ← Export
```

## Data Flow

### Markdown Syntax → Python Objects

```
Markdown with tags
    ↓
YAML frontmatter parsed → PresentationConfig
    ↓
Slide blocks parsed → SlideConfig[]
    ↓
Content parsed → Python dicts/lists
    ↓
PowerPointGenerator → python-pptx objects
    ↓
Presentation saved to disk
```

### MCP Tool Calls → Presentation

```
Tool call JSON
    ↓
UnifiedPowerPointMCPServer.handle_call_tool()
    ↓
Route to specific handler
    ↓
Handler creates/modifies Presentation object
    ↓
Store in shared presentations dict
    ↓
Return TextContent response
```

## Module Dependencies

```
server.py
    ├── imports: md2ppt.py
    │   └── requires: python-pptx, PyYAML, Pillow
    │
    ├── imports: ppt_mcp.py
    │   └── requires: python-pptx, numpy, mcp
    │
    └── imports: material_design.py
        └── requires: python-pptx, colorsys

All modules
    └── require: mcp (Model Context Protocol)
```

## State Management

```
UnifiedPowerPointMCPServer
    │
    ├── self.presentations: Dict[str, Presentation]
    │   └── Shared across all modules
    │
    ├── self.markdown_converter: MarkdownToPowerPoint
    │   └── Independent instance
    │
    ├── self.ppt_server: ExtendedPowerPointServer
    │   └── References self.presentations
    │
    ├── self.material_themes: MaterialDesignThemes
    │   └── Stateless theme generator
    │
    └── self.material_advisor: MaterialDesignAdvisor
        └── Stateless advice generator
```

## Error Handling Flow

```
LLM calls tool
    ↓
server.handle_call_tool()
    ↓
try:
    Route to handler
        ↓
    Execute tool logic
        ↓
    Return success response
    
except Exception as e:
    ↓
    Catch error
        ↓
    Format error message
        ↓
    Return error response
            ↓
            LLM receives error and can retry
```

## Export Formats

```
Presentation Object
    ↓
    ├─→ save_presentation() → .pptx
    │
    ├─→ export_as_pdf() → .pdf (requires additional libs)
    │
    ├─→ export_slides_as_images() → .png/.jpg
    │
    └─→ generate_handouts() → print-ready .pptx
```

## Configuration Files

```
config.json
    ├── server metadata
    ├── template definitions
    └── material theme list

.mcp.json
    ├── MCP protocol config
    ├── tool categories
    └── install instructions

requirements.txt
    └── Python dependencies

__init__.py
    └── Package exports
```

## Tool Categories

```
Unified Server Tools (35+)
    │
    ├── Markdown (3)
    │   ├── convert_markdown_to_pptx
    │   ├── convert_markdown_file_to_pptx
    │   └── validate_markdown_presentation
    │
    ├── Presentation Management (4)
    │   ├── create_presentation
    │   ├── save_presentation
    │   ├── merge_presentations
    │   └── export_as_pdf
    │
    ├── Basic Slides (3)
    │   ├── add_title_slide
    │   ├── add_content_slide
    │   └── add_chart_slide
    │
    ├── Advanced Slides (6)
    │   ├── add_smart_art
    │   ├── add_timeline_slide
    │   ├── add_comparison_slide
    │   ├── add_quote_slide
    │   ├── add_agenda_slide
    │   └── add_swot_analysis
    │
    ├── Material Design (4)
    │   ├── apply_material_theme
    │   ├── get_material_color_palette
    │   ├── get_design_advice
    │   └── check_accessibility
    │
    └── Enhancements (3)
        ├── add_slide_notes
        ├── add_footer
        └── add_qr_code
```

## Testing Flow

```
test_integration.py
    │
    ├─→ Test 1: Markdown Conversion
    │   └── MarkdownToPowerPoint.convert()
    │
    ├─→ Test 2: MCP Tools
    │   └── ExtendedPowerPointServer methods
    │
    ├─→ Test 3: Material Design
    │   └── MaterialDesignThemes & Advisor
    │
    └─→ Test 4: Unified Server
        └── UnifiedPowerPointMCPServer integration
```

## Deployment Options

```
Option 1: Direct Python
    python server.py

Option 2: Module
    python -m md2ppt.server

Option 3: MCP Configuration
    Add to LLM config → Auto-start

Option 4: Docker (future)
    docker run md2ppt-server
```
