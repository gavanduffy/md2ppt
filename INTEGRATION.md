# Integration Guide for Unified PowerPoint MCP Server

## Overview

The Unified PowerPoint MCP Server combines three complementary approaches for AI-driven presentation creation:

1. **Markdown Conversion** (`md2ppt.py`) - Natural language to PowerPoint
2. **MCP Tools** (`ppt_mcp.py`) - Granular programmatic control
3. **Material Design** (`material_design.py`) - Professional theming and design

## Architecture

```
┌─────────────────────────────────────────────────┐
│         UnifiedPowerPointMCPServer              │
│                (server.py)                       │
├─────────────────────────────────────────────────┤
│                                                  │
│  ┌──────────────┐  ┌──────────────┐  ┌────────┐│
│  │  Markdown    │  │  MCP Tools   │  │Material││
│  │  Converter   │  │  35+ Tools   │  │ Design ││
│  │  md2ppt.py   │  │ ppt-mcp.py   │  │ .py    ││
│  └──────────────┘  └──────────────┘  └────────┘│
│                                                  │
└──────────────────┬───────────────────────────────┘
                   │
                   ▼
            ┌──────────────┐
            │  PowerPoint  │
            │   (.pptx)    │
            └──────────────┘
```

## Module Integration

### 1. Import Structure

All modules are now properly integrated with exports:

```python
# From unified server
from server import UnifiedPowerPointMCPServer

# Individual components (if needed)
from md2ppt import MarkdownToPowerPoint
from ppt_mcp import ExtendedPowerPointServer
from material_design import MaterialDesignThemes
```

### 2. Shared State

The unified server shares presentation state across all three modules:

```python
class UnifiedPowerPointMCPServer:
    def __init__(self):
        self.ppt_server = ExtendedPowerPointServer()
        self.presentations = self.ppt_server.presentations  # Shared!
        self.markdown_converter = MarkdownToPowerPoint()
        self.material_advisor = MaterialDesignAdvisor()
```

### 3. Tool Routing

The unified server routes tool calls to appropriate handlers:

```python
@self.server.call_tool()
async def handle_call_tool(name: str, arguments: Dict):
    # Markdown tools
    if name == "convert_markdown_to_pptx":
        return await self.convert_markdown_to_pptx(arguments)
    
    # Material Design tools
    elif name == "apply_material_theme":
        return await self.apply_material_theme(arguments)
    
    # Delegate to ppt_server for other tools
    else:
        handler = getattr(self.ppt_server, name, None)
        return await handler(arguments)
```

## Usage Patterns

### Pattern 1: Pure Markdown

Best for quick content generation:

```python
server = UnifiedPowerPointMCPServer()

result = await server.convert_markdown_to_pptx({
    "markdown_content": markdown_text,
    "output_path": "presentation.pptx"
})
```

### Pattern 2: Pure MCP Tools

Best for precise control:

```python
server = UnifiedPowerPointMCPServer()

# Create and build presentation
await server.create_presentation({"presentation_id": "demo"})
await server.add_title_slide({...})
await server.add_chart_slide({...})
await server.save_presentation({...})
```

### Pattern 3: Hybrid Workflow

Best for complex presentations:

```python
server = UnifiedPowerPointMCPServer()

# 1. Start with markdown for content
await server.convert_markdown_to_pptx({...})

# 2. Enhance with specific tools
await server.add_swot_analysis({...})
await server.add_timeline_slide({...})

# 3. Apply Material Design polish
await server.apply_material_theme({
    "presentation_id": "demo",
    "theme_name": "google_blue"
})

# 4. Check accessibility
await server.check_accessibility({
    "presentation_id": "demo",
    "standard": "WCAG_AA"
})
```

## LLM Integration

### For Claude/GPT-4

The unified server exposes all tools through a single MCP interface:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["-m", "md2ppt.server"],
      "cwd": "/path/to/md2ppt"
    }
  }
}
```

### Tool Categories

LLMs can choose tools based on task:

| Task Type | Recommended Tools |
|-----------|------------------|
| Quick draft | `convert_markdown_to_pptx` |
| Data visualization | `add_chart_slide`, `add_infographic_slide` |
| Project planning | `add_gantt_chart`, `add_timeline_slide` |
| Analysis | `add_swot_analysis`, `add_comparison_slide` |
| Branding | `apply_material_theme`, `add_footer` |
| Accessibility | `check_accessibility`, `get_design_advice` |

## Error Handling

All modules use consistent error handling:

```python
try:
    result = await server.some_tool(args)
    return [types.TextContent(
        type="text",
        text=json.dumps(result, indent=2)
    )]
except Exception as e:
    return [types.TextContent(
        type="text",
        text=f"Error: {str(e)}"
    )]
```

## Testing

Run the integration test suite:

```bash
python test_integration.py
```

This tests:
- ✓ Module imports
- ✓ Markdown conversion
- ✓ MCP tools
- ✓ Material Design
- ✓ Unified server integration

## Dependencies

All dependencies are in `requirements.txt`:

```bash
pip install -r requirements.txt
```

Core dependencies:
- `mcp` - MCP protocol
- `python-pptx` - PowerPoint generation
- `Pillow` - Image processing
- `numpy` - Numerical operations
- `pyyaml` - YAML parsing
- `markdown` - Markdown processing

## Configuration

Server configuration in `config.json`:
- Templates (corporate, creative, academic, minimalist)
- Material Design themes
- Server capabilities

MCP configuration in `.mcp.json`:
- Tool categories
- Install instructions
- Capabilities

## Best Practices

### 1. Module Independence

Each module can work independently:

```python
# Direct markdown usage
from md2ppt import MarkdownToPowerPoint
converter = MarkdownToPowerPoint()
await converter.convert(content, "out.pptx")

# Direct MCP usage
from ppt_mcp import ExtendedPowerPointServer
server = ExtendedPowerPointServer()
await server.create_presentation({...})
```

### 2. Presentation IDs

Use consistent IDs across tools:

```python
# Create with ID
await server.create_presentation({"presentation_id": "q4_report"})

# Reference same ID
await server.add_title_slide({"presentation_id": "q4_report", ...})
await server.apply_material_theme({"presentation_id": "q4_report", ...})
```

### 3. Error Recovery

Handle errors gracefully:

```python
# Validate before converting
result = await server.validate_markdown_presentation({
    "markdown_content": content
})

if result["valid"]:
    await server.convert_markdown_to_pptx({...})
```

## Extending the Server

### Adding New Tools

1. Add tool definition in `handle_list_tools()`
2. Add handler method
3. Update `.mcp.json` tool list

### Adding New Themes

1. Add theme to `MaterialDesignThemes.THEMES`
2. Update `config.json` material_themes list

### Adding New Markdown Features

1. Update `MarkdownPresentationParser.PATTERNS`
2. Add parsing method
3. Add generation method in `PowerPointGenerator`

## Troubleshooting

### Import Errors

```bash
# Install all dependencies
pip install -r requirements.txt

# Test imports
python -c "from server import UnifiedPowerPointMCPServer"
```

### MCP Server Not Starting

```bash
# Check Python path
which python

# Run with verbose output
python -m md2ppt.server --verbose
```

### Presentation Not Saving

```python
# Check presentation exists
if "presentation_id" not in server.presentations:
    print("Presentation not found!")

# Check file path
output_path = Path(file_path)
output_path.parent.mkdir(parents=True, exist_ok=True)
```

## Performance

### Memory Usage

Each presentation stays in memory until saved. For large batches:

```python
# Create and save immediately
await server.create_presentation({...})
# ... add slides ...
await server.save_presentation({...})

# Clear from memory after saving
del server.presentations["presentation_id"]
```

### Concurrent Operations

The server supports concurrent tool calls:

```python
# Can process multiple presentations simultaneously
await asyncio.gather(
    server.create_presentation({"presentation_id": "deck1"}),
    server.create_presentation({"presentation_id": "deck2"}),
    server.create_presentation({"presentation_id": "deck3"})
)
```

## License

MIT License - See LICENSE file for details
