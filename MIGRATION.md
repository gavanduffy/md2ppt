# Migration Guide: Moving to Unified Server

This guide helps you transition from the old three-server setup to the new unified server.

## 🔄 What Changed

### Before (3 Separate Servers)
```
server.py           # Orchestration server
├── md2ppt.py      # Markdown conversion
├── ppt-mcp.py     # MCP tools
└── material-design.py  # Material Design
```

### After (1 Unified Server)
```
unified_pptx_server.py  # Everything in one place
```

## ✅ Benefits

1. **Simpler** - One server instead of three
2. **Faster** - No inter-server communication overhead
3. **Easier** - Single configuration file
4. **Cleaner** - Shared presentation state
5. **Maintainable** - One codebase to maintain

## 🚀 Quick Migration

### Step 1: Update Configuration

**Old `config.json`:**
```json
{
  "mcpServers": {
    "unified-powerpoint": {
      "command": "python",
      "args": ["-m", "md2ppt.server"]
    }
  }
}
```

**New `config.json`:**
```json
{
  "mcpServers": {
    "unified-powerpoint": {
      "command": "python3",
      "args": ["unified_pptx_server.py"]
    }
  }
}
```

### Step 2: Update Import Statements

**Old:**
```python
from server import UnifiedPowerPointMCPServer
```

**New:**
```python
from unified_pptx_server import UnifiedPowerPointServer
```

### Step 3: API Changes (Minimal)

The API is mostly the same. Key changes:

#### Markdown Conversion
```python
# Old
from md2ppt import MarkdownToPowerPoint
converter = MarkdownToPowerPoint()
await converter.convert(markdown_content, output_path)

# New (simpler)
server = UnifiedPowerPointServer()
await server.convert_markdown_to_pptx({
    "markdown_content": markdown_content,
    "output_path": output_path
})
```

#### MCP Tools
```python
# Old
from ppt_mcp import ExtendedPowerPointServer
server = ExtendedPowerPointServer()
await server.create_presentation({"presentation_id": "test"})

# New (same API)
server = UnifiedPowerPointServer()
await server.create_presentation({"presentation_id": "test"})
```

#### Material Design
```python
# Old
from material_design import MaterialDesignThemes
themes = MaterialDesignThemes.get_themes()

# New (integrated)
server = UnifiedPowerPointServer()
themes = server.material_themes
await server.apply_material_theme({
    "presentation_id": "test",
    "theme_name": "google_blue"
})
```

## 📝 Tool Name Changes

All tool names remain the same! No changes needed to existing tool calls.

## 🧪 Testing Your Migration

```bash
# Run the unified test
python3 test_unified.py

# All tests should pass (9/9)
```

## 🔍 Key Differences

### 1. Single Server Instance
**Before:**
```python
# Had to coordinate between servers
md2ppt_server = MarkdownToPowerPoint()
ppt_server = ExtendedPowerPointServer()
material_server = MaterialDesignPowerPointExtension(ppt_server)
```

**After:**
```python
# One server does everything
server = UnifiedPowerPointServer()
```

### 2. Shared Presentation State
**Before:**
```python
# Presentations were isolated between modules
converter.convert(markdown, "output.pptx")
# Can't access the presentation object to modify it
```

**After:**
```python
# Presentations are accessible across all methods
await server.convert_markdown_to_pptx({
    "markdown_content": markdown,
    "output_path": "output.pptx",
    "presentation_id": "my_pres"  # Store it!
})

# Now you can enhance it
await server.add_chart_slide({"presentation_id": "my_pres", ...})
await server.apply_material_theme({"presentation_id": "my_pres", ...})
```

### 3. Simplified Imports
**Before:**
```python
import importlib.util

# Complex imports for hyphenated filenames
spec = importlib.util.spec_from_file_location("ppt_mcp", "ppt-mcp.py")
ppt_mcp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ppt_mcp)
```

**After:**
```python
# Simple, standard imports
from unified_pptx_server import UnifiedPowerPointServer
```

## 🎯 Migration Checklist

- [ ] Update `config.json` to point to `unified_pptx_server.py`
- [ ] Update import statements in your code
- [ ] Update any direct file references (e.g., `ppt-mcp.py` → `unified_pptx_server.py`)
- [ ] Run `python3 test_unified.py` to verify
- [ ] Test your existing workflows
- [ ] Update documentation/scripts that reference old files

## 🐛 Troubleshooting

### Issue: "Module not found"
**Solution:** Make sure you're importing from `unified_pptx_server`:
```python
from unified_pptx_server import UnifiedPowerPointServer
```

### Issue: "Presentation not found"
**Solution:** Make sure you're using the same `presentation_id` across calls:
```python
await server.create_presentation({"presentation_id": "my_pres"})
await server.add_title_slide({"presentation_id": "my_pres", ...})  # Same ID
```

### Issue: Old files still referenced
**Solution:** Search your codebase for references to old files:
```bash
grep -r "server.py\|ppt-mcp.py\|md2ppt.py\|material-design.py" .
```

## 📚 Further Reading

- See `UNIFIED_README.md` for complete API documentation
- See `test_unified.py` for working examples
- See original files (`ppt-mcp.py`, etc.) for implementation details (if needed)

## ❓ FAQ

**Q: Can I still use the old three-server setup?**  
A: Yes, the old files are still present for reference, but the unified server is recommended.

**Q: Are all features supported?**  
A: Yes! The unified server includes all features from all three original modules.

**Q: Is the API backward compatible?**  
A: Yes, the tool names and parameters are identical. Only the server instance is different.

**Q: What about performance?**  
A: The unified server is actually faster because there's no inter-server communication overhead.

**Q: Can I contribute?**  
A: Yes! The unified server is easier to contribute to since it's all in one file.

## 🎉 You're Ready!

The unified server is simpler, faster, and easier to use. Happy presenting! 🚀
