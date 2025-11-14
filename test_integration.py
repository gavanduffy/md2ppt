#!/usr/bin/env python3
"""
Test script for the Unified PowerPoint MCP Server
Demonstrates all three interaction methods
"""

import asyncio
import json
from pathlib import Path

# Test imports
try:
    from server import UnifiedPowerPointMCPServer
    print("✓ Server module imported successfully")
except ImportError as e:
    print(f"✗ Failed to import server: {e}")
    exit(1)

try:
    from md2ppt import MarkdownToPowerPoint
    print("✓ Markdown converter imported successfully")
except ImportError as e:
    print(f"✗ Failed to import md2ppt: {e}")
    exit(1)

try:
    from ppt_mcp import ExtendedPowerPointServer
    print("✓ PPT MCP server imported successfully")
except ImportError as e:
    print(f"✗ Failed to import ppt_mcp: {e}")
    exit(1)

try:
    from material_design import MaterialDesignThemes, MaterialDesignAdvisor
    print("✓ Material Design module imported successfully")
except ImportError as e:
    print(f"✗ Failed to import material_design: {e}")
    exit(1)


async def test_markdown_conversion():
    """Test markdown to PowerPoint conversion"""
    print("\n" + "="*60)
    print("TEST 1: Markdown Conversion")
    print("="*60)
    
    converter = MarkdownToPowerPoint()
    
    markdown_content = """---
title: Test Presentation
author: Test User
theme: corporate
---

# Welcome
## Test Subtitle

---
<!-- slide: content -->

### Key Points

- Point 1
- Point 2
- Point 3
"""
    
    try:
        result = await converter.convert(markdown_content, "/tmp/test_markdown.pptx")
        print(f"✓ Markdown conversion successful")
        print(f"  Output: {result}")
    except Exception as e:
        print(f"✗ Markdown conversion failed: {e}")


async def test_mcp_tools():
    """Test MCP tools"""
    print("\n" + "="*60)
    print("TEST 2: MCP Tools")
    print("="*60)
    
    server = ExtendedPowerPointServer()
    
    try:
        # Create presentation
        result = await server.create_presentation({
            "presentation_id": "test",
            "template": "corporate",
            "aspect_ratio": "16:9"
        })
        print(f"✓ Created presentation: {result[0].text}")
        
        # Add title slide
        result = await server.add_title_slide({
            "presentation_id": "test",
            "title": "MCP Test Presentation",
            "subtitle": "Testing Tool Integration"
        })
        print(f"✓ Added title slide: {result[0].text}")
        
        # Add content slide
        result = await server.add_content_slide({
            "presentation_id": "test",
            "title": "Test Content",
            "content": ["Item 1", "Item 2", "Item 3"]
        })
        print(f"✓ Added content slide: {result[0].text}")
        
        # Save presentation
        result = await server.save_presentation({
            "presentation_id": "test",
            "file_path": "/tmp/test_mcp.pptx"
        })
        print(f"✓ Saved presentation: {result[0].text}")
        
    except Exception as e:
        print(f"✗ MCP tools test failed: {e}")


async def test_material_design():
    """Test Material Design integration"""
    print("\n" + "="*60)
    print("TEST 3: Material Design")
    print("="*60)
    
    try:
        # Test color palette generation
        advisor = MaterialDesignAdvisor()
        advice = advisor.get_color_advice("2196F3")
        print(f"✓ Generated color advice")
        print(f"  Psychology: {advice['color_psychology'][:50]}...")
        print(f"  Combinations: {len(advice['recommended_combinations'])} palettes")
        
        # Test theme creation
        themes = MaterialDesignThemes()
        theme = themes.get_material_you_theme("FF5722")
        print(f"✓ Created Material You theme: {theme.name}")
        print(f"  Primary: #{theme.primary_color}")
        print(f"  Secondary: #{theme.secondary_color}")
        
        # Test accessibility checking
        accessibility = advisor._check_accessibility("2196F3")
        print(f"✓ Accessibility check complete")
        print(f"  WCAG AA: {accessibility['wcag_aa_normal']}")
        print(f"  Best text color: #{accessibility['best_text_color']}")
        
    except Exception as e:
        print(f"✗ Material Design test failed: {e}")


async def test_unified_server():
    """Test unified server integration"""
    print("\n" + "="*60)
    print("TEST 4: Unified Server Integration")
    print("="*60)
    
    try:
        server = UnifiedPowerPointMCPServer()
        print("✓ Unified server initialized")
        
        # Test markdown conversion through unified server
        result = await server.convert_markdown_to_pptx({
            "markdown_content": "# Test\n\nSimple test",
            "output_path": "/tmp/test_unified.pptx"
        })
        print(f"✓ Unified markdown conversion: {result[0].text[:50]}...")
        
        # Test Material Design advice through unified server
        result = await server.get_design_advice({
            "advice_type": "color",
            "context": "corporate"
        })
        print(f"✓ Unified design advice: {result[0].text[:50]}...")
        
    except Exception as e:
        print(f"✗ Unified server test failed: {e}")


async def main():
    """Run all tests"""
    print("\n" + "="*60)
    print("UNIFIED POWERPOINT MCP SERVER - TEST SUITE")
    print("="*60)
    
    await test_markdown_conversion()
    await test_mcp_tools()
    await test_material_design()
    await test_unified_server()
    
    print("\n" + "="*60)
    print("TEST SUITE COMPLETE")
    print("="*60)
    print("\nCheck /tmp/ directory for generated presentation files:")
    print("  - test_markdown.pptx")
    print("  - test_mcp.pptx")
    print("  - test_unified.pptx")


if __name__ == "__main__":
    asyncio.run(main())
