#!/usr/bin/env python3
"""
Quick test of the unified PowerPoint server
"""

import asyncio
from pathlib import Path
import sys

# Import the unified server
from unified_pptx_server import UnifiedPowerPointServer

async def test_unified_server():
    """Test the unified server functionality"""
    
    print("="*70)
    print("TESTING UNIFIED POWERPOINT SERVER")
    print("="*70)
    
    server = UnifiedPowerPointServer()
    test_dir = Path("test_output_unified")
    test_dir.mkdir(exist_ok=True)
    
    tests_passed = 0
    tests_failed = 0
    
    # Test 1: Create presentation
    print("\n[1] Testing: Create presentation...")
    try:
        result = await server.create_presentation({
            "presentation_id": "test1",
            "template": "corporate"
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 2: Add title slide
    print("\n[2] Testing: Add title slide...")
    try:
        result = await server.add_title_slide({
            "presentation_id": "test1",
            "title": "Unified Server Test",
            "subtitle": "All features in one place"
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 3: Add content slide
    print("\n[3] Testing: Add content slide...")
    try:
        result = await server.add_content_slide({
            "presentation_id": "test1",
            "title": "Features",
            "content": [
                "Markdown conversion",
                "Granular MCP tools",
                "Material Design themes",
                "Accessibility checking",
                "Charts and SmartArt"
            ]
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 4: Add chart
    print("\n[4] Testing: Add chart slide...")
    try:
        result = await server.add_chart_slide({
            "presentation_id": "test1",
            "title": "Sample Data",
            "chart_type": "column",
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
                {"name": "Sales", "values": [100, 150, 200, 180]}
            ]
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 5: Apply Material theme
    print("\n[5] Testing: Apply Material Design theme...")
    try:
        result = await server.apply_material_theme({
            "presentation_id": "test1",
            "theme_name": "google_blue"
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 6: Save presentation
    print("\n[6] Testing: Save presentation...")
    try:
        output_file = test_dir / "test_unified_all_features.pptx"
        result = await server.save_presentation({
            "presentation_id": "test1",
            "file_path": str(output_file)
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 7: Markdown conversion
    print("\n[7] Testing: Markdown to PPTX conversion...")
    try:
        markdown_content = """---
title: Markdown Test
author: Unified Server
theme: corporate
---

# Welcome
## This is a test

---

# Features

- Easy markdown syntax
- Automatic slide generation
- Professional templates

---

# Charts

## Data visualization made simple
"""
        output_file = test_dir / "test_unified_markdown.pptx"
        result = await server.convert_markdown_to_pptx({
            "markdown_content": markdown_content,
            "output_path": str(output_file)
        })
        print(f"✅ {result[0].text}")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 8: Color palette generation
    print("\n[8] Testing: Material color palette generation...")
    try:
        result = await server.get_material_color_palette({
            "seed_color": "4CAF50"
        })
        print(f"✅ Generated palette: {result[0].text[:100]}...")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Test 9: Accessibility check
    print("\n[9] Testing: Accessibility checking...")
    try:
        result = await server.check_accessibility({
            "background_color": "FFFFFF",
            "text_color": "000000"
        })
        print(f"✅ Accessibility check: {result[0].text[:100]}...")
        tests_passed += 1
    except Exception as e:
        print(f"❌ Failed: {e}")
        tests_failed += 1
    
    # Summary
    print("\n" + "="*70)
    print("TEST SUMMARY")
    print("="*70)
    print(f"✅ Passed: {tests_passed}/9")
    print(f"❌ Failed: {tests_failed}/9")
    print(f"📈 Success Rate: {(tests_passed/9)*100:.1f}%")
    print("\n📁 Output files saved to: test_output_unified/")
    print("="*70)
    
    return tests_failed == 0

if __name__ == "__main__":
    success = asyncio.run(test_unified_server())
    sys.exit(0 if success else 1)
