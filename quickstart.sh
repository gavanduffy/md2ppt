#!/bin/bash

# Quick Start Script for Unified PowerPoint MCP Server

set -e

echo "=================================================="
echo "Unified PowerPoint MCP Server - Quick Start"
echo "=================================================="
echo ""

# Check Python version
echo "🔍 Checking Python version..."
python3 --version || { echo "❌ Python 3 not found. Please install Python 3.8+"; exit 1; }
echo "✅ Python found"
echo ""

# Install dependencies
echo "📦 Installing dependencies..."
if [ -f requirements.txt ]; then
    pip3 install -r requirements.txt
    echo "✅ Dependencies installed"
else
    echo "❌ requirements.txt not found"
    exit 1
fi
echo ""

# Verify imports
echo "🔍 Verifying module imports..."
python3 -c "from server import UnifiedPowerPointMCPServer" 2>/dev/null && echo "✅ server.py" || echo "❌ server.py import failed"
python3 -c "from md2ppt import MarkdownToPowerPoint" 2>/dev/null && echo "✅ md2ppt.py" || echo "❌ md2ppt.py import failed"
python3 -c "from ppt_mcp import ExtendedPowerPointServer" 2>/dev/null && echo "✅ ppt-mcp.py" || echo "❌ ppt-mcp.py import failed"
python3 -c "from material_design import MaterialDesignThemes" 2>/dev/null && echo "✅ material-design.py" || echo "❌ material-design.py import failed"
echo ""

# Run tests
echo "🧪 Running integration tests..."
if [ -f test_integration.py ]; then
    python3 test_integration.py
    echo ""
else
    echo "⚠️  test_integration.py not found, skipping tests"
    echo ""
fi

# Create example output directory
echo "📁 Creating output directory..."
mkdir -p /tmp/md2ppt_output
echo "✅ Output directory: /tmp/md2ppt_output"
echo ""

# Display configuration
echo "⚙️  Server Configuration:"
if [ -f config.json ]; then
    cat config.json | python3 -m json.tool | head -20
else
    echo "⚠️  config.json not found"
fi
echo ""

# Display next steps
echo "=================================================="
echo "✅ Setup Complete!"
echo "=================================================="
echo ""
echo "🚀 Start the MCP Server:"
echo "   python3 server.py"
echo ""
echo "🧪 Run tests:"
echo "   python3 test_integration.py"
echo ""
echo "📖 Read documentation:"
echo "   - README.md - User guide"
echo "   - INTEGRATION.md - Technical details"
echo "   - INTEGRATION_SUMMARY.md - Quick reference"
echo ""
echo "🔧 Configure with Claude Desktop:"
echo "   Add to claude_desktop_config.json:"
echo '   {'
echo '     "mcpServers": {'
echo '       "powerpoint": {'
echo '         "command": "python3",'
echo '         "args": ["'$(pwd)'/server.py"]'
echo '       }'
echo '     }'
echo '   }'
echo ""
echo "📝 Example Usage:"
echo "   # Method 1: Markdown"
echo "   convert_markdown_to_pptx(content, output.pptx)"
echo ""
echo "   # Method 2: MCP Tools"
echo "   create_presentation(id) → add_slides() → save()"
echo ""
echo "   # Method 3: Material Design"
echo "   apply_material_theme() → check_accessibility()"
echo ""
echo "=================================================="
