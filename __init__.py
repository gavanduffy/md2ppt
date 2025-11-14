"""
Unified PowerPoint MCP Server

A comprehensive Model Context Protocol server for creating PowerPoint presentations.
Combines markdown conversion, granular MCP tools, and Material Design theming.
"""

__version__ = "1.0.0"
__author__ = "md2ppt Team"

# Import main components
try:
    from .server import UnifiedPowerPointMCPServer
    from .md2ppt import MarkdownToPowerPoint
    from .ppt_mcp import ExtendedPowerPointServer
    from .material_design import MaterialDesignThemes, MaterialDesignAdvisor
    
    __all__ = [
        'UnifiedPowerPointMCPServer',
        'MarkdownToPowerPoint',
        'ExtendedPowerPointServer',
        'MaterialDesignThemes',
        'MaterialDesignAdvisor'
    ]
except ImportError as e:
    print(f"Warning: Some modules could not be imported: {e}")
    print("Make sure all dependencies are installed: pip install -r requirements.txt")
