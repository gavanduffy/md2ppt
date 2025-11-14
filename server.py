#!/usr/bin/env python3
"""
Unified MCP Server for PowerPoint Creation
Integrates markdown conversion, MCP tools, and Material Design
"""

import asyncio
import json
from pathlib import Path
from typing import Any, Dict, List, Optional
import importlib.util

from mcp.server.models import InitializationOptions
from mcp.server import NotificationOptions, Server
from mcp.server.stdio import stdio_server
import mcp.types as types

# Import the three main components
from md2ppt import MarkdownToPowerPoint, MarkdownPresentationParser, PowerPointGenerator

# Import ppt-mcp using importlib (hyphenated filename)
spec = importlib.util.spec_from_file_location("ppt_mcp", str(Path(__file__).parent / "ppt-mcp.py"))
ppt_mcp_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ppt_mcp_module)
ExtendedPowerPointServer = ppt_mcp_module.ExtendedPowerPointServer

# Import material-design using importlib (hyphenated filename)
spec2 = importlib.util.spec_from_file_location("material_design", str(Path(__file__).parent / "material-design.py"))
material_design_module = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(material_design_module)
MaterialDesignThemes = material_design_module.MaterialDesignThemes
MaterialDesignAdvisor = material_design_module.MaterialDesignAdvisor
MaterialDesignPowerPointExtension = material_design_module.MaterialDesignPowerPointExtension
MaterialColorPalette = material_design_module.MaterialColorPalette


class UnifiedPowerPointMCPServer:
    """
    Unified MCP Server combining:
    1. Markdown-to-PowerPoint conversion
    2. Granular MCP tools for presentation building
    3. Material Design theming and best practices
    """

    def __init__(self):
        self.server = Server("unified-powerpoint-server")
        
        # Initialize component modules
        self.markdown_converter = MarkdownToPowerPoint()
        self.ppt_server = ExtendedPowerPointServer()
        self.material_advisor = MaterialDesignAdvisor()
        
        # Share presentations dict across components
        self.presentations = self.ppt_server.presentations
        
        # Material Design themes
        self.material_themes = MaterialDesignThemes()
        
        self.setup_handlers()

    def setup_handlers(self):
        """Setup all MCP handlers"""

        @self.server.list_tools()
        async def handle_list_tools() -> list[types.Tool]:
            """List all available tools from all three modules"""
            
            tools = [
                # ============================================
                # MARKDOWN CONVERSION TOOLS
                # ============================================
                types.Tool(
                    name="convert_markdown_to_pptx",
                    description="Convert Markdown content with PowerPoint tags to PPTX. Supports YAML frontmatter, slide types, charts, tables, timelines, and more. Best for natural language generation.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "markdown_content": {
                                "type": "string",
                                "description": "Markdown content with PowerPoint-specific tags and syntax"
                            },
                            "output_path": {
                                "type": "string",
                                "description": "Path for output .pptx file"
                            },
                            "template": {
                                "type": "string",
                                "description": "Optional template to apply: default, corporate, creative, academic",
                                "default": "default"
                            }
                        },
                        "required": ["markdown_content", "output_path"]
                    }
                ),

                types.Tool(
                    name="convert_markdown_file_to_pptx",
                    description="Convert Markdown file to PowerPoint presentation",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "markdown_file": {
                                "type": "string",
                                "description": "Path to markdown (.md) file"
                            },
                            "output_path": {
                                "type": "string",
                                "description": "Path for output .pptx file"
                            }
                        },
                        "required": ["markdown_file", "output_path"]
                    }
                ),

                types.Tool(
                    name="validate_markdown_presentation",
                    description="Validate Markdown presentation syntax before conversion",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "markdown_content": {
                                "type": "string",
                                "description": "Markdown content to validate"
                            }
                        },
                        "required": ["markdown_content"]
                    }
                ),

                # ============================================
                # PRESENTATION MANAGEMENT TOOLS
                # ============================================
                types.Tool(
                    name="create_presentation",
                    description="Create a new PowerPoint presentation with advanced options",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "template": {"type": "string"},
                            "aspect_ratio": {
                                "type": "string",
                                "description": "16:9, 4:3, custom",
                                "default": "16:9"
                            },
                            "orientation": {
                                "type": "string",
                                "description": "landscape, portrait",
                                "default": "landscape"
                            }
                        },
                        "required": ["presentation_id"]
                    }
                ),

                types.Tool(
                    name="save_presentation",
                    description="Save the presentation to a file",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "file_path": {"type": "string"}
                        },
                        "required": ["presentation_id", "file_path"]
                    }
                ),

                # ============================================
                # BASIC SLIDE TOOLS
                # ============================================
                types.Tool(
                    name="add_title_slide",
                    description="Add a title slide to the presentation",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "subtitle": {"type": "string"},
                            "author": {"type": "string"}
                        },
                        "required": ["presentation_id", "title"]
                    }
                ),

                types.Tool(
                    name="add_content_slide",
                    description="Add a content slide with title and bullet points",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "content": {
                                "type": "array",
                                "items": {"type": "string"}
                            },
                            "layout": {
                                "type": "string",
                                "description": "Layout: bullets, two_column, comparison",
                                "default": "bullets"
                            }
                        },
                        "required": ["presentation_id", "title", "content"]
                    }
                ),

                types.Tool(
                    name="add_chart_slide",
                    description="Add a slide with a chart (bar, column, line, pie, scatter, bubble, radar)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "chart_type": {
                                "type": "string",
                                "description": "Chart type: bar, column, line, pie, scatter, bubble, radar, waterfall"
                            },
                            "categories": {
                                "type": "array",
                                "items": {"type": "string"}
                            },
                            "series_data": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "name": {"type": "string"},
                                        "values": {
                                            "type": "array",
                                            "items": {"type": "number"}
                                        }
                                    }
                                }
                            }
                        },
                        "required": ["presentation_id", "title", "chart_type", "categories", "series_data"]
                    }
                ),

                # ============================================
                # ADVANCED SLIDE TOOLS
                # ============================================
                types.Tool(
                    name="add_smart_art",
                    description="Add SmartArt diagram (process, cycle, hierarchy, pyramid)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "smart_art_type": {
                                "type": "string",
                                "description": "process, cycle, hierarchy, relationship, matrix, pyramid"
                            },
                            "items": {
                                "type": "array",
                                "items": {"type": "string"}
                            },
                            "color_scheme": {"type": "string"}
                        },
                        "required": ["presentation_id", "smart_art_type", "items"]
                    }
                ),

                types.Tool(
                    name="add_timeline_slide",
                    description="Add a timeline slide with events",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "events": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "date": {"type": "string"},
                                        "title": {"type": "string"},
                                        "description": {"type": "string"}
                                    }
                                }
                            },
                            "style": {
                                "type": "string",
                                "description": "horizontal, vertical, curved, zigzag"
                            }
                        },
                        "required": ["presentation_id", "title", "events"]
                    }
                ),

                types.Tool(
                    name="add_comparison_slide",
                    description="Add a comparison slide with multiple columns",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "items": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "header": {"type": "string"},
                                        "features": {"type": "array", "items": {"type": "string"}},
                                        "highlight": {"type": "boolean"}
                                    }
                                }
                            }
                        },
                        "required": ["presentation_id", "title", "items"]
                    }
                ),

                types.Tool(
                    name="add_quote_slide",
                    description="Add a quote/testimonial slide",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "quote": {"type": "string"},
                            "author": {"type": "string"},
                            "title": {"type": "string"}
                        },
                        "required": ["presentation_id", "quote", "author"]
                    }
                ),

                types.Tool(
                    name="add_agenda_slide",
                    description="Add an agenda/table of contents slide",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string", "default": "Agenda"},
                            "sections": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "title": {"type": "string"},
                                        "duration": {"type": "string"},
                                        "presenter": {"type": "string"}
                                    }
                                }
                            }
                        },
                        "required": ["presentation_id", "sections"]
                    }
                ),

                types.Tool(
                    name="add_swot_analysis",
                    description="Add a SWOT analysis slide",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "title": {"type": "string"},
                            "strengths": {"type": "array", "items": {"type": "string"}},
                            "weaknesses": {"type": "array", "items": {"type": "string"}},
                            "opportunities": {"type": "array", "items": {"type": "string"}},
                            "threats": {"type": "array", "items": {"type": "string"}}
                        },
                        "required": ["presentation_id", "strengths", "weaknesses", "opportunities", "threats"]
                    }
                ),

                # ============================================
                # MATERIAL DESIGN TOOLS
                # ============================================
                types.Tool(
                    name="apply_material_theme",
                    description="Apply Material Design theme (material_baseline, material_dark, google_blue, spotify_green, notion_minimal, or custom with seed color)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "theme_name": {
                                "type": "string",
                                "description": "Theme: material_baseline, material_dark, google_blue, spotify_green, notion_minimal, or custom"
                            },
                            "seed_color": {
                                "type": "string",
                                "description": "Hex color for Material You dynamic theme (if theme is 'custom')"
                            },
                            "dark_mode": {"type": "boolean", "default": False}
                        },
                        "required": ["presentation_id", "theme_name"]
                    }
                ),

                types.Tool(
                    name="get_material_color_palette",
                    description="Get Material Design color palette suggestions (complementary, analogous, triadic, monochromatic)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "base_color": {"type": "string", "description": "Hex color code"},
                            "palette_type": {
                                "type": "string",
                                "description": "complementary, analogous, triadic, monochromatic"
                            }
                        },
                        "required": ["base_color"]
                    }
                ),

                types.Tool(
                    name="get_design_advice",
                    description="Get Material Design advice (color, typography, layout, spacing, animation, accessibility)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "advice_type": {
                                "type": "string",
                                "description": "Type: color, typography, layout, spacing, animation, accessibility"
                            },
                            "context": {
                                "type": "string",
                                "description": "Context: corporate, educational, creative, technical, marketing"
                            }
                        },
                        "required": ["advice_type"]
                    }
                ),

                types.Tool(
                    name="check_accessibility",
                    description="Check presentation accessibility against WCAG standards",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "standard": {
                                "type": "string",
                                "description": "Standard: WCAG_AA, WCAG_AAA",
                                "default": "WCAG_AA"
                            }
                        },
                        "required": ["presentation_id"]
                    }
                ),

                # ============================================
                # ENHANCEMENT TOOLS
                # ============================================
                types.Tool(
                    name="add_slide_notes",
                    description="Add speaker notes to a slide",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "slide_index": {"type": "integer"},
                            "notes": {"type": "string"}
                        },
                        "required": ["presentation_id", "slide_index", "notes"]
                    }
                ),

                types.Tool(
                    name="add_footer",
                    description="Add footer to all slides",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "footer_text": {"type": "string"},
                            "show_slide_number": {"type": "boolean"},
                            "show_date": {"type": "boolean"}
                        },
                        "required": ["presentation_id"]
                    }
                ),

                types.Tool(
                    name="add_qr_code",
                    description="Add QR code to slide",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "slide_index": {"type": "integer"},
                            "data": {"type": "string", "description": "URL or text to encode"},
                            "position": {
                                "type": "string",
                                "description": "center, top-right, bottom-right",
                                "default": "bottom-right"
                            }
                        },
                        "required": ["presentation_id", "slide_index", "data"]
                    }
                ),

                # ============================================
                # EXPORT TOOLS
                # ============================================
                types.Tool(
                    name="export_as_pdf",
                    description="Export presentation as PDF",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "presentation_id": {"type": "string"},
                            "output_path": {"type": "string"}
                        },
                        "required": ["presentation_id", "output_path"]
                    }
                ),

                types.Tool(
                    name="merge_presentations",
                    description="Merge multiple presentations into one",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "target_id": {"type": "string"},
                            "source_ids": {
                                "type": "array",
                                "items": {"type": "string"}
                            }
                        },
                        "required": ["target_id", "source_ids"]
                    }
                ),
            ]

            return tools

        @self.server.call_tool()
        async def handle_call_tool(
            name: str, arguments: Optional[Dict[str, Any]]
        ) -> list[types.TextContent]:
            """Route tool calls to appropriate handler"""

            try:
                # Markdown conversion tools
                if name == "convert_markdown_to_pptx":
                    return await self.convert_markdown_to_pptx(arguments)
                elif name == "convert_markdown_file_to_pptx":
                    return await self.convert_markdown_file_to_pptx(arguments)
                elif name == "validate_markdown_presentation":
                    return await self.validate_markdown_presentation(arguments)
                
                # Material Design tools
                elif name == "apply_material_theme":
                    return await self.apply_material_theme(arguments)
                elif name == "get_material_color_palette":
                    return await self.get_material_color_palette(arguments)
                elif name == "get_design_advice":
                    return await self.get_design_advice(arguments)
                elif name == "check_accessibility":
                    return await self.check_accessibility(arguments)
                
                # Delegate to ppt_server for all other tools
                else:
                    handler = getattr(self.ppt_server, name, None)
                    if handler and callable(handler):
                        return await handler(arguments)
                    else:
                        raise ValueError(f"Unknown tool: {name}")

            except Exception as e:
                return [types.TextContent(
                    type="text",
                    text=f"Error executing {name}: {str(e)}"
                )]

    # ============================================
    # MARKDOWN CONVERSION HANDLERS
    # ============================================

    async def convert_markdown_to_pptx(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Convert Markdown content to PowerPoint"""
        try:
            result = await self.markdown_converter.convert(
                args["markdown_content"],
                args["output_path"]
            )
            
            return [types.TextContent(
                type="text",
                text=json.dumps(result, indent=2)
            )]
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, indent=2)
            )]

    async def convert_markdown_file_to_pptx(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Convert Markdown file to PowerPoint"""
        try:
            result = await self.markdown_converter.convert_file(
                args["markdown_file"],
                args["output_path"]
            )
            
            return [types.TextContent(
                type="text",
                text=json.dumps(result, indent=2)
            )]
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, indent=2)
            )]

    async def validate_markdown_presentation(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Validate Markdown presentation syntax"""
        try:
            config, slides = self.markdown_converter.parser.parse(args["markdown_content"])
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "valid": True,
                    "slide_count": len(slides),
                    "title": config.title,
                    "slides": [{"type": s.type.value, "title": s.title} for s in slides]
                }, indent=2)
            )]
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "valid": False,
                    "error": str(e)
                }, indent=2)
            )]

    # ============================================
    # MATERIAL DESIGN HANDLERS
    # ============================================

    async def apply_material_theme(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Apply Material Design theme"""
        pres_id = args["presentation_id"]
        
        if pres_id not in self.presentations:
            raise ValueError(f"Presentation '{pres_id}' not found")

        theme_name = args["theme_name"]
        
        # Get or create theme
        if theme_name == "custom" and "seed_color" in args:
            theme = self.material_themes.get_material_you_theme(args["seed_color"])
        else:
            theme = self.material_themes.THEMES.get(theme_name)
            if not theme:
                raise ValueError(f"Unknown theme: {theme_name}")

        # Apply to presentation
        prs = self.presentations[pres_id]
        extension = MaterialDesignPowerPointExtension(self.ppt_server)
        
        for slide in prs.slides:
            extension._apply_theme_to_slide(slide, theme)

        return [types.TextContent(
            type="text",
            text=f"Applied Material Design theme '{theme.name}' to presentation '{pres_id}'"
        )]

    async def get_material_color_palette(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Get Material Design color palette"""
        base_color = args["base_color"].lstrip('#')
        palette_type = args.get("palette_type", "complementary")
        
        advice = self.material_advisor.get_color_advice(base_color)
        
        # Filter to requested palette type if specified
        if palette_type:
            combinations = [c for c in advice["recommended_combinations"] 
                          if c["name"].lower().replace(" ", "_") == palette_type]
            if combinations:
                advice["recommended_combinations"] = combinations

        return [types.TextContent(
            type="text",
            text=json.dumps(advice, indent=2)
        )]

    async def get_design_advice(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Get Material Design advice"""
        advice_type = args["advice_type"]
        context = args.get("context", "general")
        
        extension = MaterialDesignPowerPointExtension(self.ppt_server)
        result = await extension.get_design_advice({
            "advice_type": advice_type,
            "context": context
        })
        
        return result

    async def check_accessibility(self, args: Dict[str, Any]) -> list[types.TextContent]:
        """Check presentation accessibility"""
        pres_id = args["presentation_id"]
        standard = args.get("standard", "WCAG_AA")
        
        if pres_id not in self.presentations:
            raise ValueError(f"Presentation '{pres_id}' not found")

        prs = self.presentations[pres_id]
        issues = []
        
        # Check contrast ratios, text size, etc.
        for i, slide in enumerate(prs.slides):
            slide_issues = []
            
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    # Check text size
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.font.size and run.font.size.pt < 12:
                                slide_issues.append({
                                    "type": "text_size",
                                    "issue": f"Text size {run.font.size.pt}pt is below recommended minimum of 12pt"
                                })
            
            if slide_issues:
                issues.append({
                    "slide_index": i,
                    "issues": slide_issues
                })

        result = {
            "presentation_id": pres_id,
            "standard": standard,
            "total_slides": len(prs.slides),
            "slides_with_issues": len(issues),
            "issues": issues,
            "compliant": len(issues) == 0
        }

        return [types.TextContent(
            type="text",
            text=json.dumps(result, indent=2)
        )]

    async def run(self):
        """Run the unified MCP server"""
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                InitializationOptions(
                    server_name="unified-powerpoint-mcp",
                    server_version="1.0.0",
                    capabilities=self.server.get_capabilities(
                        notification_options=NotificationOptions(),
                        experimental_capabilities={},
                    ),
                ),
            )


async def main():
    """Main entry point"""
    server = UnifiedPowerPointMCPServer()
    await server.run()


if __name__ == "__main__":
    asyncio.run(main())
