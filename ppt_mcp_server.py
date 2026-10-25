"""PowerPoint MCP Server.

Exposes PowerPoint presentation creation tools via the Model Context Protocol (MCP)
over stdio transport, allowing Google Antigravity agents to generate presentations.
"""

import os
import sys
import json
import uuid
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from mcp.server.mcpserver import MCPServer
from ppt_builder import PowerPointBuilder, THEMES

server = MCPServer("powerpoint_generator")

OUTPUT_DIR = Path(__file__).resolve().parent / "generated_presentations"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def sanitize_filename(name: str) -> str:
    """Sanitizes presentation filename."""
    clean = re.sub(r'[^a-zA-Z0-9_\-]', '_', name).strip('_')
    return clean[:50] or f"presentation_{uuid.uuid4().hex[:8]}"


@server.tool()
def get_available_themes() -> str:
    """Returns a list of available presentation design themes and their color palettes.
    
    Available themes: modern_indigo, tech_emerald, executive_slate, crimson_bold, cyber_dark.
    """
    themes_info = {
        key: {
            "name": val["name"],
            "description": val["description"],
            "primary": f"#{val['primary']}",
            "secondary": f"#{val['secondary']}",
            "accent": f"#{val['accent']}",
        }
        for key, val in THEMES.items()
    }
    return json.dumps(themes_info, indent=2)


@server.tool()
def create_presentation(
    title: str,
    subtitle: str = "",
    theme: str = "modern_indigo",
    slides: Union[str, List[Dict[str, Any]], None] = None,
    output_filename: str = "",
    presenter: str = "Google Antigravity AI",
    date: str = "",
) -> str:
    """Generates a complete, professional PowerPoint presentation (.pptx) based on structured slide content.

    Args:
        title: Main presentation title displayed on the hero slide.
        subtitle: Subtitle or tagline for the presentation.
        theme: Theme name (modern_indigo, tech_emerald, executive_slate, crimson_bold, cyber_dark).
        slides: List of slide objects or JSON string. Each slide object must contain:
            - type: "agenda" | "cards" | "two_column" | "bullets" | "conclusion"
            - title: Slide headline (e.g., "Market Analysis")
            - category: Short uppercase category/kicker tag (e.g., "STRATEGY")
            - subtitle: Context or brief sentence below title
            - speaker_notes: Detailed speaker notes for the presenter
            - Type-specific fields:
                - For "agenda": items: list of {"title": str, "description": str}
                - For "cards": cards: list of {"metric": str (e.g. '85%', '3x'), "title": str, "description": str}
                - For "two_column": col1: {"title": str, "points": [str]}, col2: {"title": str, "points": [str]}
                - For "bullets": points: list of {"title": str, "description": str}
                - For "conclusion": takeaways: list of str, call_to_action: str
        output_filename: Desired output filename without extension (optional).
        presenter: Presenter name or organization.
        date: Presentation date string.

    Returns:
        JSON string containing the output filepath, slide count, and summary.
    """
    try:
        # If slides is passed as a JSON string, deserialize it
        if isinstance(slides, str):
            try:
                slides = json.loads(slides)
            except Exception:
                slides = None

        if not slides or not isinstance(slides, list):
            slides = [
                {
                    "type": "bullets",
                    "title": "Core Insights",
                    "category": "SYNTHESIS",
                    "points": [{"title": "Executive Summary", "description": "Synthesized via Google Antigravity & MCP"}],
                    "speaker_notes": "Welcome to this presentation."
                }
            ]

        # Normalize inner structures if any were serialized as JSON strings
        for s in slides:
            if isinstance(s, dict):
                for subkey in ("items", "cards", "points", "takeaways", "col1", "col2"):
                    if subkey in s and isinstance(s[subkey], str):
                        try:
                            s[subkey] = json.loads(s[subkey])
                        except Exception:
                            pass

        # Clean output filename
        if not output_filename:
            base_name = sanitize_filename(title)
            output_filename = f"{base_name}_{uuid.uuid4().hex[:6]}"
        else:
            output_filename = sanitize_filename(output_filename)

        if not output_filename.endswith(".pptx"):
            full_filename = f"{output_filename}.pptx"
        else:
            full_filename = output_filename

        output_path = OUTPUT_DIR / full_filename

        spec = {
            "title": title,
            "subtitle": subtitle,
            "presenter": presenter,
            "date": date,
            "title_notes": f"Welcome to the presentation: {title}.",
            "slides": slides,
        }

        builder = PowerPointBuilder(theme_name=theme)
        final_file_path = builder.build_from_spec(spec, str(output_path))

        slide_summaries = []
        slide_summaries.append({"slide": 1, "type": "title", "title": title})
        for idx, s in enumerate(slides, start=2):
            if isinstance(s, dict):
                slide_summaries.append({
                    "slide": idx,
                    "type": s.get("type", "bullets"),
                    "title": s.get("title", f"Slide {idx}"),
                    "category": s.get("category", "")
                })

        result = {
            "status": "success",
            "message": "PowerPoint presentation generated successfully.",
            "file_path": final_file_path,
            "filename": full_filename,
            "total_slides": len(slide_summaries),
            "theme": builder.theme_name,
            "slides": slide_summaries
        }
        return json.dumps(result, indent=2)

    except Exception as e:
        import traceback
        err_msg = f"Error generating PowerPoint: {str(e)}\n{traceback.format_exc()}"
        return json.dumps({
            "status": "error",
            "message": err_msg
        }, indent=2)


if __name__ == "__main__":
    server.run("stdio")
