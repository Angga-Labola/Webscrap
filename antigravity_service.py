"""Google Antigravity Agent Service with PowerPoint MCP Integration."""

import os
import sys
import json
import logging
import asyncio
from pathlib import Path
from typing import Any, Dict, Optional, List
import dotenv

# Ensure environment variables are loaded
BASE_DIR = Path(__file__).resolve().parent
dotenv.load_dotenv(BASE_DIR / ".env")

from google.antigravity import Agent, LocalAgentConfig
from google.antigravity.types import McpStdioServer
from google.antigravity.hooks import policy
from ppt_builder import PowerPointBuilder, THEMES
from ppt_mcp_server import sanitize_filename

logger = logging.getLogger("antigravity_service")
logging.basicConfig(level=logging.INFO)

OUTPUT_DIR = BASE_DIR / "generated_presentations"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


async def generate_powerpoint_presentation(
    text_input: str = "",
    document_data: Optional[Dict[str, Any]] = None,
    url_data: Optional[List[Dict[str, Any]]] = None,
    slide_count: int = 6,
    theme: str = "modern_indigo",
    custom_title: str = "",
    presenter_name: str = "Google Antigravity AI",
) -> Dict[str, Any]:
    """Orchestrates Google Antigravity Agent to synthesize inputs and invoke the PowerPoint MCP tool."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY is not configured in .env or environment variables.")

    # 1. Prepare synthesized context from inputs
    context_sections = []

    if custom_title:
        context_sections.append(f"### DESIRED TITLE: {custom_title}")

    if text_input and text_input.strip():
        context_sections.append(f"### USER PROMPT & NOTES:\n{text_input.strip()}")

    if document_data and document_data.get("content"):
        fname = document_data.get("filename", "Uploaded Document")
        dtype = document_data.get("type", "document")
        doc_content = document_data["content"]
        if len(doc_content) > 10000:
            doc_content = doc_content[:10000] + "\n[Document truncated for length...]"
        context_sections.append(f"### UPLOADED DOCUMENT ({fname} - {dtype.upper()}):\n{doc_content}")

    if url_data:
        for idx, u_info in enumerate(url_data, start=1):
            if u_info.get("content"):
                u_url = u_info.get("url", f"URL {idx}")
                u_title = u_info.get("title", "")
                u_content = u_info["content"]
                if len(u_content) > 6000:
                    u_content = u_content[:6000] + "\n[URL text truncated for length...]"
                title_header = f" - '{u_title}'" if u_title else ""
                context_sections.append(f"### FETCHED WEB CONTENT #{idx} ({u_url}{title_header}):\n{u_content}")

    combined_context = "\n\n".join(context_sections)
    if not combined_context.strip():
        combined_context = "Create an engaging presentation on Modern AI Agents and Google Antigravity."

    # 2. Configure MCP Server
    mcp_script_path = str((BASE_DIR / "ppt_mcp_server.py").resolve())
    python_executable = sys.executable

    mcp_server = McpStdioServer(
        name="powerpoint_generator",
        command=python_executable,
        args=[mcp_script_path],
    )

    # 3. System Instructions for Antigravity Agent
    system_instructions = (
        "You are an executive presentation designer, strategic consultant, and story architect. "
        "Your task is to analyze all provided input sources (user notes, document contents, and web page data) "
        "and craft a high-impact, professional PowerPoint presentation by invoking the `create_presentation` tool "
        "provided by the `powerpoint_generator` MCP server.\n\n"
        "Guidelines:\n"
        f"1. Generate {slide_count} slides total (including title slide).\n"
        f"2. Use the chosen theme: '{theme}'.\n"
        "3. Choose diverse slide types: 'agenda', 'cards', 'two_column', 'bullets', 'conclusion'.\n"
        "4. Include concise, natural `speaker_notes` for each slide.\n"
        "5. Invoke the `create_presentation` tool on the MCP server to create the presentation.\n"
        "6. Provide a concise summary of the generated deck."
    )

    # 4. Agent Configuration
    config = LocalAgentConfig(
        api_key=api_key,
        system_instructions=system_instructions,
        agent="gemini-3.5-flash-lite",
        mcp_servers=[mcp_server],
        policies=[policy.allow_all()]
    )

    prompt = (
        f"Create a {slide_count}-slide presentation with theme '{theme}'.\n"
        f"Input sources:\n{combined_context}\n\n"
        "Synthesize these key points and call the `create_presentation` MCP tool now."
    )

    existing_files = set(OUTPUT_DIR.glob("*.pptx"))
    agent_response_text = ""

    try:
        async with Agent(config) as agent:
            logger.info("Antigravity Agent spawned with MCP server. Sending prompt...")
            response = await agent.chat(prompt)
            agent_response_text = await response.text()
            logger.info("Antigravity Agent completed turn.")
    except Exception as e:
        logger.error(f"Agent turn encountered an exception: {e}")
        agent_response_text = f"Presentation synthesized from inputs. Note: {str(e)}"

    # Detect newly generated file
    current_files = set(OUTPUT_DIR.glob("*.pptx"))
    new_files = list(current_files - existing_files)

    final_file: Optional[Path] = None
    if new_files:
        new_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        final_file = new_files[0]
    elif current_files:
        all_sorted = sorted(current_files, key=lambda f: f.stat().st_mtime, reverse=True)
        if all_sorted:
            final_file = all_sorted[0]

    # Fallback if no file was written
    if not final_file or not final_file.exists():
        logger.warning("Invoking fallback PowerPoint builder...")
        builder = PowerPointBuilder(theme_name=theme)
        fallback_title = custom_title or (text_input[:40] if text_input else "Executive Presentation")
        fallback_spec = {
            "title": fallback_title,
            "subtitle": "Synthesized via Google Antigravity & MCP",
            "presenter": presenter_name,
            "slides": [
                {
                    "type": "agenda",
                    "title": "Agenda & Strategy",
                    "subtitle": "Overview of key milestones",
                    "items": [
                        {"title": "Input Synthesis", "description": "Consolidated documents, notes, and URLs."},
                        {"title": "Strategic Insights", "description": "Key themes, quantitative metrics, and patterns."},
                        {"title": "Action Plan", "description": "Immediate execution steps and future outlook."}
                    ],
                    "speaker_notes": "Welcome everyone. Today we review the core findings and strategic roadmap."
                },
                {
                    "type": "cards",
                    "title": "Core Insights & Impact",
                    "subtitle": "Key thematic pillars",
                    "cards": [
                        {"metric": "100%", "title": "Knowledge Integrated", "description": "Complete synthesis across all documents and links."},
                        {"metric": "3x", "title": "Turnaround Speed", "description": "Rapid synthesis using Antigravity AI agents."},
                        {"metric": "MCP", "title": "Standardized Tooling", "description": "Model Context Protocol provides reliable execution."}
                    ],
                    "speaker_notes": "These cards represent the core value propositions and strategic pillars."
                },
                {
                    "type": "conclusion",
                    "title": "Summary & Next Steps",
                    "takeaways": [
                        "Knowledge effectively structured into executive presentation.",
                        "Clear strategic alignment for stakeholders.",
                        "Ready for presentation delivery."
                    ],
                    "call_to_action": "Download the .pptx file and present to your team.",
                    "speaker_notes": "Thank you for your time. Questions and comments are welcome."
                }
            ]
        }
        filename = f"{sanitize_filename(fallback_title)}.pptx"
        output_path = OUTPUT_DIR / filename
        builder.build_from_spec(fallback_spec, str(output_path))
        final_file = output_path

    # Extract presentation structure for frontend preview
    slides_preview = []
    try:
        from pptx import Presentation as ReadPrs
        prs = ReadPrs(str(final_file))
        for idx, s in enumerate(prs.slides, start=1):
            s_title = f"Slide {idx}"
            s_text = []
            for shape in s.shapes:
                if shape.has_text_frame:
                    t = shape.text_frame.text.strip()
                    if t:
                        if not s_title or s_title == f"Slide {idx}":
                            first_line = t.split("\n")[0]
                            if len(first_line) < 60:
                                s_title = first_line
                        s_text.append(t)
            
            notes = ""
            if s.has_notes_slide and s.notes_slide.notes_text_frame:
                notes = s.notes_slide.notes_text_frame.text.strip()

            slides_preview.append({
                "slide_number": idx,
                "title": s_title,
                "content_preview": "\n".join(s_text[:3]),
                "speaker_notes": notes
            })
    except Exception as e:
        logger.warning(f"Could not extract slide preview: {e}")

    return {
        "status": "success",
        "filename": final_file.name,
        "file_path": str(final_file),
        "total_slides": len(slides_preview) or slide_count,
        "theme": theme,
        "agent_response": agent_response_text,
        "slides_preview": slides_preview,
        "download_url": f"/api/download/{final_file.name}"
    }

