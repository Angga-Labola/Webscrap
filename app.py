"""Flask Application connecting to Google Antigravity and PowerPoint MCP."""

import os
import sys
import uuid
import asyncio
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import dotenv

BASE_DIR = Path(__file__).resolve().parent
dotenv.load_dotenv(BASE_DIR / ".env")

from ppt_builder import THEMES
from extractors import extract_document_text, extract_url_text
from antigravity_service import generate_powerpoint_presentation, OUTPUT_DIR

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "antigravity-secret-key-2026")
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB max upload

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "docx", "pptx", "txt", "md", "csv", "json"}


def is_allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    """Renders the main dashboard."""
    return render_template("index.html", themes=THEMES)


@app.route("/api/themes", methods=["GET"])
def get_themes():
    """Returns available design themes with color swatches."""
    return jsonify({
        "status": "success",
        "themes": THEMES
    })


@app.route("/api/health", methods=["GET"])
def health():
    """Health check endpoint."""
    api_key = os.getenv("GEMINI_API_KEY")
    return jsonify({
        "status": "healthy",
        "antigravity_ready": bool(api_key),
        "key_configured": bool(api_key),
        "output_directory": str(OUTPUT_DIR),
        "active_themes": list(THEMES.keys())
    })


@app.route("/api/recent", methods=["GET"])
def recent_presentations():
    """Returns recently generated presentations."""
    files = list(OUTPUT_DIR.glob("*.pptx"))
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    recent = []
    for f in files[:10]:
        recent.append({
            "filename": f.name,
            "size_kb": round(f.stat().st_size / 1024, 1),
            "modified": f.stat().st_mtime,
            "download_url": f"/api/download/{f.name}"
        })
    return jsonify({"status": "success", "presentations": recent})


@app.route("/api/generate", methods=["POST"])
def generate():
    """Receives text, document, and URLs, orchestrates Google Antigravity with MCP, and returns PowerPoint."""
    try:
        text_input = request.form.get("text", "").strip()
        url_input_raw = request.form.get("urls", "").strip()
        theme = request.form.get("theme", "modern_indigo").strip()
        custom_title = request.form.get("title", "").strip()
        presenter = request.form.get("presenter", "Google Antigravity AI").strip()
        
        try:
            slide_count = int(request.form.get("slide_count", 6))
            slide_count = max(3, min(slide_count, 15))
        except (ValueError, TypeError):
            slide_count = 6

        # 1. Process Document Upload if present
        document_data = None
        if "document" in request.files:
            file = request.files["document"]
            if file and file.filename and is_allowed_file(file.filename):
                safe_name = secure_filename(file.filename)
                saved_name = f"{uuid.uuid4().hex[:8]}_{safe_name}"
                upload_path = UPLOAD_DIR / saved_name
                file.save(str(upload_path))
                
                # Extract text
                document_data = extract_document_text(str(upload_path), original_filename=safe_name)

        # 2. Process URLs if present
        url_data_list = []
        if url_input_raw:
            # URLs can be comma-separated or newline-separated
            raw_urls = [u.strip() for u in url_input_raw.replace(",", "\n").splitlines() if u.strip()]
            for u in raw_urls[:5]:  # Limit to 5 URLs max
                url_res = extract_url_text(u)
                url_data_list.append(url_res)

        # Validate that at least one input source is provided
        has_text = bool(text_input)
        has_doc = bool(document_data and document_data.get("content"))
        has_url = any(u.get("success") and u.get("content") for u in url_data_list)

        if not (has_text or has_doc or has_url):
            return jsonify({
                "status": "error",
                "message": "Please provide at least one input source: Text, Document file, or a URL."
            }), 400

        # 3. Call Google Antigravity Agent with MCP
        result = asyncio.run(
            generate_powerpoint_presentation(
                text_input=text_input,
                document_data=document_data,
                url_data=url_data_list,
                slide_count=slide_count,
                theme=theme,
                custom_title=custom_title,
                presenter_name=presenter,
            )
        )

        # Add ingestion summary to response
        result["inputs_summary"] = {
            "has_text": has_text,
            "text_length": len(text_input) if has_text else 0,
            "has_document": has_doc,
            "document_name": document_data["filename"] if has_doc else None,
            "document_chars": len(document_data["content"]) if has_doc else 0,
            "urls_processed": len([u for u in url_data_list if u.get("success")]),
        }

        return jsonify(result)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


@app.route("/api/download/<filename>", methods=["GET"])
def download_presentation(filename: str):
    """Downloads a generated PowerPoint file."""
    clean_name = os.path.basename(filename)
    if ".." in clean_name or "/" in clean_name or "\\" in clean_name:
        return jsonify({"status": "error", "message": "Invalid filename"}), 400

    target_path = OUTPUT_DIR / clean_name
    if not target_path.exists():
        # Fallback check for sanitized name
        safe_name = secure_filename(filename)
        target_path = OUTPUT_DIR / safe_name
        if not target_path.exists():
            return jsonify({"status": "error", "message": "File not found"}), 404
        clean_name = safe_name

    return send_from_directory(
        str(OUTPUT_DIR),
        clean_name,
        as_attachment=True,
        download_name=clean_name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    print(f"\n==================================================")
    print(f"  Google Antigravity & MCP PowerPoint Generator   ")
    print(f"  Server starting at: http://127.0.0.1:{port}       ")
    print(f"==================================================\n")
    app.run(host="0.0.0.0", port=port, debug=True)
