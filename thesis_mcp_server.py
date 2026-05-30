"""
Thesis MCP Server — Word document processing tools for thesis format review.

Provides three tools:
  1. extract_doc_structure — Read a .docx/.doc file and extract full structure, formatting, and content
  2. convert_doc_format   — Convert between .doc and .docx using LibreOffice
  3. apply_corrections     — Apply formatting and content corrections to a .docx file

Supports both local file paths and URLs (HTTP/HTTPS presigned URLs from Nexent/MinIO).
URL-based files are downloaded to a temp directory before processing.

Uses FastMCP (StreamableHTTP transport) to expose tools via the MCP protocol.
"""

import json
import logging
import os
import shutil
import subprocess
import tempfile
import urllib.parse
from pathlib import Path
from typing import Any

import boto3
import requests
from botocore.config import Config as BotoConfig
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Emu
from fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("thesis-mcp")

# ---------------------------------------------------------------------------
# FastMCP server instance (StreamableHTTP on port 8899)
# ---------------------------------------------------------------------------
mcp = FastMCP("thesis-reviewer")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Directory for output files (corrected documents)
# Override via env var THESIS_OUTPUT_DIR
OUTPUT_DIR = os.getenv("THESIS_OUTPUT_DIR", os.path.join(os.path.dirname(__file__), "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# MinIO / S3 configuration for uploading corrected files (optional)
# If MINIO_ENDPOINT is not set, MinIO upload is skipped.
MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "")
MINIO_BUCKET = os.getenv("MINIO_BUCKET", "nexent")
MINIO_CORRECTIONS_PREFIX = os.getenv("MINIO_CORRECTIONS_PREFIX", "corrections")
# For hostname replacement in presigned URLs (e.g., localhost → external IP)
MINIO_EXTERNAL_HOST = os.getenv("MINIO_EXTERNAL_HOST", "")
# Server listen port
SERVER_PORT = int(os.getenv("THESIS_MCP_PORT", "8899"))

# Lazy-initialized S3 client
_s3_client = None


def _get_s3_client():
    """Get or create the S3 client for MinIO. Returns None if MinIO not configured."""
    if not MINIO_ENDPOINT:
        return None
    global _s3_client
    if _s3_client is None:
        _s3_client = boto3.client(
            "s3",
            endpoint_url=MINIO_ENDPOINT,
            aws_access_key_id=MINIO_ACCESS_KEY,
            aws_secret_access_key=MINIO_SECRET_KEY,
            config=BotoConfig(signature_version="s3v4"),
            region_name="us-east-1",
        )
    return _s3_client


def _upload_to_minio(local_path: str, object_name: str | None = None) -> str | None:
    """
    Upload a file to MinIO and return a presigned download URL (valid 24 hours).

    Args:
        local_path: Local file path to upload.
        object_name: Object name in MinIO. Defaults to corrections/<basename>.

    Returns:
        Presigned URL for direct download, or None if MinIO is not configured.
    """
    if not MINIO_ENDPOINT:
        logger.info("MinIO not configured, skipping upload.")
        return None

    if object_name is None:
        object_name = f"{MINIO_CORRECTIONS_PREFIX}/{os.path.basename(local_path)}"

    s3 = _get_s3_client()
    if s3 is None:
        return None

    logger.info("Uploading to MinIO: %s -> %s/%s", local_path, MINIO_BUCKET, object_name)

    s3.upload_file(
        local_path,
        MINIO_BUCKET,
        object_name,
        ExtraArgs={"ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
    )

    presigned_url = s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": MINIO_BUCKET, "Key": object_name},
        ExpiresIn=86400,  # 24 hours
    )
    # Replace localhost with external host if configured
    if MINIO_EXTERNAL_HOST:
        presigned_url = presigned_url.replace("localhost", MINIO_EXTERNAL_HOST)

    logger.info("Upload complete. Presigned URL valid for 24h.")
    return presigned_url

# Track temp files for cleanup
_temp_files: list = []

ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: "LEFT",
    WD_ALIGN_PARAGRAPH.CENTER: "CENTER",
    WD_ALIGN_PARAGRAPH.RIGHT: "RIGHT",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY",
    WD_ALIGN_PARAGRAPH.DISTRIBUTE: "DISTRIBUTE",
}

SPECIAL_KEYWORDS = {
    "cover_page": ["封面", "论文题目", "学校"],
    "abstract_cn": ["摘要", "中文摘要"],
    "abstract_en": ["Abstract", "ABSTRACT"],
    "toc": ["目录", "目  录"],
    "references": ["参考文献"],
    "acknowledgments": ["致谢"],
    "appendix": ["附录"],
}

# ---------------------------------------------------------------------------
# URL/file helpers
# ---------------------------------------------------------------------------

def _is_url(path: str) -> bool:
    """Check if a path is an HTTP/HTTPS URL."""
    return path.startswith("http://") or path.startswith("https://")


def _resolve_file_path(file_path: str) -> str:
    """
    Resolve a file path — if it's a URL, download to a temp directory and
    return the local path. If it's already a local path, return as-is.

    Supported URL patterns from Nexent:
      1. Northbound proxy: http://localhost:5013/api/nb/v1/file/fetch?presigned_url=...
         → Downloaded as-is (northbound handles MinIO access via Docker DNS).
      2. Direct presigned URL: http://nexent-minio:9000/nexent/...?X-Amz-...
         → nexent-minio:9000 rewritten to localhost:9010 (host access).
    """
    if _is_url(file_path):
        download_url = file_path

        # If this is a direct MinIO presigned URL (not going through northbound),
        # rewrite the Docker hostname for host-level access.
        if "nexent-minio:9000" in download_url and "/api/nb/v1/file/fetch" not in download_url:
            download_url = download_url.replace("nexent-minio:9000", "localhost:9010")

        logger.info("Downloading file from URL: %s", download_url)
        try:
            response = requests.get(download_url, timeout=60, stream=True)
            response.raise_for_status()

            # Try to get filename from Content-Disposition header
            cd = response.headers.get("Content-Disposition", "")
            filename = None
            if "filename=" in cd:
                import re
                match = re.search(r'filename[*]?=["\']?([^"\';\s]+)', cd)
                if match:
                    filename = urllib.parse.unquote(match.group(1))

            if not filename:
                # Fallback 1: extract filename from presigned_url query param
                parsed = urllib.parse.urlparse(file_path)
                query = urllib.parse.parse_qs(parsed.query)
                if "presigned_url" in query:
                    presigned = query["presigned_url"][0]
                    presigned_path = urllib.parse.urlparse(presigned).path
                    filename = os.path.basename(presigned_path)

            if not filename:
                # Fallback 2: derive from URL path or use generic name
                parsed = urllib.parse.urlparse(file_path)
                filename = os.path.basename(parsed.path) or "downloaded_file.docx"

            # Ensure the filename has a recognized extension
            if not os.path.splitext(filename)[1]:
                filename += ".docx"

            # Ensure temp dir exists
            temp_dir = tempfile.mkdtemp(prefix="thesis_mcp_")
            local_path = os.path.join(temp_dir, filename)

            with open(local_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            _temp_files.append(local_path)
            logger.info("Downloaded to: %s (%d bytes)", local_path, os.path.getsize(local_path))
            return local_path

        except requests.RequestException as e:
            raise RuntimeError(f"Failed to download file from URL: {e}")

    # Local path — verify it exists
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    return os.path.abspath(file_path)


def _get_paragraph_format_info(para: Any) -> dict:
    """Extract formatting metadata from a single paragraph."""
    pf = para.paragraph_format
    fmt: dict = {}

    # Line spacing
    if pf.line_spacing is not None:
        try:
            fmt["line_spacing"] = round(float(pf.line_spacing), 2)
        except (TypeError, ValueError):
            fmt["line_spacing"] = str(pf.line_spacing)

    # Spacing before / after (in points)
    if pf.space_before is not None:
        fmt["space_before_pt"] = round(pf.space_before.pt, 1)
    if pf.space_after is not None:
        fmt["space_after_pt"] = round(pf.space_after.pt, 1)

    # First-line indent (in cm)
    if pf.first_line_indent is not None:
        fmt["first_line_indent_cm"] = round(pf.first_line_indent.cm, 2)

    # Left / right indent
    if pf.left_indent is not None:
        fmt["left_indent_cm"] = round(pf.left_indent.cm, 2)
    if pf.right_indent is not None:
        fmt["right_indent_cm"] = round(pf.right_indent.cm, 2)

    # Alignment
    if pf.alignment is not None:
        fmt["alignment"] = ALIGN_MAP.get(pf.alignment, str(pf.alignment))

    # Font info from the first run
    if para.runs:
        run = para.runs[0]
        fmt["font_name"] = run.font.name if run.font.name else None
        if run.font.size:
            try:
                fmt["font_size"] = round(run.font.size.pt, 1)
            except Exception:
                fmt["font_size"] = None
        fmt["bold"] = run.font.bold
        fmt["italic"] = run.font.italic
        fmt["underline"] = run.font.underline

    # Style name
    if para.style:
        fmt["style_name"] = para.style.name

    # Heading level (via outline level or style)
    fmt["outline_level"] = _get_outline_level(para)

    return fmt


def _get_outline_level(para: Any) -> int | None:
    """Return the outline level of a paragraph (0-based), or None if body text."""
    pPr = para._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    outline_lvl = pPr.find(qn("w:outlineLvl"))
    if outline_lvl is not None:
        return int(outline_lvl.get(qn("w:val")))
    # Fallback: detect heading style
    style_name = (para.style.name if para.style else "").lower()
    if "heading" in style_name or "标题" in style_name:
        for word in style_name.replace("_", " ").split():
            if word.isdigit():
                return int(word) - 1
        return 0
    return None


def _detect_special_elements(paragraphs: list) -> dict:
    """Detect special elements (cover, abstract, toc, references, etc.) by keyword."""
    found: dict = {}
    for idx, para in enumerate(paragraphs):
        text = para.get("text", "")
        for element, keywords in SPECIAL_KEYWORDS.items():
            if element not in found:
                if any(kw in text for kw in keywords):
                    found[element] = {"exists": True, "paragraph_index": idx}

    for element in SPECIAL_KEYWORDS:
        if element not in found:
            found[element] = {"exists": False}
    return found


def _cm_to_value(emu_val: Any) -> float | None:
    """Convert an EMU or other value to cm, safely."""
    if emu_val is None:
        return None
    try:
        return round(Emu(int(emu_val)).cm, 2)
    except Exception:
        return None


def _convert_doc_to_docx(file_path: str) -> str:
    """Convert a .doc file to .docx using LibreOffice. Returns the new file path."""
    file_path = os.path.abspath(file_path)
    output_dir = os.path.dirname(file_path)
    logger.info("Converting %s to docx via LibreOffice ...", file_path)
    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "docx",
            "--outdir", output_dir,
            file_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    base = os.path.splitext(os.path.basename(file_path))[0]
    converted = os.path.join(output_dir, base + ".docx")
    if not os.path.exists(converted):
        raise FileNotFoundError(f"Converted file not found: {converted}")
    logger.info("Conversion succeeded: %s", converted)
    return converted


def _read_document(file_path: str) -> tuple[Document, str]:
    """
    Resolve file (URL or local), read a .docx (or auto-convert .doc),
    and return (Document, resolved_local_path).
    """
    file_path = _resolve_file_path(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".doc":
        file_path = _convert_doc_to_docx(file_path)
    elif ext != ".docx":
        raise ValueError(f"Unsupported file format: {ext}. Supported: .doc, .docx")
    return Document(file_path), file_path


# ===================================================================
# Tool 1: extract_doc_structure
# ===================================================================

@mcp.tool(
    name="extract_doc_structure",
    description=(
        "Extract the complete structure, formatting metadata, and text content "
        "from a Word document (.docx or .doc). Accepts either a local file path "
        "or an HTTP URL (presigned S3 URL, northbound fetch URL, etc.). "
        "Returns a JSON object with document-level formatting, section hierarchy, "
        "per-paragraph format details (font, size, alignment, spacing, indentation, "
        "style), tables, and special elements (cover page, abstracts, TOC, "
        "references, etc.)."
    ),
)
def extract_doc_structure(file_path: str) -> str:
    """
    Extract complete structure and content from a Word document.

    Args:
        file_path: Local file path OR HTTP URL to the .docx/.doc file.

    Returns:
        JSON string with the full document structure.
    """
    logger.info("extract_doc_structure called: %s", file_path)
    doc, _ = _read_document(file_path)

    # --- Document-level format ---
    doc_format: dict = {}
    if doc.sections:
        section = doc.sections[0]
        doc_format["page_size"] = (
            f"{section.page_width.cm:.1f}x{section.page_height.cm:.1f}cm"
            if section.page_width and section.page_height else None
        )
        doc_format["page_margins"] = {
            "top_cm": _cm_to_value(section.top_margin),
            "bottom_cm": _cm_to_value(section.bottom_margin),
            "left_cm": _cm_to_value(section.left_margin),
            "right_cm": _cm_to_value(section.right_margin),
        }
    # Default font from document-level style
    default_style = doc.styles["Normal"]
    if default_style.font.name:
        doc_format["default_font_name"] = default_style.font.name
    if default_style.font.size:
        doc_format["default_font_size"] = round(default_style.font.size.pt, 1)

    # --- Core properties ---
    core_props = doc.core_properties
    metadata = {
        "title": core_props.title or "",
        "author": core_props.author or "",
        "subject": core_props.subject or "",
        "created": str(core_props.created) if core_props.created else "",
    }

    # --- Paragraph-by-paragraph extraction ---
    all_paragraphs: list = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue  # skip fully empty paragraphs
        all_paragraphs.append({
            "paragraph_index": idx,
            "text": para.text,
            "format": _get_paragraph_format_info(para),
        })

    # --- Build section hierarchy ---
    current_section: dict | None = None
    sections: list = []
    for para_data in all_paragraphs:
        outline_lvl = para_data["format"].get("outline_level")
        if outline_lvl is not None and outline_lvl <= 3:
            current_section = {
                "level": outline_lvl + 1,
                "heading": para_data["text"],
                "heading_paragraph_index": para_data["paragraph_index"],
                "heading_format": para_data["format"],
                "paragraphs": [],
                "tables": [],
                "images": [],
                "content_summary": "",
            }
            sections.append(current_section)
        elif current_section is not None:
            text = para_data["text"]
            if len(current_section["content_summary"]) < 200:
                remaining = 200 - len(current_section["content_summary"])
                current_section["content_summary"] += text[:remaining] + " "
            current_section["paragraphs"].append(para_data)
        else:
            if not sections:
                preamble = {
                    "level": 0,
                    "heading": "[Preamble / Pre-content]",
                    "heading_paragraph_index": 0,
                    "heading_format": {},
                    "paragraphs": [],
                    "tables": [],
                    "images": [],
                    "content_summary": "",
                }
                sections.append(preamble)
            sections[0]["paragraphs"].append(para_data)

    # --- Tables ---
    all_tables: list = []
    for t_idx, table in enumerate(doc.tables):
        table_info = {
            "table_index": t_idx,
            "rows": len(table.rows),
            "columns": len(table.columns),
            "headers": [],
            "caption": "",
        }
        if table.rows:
            table_info["headers"] = [
                cell.text.strip()[:100] for cell in table.rows[0].cells
            ]
        all_tables.append(table_info)

    # --- Special elements ---
    special_elements = _detect_special_elements(all_paragraphs)

    # --- Build result ---
    result = {
        "metadata": metadata,
        "document_format": doc_format,
        "sections": sections,
        "tables": all_tables,
        "special_elements": special_elements,
        "total_paragraphs": len(all_paragraphs),
        "total_sections": len(sections),
    }

    logger.info("Extracted %d sections, %d paragraphs", len(sections), len(all_paragraphs))
    return json.dumps(result, ensure_ascii=False, indent=2)


# ===================================================================
# Tool 2: convert_doc_format
# ===================================================================

@mcp.tool(
    name="convert_doc_format",
    description=(
        "Convert a Word document between .doc and .docx formats using LibreOffice. "
        "Accepts either a local file path or an HTTP URL. "
        "Use this when you need to process an old-format .doc file with other tools."
    ),
)
def convert_doc_format(file_path: str, target_format: str = "docx") -> str:
    """
    Convert a Word document between .doc and .docx formats.

    Args:
        file_path: Local file path OR HTTP URL to the source file.
        target_format: Target format — "docx" or "doc". Default is "docx".

    Returns:
        JSON string with the original and converted file paths.
    """
    logger.info("convert_doc_format called: %s -> %s", file_path, target_format)
    file_path = _resolve_file_path(file_path)
    target_format = target_format.lower().strip(".")

    if target_format not in ("doc", "docx"):
        raise ValueError(f"target_format must be 'doc' or 'docx', got '{target_format}'")

    source_ext = os.path.splitext(file_path)[1].lower()
    if source_ext == f".{target_format}":
        return json.dumps({
            "success": True,
            "original_path": file_path,
            "converted_path": file_path,
            "original_format": source_ext,
            "target_format": f".{target_format}",
            "file_size_kb": round(os.path.getsize(file_path) / 1024, 1),
            "note": "File already in target format, no conversion needed.",
        }, ensure_ascii=False)

    output_dir = os.path.dirname(file_path)
    logger.info("Running LibreOffice conversion ...")
    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", target_format,
            "--outdir", output_dir,
            file_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    base = os.path.splitext(os.path.basename(file_path))[0]
    converted_path = os.path.join(output_dir, base + f".{target_format}")

    if not os.path.exists(converted_path):
        raise FileNotFoundError(f"Converted file not found: {converted_path}")

    response = {
        "success": True,
        "original_path": file_path,
        "converted_path": converted_path,
        "original_format": source_ext,
        "target_format": f".{target_format}",
        "file_size_kb": round(os.path.getsize(converted_path) / 1024, 1),
    }
    logger.info("Conversion complete: %s", converted_path)
    return json.dumps(response, ensure_ascii=False, indent=2)


# ===================================================================
# Tool 3: apply_corrections
# ===================================================================

@mcp.tool(
    name="apply_corrections",
    description=(
        "Apply formatting and content corrections to a Word (.docx) document. "
        "Accepts either a local file path or an HTTP URL for the original document. "
        "Accepts a JSON corrections specification with paragraph-level format "
        "changes, document-level settings, content additions, and content "
        "modifications. Saves the corrected document to a new file in the output "
        "directory and returns the path."
    ),
)
def apply_corrections(
    original_path: str,
    corrections_json: str,
    output_filename: str,
) -> str:
    """
    Apply corrections to a Word document and save a corrected copy.

    Args:
        original_path: Local file path OR HTTP URL to the original .docx file.
        corrections_json: JSON string describing corrections to apply.
        output_filename: Name for the output file (basename only, e.g. "corrected.docx").

    Returns:
        JSON string with the output path and summary of changes applied.
    """
    logger.info("apply_corrections called: %s", original_path)

    # Parse corrections
    try:
        corrections = json.loads(corrections_json)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid corrections_json: {e}")

    # Open document (handles both URL and local path)
    original_path = _resolve_file_path(original_path)
    doc = Document(original_path)

    stats = {
        "document_level_changes": 0,
        "format_corrections": 0,
        "content_additions": 0,
        "content_modifications": 0,
    }
    warnings: list = []

    # --- Document-level changes ---
    doc_level = corrections.get("document_level", {})
    if doc_level and doc.sections:
        sec = doc.sections[0]
        margins = doc_level.get("page_margins", {})
        if margins:
            if "top_cm" in margins:
                sec.top_margin = Cm(float(margins["top_cm"]))
            if "bottom_cm" in margins:
                sec.bottom_margin = Cm(float(margins["bottom_cm"]))
            if "left_cm" in margins:
                sec.left_margin = Cm(float(margins["left_cm"]))
            if "right_cm" in margins:
                sec.right_margin = Cm(float(margins["right_cm"]))
            stats["document_level_changes"] = 1

    # --- Paragraph format corrections ---
    para_count = len(doc.paragraphs)
    para_corrections = corrections.get("paragraph_corrections", [])
    for corr in para_corrections:
        p_idx = corr.get("paragraph_index")
        changes = corr.get("changes", {})
        if p_idx is None or p_idx >= para_count:
            warnings.append(f"Paragraph index {p_idx} does not exist, skipped.")
            continue

        para = doc.paragraphs[p_idx]
        _apply_paragraph_format_changes(para, changes)
        stats["format_corrections"] += 1

    # --- Content additions ---
    content_additions = corrections.get("content_additions", [])
    for add in content_additions:
        after_idx = add.get("after_paragraph_index")
        text = add.get("text", "")
        style_name = add.get("style", "Normal")

        if after_idx is None or after_idx >= para_count:
            warnings.append(f"Insertion point {after_idx} does not exist, skipped.")
            continue

        ref_para = doc.paragraphs[after_idx]
        new_para = _insert_paragraph_after(ref_para, text, style_name)
        if new_para:
            stats["content_additions"] += 1

    # --- Content modifications ---
    content_modifications = corrections.get("content_modifications", [])
    for mod in content_modifications:
        p_idx = mod.get("paragraph_index")
        new_text = mod.get("new_text", "")

        if p_idx is None or p_idx >= para_count:
            warnings.append(f"Paragraph index {p_idx} does not exist, skipped.")
            continue

        para = doc.paragraphs[p_idx]
        for run in para.runs:
            run.text = ""
        if para.runs:
            para.runs[0].text = new_text
        else:
            para.add_run(new_text)
        stats["content_modifications"] += 1

    # --- Save to output directory ---
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    doc.save(output_path)

    # --- Upload to MinIO for user download ---
    download_url = None
    try:
        download_url = _upload_to_minio(output_path)
    except Exception as e:
        logger.warning("Failed to upload corrected file to MinIO: %s", e)
        # File is still available locally — fall through

    result = {
        "success": True,
        "output_path": output_path,
        "download_url": download_url,
        "note": (
            "Corrected file has been saved. Use the download_url to download it. "
            "The download link is valid for 24 hours."
            if download_url
            else "Corrected file saved locally but upload to MinIO failed. "
                 "Please contact the administrator to retrieve the file."
        ),
        "changes_summary": stats,
        "total_changes_applied": sum(stats.values()),
        "warnings": warnings,
    }
    logger.info("Corrections applied, saved to: %s", output_path)
    return json.dumps(result, ensure_ascii=False, indent=2)


def _apply_paragraph_format_changes(para: Any, changes: dict) -> None:
    """Apply formatting changes to a single paragraph."""
    pf = para.paragraph_format

    if "line_spacing" in changes:
        pf.line_spacing = float(changes["line_spacing"])
    if "space_before_pt" in changes:
        pf.space_before = Pt(float(changes["space_before_pt"]))
    if "space_after_pt" in changes:
        pf.space_after = Pt(float(changes["space_after_pt"]))
    if "first_line_indent_cm" in changes:
        pf.first_line_indent = Cm(float(changes["first_line_indent_cm"]))
    if "left_indent_cm" in changes:
        pf.left_indent = Cm(float(changes["left_indent_cm"]))
    if "right_indent_cm" in changes:
        pf.right_indent = Cm(float(changes["right_indent_cm"]))

    align_str = changes.get("alignment")
    if align_str:
        align_map = {
            "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
            "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
            "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
            "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if align_str.upper() in align_map:
            pf.alignment = align_map[align_str.upper()]

    font_changes = {}
    for key in ("font_name", "font_size", "bold", "italic", "underline"):
        if key in changes:
            font_changes[key] = changes[key]

    if font_changes:
        if not para.runs:
            para.add_run("")
        for run in para.runs:
            if "font_name" in font_changes and font_changes["font_name"]:
                run.font.name = font_changes["font_name"]
            if "font_size" in font_changes and font_changes["font_size"]:
                run.font.size = Pt(float(font_changes["font_size"]))
            if "bold" in font_changes:
                run.font.bold = bool(font_changes["bold"])
            if "italic" in font_changes:
                run.font.italic = bool(font_changes["italic"])
            if "underline" in font_changes:
                run.font.underline = bool(font_changes["underline"])


def _insert_paragraph_after(ref_para: Any, text: str, style_name: str) -> Any | None:
    """Insert a new paragraph after the reference paragraph with given text and style."""
    try:
        new_element = ref_para._element.makeelement(qn("w:p"), {})
        ref_para._element.addnext(new_element)

        from docx.text.paragraph import Paragraph
        new_para = Paragraph(new_element, ref_para._parent)

        try:
            new_para.style = new_para.part.document.styles[style_name]
        except KeyError:
            new_para.style = new_para.part.document.styles["Normal"]

        new_para.add_run(text)
        return new_para
    except Exception as e:
        logger.warning("Failed to insert paragraph after index: %s", e)
        return None


# ===================================================================
# Entrypoint
# ===================================================================

if __name__ == "__main__":
    logger.info("Starting Thesis MCP Server on port %d ...", SERVER_PORT)
    mcp.run(transport="streamable-http", host="0.0.0.0", port=SERVER_PORT)
