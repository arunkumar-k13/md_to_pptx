"""REST API Route Definitions for Presentation Generation."""

from __future__ import annotations
import logging
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse, RedirectResponse
from starlette.background import BackgroundTask

from md_to_pptx.api.schemas import ErrorResponse, HealthResponse
from md_to_pptx.config.settings_loader import load_settings
from md_to_pptx.core.exporters.environment import PdfEnvironmentDetector
from md_to_pptx.core.markdown_generator import generate_markdown_from_prompt
from md_to_pptx.core.markdown_parser import MarkdownParser, restore_markdown_structure
from md_to_pptx.main import generate_presentation
from md_to_pptx.utils.logger import setup_logger

logger = setup_logger(__name__)

router = APIRouter()

ALLOWED_EXTENSIONS = {".md", ".markdown", ".txt"}
ALLOWED_FORMATS = {"pptx", "pdf"}
ALLOWED_EXPORTERS = {"libreoffice"}
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
DEFAULT_TEMPLATE_PATH = Path(os.getenv("DEFAULT_PPTX_TEMPLATE", str(PROJECT_ROOT / "corporate_template.pptx"))).resolve()



def cleanup_directory(dir_path: str) -> None:
    """Safely remove a temporary workspace directory after response completion."""
    try:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path, ignore_errors=True)
            logger.debug("Successfully cleaned up API temporary directory: %s", dir_path)
    except Exception as err:
        logger.warning("Failed to clean up API temporary directory %s: %s", dir_path, err)


def sanitize_topic_filename(topic: str, fallback_stem: str, file_ext: str) -> str:
    """Derive a safe, clean topic-based attachment filename.

    Args:
        topic: Extracted document title string.
        fallback_stem: Original uploaded file stem.
        file_ext: Target file extension ('.pptx' or '.pdf').

    Returns:
        Sanitized filename string (e.g. 'AI_Innovations_in_Life_Sciences.pptx').
    """
    raw_name = topic.strip() if (topic and topic.strip() and topic.strip() != "Untitled Presentation") else fallback_stem.strip()
    
    # Remove invalid characters and sanitize whitespace
    clean_name = re.sub(r"[^\w\s\-]", "", raw_name)
    clean_name = re.sub(r"[\s_]+", "_", clean_name).strip("_")

    generic_names = {"generated", "output", "presentation", "untitled", "sample", "temp", "file"}
    if not clean_name or clean_name.lower() in generic_names:
        clean_name = f"Presentation_{re.sub(r'[^\w\-]', '_', fallback_stem)}" if fallback_stem else "Executive_Presentation"

    # Truncate topic filename to a safe max length of 60 characters
    clean_name = clean_name[:60].rstrip("_")

    return f"{clean_name}{file_ext}"


@router.get("/", include_in_schema=False)
def root_redirect() -> RedirectResponse:
    """Redirect root GET / requests to the interactive OpenAPI documentation (/docs)."""
    return RedirectResponse(url="/docs")


@router.get(
    "/health",
    response_model=HealthResponse,
    summary="Service Health Check",
    description="Returns basic operational status of the Markdown-to-PowerPoint Generator service.",
)
def health_check() -> HealthResponse:
    """Health check endpoint returning service status."""
    return HealthResponse(status="ok", service="markdown-to-powerpoint-generator")


async def _process_presentation_request(
    file: Optional[UploadFile] = None,
    markdown_text: Optional[str] = None,
    template: Optional[UploadFile] = None,
    format_clean: str = "pptx",
    preferred_exporter: str = "libreoffice",
) -> FileResponse:
    """Internal helper processing presentation requests for general, dedicated, and testing API endpoints."""
    # 1. Validate Requested Format
    if format_clean not in ALLOWED_FORMATS:
        logger.warning("API rejected invalid format request '%s'", format_clean)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={
                "error": "invalid_format",
                "message": f"Format '{format_clean}' is not supported. Allowed values: 'pptx', 'pdf'",
            },
        )

    # 2. Validate Preferred Exporter
    exporter_clean = preferred_exporter.strip().lower() if preferred_exporter else "libreoffice"
    if exporter_clean not in ALLOWED_EXPORTERS:
        exporter_clean = "libreoffice"

    # 3. Setup Secure Temporary Working Directory
    temp_dir = tempfile.mkdtemp(prefix="md2pptx_api_")
    temp_dir_path = Path(temp_dir).resolve()

    # 4. Resolve Corporate Template Path (Uploaded File vs Configured Default)
    tmpl_path = DEFAULT_TEMPLATE_PATH

    if template and hasattr(template, "filename") and template.filename:
        tmpl_ext = Path(template.filename).suffix.lower()
        if tmpl_ext != ".pptx":
            cleanup_directory(temp_dir)
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail={
                    "error": "invalid_template_file",
                    "message": f"Uploaded template extension '{tmpl_ext}' is not supported. Must be a .pptx file.",
                },
            )
        
        tmpl_content = await template.read()
        if tmpl_content:
            safe_tmpl_name = re.sub(r"[^\w\.\-]", "_", Path(template.filename).name)
            uploaded_tmpl_path = temp_dir_path / f"custom_{safe_tmpl_name}"
            uploaded_tmpl_path.write_bytes(tmpl_content)
            tmpl_path = uploaded_tmpl_path

    if not tmpl_path.is_file():
        cleanup_directory(temp_dir)
        logger.error("Default corporate template file not found at %s", tmpl_path)
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={
                "error": "template_not_found",
                "message": f"Required template '{tmpl_path.name}' does not exist on the server.",
            },
        )

    # 5. Validate PDF Export Environment if PDF requested
    if format_clean == "pdf":
        cfg = load_settings()
        env_report = PdfEnvironmentDetector.detect(preferred_exporter=exporter_clean, settings=cfg)
        if not env_report.pdf_available:
            cleanup_directory(temp_dir)
            logger.warning("PDF requested via API but LibreOffice engine is unavailable on host machine.")
            raise HTTPException(
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
                detail={
                    "error": "pdf_export_unavailable",
                    "message": "No supported PDF export engine (LibreOffice) is available on the server.",
                },
            )

    try:
        original_filename = "uploaded.md"
        file_content: Optional[bytes] = None

        if file and hasattr(file, "filename") and file.filename:
            original_filename = file.filename
            file_ext = Path(original_filename).suffix.lower()

            if file_ext not in ALLOWED_EXTENSIONS:
                cleanup_directory(temp_dir)
                logger.warning("API rejected upload with invalid extension '%s': %s", file_ext, original_filename)
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail={
                        "error": "invalid_file_type",
                        "message": f"Uploaded file extension '{file_ext}' is not supported. Allowed extensions: .md, .markdown, .txt",
                    },
                )
            file_content = await file.read()

        elif markdown_text and isinstance(markdown_text, str) and markdown_text.strip():
            clean_text = markdown_text.strip()
            original_filename = "test_input.md"
            file_content = clean_text.encode("utf-8")

        else:
            cleanup_directory(temp_dir)
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail={
                    "error": "missing_file_or_text",
                    "message": "Request must include either an uploaded Markdown 'file' (.md, .markdown, or .txt) or a 'markdown_text' string.",
                },
            )

        if not file_content or not file_content.strip():
            cleanup_directory(temp_dir)
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail={
                    "error": "empty_file",
                    "message": "The Markdown input file or text content contains 0 bytes.",
                },
            )

        max_size_mb = int(os.getenv("MAX_UPLOAD_SIZE_MB", "50"))
        max_size_bytes = max_size_mb * 1024 * 1024
        if len(file_content) > max_size_bytes:
            cleanup_directory(temp_dir)
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail={
                    "error": "file_too_large",
                    "message": f"File size ({len(file_content) / (1024*1024):.1f} MB) exceeds maximum allowed size ({max_size_mb} MB).",
                },
            )

        safe_input_name = re.sub(r"[^\w\.\-]", "_", Path(original_filename).name)
        temp_md_path = temp_dir_path / safe_input_name
        temp_md_path.write_bytes(file_content)

        # 6. Execute Presentation Generation Pipeline
        temp_output_dir = temp_dir_path / "output"
        temp_output_dir.mkdir(parents=True, exist_ok=True)

        logger.info("API processing presentation request for '%s' (Format: %s)...", original_filename, format_clean.upper())

        report = generate_presentation(
            markdown_path=temp_md_path,
            template_path=tmpl_path,
            output_dir=temp_output_dir,
            export_pdf=(format_clean == "pdf"),
            preferred_exporter=exporter_clean,
        )

        # Determine Target Output Path
        if format_clean == "pdf":
            target_path = Path(report.pdf_path) if report.pdf_path else None
            media_type = "application/pdf"
            out_ext = ".pdf"
        else:
            target_path = Path(report.pptx_path) if report.pptx_path else None
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            out_ext = ".pptx"

        if not target_path or not target_path.is_file():
            raise RuntimeError(f"Pipeline completed but target {format_clean.upper()} file was not generated.")

        # Derive Topic-based Attachment Filename
        doc_ast = MarkdownParser().parse_file(temp_md_path)
        attachment_filename = sanitize_topic_filename(
            topic=doc_ast.title,
            fallback_stem=temp_md_path.stem,
            file_ext=out_ext,
        )

        logger.info("API successfully generated presentation '%s' -> attachment '%s'", original_filename, attachment_filename)

        # Return downloadable FileResponse with Starlette BackgroundTask for automatic cleanup
        return FileResponse(
            path=str(target_path.resolve()),
            media_type=media_type,
            filename=attachment_filename,
            background=BackgroundTask(cleanup_directory, temp_dir),
        )

    except HTTPException:
        cleanup_directory(temp_dir)
        raise
    except Exception as err:
        logger.error("API presentation generation failed for '%s': %s", original_filename, err, exc_info=True)
        cleanup_directory(temp_dir)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail={
                "error": "generation_failed",
                "message": "Presentation generation failed.",
                "detail": f"Generation pipeline error: {str(err)}",
            },
        )


@router.post(
    "/api/v1/generate",
    summary="Generate PowerPoint or PDF Presentation",
    description="Accepts a Markdown file as a multipart/form-data upload, processes it through the corporate presentation pipeline, and returns the generated PPTX or PDF file attachment.",
    responses={
        200: {
            "description": "Generated presentation file (PPTX or PDF) returned as a downloadable attachment.",
            "content": {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation": {
                    "schema": {"type": "string", "format": "binary"}
                },
                "application/pdf": {
                    "schema": {"type": "string", "format": "binary"}
                },
            },
        },
        400: {"model": ErrorResponse, "description": "Invalid request (unsupported file extension or format)."},
        404: {"model": ErrorResponse, "description": "Required corporate template file not found."},
        422: {"model": ErrorResponse, "description": "Unprocessable request parameters."},
        500: {"model": ErrorResponse, "description": "Unexpected presentation generation error."},
        503: {"model": ErrorResponse, "description": "Requested PDF export unavailable (no PDF engine installed)."},
    },
)
async def generate_presentation_endpoint(
    file: UploadFile = File(..., description="Input Markdown file (.md, .markdown, or .txt)."),
    template: Optional[UploadFile] = File(None, description="Optional custom PowerPoint template (.pptx) file upload."),
    format: Optional[str] = Form(None, description="Target output format ('pptx' or 'pdf'). Defaults to 'pptx' if empty."),
    preferred_exporter: Optional[str] = Form(None, description="Preferred PDF exporter engine ('libreoffice')."),
) -> FileResponse:
    """REST API endpoint converting uploaded Markdown into an MC-branded PPTX or PDF presentation."""
    format_clean = format.strip().lower() if (format and isinstance(format, str) and format.strip() and format.strip().lower() != "string") else "pptx"
    exporter_clean = preferred_exporter.strip().lower() if (preferred_exporter and isinstance(preferred_exporter, str) and preferred_exporter.strip()) else "libreoffice"

    if format_clean not in ALLOWED_FORMATS:
        logger.warning("API rejected invalid format request '%s'", format)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={
                "error": "invalid_format",
                "message": f"Format '{format}' is not supported. Allowed values: 'pptx', 'pdf'",
            },
        )

    return await _process_presentation_request(
        file=file,
        template=template,
        format_clean=format_clean,
        preferred_exporter=exporter_clean,
    )


