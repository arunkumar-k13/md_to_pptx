"""FastAPI Application Entrypoint for Markdown-to-PowerPoint REST API."""

from __future__ import annotations
import logging
from typing import Any, Dict

from fastapi import FastAPI, HTTPException, Request, status
from fastapi.openapi.utils import get_openapi
from fastapi.responses import JSONResponse

from md_to_pptx.api.routes import router
from md_to_pptx.utils.logger import setup_logger

logger = setup_logger(__name__)

app = FastAPI(
    title="Markdown to PowerPoint Generator API",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
)


def custom_openapi() -> Dict[str, Any]:
    """Custom OpenAPI schema generator ensuring binary media types for 200 response."""
    if app.openapi_schema:
        return app.openapi_schema

    openapi_schema = get_openapi(
        title=app.title,
        version=app.version,
        routes=app.routes,
    )

    # Clean up 200 response schema for POST /api/v1/generate to explicitly document binary PPTX and PDF output
    paths = openapi_schema.get("paths", {})
    generate_post = paths.get("/api/v1/generate", {}).get("post", {})
    responses = generate_post.get("responses", {})

    if "200" in responses:
        responses["200"] = {
            "description": "Generated presentation file (PPTX or PDF) returned as a downloadable attachment.",
            "content": {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation": {
                    "schema": {
                        "type": "string",
                        "format": "binary",
                    },
                    "example": "",
                },
                "application/pdf": {
                    "schema": {
                        "type": "string",
                        "format": "binary",
                    },
                    "example": "",
                },
            },
        }

    app.openapi_schema = openapi_schema
    return app.openapi_schema


app.openapi = custom_openapi


@app.exception_handler(HTTPException)
async def custom_http_exception_handler(request: Request, exc: HTTPException) -> JSONResponse:
    """Ensure consistent structured JSON error responses across all HTTP exception handlers."""
    detail = exc.detail
    if isinstance(detail, dict):
        content = detail
    else:
        content = {
            "error": "http_error",
            "message": "An HTTP error occurred.",
            "detail": str(detail),
        }
    return JSONResponse(status_code=exc.status_code, content=content)


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception) -> JSONResponse:
    """Catch-all exception handler suppressing stack traces while logging server errors."""
    logger.error("Unhandled API exception on %s: %s", request.url, exc, exc_info=True)
    content = {
        "error": "internal_server_error",
        "message": "An unexpected server error occurred.",
        "detail": "Generation pipeline encountered an unhandled exception.",
    }
    return JSONResponse(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, content=content)


# Mount REST API Routes
app.include_router(router)
