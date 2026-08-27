"""Pydantic Schema Definitions for REST API Requests and Responses."""

from __future__ import annotations
from typing import Optional
from pydantic import BaseModel, Field


class HealthResponse(BaseModel):
    """Service health check response payload."""

    status: str = Field("ok", description="Service operational status.")
    service: str = Field("markdown-to-powerpoint-generator", description="Service identifier string.")


class ErrorResponse(BaseModel):
    """Standardized API error response payload."""

    error: str = Field(..., description="Machine-readable error type string.")
    message: str = Field(..., description="Summary of the error condition.")
    detail: Optional[str] = Field(None, description="Detailed user-safe explanation.")
