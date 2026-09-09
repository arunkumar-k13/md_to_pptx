"""Unit and integration test suite for FastAPI REST API endpoints."""

import io
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi.testclient import TestClient
from pptx import Presentation as PPTXPresentation

from md_to_pptx.api.app import app

client = TestClient(app)


class TestAPIEndpoints(unittest.TestCase):
    """Test cases validating FastAPI REST endpoints including /api/v1/generate and /api/v1/markdown-to-pptx."""

    def setUp(self):
        self.sample_md_content = """# Executive Test Presentation

## Strategic Overview
- First key milestone achieved in Q3
- Operations efficiency improved by 25%

## Data Metrics
| Metric | Q1 | Q2 | Q3 |
| :--- | :--- | :--- | :--- |
| Revenue | $1.2M | $1.5M | $1.8M |
| Growth | +10% | +15% | +20% |
"""
        self.template_path = Path("corporate_template.pptx").resolve()
        self.assertTrue(self.template_path.is_file(), "corporate_template.pptx must exist for API testing.")

    def test_health_check_endpoint(self):
        """GET /health returns 200 OK and status ok."""
        response = client.get("/health")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data.get("status"), "ok")
        self.assertEqual(data.get("service"), "markdown-to-powerpoint-generator")

    def test_markdown_to_pptx_endpoint_simple_markdown(self):
        """POST /api/v1/generate generates a valid downloadable PPTX attachment from simple Markdown."""
        files = {"file": ("simple.md", io.BytesIO(b"# Simple Slide\n\n- Bullet point one\n- Bullet point two\n"), "text/markdown")}
        response = client.post("/api/v1/generate", files=files)
        
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.assertTrue(len(response.content) > 1000)

        # Verify returned binary is a valid PowerPoint presentation
        prs = PPTXPresentation(io.BytesIO(response.content))
        self.assertGreaterEqual(len(prs.slides), 1)

    def test_markdown_to_pptx_endpoint_content_heavy_markdown_with_tables(self):
        """POST /api/v1/generate handles content-heavy Markdown with headings and tables."""
        files = {"file": ("executive.md", io.BytesIO(self.sample_md_content.encode("utf-8")), "text/markdown")}
        response = client.post("/api/v1/generate", files=files)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        
        prs = PPTXPresentation(io.BytesIO(response.content))
        self.assertGreaterEqual(len(prs.slides), 3)

    def test_markdown_to_pptx_endpoint_custom_template_upload(self):
        """POST /api/v1/generate accepts an optional custom PPTX template file upload."""
        tmpl_bytes = self.template_path.read_bytes()
        files = {
            "file": ("custom_test.md", io.BytesIO(b"# Custom Template Test\n\n- Content point\n"), "text/markdown"),
            "template": ("custom.pptx", io.BytesIO(tmpl_bytes), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        }
        response = client.post("/api/v1/generate", files=files)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        prs = PPTXPresentation(io.BytesIO(response.content))
        self.assertGreaterEqual(len(prs.slides), 1)

    def test_markdown_to_pptx_does_not_invoke_pdf_exporter(self):
        """POST /api/v1/generate generates ONLY PPTX and never calls PDF / LibreOffice exporter when format='pptx'."""
        files = {"file": ("nopdf.md", io.BytesIO(b"# No PDF Test\n\n- Paragraph content\n"), "text/markdown")}
        
        from md_to_pptx.main import generate_presentation
        with patch("md_to_pptx.api.routes.generate_presentation", side_effect=generate_presentation) as spy_gen:
            response = client.post("/api/v1/generate", files=files)
            self.assertEqual(response.status_code, 200)
            spy_gen.assert_called_once()
            kwargs = spy_gen.call_args.kwargs
            self.assertFalse(kwargs.get("export_pdf"), "export_pdf must be False when format is pptx")

    def test_markdown_to_pptx_invalid_file_extension(self):
        """POST /api/v1/generate rejects unsupported input file extensions with 400 Bad Request."""
        files = {"file": ("invalid.pdf", io.BytesIO(b"%PDF-1.4 header"), "application/pdf")}
        response = client.post("/api/v1/generate", files=files)
        
        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("invalid_file_type", str(data))

    def test_markdown_to_pptx_invalid_template_extension(self):
        """POST /api/v1/generate rejects non-PPTX template files with 400 Bad Request."""
        files = {
            "file": ("test.md", io.BytesIO(b"# Test\n"), "text/markdown"),
            "template": ("bad_template.docx", io.BytesIO(b"word content"), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        }
        response = client.post("/api/v1/generate", files=files)

        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("invalid_template_file", str(data))

    def test_markdown_to_pptx_empty_file(self):
        """POST /api/v1/markdown-to-pptx rejects empty file uploads with 400 Bad Request."""
        files = {"file": ("empty.md", io.BytesIO(b""), "text/markdown")}
        response = client.post("/api/v1/generate", files=files)

        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("empty_file", str(data))

    def test_existing_generate_endpoint_pptx_and_pdf(self):
        """POST /api/v1/generate continues to work for both PPTX and PDF formats without breaking."""
        files_pptx = {"file": ("existing.md", io.BytesIO(b"# Existing Endpoint PPTX\n\n- Content\n"), "text/markdown")}
        response_pptx = client.post("/api/v1/generate", files=files_pptx, data={"format": "pptx"})
        self.assertEqual(response_pptx.status_code, 200)
        self.assertEqual(
            response_pptx.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

        files_bad = {"file": ("existing.md", io.BytesIO(b"# Existing Endpoint PPTX\n\n- Content\n"), "text/markdown")}
        response_bad = client.post("/api/v1/generate", files=files_bad, data={"format": "docx"})
        self.assertEqual(response_bad.status_code, 400)
        res_data = response_bad.json()
        self.assertIn("invalid_format", str(res_data))



if __name__ == "__main__":
    unittest.main()
