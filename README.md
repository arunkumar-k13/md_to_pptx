# md_to_pptx — Markdown to PowerPoint & PDF Generation Service

`md_to_pptx` is a REST API, CLI, and Python library service that converts Markdown documents into Molecular Connections (MC) corporate branded PowerPoint (`.pptx`) presentations and PDF (`.pdf`) documents.

---

## 1. Overview

`md_to_pptx` transforms plain Markdown files (`.md`, `.markdown`, `.txt`) into structured presentation slides. It automates content parsing, presentation design, template introspection, brand layout matching, text overflow pagination, layout boundary protection, and PDF export.

---

## 2. Key Features

- **Markdown → PPTX & PDF**: Converts raw Markdown into branded PowerPoint presentations and PDF documents.
- **Corporate & Custom PPTX Templates**: Uses `corporate_template.pptx` by default or dynamically introspects uploaded custom `.pptx` templates.
- **Dynamic Template Introspection**: Automatically derives slide layout catalogs, placeholder coordinates, and master slide styles.
- **Brand Slide Manager**: Preserves pre-existing corporate branded slides (`BRAND_COVER`, `BRAND_CASE_STUDY`, `BRAND_ABOUT_COMPANY`, `BRAND_THANK_YOU`, `BRAND_DISCLAIMER`).
- **Overflow & Pagination Engine**: Estimates text frame height based on font sizes, padding, and line wrapping, paginating long bullet lists into continuation slides (`"{Title} (Continued)"`).
- **Layout Boundary Protection**: Enforces physical boundaries (`usable_bottom`) above footer placeholders to prevent text from crossing into reserved slide areas.
- **Table Word-Wrap Protection**: Column width calculations enforce word length thresholds to eliminate hyphenated word splits (e.g. `Framework` -> `Framewo` / `rk`).
- **LibreOffice PDF Export**: Headless PDF conversion via LibreOffice (`soffice`), eliminating desktop Microsoft PowerPoint software dependencies.
- **Open XML Validation**: Validates generated `.pptx` ZIP structures before file delivery.
- **Multi-Interface Access**: Accessible via FastAPI REST API, CLI, or Python SDK.

---

## 3. Architecture

```
Client / HTTP POST
  |
  v
FastAPI REST API (/api/v1/generate)
  |
  v
Markdown Parser
  |
  v
Content Analysis / Slide Planning
  |
  v
Template Analyzer
  |
  v
Brand Slide Manager
  |
  v
Overflow + Pagination Engine
  |
  v
PPTX Generator
  |
  +----> PPTX Presentation Output
  |
  +----> LibreOffice Exporter ----> PDF Document Output
```

---

## 4. How the Pipeline Works

1. **Markdown Parser**: Parses input `.md` files into an Abstract Syntax Tree (AST) containing document metadata, titles, section headers, bullet lists, tables, and quote blocks.
2. **Content Analysis & Slide Planning**: Maps AST sections into structured slide intents (`TITLE_COVER`, `BULLET_CONTENT`, `BRAND_CASE_STUDY`, etc.) based on heading levels and element types.
3. **Template Analyzer**: Introspects the target PowerPoint template (`.pptx`), mapping available slide layouts, placeholder dimensions, and reserved header/footer bounds.
4. **Brand Slide Manager**: Matches content slides against pre-existing template slides to preserve custom corporate graphics, logos, and layouts.
5. **Overflow & Pagination Engine**: Evaluates content heights against template usable content bounds (`usable_bottom`). If content exceeds available space, it splits the block cleanly into continuation slides.
6. **PPTX Generator**: Populates slide text frames, applies proportional typography scaling, and renders tables and layout shapes.
7. **LibreOffice Exporter**: (Optional) Invokes headless LibreOffice (`soffice`) to convert the generated `.pptx` into a `.pdf` document.

---

## 5. Project Structure

```
DOJO/
├── README.md                      # Single source of truth documentation
├── requirements.txt               # Primary runtime dependency specification
├── .env.example                   # Environment configuration template
├── .gitignore                     # Git exclusion rules
├── corporate_template.pptx        # Default Molecular Connections corporate template
├── run_api.bat                    # Windows API launcher script
├── run_api.sh                     # Linux/macOS API launcher script
├── output/                        # Directory for generated outputs (contains .gitkeep)
├── samples/                       # Sample Markdown input files
│   ├── generated.md
│   └── sample_executive_presentation.md
└── md_to_pptx/                    # Core Python Package
    ├── api/                       # FastAPI REST API App, Routes & Schemas
    │   ├── app.py
    │   ├── routes.py
    │   └── schemas.py
    ├── config/                    # Settings Loader & Settings YAML
    │   ├── settings.yaml
    │   └── settings_loader.py
    ├── core/                      # Presentation Core Pipeline
    │   ├── ast_nodes.py
    │   ├── brand_slide_manager.py
    │   ├── content_analyzer.py
    │   ├── design_engine.py
    │   ├── markdown_parser.py
    │   ├── presentation_model.py
    │   ├── slide_planner.py
    │   ├── exporters/             # PDF Exporters (LibreOffice & Legacy COM Stub)
    │   │   ├── base_exporter.py
    │   │   ├── com_exporter.py
    │   │   ├── environment.py
    │   │   ├── exporter_factory.py
    │   │   └── libreoffice_exporter.py
    │   ├── generator/             # PowerPoint Generator & Shape Renderers
    │   │   ├── pptx_generator.py
    │   │   └── shape_renderers.py
    │   ├── overflow/              # Spatial Overflow Analyzer & Paginator
    │   │   ├── overflow_analyzer.py
    │   │   ├── overflow_strategy.py
    │   │   └── paginator.py
    │   ├── qa/                    # Package & Visual QA Validators
    │   │   ├── package_validator.py
    │   │   └── visual_qa.py
    │   └── template/              # Template Introspection & Resolvers
    │       ├── layout_resolver.py
    │       ├── template_analyzer.py
    │       ├── template_cache.py
    │       └── template_validator.py
    ├── reporting/                 # Diagnostic Context & Generation Reports
    │   ├── diagnostic_context.py
    │   └── generation_report.py
    ├── tools/                     # DOJO RAG Utility (Optional)
    │   └── dojo_markdown_client.py
    └── utils/                     # Logger & Path Utilities
        ├── logger.py
        └── path_utils.py
```

---

## 6. System Requirements

- **Python Version**: Python 3.9+ (Tested on Python 3.14).
- **Operating System**: Windows, Linux, or macOS.
- **Microsoft PowerPoint Software**: **NOT required**. PowerPoint COM automation is disabled in production.
- **PDF Export Requirement**: **LibreOffice** (`soffice`) must be installed on the host machine for PDF generation.

---

## 7. Installation

Primary runtime dependency installation:

```bash
# 1. Create virtual environment
python -m venv venv

# Windows activate
venv\Scripts\activate
# Linux/macOS activate
source venv/bin/activate

# 2. Install primary dependencies
pip install -r requirements.txt
```

---

## 8. Configuration / Environment Variables

The application can be configured via environment variables or a `.env` file (see `.env.example`):

| Variable | Default | Description |
| :--- | :--- | :--- |
| `DEFAULT_PPTX_TEMPLATE` | `corporate_template.pptx` | Path to default PowerPoint corporate template. |
| `LIBREOFFICE_BINARY` | Auto-detected | Path to LibreOffice executable (`soffice` / `soffice.exe`). |
| `MAX_UPLOAD_SIZE_MB` | `50` | Maximum allowed REST API upload file size in MB. |
| `DOJO_ENDPOINT` | Configured URL | Optional DOJO RAG service API URL. |
| `DOJO_LOGIN_ID` | `PowerPointTest_IT` | Optional DOJO login ID. |
| `DOJO_PROJECT` | `dojo` | Optional DOJO project tag. |
| `DOJO_MODEL` | `gemini-3-flash-preview` | Optional DOJO model ID. |

---

## 9. Running the FastAPI Server

Start the development server using Uvicorn:

```bash
python -m uvicorn md_to_pptx.api.app:app --host 127.0.0.1 --port 8000 --reload
```

Or run launcher scripts:
- **Windows**: `run_api.bat`
- **Linux / macOS**: `./run_api.sh`

---

## 10. API Documentation

Interactive API documentation is automatically generated by FastAPI:
- **Swagger UI**: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)
- **ReDoc**: [http://127.0.0.1:8000/redoc](http://127.0.0.1:8000/redoc)
- **OpenAPI Schema**: [http://127.0.0.1:8000/openapi.json](http://127.0.0.1:8000/openapi.json)

---

## 11. API Endpoints

### `POST /api/v1/markdown-to-pptx` (Dedicated PPTX Conversion)

Converts an uploaded Markdown file directly into a PowerPoint (`.pptx`) presentation attachment without invoking PDF or LibreOffice conversion.

#### Parameters (`multipart/form-data`)

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `file` | File | **Yes** | — | Input Markdown file (`.md`, `.markdown`, `.txt`). |
| `template` | File | No | Server Default | Optional custom PowerPoint template file (`.pptx`). |

#### cURL Example

```bash
curl.exe -X POST http://localhost:8000/api/v1/markdown-to-pptx \
  -F "file=@samples/generated.md" \
  -o presentation.pptx
```

---

### `POST /api/v1/test/markdown-to-pptx` (Testing Endpoint: File or Text String)

Dedicated testing endpoint accepting EITHER an uploaded Markdown file (`.md`, `.markdown`, `.txt`) OR a raw Markdown text string (`markdown_text`) directly, along with an optional custom PowerPoint template (`.pptx`).

#### Parameters (`multipart/form-data`)

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `file` | File | No | — | Optional input Markdown file (`.md`, `.markdown`, `.txt`). |
| `markdown_text` | Text | No | — | Optional raw Markdown text string content (multiline text block). |
| `template` | File | No | Server Default | Optional custom PowerPoint template file (`.pptx`). |

#### cURL Examples

**Using raw Markdown text string**:
```bash
curl.exe -X POST http://localhost:8000/api/v1/test/markdown-to-pptx \
  -F "markdown_text=# Sample Title`n`n- Point 1`n- Point 2" \
  -o presentation.pptx
```

**Using file upload**:
```bash
curl.exe -X POST http://localhost:8000/api/v1/test/markdown-to-pptx \
  -F "file=@samples/generated.md" \
  -o presentation.pptx
```

---

### `POST /api/v1/generate` (General Presentation & PDF Endpoint)

Converts an uploaded Markdown file into an MC-branded PowerPoint presentation or PDF document depending on the `format` parameter.

#### Parameters (`multipart/form-data`)

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `file` | File | **Yes** | — | Input Markdown file (`.md`, `.markdown`, `.txt`). |
| `template` | File | No | Server Default | Optional custom PowerPoint template file (`.pptx`). |
| `format` | Form | No | `pptx` | Requested output format: `'pptx'` or `'pdf'`. |
| `preferred_exporter` | Form | No | `libreoffice` | Preferred PDF exporter engine (`'libreoffice'`). |

#### cURL Example

```bash
curl.exe -X POST http://localhost:8000/api/v1/generate \
  -F "file=@samples/generated.md" \
  -F "format=pptx" \
  -o presentation.pptx
```

---

### `GET /health` (Service Health Check)

Returns basic operational status:
```json
{
  "status": "ok",
  "service": "markdown-to-powerpoint-generator"
}
```

---

## 13. Custom PowerPoint Templates

The pipeline introspects supplied `.pptx` templates and adapts generated content to available layouts and boundaries. Corporate branded slides can be preserved when recognized by the Brand Slide Manager.

Upload a custom template via API:
```bash
curl -X POST http://127.0.0.1:8000/api/v1/generate \
  -F "file=@samples/generated.md" \
  -F "template=@my_custom_template.pptx" \
  -F "format=pptx" \
  -o custom_presentation.pptx
```

---

## 14. PPTX Generation

The presentation engine constructs Open XML PowerPoint files using `python-pptx`:
- Evaluates title/body placeholder geometries.
- Sets paragraph formatting, bullet points, font sizes, and line spacing.
- Renders tables with word-length column floors.
- Validates the generated ZIP file structure before completion.

---

## 15. PDF Generation with LibreOffice

- PowerPoint `.pptx` is generated first.
- Headless LibreOffice (`soffice --headless --convert-to pdf`) converts `.pptx` to `.pdf`.
- Microsoft PowerPoint software is **NOT** required.
- If LibreOffice is not installed when PDF format is requested, the API returns HTTP `503 Service Unavailable`.

---

## 16. CLI Usage

Generate presentations directly from the command line:

```bash
# Generate PPTX using corporate template
python -m md_to_pptx.main samples/generated.md --template corporate_template.pptx --output output/

# Generate PPTX and PDF via LibreOffice
python -m md_to_pptx.main samples/generated.md --template corporate_template.pptx --output output/ --pdf
```

---

## 17. Python Library Usage

Call presentation generation directly from Python code:

```python
from pathlib import Path
from md_to_pptx.main import generate_presentation

report = generate_presentation(
    markdown_path=Path("samples/generated.md"),
    template_path=Path("corporate_template.pptx"),
    output_dir=Path("output"),
    export_pdf=True,
)

print(f"Generated PPTX: {report.pptx_path}")
print(f"Generated PDF:  {report.pdf_path}")
```

---

## 18. Optional DOJO Integration

DOJO integration is an **optional development utility**. The core presentation generation engine functions completely independently of DOJO.

DOJO utility script: [`md_to_pptx/tools/dojo_markdown_client.py`](file:///c:/Users/arun.kumar/Desktop/MC/DOJO/md_to_pptx/tools/dojo_markdown_client.py)
```bash
python -m md_to_pptx.tools.dojo_markdown_client -p "Summarize AI Innovations" -o samples/generated.md
```

---

## 19. Security / File Handling

- **File Validation**: Validates file extensions (`.md`, `.markdown`, `.txt`, `.pptx`).
- **Path Traversal Protection**: Input filenames are sanitized before writing to temporary workspaces.
- **Upload Size Limits**: Enforces `MAX_UPLOAD_SIZE_MB` (default 50 MB).
- **Cleanup**: Temporary workspace directories are automatically purged after API response completion using Starlette background tasks.

---

## 20. Testing

Run the automated test suite to verify system integrity:

```bash
# Run API endpoint tests
python -m unittest md_to_pptx/tests/test_api.py

# Run complete test suite
python -m unittest discover -s md_to_pptx/tests
```

---

## 21. Troubleshooting

- **PDF Export Fails (HTTP 503 / RuntimeError)**: Ensure LibreOffice is installed and `soffice` is in your system PATH or configured via `LIBREOFFICE_BINARY`.
- **Template File Not Found (HTTP 404)**: Ensure `corporate_template.pptx` is present in the project root or specified via `DEFAULT_PPTX_TEMPLATE`.
- **Text Overflows Footer**: The engine uses template-aware bounds (`usable_bottom`). Verify that placeholder bounds in custom templates are marked correctly.

---

## 22. Development History / Major Changes

- Added Markdown-to-PowerPoint generation pipeline.
- Added FastAPI REST API endpoint (`POST /api/v1/generate`).
- Added corporate template introspection & layout resolver.
- Added corporate Brand Slide Manager for pre-existing slides.
- Added spatial overflow detection and pagination engine.
- Added layout boundary protection (`usable_bottom`) above footer region.
- Added table word-length column floors to prevent hyphenated word breaks.
- Added real alternate template regression testing (`test_alternate_template.pptx`).
- Added headless LibreOffice PDF exporter (`LibreOfficeExporter`).
- Removed Microsoft PowerPoint COM as a production requirement.
- Added environment variable configuration and upload size validation.

---


