# md_to_pptx — Markdown to PowerPoint & PDF Generation Service

`md_to_pptx` is a FastAPI REST API, CLI, and Python library service that converts Markdown documents into Molecular Connections (MC) corporate branded PowerPoint (`.pptx`) presentations and PDF (`.pdf`) documents.

---

## 1. Overview

`md_to_pptx` transforms Markdown files (`.md`, `.markdown`, `.txt`) into structured presentation slides. It automates content parsing, LaTeX mathematical formatting, presentation design, template introspection, brand layout matching, text overflow pagination, layout boundary protection, and PDF export.

---

## 2. Key Features

- **Markdown → PPTX & PDF**: Converts Markdown into branded PowerPoint presentations and PDF documents.
- **Corporate & Custom PPTX Templates**: Uses `corporate_template.pptx` by default or dynamically introspects uploaded custom `.pptx` templates.
- **LaTeX & Mathematical Expressions**: Parses inline and block LaTeX equations, fractions, roots, matrices, subscripts/superscripts, and scientific notation.
- **Brand Slide Manager**: Preserves corporate branded slides (`BRAND_COVER`, `BRAND_CASE_STUDY`, `BRAND_ABOUT_COMPANY`, `BRAND_THANK_YOU`, `BRAND_DISCLAIMER`).
- **Calibrated Overflow & Pagination Engine**: Evaluates content heights against usable slide bounds, keeping standard lists (up to 10 items) and multi-heading sections intact while splitting overcrowded slides cleanly into `(Continued)` slides above the footer line.
- **Table Word-Wrap Protection**: Column width calculations enforce word length thresholds to prevent mid-word hyphenation splits.
- **LibreOffice PDF Export**: Headless PDF conversion via LibreOffice (`soffice`), eliminating desktop Microsoft PowerPoint software dependencies on Linux servers.
- **Open XML Validation**: Validates generated `.pptx` package structures before file delivery.

---

## 3. Architecture

```text
Client / HTTP POST (/api/v1/generate)
  |
  v
Markdown Parser & LaTeX Formatter
  |
  v
Content Analyzer & Slide Planner
  |
  v
Template Introspector & Brand Slide Manager
  |
  v
Spatial Overflow & Pagination Engine
  |
  v
PPTX Generator
  |
  +----> PPTX Presentation Output (.pptx)
  |
  +----> LibreOffice Exporter ----> PDF Document Output (.pdf)
```

---

## 4. Project Structure

```text
DOJO/
├── README.md                      # Documentation
├── requirements.txt               # Runtime & test dependencies
├── Dockerfile                     # Production container definition
├── .dockerignore                  # Docker build exclusions
├── .env.example                   # Environment configuration template
├── .gitignore                     # Git exclusion rules
├── corporate_template.pptx        # Default Molecular Connections corporate template
├── run_api.bat                    # Windows API launcher script
├── run_api.sh                     # Linux/macOS API launcher script
├── output/                        # Directory for CLI generated outputs
├── samples/                       # Sample Markdown & stress-test inputs
└── md_to_pptx/                    # Core Python Package
    ├── main.py                    # CLI & Pipeline Orchestrator
    ├── api/                       # FastAPI App, Routes & Schemas
    │   ├── app.py
    │   ├── routes.py
    │   └── schemas.py
    ├── config/                    # Settings Loader & YAML Configuration
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
    │   ├── exporters/             # PDF Exporters (LibreOffice)
    │   ├── generator/             # PowerPoint Generator & Shape Renderers
    │   ├── overflow/              # Spatial Overflow Analyzer & Paginator
    │   ├── qa/                    # Package & Visual QA Validators
    │   └── template/              # Template Introspection & Layout Resolvers
    ├── reporting/                 # Diagnostic Context & Generation Reports
    ├── tests/                     # Automated Unit & Regression Test Suites
    └── utils/                     # Logger & Path Utilities
```

---

## 5. System Requirements

- **Python Version**: Python 3.9+
- **Operating System**: Linux, Windows, or macOS
- **Microsoft PowerPoint Software**: **NOT required**
- **PDF Export Requirement**: **LibreOffice** (`soffice`) must be installed on the host machine (or container) when `format=pdf` is requested

---

## 6. Installation & Running the Server

### Option A: Python Virtual Environment

```bash
# 1. Create and activate virtual environment
python -m venv venv
# Linux/macOS:
source venv/bin/activate
# Windows:
venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the FastAPI server on port 9500
python -m uvicorn md_to_pptx.api.app:app --host 0.0.0.0 --port 9500
```

Or use the provided launcher scripts:
- **Linux / macOS**: `./run_api.sh`
- **Windows**: `run_api.bat`

### Option B: Docker Container

```bash
# Build image
docker build -t md-to-pptx .

# Run container on port 9500
docker run -d -p 9500:9500 --name md-to-pptx-service md-to-pptx
```

---

## 7. API Documentation

When the server is running, interactive documentation is available at:
- **Swagger UI**: `http://<HOST_IP>:9500/docs`
- **ReDoc**: `http://<HOST_IP>:9500/redoc`
- **OpenAPI JSON**: `http://<HOST_IP>:9500/openapi.json`

---

## 8. REST API Endpoints

### 1. `POST /api/v1/generate` — Generate PowerPoint or PDF Presentation

Converts an uploaded Markdown file into an MC-branded PowerPoint presentation (`.pptx`) or PDF document (`.pdf`).

#### Request Parameters (`multipart/form-data`)

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `file` | File | **Yes** | — | Input Markdown file (`.md`, `.markdown`, `.txt`). |
| `template` | File | No | `corporate_template.pptx` | Optional custom PowerPoint template file (`.pptx`). |
| `format` | Form (String) | No | `pptx` | Output format: `'pptx'` or `'pdf'`. |
| `preferred_exporter` | Form (String) | No | `libreoffice` | PDF exporter engine (`'libreoffice'`). |

#### Example Request (cURL)

```bash
curl -X POST http://localhost:9500/api/v1/generate \
  -F "file=@samples/sample_executive_presentation.md" \
  -F "format=pptx" \
  -o output_presentation.pptx
```

#### Example Request (Python `requests`)

```python
import requests

url = "http://localhost:9500/api/v1/generate"
with open("samples/sample_executive_presentation.md", "rb") as f:
    response = requests.post(
        url,
        files={"file": ("sample.md", f, "text/markdown")},
        data={"format": "pptx"},
    )

if response.status_code == 200:
    with open("generated_deck.pptx", "wb") as out:
        out.write(response.content)
```

---

### 2. `GET /health` — Service Health Check

Returns the operational status of the service:

```json
{
  "status": "ok",
  "service": "markdown-to-powerpoint-generator"
}
```

---

## 9. Environment Variables

| Variable | Default | Description |
| :--- | :--- | :--- |
| `DEFAULT_PPTX_TEMPLATE` | `corporate_template.pptx` | Path to default PowerPoint corporate template. |
| `LIBREOFFICE_BINARY` | Auto-detected (`soffice`) | Path to LibreOffice executable for PDF export. |
| `MAX_UPLOAD_SIZE_MB` | `50` | Maximum allowed upload file size in MB. |

---

## 10. CLI Usage

```bash
# Generate PPTX from CLI
python -m md_to_pptx.main samples/sample_executive_presentation.md --template corporate_template.pptx --output output/

# Generate PPTX and PDF from CLI
python -m md_to_pptx.main samples/sample_executive_presentation.md --template corporate_template.pptx --output output/ --pdf
```

---

## 11. Running Automated Tests

```bash
python -m unittest discover -s md_to_pptx/tests -v
```
