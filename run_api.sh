#!/bin/sh
echo "Starting Molecular Connections Markdown-to-PowerPoint REST API..."
echo "Documentation available at http://127.0.0.1:8000/docs"
echo ""
python -m uvicorn md_to_pptx.api.app:app --host 127.0.0.1 --port 8000 --reload
