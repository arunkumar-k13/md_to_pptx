#!/bin/sh
echo "Starting Molecular Connections Markdown-to-PowerPoint REST API..."
echo "Swagger Documentation available at http://0.0.0.0:9500/docs"
echo ""
python -m uvicorn md_to_pptx.api.app:app --host 0.0.0.0 --port 9500
