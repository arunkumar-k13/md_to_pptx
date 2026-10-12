@echo off
echo Starting Molecular Connections Markdown-to-PowerPoint REST API...
echo Swagger Documentation available at http://127.0.0.1:9500/docs
echo.
python -m uvicorn md_to_pptx.api.app:app --host 0.0.0.0 --port 9500
pause
