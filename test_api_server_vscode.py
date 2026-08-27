"""VSCode Live Server Testing Script for Markdown-to-PPTX.

Use this script to test sending raw Markdown strings via HTTP requests
against a running Uvicorn server instance (e.g. http://localhost:8000).
Generated PowerPoint files will be saved in 'vscode_outputs/'.
"""

import requests
from pathlib import Path

SERVER_URL = "http://localhost:8000"
OUTPUT_DIR = Path("vscode_outputs").resolve()
OUTPUT_DIR.mkdir(exist_ok=True)

# Example Raw Markdown String
MY_CUSTOM_MARKDOWN = """# Executive Briefing: AI Transformation in Healthcare

## Slide 1: Executive Summary
* **Primary Focus:** Deploying machine learning models in clinical diagnostics.
* **Core Goal:** Reduce diagnostic turnaround time by 45%.
* **Target Audience:** Chief Medical Officers & IT Leadership.

---

## Slide 2: Key Milestones & Metrics
| Milestone | Status | Target Date | Impact |
| :--- | :--- | :--- | :--- |
| Model Validation | Complete | 2026-06-30 | 99.4% Accuracy |
| EHR Integration | In Progress | 2026-09-15 | Real-time alerts |
| Clinical Pilot | Planned | 2026-12-01 | 5 Hospital Sites |

---

## Slide 3: Code Implementation
```python
def process_patient_record(record_id: str) -> dict:
    # Query diagnostic model pipeline
    result = model.predict(record_id)
    return {"status": "SUCCESS", "confidence": 0.98}
```
"""


def test_live_server(markdown_text: str, filename_stem: str = "live_test") -> None:
    """Send a raw Markdown string to live FastAPI server via HTTP POST."""
    endpoint_url = f"{SERVER_URL}/api/v1/markdown-to-pptx"
    print(f"Sending HTTP POST to live server endpoint: {endpoint_url}")

    # Pass raw string in-memory as a multipart file upload
    files = {
        "file": (f"{filename_stem}.md", markdown_text.encode("utf-8"), "text/markdown")
    }

    try:
        response = requests.post(endpoint_url, files=files)
        if response.status_code == 200:
            out_path = OUTPUT_DIR / f"{filename_stem}_live.pptx"
            out_path.write_bytes(response.content)
            print(f"SUCCESS! Received PPTX from server. Saved to:\n  -> {out_path}")
        else:
            print(f"SERVER ERROR ({response.status_code}): {response.text}")
    except requests.exceptions.ConnectionError:
        print(f"CONNECTION ERROR: Could not connect to {SERVER_URL}.")
        print("Make sure uvicorn is running: python -m uvicorn md_to_pptx.api.app:app --host 0.0.0.0 --port 8000")


if __name__ == "__main__":
    print("Testing live server with raw Markdown string input...\n")
    test_live_server(MY_CUSTOM_MARKDOWN, "custom_live_briefing")
