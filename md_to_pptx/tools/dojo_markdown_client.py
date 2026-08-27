"""DOJO Endpoint Markdown Client Utility.

Standalone development utility to communicate with the DOJO API endpoint,
extract Markdown responses, and save them locally to 'samples/generated.md'
or 'temp/generated.md' for converter validation.

This tool operates completely independently of the PowerPoint converter engine.
"""

from __future__ import annotations
import argparse
import json
import logging
from pathlib import Path
import os
from typing import Optional, Tuple, Union
import requests

from md_to_pptx.utils.path_utils import get_unique_filepath

logger = logging.getLogger(__name__)

DEFAULT_ENDPOINT = os.getenv("DOJO_ENDPOINT", "https://mc-apps.molecularconnections.dev/MCPROMPTPILOT_API/api/v1/dojo_chat_pdf")
DEFAULT_LOGIN_ID = os.getenv("DOJO_LOGIN_ID", "PowerPointTest_IT")
DEFAULT_PROJECT = os.getenv("DOJO_PROJECT", "dojo")
DEFAULT_MODEL = os.getenv("DOJO_MODEL", "gemini-3-flash-preview")


class DojoMarkdownClient:
    """HTTP Client communicating with DOJO API to fetch and extract Markdown."""

    def __init__(
        self,
        endpoint_url: str = DEFAULT_ENDPOINT,
        timeout_seconds: int = 120,
    ) -> None:
        """Initialize DOJO API Markdown Client.

        Args:
            endpoint_url: DOJO API endpoint URL.
            timeout_seconds: HTTP request timeout in seconds.
        """
        self.endpoint_url = endpoint_url
        self.timeout_seconds = timeout_seconds

    def fetch_markdown(
        self,
        prompt: str,
        input_text: str = "string",
        file_path: Optional[Union[str, Path]] = None,
        model: str = DEFAULT_MODEL,
        project: str = DEFAULT_PROJECT,
        login_id: str = DEFAULT_LOGIN_ID,
        output_file_path: Union[str, Path] = "samples/generated.md",
    ) -> Tuple[str, Path]:
        """Send a multipart/form-data request to DOJO endpoint and extract Markdown.

        Args:
            prompt: Prompt instruction string for DOJO.
            input_text: Additional input text (default: 'string').
            file_path: Optional path to attached file for analysis.
            model: DOJO AI model identifier.
            project: Project tag identifier.
            login_id: Client authentication login ID.
            output_file_path: Target path for saving extracted .md file.

        Returns:
            Tuple of (extracted_markdown_content_string, saved_file_path_object).
        """
        # If input_text is default 'string' but prompt is provided, pass prompt as inputText
        # to ensure DOJO RAG endpoint receives the target topic context
        effective_input_text = input_text
        effective_prompt = prompt
        if input_text == "string" and prompt:
            effective_input_text = prompt
            effective_prompt = "Generate a structured slide presentation in Markdown format with titles (#), sections (##), bullet points, and tables."

        # Construct multipart/form-data fields dictionary for requests
        # Spring Boot endpoint requires Content-Type: multipart/form-data for ALL requests
        files_dict = {
            "loginId": (None, login_id),
            "inputText": (None, effective_input_text),
            "prompt": (None, effective_prompt),
            "project": (None, project),
            "model": (None, model),
        }

        opened_file = None
        if file_path:
            p = Path(file_path).resolve()
            if not p.is_file():
                raise FileNotFoundError(f"File not found for upload: {file_path}")
            opened_file = open(p, "rb")
            files_dict["file"] = (p.name, opened_file, "application/octet-stream")

        print(f"[OK] Request sent to DOJO endpoint ({self.endpoint_url})")
        logger.info("Sending request to DOJO endpoint: %s", self.endpoint_url)

        try:
            response = requests.post(
                self.endpoint_url,
                files=files_dict,
                timeout=self.timeout_seconds,
            )
            response.raise_for_status()
        except requests.exceptions.Timeout as err:
            logger.warning("Request to DOJO API timed out after %d seconds. Retrying once...", self.timeout_seconds)
            print(f"[WARN] Request timed out after {self.timeout_seconds}s. Retrying once with extended timeout...")
            if opened_file and not opened_file.closed:
                opened_file.seek(0)
            response = requests.post(
                self.endpoint_url,
                files=files_dict,
                timeout=self.timeout_seconds * 2,
            )
            response.raise_for_status()
        finally:
            if opened_file:
                opened_file.close()

        raw_body = response.text
        print(f"[OK] Response received (Status Code: {response.status_code})")
        logger.info("Response received from DOJO API (Status %d)", response.status_code)

        # Automatically detect JSON vs raw Markdown
        extracted_md = self.extract_markdown_from_response(raw_body)
        print(f"[OK] Markdown length: {len(extracted_md):,} characters ({len(extracted_md.splitlines())} lines)")

        # Save to output file
        raw_out_p = Path(output_file_path).resolve()
        out_p = get_unique_filepath(raw_out_p)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text(extracted_md, encoding="utf-8")
        print(f"[OK] File saved to {out_p.name} ({out_p})")

        return extracted_md, out_p

    @staticmethod
    def extract_markdown_from_response(raw_body: str) -> str:
        """Extract Markdown text whether response is raw Markdown or JSON payload.

        Args:
            raw_body: Raw HTTP response text string.

        Returns:
            Clean Markdown content string.
        """
        body_trimmed = raw_body.strip()
        if not body_trimmed:
            return ""

        try:
            parsed_json = json.loads(body_trimmed)
            val = None
            if isinstance(parsed_json, dict):
                # Search common keys for markdown payload (including DOJO's llmResponse key)
                for key in ("llmResponse", "markdown", "response", "content", "text", "data", "output", "result"):
                    if key in parsed_json and isinstance(parsed_json[key], str):
                        val = parsed_json[key]
                        break

                if val is None:
                    # Search nested dictionaries
                    for parent_key in ("data", "result", "payload"):
                        if parent_key in parsed_json and isinstance(parsed_json[parent_key], dict):
                            sub_dict = parsed_json[parent_key]
                            for sub_key in ("llmResponse", "markdown", "response", "content", "text", "output"):
                                if sub_key in sub_dict and isinstance(sub_dict[sub_key], str):
                                    val = sub_dict[sub_key]
                                    break
                            if val:
                                break

                if val is None:
                    return json.dumps(parsed_json, indent=2)
            elif isinstance(parsed_json, str):
                val = parsed_json

            if val:
                val = val.strip()
                # Unescape if payload is double JSON-encoded (e.g. "\"# Title...\"")
                if val.startswith('"') and val.endswith('"'):
                    try:
                        val = json.loads(val)
                    except Exception:
                        val = val[1:-1]
                return val
        except Exception:
            pass

        # Return raw body if not JSON
        return raw_body


def main() -> None:
    """CLI entrypoint for testing DOJO Markdown client independently."""
    parser = argparse.ArgumentParser(description="DOJO Markdown Client Development Utility")
    parser.add_argument("-p", "--prompt", required=True, help="Prompt instruction string")
    parser.add_argument("-i", "--input-text", default="string", help="Input text string")
    parser.add_argument("-f", "--file", help="Optional path to attached file")
    parser.add_argument("-m", "--model", default=DEFAULT_MODEL, help="DOJO model identifier")
    parser.add_argument("-o", "--output", default="samples/generated.md", help="Target output file path")
    parser.add_argument("--timeout", type=int, default=180, help="HTTP request timeout in seconds (default: 180)")

    args = parser.parse_args()

    client = DojoMarkdownClient(timeout_seconds=args.timeout)
    client.fetch_markdown(
        prompt=args.prompt,
        input_text=args.input_text,
        file_path=args.file,
        model=args.model,
        output_file_path=args.output,
    )


if __name__ == "__main__":
    main()
