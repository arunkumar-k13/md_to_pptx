"""VSCode Manual Testing Script for Markdown-to-PPTX Generation.

Run this script directly inside VSCode to test raw Markdown strings without using a browser or Swagger UI.
Generated PowerPoint files will be saved in the 'vscode_outputs' directory.
"""

import io
from pathlib import Path
import tempfile
from pptx import Presentation as PPTXPresentation

from md_to_pptx.api.app import app
from fastapi.testclient import TestClient

client = TestClient(app)
OUTPUT_DIR = Path("vscode_outputs").resolve()
OUTPUT_DIR.mkdir(exist_ok=True)


# ==============================================================================
# SAMPLE 1: Standard Presentation Elements Test
# ==============================================================================
SAMPLE_1_STANDARD = """# Slide 1: Main Title
## Subtitle or Presenter Name
This is a test of standard body text on a title slide.

---

# Slide 2: Heading Hierarchy
## Level 2 Heading
### Level 3 Heading
#### Level 4 Heading
Testing how different heading levels are mapped to slide titles or nested text boxes.

---

# Slide 3: Text Formatting
- **Bold Text** for emphasis.
- *Italic Text* for secondary emphasis.
- ~~Strikethrough~~ to test deletion styles.
- `Inline Code` for technical terms.
- **_Combined Bold and Italic_**

---

# Slide 4: Lists and Nesting
1. First Ordered Item
2. Second Ordered Item
 - Nested Unordered Item A
 - Nested Unordered Item B
 1. Deeply Nested Item 1
3. Third Ordered Item

---

# Slide 5: Tables
| Feature | Supported | Notes |
| :--- | :---: | :--- |
| Tables | Yes | Testing alignment |
| Images | Yes | Testing scaling |
| Links | Yes | [Click Here](https://example.com) |

---

# Slide 6: Code Blocks
```python
def test_conversion():
    markdown = "PPTX"
    print(f"Converting {markdown}...")
    return True
```

---

# Slide 7: Blockquotes and Rules
> "This is a blockquote to test indentation and stylized callouts within a slide layout."

---

# Slide 8: Images and Links
![Test Image](https://via.placeholder.com/150)
Check the link functionality: [Markdown Guide](https://www.markdownguide.org)
"""


# ==============================================================================
# SAMPLE 2: Research Integrity Presentation Material
# ==============================================================================
SAMPLE_2_RESEARCH_INTEGRITY = """# Make Detailed Presentation Material On Latest Trends In Research Integrity: Strategic Overview & Key Insights

## Slide 1: Executive Summary
*   **Primary Focus:** Comprehensive overview addressing make a detailed presentation material on latest trends in research integrity.
*   **Core Objectives:** Align strategic goals, optimize execution, and deliver measurable impact.
*   **Key Value Proposition:** Accelerate transformation with structured framework and operational excellence.

---

## Slide 2: Market Context & Strategic Drivers
**Key Drivers:**
*   **Industry Alignment:** Rapid evolution and demand for research integrity solutions.
*   **Operational Efficiency:** Streamlining workflows to enhance speed and consistency.
*   **Risk Mitigation:** Proactive identification and handling of potential bottlenecks.
*   **Stakeholder Engagement:** Continuous feedback loops ensuring alignment across teams.

---

## Slide 3: Core Implementation Pillars
1.  **Phase 1 - Assessment & Planning:** Baseline evaluation and goal setting.
2.  **Phase 2 - Execution & Integration:** Deployment of core capabilities and processes.
3.  **Phase 3 - Optimization & Scaling:** Continuous monitoring, refinement, and expansion.
4.  **Phase 4 - Governance & Compliance:** Establishing metrics, quality controls, and auditing.

---

## Slide 4: Key Metrics & Deliverables
| Objective Area | Target Metric | Expected Outcome |
| :--- | :--- | :--- |
| **Strategy & Adoption** | 100% Alignment | Full team onboarding on research integrity |
| **Efficiency Gain** | +30% Improvement | Accelerated cycle times and lower overhead |
| **Quality Assurance** | Zero Critical Defect | Standardized compliance & governance |
| **ROI & Growth** | Measurable Impact | High ROI and sustained operational growth |

---

## Slide 5: Strategic Outcomes & Next Steps
*   **Immediate Action Item:** Finalize implementation roadmap.
*   **Short-term Goal:** Launch pilot initiatives and gather operational feedback.
*   **Long-term Vision:** Scale capabilities enterprise-wide to drive sustainable value.
*   **Conclusion:** Strategic alignment empowers continuous innovation and long-term success.
"""


# ==============================================================================
# SAMPLE 3: Technical Stress-Test (LaTeX Math & Unicode)
# ==============================================================================
SAMPLE_3_LATEX_MATH = """# Technical Stress-Test: LaTeX and Markdown for PPTX Generation

## Purpose
This presentation tests Markdown parsing, newline-sensitive structure, LaTeX rendering, mathematical expressions, tables, lists, code blocks, Unicode characters, special characters, and Markdown-to-PPTX conversion.

---

## I. Fundamental Arithmetic & Algebra
1. Simple inline: $x + y = z$
2. Subscripts/Superscripts: $x_i^2 + y_j^3$
3. Basic Fractions: $\\frac{1}{2}$
4. Square Roots: $\\sqrt{x}$
5. Parentheses scaling: $\\left( \\frac{x}{y} \\right)$

---

## II. Calculus & Analysis
1. Limits: $\\lim_{x \\to 0} f(x)$
2. Derivatives: $\\frac{df}{dx}$
3. Integrals: $\\int_{a}^{b} x^2 dx$
4. Summation: $\\sum_{k=0}^n k^2$

---

## III. Linear Algebra & Set Theory
1. Matrix: $\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}$
2. Sets: $\\{x \\in \\mathbb{R} \\mid x > 0\\}$
3. Implications: $P \\implies Q$
4. Symbols: $\\alpha, \\beta, \\gamma, \\infty, \\nabla, \\partial$
"""


# ==============================================================================
# SAMPLE 4: Comprehensive Edge Cases & Complex Content
# ==============================================================================
SAMPLE_4_EDGE_CASES = """# Markdown to PPTX Conversion Test

## 1. Basic Text Formatting
This is a normal paragraph containing standard text to establish a baseline for font size and style.

*   **This is bold text**
*   *This is italic text*
*   ***This is bold and italic text***
*   ~~This is strikethrough text~~
*   `This is inline code`

**Special Characters:** `! @ # $ % ^ & * ( ) _ + = { } [ ] < > ? / \\ | ~`

**Numbers and Values:**
*   Numbers: 1234567890
*   Dates: 2024-12-31, 15/08/2026
*   Percentages: 45%, 99.99%, 0.001%
*   Currency: $100.00, €50, £20, ₹1,50,000.75

**Unicode characters:** © ® ™ € £ ¥ ₹ ₩ ± × ÷ ≠ ≤ ≥ → ← ↑ ↓ ★ ✓ ✗
**Emoji:** 😀 🚀 🔥 ✅ ⚠️ ❌ 💡 📊 🎯

## 2. Lists & Nesting
1.  Main item
    *   Sub item
    *   Another sub item
2.  Second main item
    *   Sub item
        1.  Nested numbered item
        2.  Another numbered item

## 3. Code Blocks
```python
def calculate_average(numbers):
    return sum(numbers) / len(numbers)
```

## 4. Complex Tables
| Metric | Value | Percentage | Cost | Date | Result |
| ------- | -----: | ---------: | -----------: | ---------- | --------- |
| Users | 12,450 | 87.5% | ₹1,25,000.50 | 2026-08-18 | ✓ Pass |
| Errors | 125 | 1.2% | $1,250.75 | 2026-08-17 | ⚠ Warning |
| Success | 12,325 | 98.8% | €9,999.99 | 2026-08-16 | ✓ Pass |
"""


def test_markdown_string(sample_name: str, markdown_text: str, filename_stem: str) -> None:
    """Test converting a raw Markdown string directly into a PPTX presentation file in VSCode."""
    print(f"\n==================================================================")
    print(f"Testing: {sample_name}")
    print(f"==================================================================")

    # Send in-memory string as a file stream to FastAPI app engine
    response = client.post(
        "/api/v1/markdown-to-pptx",
        files={"file": (f"{filename_stem}.md", io.BytesIO(markdown_text.encode("utf-8")), "text/markdown")},
    )

    if response.status_code == 200:
        out_file = OUTPUT_DIR / f"{filename_stem}_presentation.pptx"
        out_file.write_bytes(response.content)

        prs = PPTXPresentation(io.BytesIO(response.content))
        print(f"SUCCESS! Created PowerPoint presentation at:")
        print(f"  -> {out_file}")
        print(f"  Total Slides Generated: {len(prs.slides)}")
        print("\n  Slide Outline:")
        for idx, slide in enumerate(prs.slides, 1):
            title = slide.shapes.title.text if slide.shapes.title else "No Title"
            print(f"    Slide {idx:2d}: '{title}' ({len(slide.shapes)} shapes)")
    else:
        print(f"FAILED! Status Code: {response.status_code}")
        print(f"  Response: {response.text}")


if __name__ == "__main__":
    print("Starting VSCode Direct String Input Testing...\n")

    # Run tests on all 4 sample Markdown strings
    test_markdown_string("Sample 1: Standard Elements", SAMPLE_1_STANDARD, "sample1_standard")
    test_markdown_string("Sample 2: Research Integrity Material", SAMPLE_2_RESEARCH_INTEGRITY, "sample2_research_integrity")
    test_markdown_string("Sample 3: LaTeX Math & Stress-Test", SAMPLE_3_LATEX_MATH, "sample3_latex_math")
    test_markdown_string("Sample 4: Comprehensive Edge Cases", SAMPLE_4_EDGE_CASES, "sample4_edge_cases")

    print("\nAll VSCode String Input Tests Complete! Check the 'vscode_outputs' directory for generated PPTX files.")
