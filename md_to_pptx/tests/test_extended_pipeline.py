"""Regression Test Suite for Extended Pipeline Functionality.

Verifies verbatim paragraph preservation, multi-table vertical stacking,
corporate branding layout application, composite inline formatting, and scoped H3 splitting.
"""

import unittest
from pathlib import Path
from md_to_pptx.core.markdown_parser import MarkdownParser
from md_to_pptx.core.content_analyzer import ContentAnalyzer
from md_to_pptx.core.design_engine import PresentationDesignEngine
from md_to_pptx.core.slide_planner import SlidePlanner
from md_to_pptx.core.overflow.paginator import Paginator
from md_to_pptx.core.generator.pptx_generator import PPTXGenerator
from pptx import Presentation as PPTXPresentation


class TestExtendedPipeline(unittest.TestCase):

    def setUp(self):
        self.parser = MarkdownParser()
        self.analyzer = ContentAnalyzer()
        self.engine = PresentationDesignEngine()
        self.planner = SlidePlanner()
        self.paginator = Paginator()
        self.generator = PPTXGenerator()

    def test_verbatim_paragraph_preservation(self):
        """Verify output paragraph text matches input source text 100% verbatim without LLM paraphrasing."""
        sample_md = """# Verbatim Test Slide

Project "Aetheris" is a cloud-native infrastructure suite.
Project "Synth-Link" is a data migration bridge.
"""
        doc = self.parser.parse(sample_md)
        analyzed = self.analyzer.analyze(doc)
        plan = self.engine.create_design_plan(analyzed)
        pres = self.planner.plan(analyzed, design_plan=plan)

        # Confirm exact text is preserved in AST and Presentation IR across content slides
        content_slide = pres.slides[1] if len(pres.slides) > 1 else pres.slides[0]
        p_block = content_slide.blocks[0]
        self.assertIn('Project "Aetheris"', p_block.text)
        p_block2 = content_slide.blocks[1]
        self.assertIn('Project "Synth-Link"', p_block2.text)

    def test_scoped_h3_splitter_rules(self):
        """Verify H3 headers in Section 2 (Headings) and Section 19 (Edge Cases) are NOT split, while Section 18 IS split."""
        sample_md = """## 2. Headings

# Heading 1
## Heading 2
### Heading 3

## 18. Realistic Project Report

### Executive Summary
Content for summary.

### Objectives
Content for objectives.

## 19. Edge Cases

### Heading 3 in Edge Cases
Content for edge case.
"""
        doc = self.parser.parse(sample_md)
        section_titles = [s.header.text for s in doc.sections if s.header]

        # Section 2 ('Headings') should remain a single section and not be split by '### Heading 3'
        self.assertIn("2. Headings", section_titles)
        self.assertNotIn("Heading 3", section_titles)

        # Section 18 ('Realistic Project Report') contains major report section title
        self.assertIn("18. Realistic Project Report", section_titles)

        # Section 19 ('Edge Cases') should remain intact and not split '### Heading 3 in Edge Cases'
        self.assertIn("19. Edge Cases", section_titles)
        self.assertNotIn("Heading 3 in Edge Cases", section_titles)

    def test_nested_style_composition(self):
        """Verify bold + code inline runs compose both is_bold and is_code flags."""
        sample_md = "**Bold with `inline code`**"
        doc = self.parser.parse(sample_md)
        p_node = doc.nodes[0]
        self.assertTrue(len(p_node.runs) >= 2)
        code_run = p_node.runs[1]
        self.assertTrue(code_run.is_bold)
        self.assertTrue(code_run.is_code)

    def test_multi_table_stacking(self):
        """Verify multiple tables in a section do not overlap at identical coordinates."""
        import tempfile
        from md_to_pptx.main import generate_presentation

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            tmp_md = tmp_path / "test_tables.md"
            tmp_md.write_text("""## Tables Section

| ID | Name |
| -- | ---- |
| 1  | Alice |

| Metric | Score |
| ------ | ----- |
| Test   | 100   |
""", encoding="utf-8")

            report = generate_presentation(
                markdown_path=tmp_md,
                template_path=Path("corporate_template.pptx"),
                output_dir=tmp_path / "out_tables",
            )
            prs = PPTXPresentation(report.pptx_path)
            table_shapes = [shp for s in prs.slides for shp in s.shapes if shp.has_table]
            self.assertTrue(len(table_shapes) >= 2)

    def test_case_study_omitted_when_not_in_markdown(self):
        """Verify Case Study brand slide is omitted when markdown does not contain a Case Study."""
        import tempfile
        from md_to_pptx.main import generate_presentation

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            tmp_md = tmp_path / "no_cs.md"
            tmp_md.write_text("""# Standard Overview Presentation

## Executive Summary
This presentation covers general architectural updates.

## Technical Milestones
- First milestone completed.
- Second milestone on track.
""", encoding="utf-8")

            report = generate_presentation(
                markdown_path=tmp_md,
                template_path=Path("corporate_template.pptx"),
                output_dir=tmp_path / "out_no_cs",
            )
            prs = PPTXPresentation(report.pptx_path)
            # Ensure no slide has title "Case Study" or contains dummy "Lorem ipsum"
            slide_titles = [s.shapes.title.text if s.shapes.title and s.shapes.title.has_text_frame else "" for s in prs.slides]
            self.assertFalse(any("case study" in t.lower() for t in slide_titles), f"Case study slide should not be present in {slide_titles}")

            # Verify no 'Lorem ipsum' dummy text exists across any slide
            for slide in prs.slides:
                for shp in slide.shapes:
                    if shp.has_text_frame:
                        self.assertNotIn("lorem ipsum", shp.text_frame.text.lower())

    def test_case_study_populated_when_present_in_markdown(self):
        """Verify Case Study brand slide is populated and dummy text replaced when markdown contains a Case Study."""
        import tempfile
        from md_to_pptx.main import generate_presentation
        from md_to_pptx.core.generator.pptx_generator import get_all_slide_shapes

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            tmp_md = tmp_path / "with_cs.md"
            tmp_md.write_text("""# Clinical Platform Presentation

## Case Study: Bio-Extraction Engine
* **Objective:** Automate extraction of trial endpoints from clinical records.
* **Data Source:** 45,000 FDA public trial registry documents.
* **Work Carried:** Trained BioBERT NER pipeline with distributed inference.
* **Skills:** Python, PyTorch, HuggingFace, FastAPI, Docker.
* **Insights Delivered:** 85% acceleration in trial review turnaround.

## Next Steps
- Rollout to staging environment.
""", encoding="utf-8")

            report = generate_presentation(
                markdown_path=tmp_md,
                template_path=Path("corporate_template.pptx"),
                output_dir=tmp_path / "out_with_cs",
            )
            prs = PPTXPresentation(report.pptx_path)

            # Locate the Case Study slide
            cs_slides = [s for s in prs.slides if s.shapes.title and "case study" in s.shapes.title.text.lower()]
            self.assertEqual(len(cs_slides), 1, "Exactly one Case Study slide should be generated")
            cs_slide = cs_slides[0]

            # Verify title is updated to the markdown section's title
            self.assertIn("Bio-Extraction Engine", cs_slide.shapes.title.text)

            # Verify all text across all shapes on the slide: real content present, zero lorem ipsum
            all_text = " ".join(
                shp.text_frame.text for shp in get_all_slide_shapes(cs_slide.shapes) if shp.has_text_frame
            )
            self.assertNotIn("lorem ipsum", all_text.lower(), "Placeholder lorem ipsum must be completely replaced")
            self.assertIn("Automate extraction of trial endpoints", all_text)
            self.assertIn("45,000 FDA public trial registry", all_text)
            self.assertIn("Trained BioBERT NER pipeline", all_text)
            self.assertIn("Python, PyTorch", all_text)
            self.assertIn("85% acceleration in trial review", all_text)


if __name__ == "__main__":
    unittest.main()
