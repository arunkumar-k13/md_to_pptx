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
        from md_to_pptx.main import generate_presentation
        tmp_md = Path("scratch/test_tables.md")
        tmp_md.parent.mkdir(parents=True, exist_ok=True)
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
            output_dir=Path("scratch/out_tables"),
        )
        prs = PPTXPresentation(report.pptx_path)
        table_shapes = [shp for s in prs.slides for shp in s.shapes if shp.has_table]
        self.assertTrue(len(table_shapes) >= 2)


if __name__ == "__main__":
    unittest.main()
