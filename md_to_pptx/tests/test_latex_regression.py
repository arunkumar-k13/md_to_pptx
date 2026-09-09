"""Unit tests verifying LaTeX math regression cases and content fidelity."""

import unittest
from md_to_pptx.core.markdown_parser import (
    MarkdownParser,
    clean_latex_and_inline_math,
    restore_markdown_structure,
)


class TestLaTeXRegression(unittest.TestCase):
    """Test suite ensuring LaTeX math expressions are never corrupted across pipeline stages."""

    def test_inline_plus_inside_math_not_split(self) -> None:
        raw = "$x + y = z$"
        restored = restore_markdown_structure(raw)
        self.assertEqual(restored, raw)
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "x + y = z")

    def test_greater_than_inside_math_not_split(self) -> None:
        raw = "$x > 0$"
        restored = restore_markdown_structure(raw)
        self.assertEqual(restored, raw)
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "x > 0")

    def test_pipe_inside_math_not_split(self) -> None:
        raw = "$\\{x \\in \\mathbb{R} \\mid x > 0\\}$"
        restored = restore_markdown_structure(raw)
        self.assertEqual(restored, raw)
        cleaned = clean_latex_and_inline_math(raw)
        self.assertIn("ℝ", cleaned)
        self.assertIn("x > 0", cleaned)

    def test_nested_fraction_parsing(self) -> None:
        raw = "$\\frac{a}{\\frac{b}{c}}$"
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "(a/(b/c))")

    def test_deeply_nested_fraction_parsing(self) -> None:
        raw = "$\\frac{a}{\\frac{b}{\\frac{c}{d}}}$"
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "(a/(b/(c/d)))")

    def test_nested_braces_in_fraction(self) -> None:
        raw = "$\\frac{x+\\{a,b\\}}{y}$"
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "(x+{a,b}/y)")

    def test_matrix_row_separators(self) -> None:
        raw_p = "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}"
        cleaned_p = clean_latex_and_inline_math(raw_p)
        self.assertEqual(cleaned_p, "[a b | c d]")

        raw_B = "\\begin{Bmatrix} 1 & 0 \\\\ 0 & 1 \\end{Bmatrix}"
        cleaned_B = clean_latex_and_inline_math(raw_B)
        self.assertEqual(cleaned_B, "[1 0 | 0 1]")

    def test_piecewise_row_separators(self) -> None:
        raw = "\\begin{cases} x & x > 0 \\\\ -x & x \\leq 0 \\end{cases}"
        cleaned = clean_latex_and_inline_math(raw)
        self.assertIn("x if x > 0", cleaned)
        self.assertIn("; -x if x ≤ 0", cleaned)

    def test_chemistry_equation(self) -> None:
        raw = "$H_2 + O_2 \\to H_2O$"
        restored = restore_markdown_structure(raw)
        self.assertEqual(restored, raw)
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "H₂ + O₂ → H₂O")

    def test_cube_root(self) -> None:
        raw = "$\\sqrt[3]{y}$"
        cleaned = clean_latex_and_inline_math(raw)
        self.assertEqual(cleaned, "∛(y)")

    def test_spacing_commands(self) -> None:
        self.assertEqual(clean_latex_and_inline_math("$x \\quad y$"), "x   y")
        self.assertEqual(clean_latex_and_inline_math("$x \\, y$"), "x   y")
        self.assertEqual(clean_latex_and_inline_math("$x \\! y$"), "x  y")


if __name__ == "__main__":
    unittest.main()
