"""Markdown Parser Module.

Converts raw Markdown text or files into a presentation-agnostic Document AST
using markdown-it-py. Fully supports GFM headings, paragraphs, formatted text,
lists, tables, images, code blocks, quotes, and horizontal rules.
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import List, Optional, Union

from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode

from md_to_pptx.core.ast_nodes import (
    ASTNode,
    BulletListNode,
    CodeBlockNode,
    Document,
    HeaderNode,
    HorizontalRuleNode,
    ImageNode,
    InlineRun,
    ListItem,
    ParagraphNode,
    QuoteNode,
    Section,
    TableCell,
    TableNode,
    TableRow,
)

logger = logging.getLogger(__name__)


import re

def shield_math_blocks(raw_text: str) -> tuple[str, list]:
    """Shield LaTeX math expressions before structure restoration."""
    placeholders = []

    def replacer(match):
        idx = len(placeholders)
        token = f"___MATH_SHIELD_{idx}___"
        placeholders.append(match.group(0))
        return token

    # 1. Shield display math $$...$$, environments (with optional $ wrappers), and \begin{...}...\end{...}
    pattern_env = r"(\$\$.*?\$\$|\$?\\begin\{[a-zA-Z]*matrix\}.*?\\end\{[a-zA-Z]*matrix\}\$?|\$?\\begin\{cases\}.*?\\end\{cases\}\$?|\$?\\begin\{aligned\}.*?\\end\{aligned\}\$?)"
    text = re.sub(pattern_env, replacer, raw_text, flags=re.DOTALL)

    # 2. Shield inline math $...$ (avoiding isolated $ or currency $100)
    pattern_inline = r"(?<!\\)\$([^\$\n]+?)(?<!\\)\$"
    text = re.sub(pattern_inline, replacer, text)

    return text, placeholders


def unshield_math_blocks(text: str, placeholders: list) -> str:
    """Restore original LaTeX math expressions from placeholders in reverse order."""
    # Reverse iteration guarantees nested placeholders are fully resolved!
    for idx in range(len(placeholders) - 1, -1, -1):
        token = f"___MATH_SHIELD_{idx}___"
        original = placeholders[idx]
        text = text.replace(token, original)
    return text


def shield_code_and_tables(raw_text: str) -> tuple[str, list]:
    """Shield inline code, fenced code blocks, and table lines before structure restoration."""
    placeholders = []
    def replacer(match):
        idx = len(placeholders)
        token = f"___CODE_SHIELD_{idx}___"
        placeholders.append(match.group(0))
        return token

    # 1. Shield fenced code blocks ```...```
    text = re.sub(r"```[\s\S]*?```", replacer, raw_text)
    # 2. Shield inline code `...`
    text = re.sub(r"`[^`\n]+`", replacer, text)
    # 3. Shield table lines containing | ... |
    text = re.sub(r"^(\|[^\n]+\|)$", replacer, text, flags=re.MULTILINE)

    return text, placeholders


def unshield_code_and_tables(text: str, placeholders: list) -> str:
    """Restore shielded inline code, code blocks, and table lines in reverse order."""
    for idx in range(len(placeholders) - 1, -1, -1):
        token = f"___CODE_SHIELD_{idx}___"
        text = text.replace(token, placeholders[idx])
    return text


def restore_markdown_structure(raw_text: str) -> str:
    """Pre-process and restore line break structure for Markdown strings.

    Ensures multiline formatting, headings, bullet lists, code blocks, and tables
    are properly separated by line breaks even if flattened by HTML forms or APIs,
    while protecting LaTeX math blocks, inline code, and tables from line splitting.
    """
    if not raw_text or not raw_text.strip():
        return raw_text

    # Strip non-printable control characters and object-replacement characters (e.g. \ufffc)
    raw_text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffc]", "", raw_text)

    text = raw_text.replace("\r\n", "\n").replace("\r", "\n")

    # Unescape literal '\n' and '\t' string sequences if present
    if "\\n" in text and "\n" not in text:
        text = text.replace("\\n", "\n").replace("\\t", "\t")

    # Shield math blocks and code/tables before applying newline/bullet/heading restoration
    shielded_text, math_placeholders = shield_math_blocks(text)
    shielded_text, code_placeholders = shield_code_and_tables(shielded_text)

    # Restore newlines before headings (#, ##, ###) ONLY when # starts a line
    shielded_text = re.sub(r"^(\s*#{1,6}\s+)", r"\n\n\1", shielded_text, flags=re.MULTILINE)

    lines = shielded_text.split("\n")
    processed: List[str] = []
    for l in lines:
        if "|" in l or l.strip().startswith("#") or "___CODE_SHIELD_" in l:
            # Preserve headings (#), table lines (|), and shielded code intact
            processed.append(l)
        else:
            # Match bullet symbols ONLY if followed by alphanumeric text or quotes (not special character strings)
            l_mod = re.sub(r"(?<![\*\w\#])(\s*)((?<!\*)\*(?!\*)\s+(?=[a-zA-Z0-9\"'\(]))", r"\n\2", l)
            l_mod = re.sub(r"(?<![\*\w\#])(\s*)(-(?=[a-zA-Z0-9\"'\(\[\{])\s+)", r"\n\2", l_mod)
            l_mod = re.sub(r"(?<![\*\w\#])(\s*)(\+\s+(?=[a-zA-Z0-9\"'\(]))", r"\n\2", l_mod)
            # Restore line breaks before numbered items at start of line or after whitespace (not within multi-digit numbers)
            l_mod = re.sub(r"(?<![\d\w])(\b\d{1,4}\.\s+(?=[a-zA-Z0-9\"'\(]))", r"\n\1", l_mod)
            l_mod = re.sub(r"(?<!\n)(\s*)(```[\w]*\s*)", r"\n\n\2", l_mod)
            l_mod = re.sub(r"(?<!\n)(\s*)(>\s+)", r"\n\n\2", l_mod)
            processed.append(l_mod)

    text = "\n".join(processed)

    # Normalize multiple blank lines to a maximum of 2 blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Unshield placeholders back to original intact form
    text = unshield_code_and_tables(text, code_placeholders)
    return unshield_math_blocks(text, math_placeholders)


def parse_balanced_braces(text: str, start_pos: int) -> tuple[str, int]:
    """Extract content within balanced curly braces {...} starting at start_pos."""
    if start_pos >= len(text) or text[start_pos] != '{':
        return "", start_pos
    depth = 0
    content = []
    i = start_pos
    while i < len(text):
        char = text[i]
        if char == '{' and (i == 0 or text[i-1] != '\\'):
            depth += 1
            if depth > 1:
                content.append(char)
        elif char == '}' and (i == 0 or text[i-1] != '\\'):
            depth -= 1
            if depth == 0:
                return "".join(content), i + 1
            else:
                content.append(char)
        else:
            content.append(char)
        i += 1
    return "".join(content), i


def clean_latex_and_inline_math(text: str) -> str:
    """Convert LaTeX math notation and inline math delimiters into clean Unicode text."""
    has_leading_space = text.startswith(" ")
    has_trailing_space = text.endswith(" ")

    result = text

    # Strip object-replacement characters, control chars, and malformed bracket/URL artifacts
    result = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffc]", "", result)
    result = re.sub(r"\]\([^\)]*\)", "", result)
    result = re.sub(r"\)+[a-zA-Z0-9_]*\)", "", result)

    # 1. Remove display math $$ delimiters and dollar wrappers
    result = re.sub(r"\$\$\s*", "", result)
    result = re.sub(r"\s*\$\$", "", result)
    result = re.sub(r"\$([^\$]+)\$", r"\1", result)

    # 2. Convert LaTeX matrices \begin{pmatrix} a & b \\ c & d \end{pmatrix} -> [a  b | c  d]
    def _clean_matrix(m):
        content = m.group(1).replace("\\\\", " | ").replace("\\", " | ").replace("&", "  ").replace("\n", " ")
        content = re.sub(r"\s+", " ", content).strip()
        return f"[{content}]"

    result = re.sub(r"\\begin\{[a-zA-Z]*matrix\}(.*?)\\end\{[a-zA-Z]*matrix\}", _clean_matrix, result, flags=re.DOTALL)

    # 3. Convert LaTeX cases \begin{cases} x & x > 0 \\ -x & x \leq 0 \end{cases}
    def _clean_cases(m):
        content = m.group(1).replace("\\\\", " ; ").replace("\\ ", " ; ").replace("&", " if ").replace("\n", " ")
        content = re.sub(r"\s+", " ", content).strip()
        return f"{{ {content} }}"

    result = re.sub(r"\\begin\{cases\}(.*?)\\end\{cases\}", _clean_cases, result, flags=re.DOTALL)

    # 4. Convert LaTeX aligned equations \begin{aligned} a &= b \\ c &= d \end{aligned}
    def _clean_aligned(m):
        content = m.group(1).replace("\\\\", " | ").replace("\\", " | ").replace("&", "").replace("\n", " ")
        content = re.sub(r"\s+", " ", content).strip()
        return f"[{content}]"

    result = re.sub(r"\\begin\{aligned\}(.*?)\\end\{aligned\}", _clean_aligned, result, flags=re.DOTALL)

    # 5. Convert recursive fractions \frac{num}{den}
    while "\\frac{" in result:
        idx = result.find("\\frac{")
        num, next_pos = parse_balanced_braces(result, idx + 5)
        if next_pos < len(result) and result[next_pos] == '{':
            den, end_pos = parse_balanced_braces(result, next_pos)
            clean_num = clean_latex_and_inline_math(num)
            clean_den = clean_latex_and_inline_math(den)
            den_fmt = f"({clean_den})" if (" " in clean_den or any(c in clean_den for c in "+-*=")) else clean_den
            result = result[:idx] + f"({clean_num}/{den_fmt})" + result[end_pos:]
        else:
            break

    # 6. Convert n-th roots \sqrt[n]{x} and \sqrt{x}
    superscript_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', 'n': 'ⁿ'}
    def _clean_nth_root(m):
        n_str = m.group(1)
        body = clean_latex_and_inline_math(m.group(2))
        if n_str == '3':
            return f"∛({body})"
        elif n_str == '4':
            return f"∜({body})"
        else:
            sup_n = "".join(superscript_map.get(c, c) for c in n_str)
            return f"{sup_n}√({body})"

    result = re.sub(r"\\sqrt\[([^\]]+)\]\{([^}]+)\}", _clean_nth_root, result)
    result = re.sub(r"\\sqrt\{([^}]+)\}", r"√(\1)", result)

    # 7. Convert Floor / Ceiling delimiters
    result = re.sub(r"\\lfloor\s*(.*?)\s*\\rfloor", r"⌊\1⌋", result)
    result = re.sub(r"\\lceil\s*(.*?)\s*\\rceil", r"⌈\1⌉", result)
    result = re.sub(r"\\lfloor\b", "⌊", result)
    result = re.sub(r"\\rfloor\b", "⌋", result)
    result = re.sub(r"\\lceil\b", "⌈", result)
    result = re.sub(r"\\rceil\b", "⌉", result)
    result = re.sub(r"\\ceil\b", "⌈", result)

    # 8. Set notation and mathbb
    result = result.replace("\\mathbb{R}", "ℝ").replace("\\mathbb{N}", "ℕ").replace("\\mathbb{Z}", "ℤ").replace("\\mathbb{Q}", "ℚ").replace("\\mathbb{C}", "ℂ")
    result = result.replace("\\{", "{").replace("\\}", "}").replace("\\mid", "|")

    # 9. Clean LaTeX big operators (sum, int, prod, lim, iint, iiint, oint)
    def _clean_big_op(m):
        op_name = m.group(1)
        sub = m.group(2) or ""
        sup = m.group(3) or ""
        op_sym = {"sum": "∑", "int": "∫", "prod": "∏", "lim": "lim", "iint": "∬", "iiint": "∭", "oint": "∮"}.get(op_name, op_name)
        sub_clean = re.sub(r"^[\{_]+|[\}]+$", "", sub).strip()
        sup_clean = re.sub(r"^[\{\^]+|[\}]+$", "", sup).strip()
        if sub_clean and sup_clean:
            return f"{op_sym}({sub_clean} to {sup_clean})"
        elif sub_clean:
            return f"{op_sym}({sub_clean})"
        return op_sym

    result = re.sub(r"\\(sum|int|prod|lim|iint|iiint|oint)(_\{?[^}\s^]+\}?)?(\^\{?[^}\s]+\}?)?", _clean_big_op, result)

    replacements = [
        (r"\^\s*\\circ\b", "°"),
        (r"\^\\circ\b", "°"),
        (r"\\partial\b", "∂"),
        (r"\\infty\b", "∞"),
        (r"\\alpha\b", "α"),
        (r"\\beta\b", "β"),
        (r"\\gamma\b", "γ"),
        (r"\\delta\b", "δ"),
        (r"\\epsilon\b", "ε"),
        (r"\\pi\b", "π"),
        (r"\\psi\b", "ψ"),
        (r"\\Psi\b", "Ψ"),
        (r"\\phi\b", "φ"),
        (r"\\Phi\b", "Φ"),
        (r"\\theta\b", "θ"),
        (r"\\Theta\b", "Θ"),
        (r"\\sigma\b", "σ"),
        (r"\\Sigma\b", "Σ"),
        (r"\\mu\b", "μ"),
        (r"\\nu\b", "ν"),
        (r"\\lambda\b", "λ"),
        (r"\\Lambda\b", "Λ"),
        (r"\\chi\b", "χ"),
        (r"\\tau\b", "τ"),
        (r"\\rho\b", "ρ"),
        (r"\\omega\b", "ω"),
        (r"\\Omega\b", "Ω"),
        (r"\\pm\b", "±"),
        (r"\\neq\b", "≠"),
        (r"\\leq?\b", "≤"),
        (r"\\geq?\b", "≥"),
        (r"\\approx\b", "≈"),
        (r"\\equiv\b", "≡"),
        (r"\\cdot\b", "·"),
        (r"\\times\b", "×"),
        (r"\\div\b", "÷"),
        (r"\\to\b", "→"),
        (r"\\implies\b", "⇒"),
        (r"\\iff\b", "⇔"),
        (r"\\forall\b", "∀"),
        (r"\\exists\b", "∃"),
        (r"\\emptyset\b", "∅"),
        (r"\\nabla\b", "∇"),
        (r"\\Gamma\b", "Γ"),
        (r"\\Delta\b", "Δ"),
        (r"\\varphi\b", "φ"),
        (r"\\varepsilon\b", "ε"),
        (r"\\aleph_0\b", "ℵ₀"),
        (r"\\aleph\b", "ℵ"),
        (r"\\langle\b", "⟨"),
        (r"\\rangle\b", "⟩"),
        (r"\\cup\b", "∪"),
        (r"\\cap\b", "∩"),
        (r"\\subseteq\b", "⊆"),
        (r"\\supset\b", "⊃"),
        (r"\\notin\b", "∉"),
        (r"\\in\b", "∈"),
        (r"\\land\b", "∧"),
        (r"\\lor\b", "∨"),
        (r"\\neg\b", "¬"),
        (r"\\therefore\b", "∴"),
        (r"\\because\b", "∵"),
        (r"\\bigoplus\b", "⊕"),
        (r"\\bigotimes\b", "⊗"),
        (r"\\Longleftarrow\b", "⇐"),
        (r"\\Longrightarrow\b", "⇒"),
        (r"\\leftrightarrow\b", "↔"),
        (r"\\uparrow\b", "↑"),
        (r"\\downarrow\b", "↓"),
        (r"\\rightarrow\b", "→"),
        (r"\\clubsuit\b", "♣"),
        (r"\\diamondsuit\b", "♦"),
        (r"\\heartsuit\b", "♥"),
        (r"\\spadesuit\b", "♠"),
        # Functions & Trig
        (r"\\sin\b", "sin"),
        (r"\\cos\b", "cos"),
        (r"\\tan\b", "tan"),
        (r"\\arcsin\b", "arcsin"),
        (r"\\arccos\b", "arccos"),
        (r"\\arctan\b", "arctan"),
        (r"\\sinh\b", "sinh"),
        (r"\\cosh\b", "cosh"),
        (r"\\tanh\b", "tanh"),
        (r"\\log_\{?([0-9a-zA-Z]+)\}?", r"log_\1"),
        (r"\\log\b", "log"),
        (r"\\ln\b", "ln"),
        (r"\\exp\b", "exp"),
        (r"\\pmod\{([^}]+)\}", r"(mod \1)"),
        (r"\\pmod\b", "mod"),
        # Dots family
        (r"\\dots\b", "…"),
        (r"\\cdots\b", "⋯"),
        (r"\\vdots\b", "⋮"),
        (r"\\ddots\b", "⋱"),
        # Relations & Symbols
        (r"\\circ\b", "°"),
        (r"\\sim\b", "~"),
        (r"\\cong\b", "≅"),
        (r"\\propto\b", "∝"),
        (r"\\parallel\b", "∥"),
        (r"\\perp\b", "⊥"),
        (r"\\triangle\b", "△"),
        (r"\\angle\b", "∠"),
        (r"\\doteq\b", "≑"),
        (r"\\prec\b", "≺"),
        (r"\\succ\b", "≻"),
        (r"\\asymp\b", "≍"),
        (r"\\ll\b", "≪"),
        (r"\\gg\b", "≫"),
    ]

    for pattern, repl in replacements:
        result = re.sub(pattern, repl, result)

    # 10. Subscript and superscript conversion
    subscript_map = {'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉', 'i': 'ᵢ', 'j': 'ⱼ', 'k': 'ₖ', 'n': 'ₙ', 'x': 'ₓ', 'y': 'ᵧ', 'a': 'ₐ', 'e': 'ₑ', 'o': 'ₒ', '+': '₊', '-': '₋', '=': '₌'}
    def _clean_subscript(m):
        val = m.group(1) or m.group(2)
        return "".join(subscript_map.get(c, c) for c in val)
    result = re.sub(r"(?<!\\)_\{([^}]+)\}|(?<!\\)_([0-9ijknaexyzo\+\-\=]+)", _clean_subscript, result)

    superscript_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', 'n': 'ⁿ', 'i': 'ⁱ', '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾'}
    def _clean_superscript(m):
        val = m.group(1) or m.group(2)
        return "".join(superscript_map.get(c, c) for c in val)
    result = re.sub(r"(?<!\\)\^\{([^}]+)\}|(?<!\\)\^([0-9ni\+\-\=\(\)])", _clean_superscript, result)

    # 11. Cleanup text tags, left/right brackets, binoms, math fonts, braces and accents
    result = re.sub(r"\\left\s*([(\[\{.|])", r"\1", result)
    result = re.sub(r"\\right\s*([)\]\}.|])", r"\1", result)
    result = re.sub(r"\\(left|right|bigl|bigr|left\.|right\.)\b", "", result)
    result = re.sub(r"\\binom\{([^}]+)\}\{([^}]+)\}", r"C(\1, \2)", result)
    result = re.sub(r"\\overbrace\{([^}]+)\}\^\{([^}]+)\}", r"(\1)^(\2)", result)
    result = re.sub(r"\\underbrace\{([^}]+)\}_\{([^}]+)\}", r"(\1)_(\2)", result)
    result = re.sub(r"\\(overbrace|underbrace)\{([^}]+)\}", r"(\2)", result)
    result = re.sub(r"\\operatorname\{([^}]+)\}", r"\1", result)
    result = re.sub(r"\\(mathbb|mathcal|mathbf|mathrm|mathit|mathsf|mathtt|makebox|mathfrak)\{([^}]+)\}", r"\2", result)
    result = re.sub(r"\\(hat|check|tilde|dot|ddot|vec|bar|overline|underline)\{([^}]+)\}", r"\2", result)
    result = re.sub(r"\\text\{([^}]+)\}", r"\1", result)
    result = re.sub(r"\\boxed\{([^}]+)\}", r"[\1]", result)
    result = re.sub(r"\\color\{[^}]+\}\{([^}]+)\}", r"\1", result)
    result = result.replace("\\quad", " ").replace("\\,", " ").replace("\\!", "").replace("\\:", " ").replace("\\;", " ")
    result = result.replace("\\rightleftharpoons", "⇌").replace("\\hbar", "ℏ")
    # 12. Balance unclosed parenthesis in equations
    if "(" in result and result.count("(") > result.count(")"):
        if "/" in result and result.rfind("/") > result.rfind("("):
            idx = result.rfind("/")
            result = result[:idx].rstrip() + ")" + result[idx:]
        else:
            result = result + ")"

    res_clean = result.strip()
    if has_leading_space and not res_clean.startswith(" "):
        res_clean = " " + res_clean
    if has_trailing_space and not res_clean.endswith(" "):
        res_clean = res_clean + " "
    return res_clean


class MarkdownParser:
    """Parser converting Markdown text or files into a structured Document AST."""

    def __init__(self) -> None:
        """Initialize the markdown-it-py parser engine with standard extensions."""
        self._md = MarkdownIt("commonmark").enable("table").enable("strikethrough")

    def parse_file(self, file_path: Union[str, Path]) -> Document:
        """Parse a Markdown file from disk into a Document AST.

        Args:
            file_path: Path to the target Markdown file.

        Returns:
            Document AST root node.

        Raises:
            FileNotFoundError: If the specified file does not exist.
        """
        path = Path(file_path)
        if not path.is_file():
            raise FileNotFoundError(f"Markdown file not found: {file_path}")
        content = path.read_text(encoding="utf-8")
        return self.parse(content)

    def parse(self, markdown_text: str) -> Document:
        """Parse Markdown text content into a Document AST.

        Args:
            markdown_text: Raw Markdown text string.

        Returns:
            Document AST root node.
        """
        if not markdown_text or not markdown_text.strip():
            logger.info("Empty markdown string passed to MarkdownParser.")
            return Document(title="Untitled Presentation", nodes=[], sections=[])

        # Restore structure and line breaks before parsing
        structured_text = restore_markdown_structure(markdown_text)

        tokens = self._md.parse(structured_text)
        tree = SyntaxTreeNode(tokens)

        nodes: List[ASTNode] = []
        for child in tree.children:
            converted = self._convert_node(child)
            if converted:
                if isinstance(converted, list):
                    nodes.extend(converted)
                else:
                    nodes.append(converted)

        sections = self._build_sections(nodes)
        document_title = self._extract_document_title(nodes)

        return Document(title=document_title, nodes=nodes, sections=sections)

    def _convert_node(self, node: SyntaxTreeNode) -> Optional[Union[ASTNode, List[ASTNode]]]:
        """Convert a markdown-it syntax tree node into an ASTNode."""
        ntype = node.type

        try:
            if ntype == "heading":
                level = int(node.tag[1]) if len(node.tag) > 1 and node.tag[1].isdigit() else 1
                text, runs = self._extract_inline_content(node)
                clean_text = re.sub(r"^Slide\s*\d+\s*[:\.\-]\s*", "", text, flags=re.IGNORECASE).strip()
                clean_text = clean_latex_and_inline_math(clean_text or text)
                return HeaderNode(level=level, text=clean_text, runs=runs)

            elif ntype == "paragraph":
                image_node = self._check_standalone_image(node)
                if image_node:
                    return image_node

                text, runs = self._extract_inline_content(node)
                if not text and not runs:
                    return None

                # Ensure missing spaces/newlines before multi-digit numbers (e.g. "dx15. Definite" -> "dx\n15. Definite")
                text = re.sub(r"([^\n\s\d])(?<!\d)(\b\d{1,4}\.\s*)", r"\1\n\2", text)

                # Check if paragraph contains multiple concatenated numbered items (e.g. "11. Limits... 12. Derivatives...")
                num_matches = list(re.finditer(r"(?<!\d)(\b\d{1,4}\.\s+)", text))
                if len(num_matches) >= 2:
                    list_items: List[ListItem] = []
                    parts = re.split(r"(?<!\d)(\b\d{1,4}\.\s+)", text)
                    current_item_text = ""
                    for p_part in parts:
                        if not p_part:
                            continue
                        if re.match(r"^\d{1,3}\.\s+$", p_part):
                            if current_item_text.strip():
                                clean_item = clean_latex_and_inline_math(current_item_text.strip())
                                list_items.append(ListItem(text=clean_item, runs=[InlineRun(text=clean_item)], level=0))
                            current_item_text = p_part
                        else:
                            current_item_text += p_part
                    if current_item_text.strip():
                        clean_item = clean_latex_and_inline_math(current_item_text.strip())
                        list_items.append(ListItem(text=clean_item, runs=[InlineRun(text=clean_item)], level=0))
                    if list_items:
                        return BulletListNode(items=list_items, is_ordered=True)

                # Split multiline paragraphs into distinct ParagraphNodes
                if "\n" in text:
                    lines = [l.strip() for l in text.split("\n") if l.strip()]
                    if len(lines) > 1:
                        split_nodes: List[ParagraphNode] = []
                        for line_str in lines:
                            clean_line = clean_latex_and_inline_math(line_str)
                            line_runs = [r for r in runs if r.text.strip() and r.text.strip() in line_str]
                            if not line_runs:
                                line_runs = [InlineRun(text=clean_line)]
                            split_nodes.append(ParagraphNode(text=clean_line, runs=line_runs))
                        return split_nodes

                clean_txt = clean_latex_and_inline_math(text)
                return ParagraphNode(text=clean_txt, runs=runs)

            elif ntype in ("bullet_list", "ordered_list"):
                is_ordered = (ntype == "ordered_list")
                items = self._extract_list_items(node, level=0)
                start_num = 1
                if is_ordered and items:
                    m = re.match(r"^(\d+)[\.\)]\s+", items[0].text.strip())
                    if m:
                        start_num = int(m.group(1))
                return BulletListNode(items=items, is_ordered=is_ordered, start_index=start_num)

            elif ntype == "table":
                return self._extract_table(node)

            elif ntype in ("fence", "code_block"):
                code_content = node.content if hasattr(node, "content") else ""
                lang = getattr(node, "info", "").strip()
                return CodeBlockNode(code=code_content.rstrip("\n"), language=lang)

            elif ntype == "blockquote":
                child_nodes: List[ASTNode] = []
                for sub_child in node.children:
                    res = self._convert_node(sub_child)
                    if res:
                        if isinstance(res, list):
                            child_nodes.extend(res)
                        else:
                            child_nodes.append(res)
                aggregated_text = "\n".join(
                    n.text for n in child_nodes if hasattr(n, "text") and n.text
                )
                aggregated_runs = []
                for n in child_nodes:
                    if hasattr(n, "runs") and n.runs:
                        aggregated_runs.extend(n.runs)
                return QuoteNode(children=child_nodes, text=aggregated_text, runs=aggregated_runs)

            elif ntype == "hr":
                return HorizontalRuleNode()

            else:
                logger.warning("Unrecognized token type '%s' handled gracefully.", ntype)
                # Extract any inline text fallback if present
                text, runs = self._extract_inline_content(node)
                if text:
                    return ParagraphNode(text=text, runs=runs)
                return None

        except Exception as err:
            logger.error("Error converting node '%s': %s", ntype, err, exc_info=True)
            return None

    def _extract_inline_content(self, node: SyntaxTreeNode) -> tuple[str, List[InlineRun]]:
        """Extract plain text and styled InlineRuns from a parent block node.

        Args:
            node: Parent syntax tree node containing inline children.

        Returns:
            Tuple of (aggregated plain text, list of InlineRun objects).
        """
        runs: List[InlineRun] = []
        full_text_parts: List[str] = []

        def _traverse(
            n: SyntaxTreeNode,
            is_bold: bool,
            is_italic: bool,
            is_code: bool,
            is_strikethrough: bool,
            url: Optional[str],
        ) -> None:
            ntype = n.type
            if ntype in ("text", "code_inline"):
                text = getattr(n, "content", "")
                if text:
                    prev_text = full_text_parts[-1] if full_text_parts else ""
                    needs_leading_space = (
                        bool(full_text_parts)
                        and prev_text != "\n"
                        and not prev_text.endswith((" ", "\n", "\t", "(", "[", "{", "\"", "'"))
                        and not text.startswith((" ", "\n", "\t", ".", ",", ":", ";", "!", "?", ")", "]", "}", "\"", "'"))
                    )
                    clean_run_text = (" " + text) if needs_leading_space else text
                    full_text_parts.append(clean_run_text)
                    runs.append(
                        InlineRun(
                            text=clean_run_text,
                            is_bold=is_bold,
                            is_italic=is_italic,
                            is_code=(is_code or ntype == "code_inline"),
                            is_strikethrough=is_strikethrough,
                            url=url,
                        )
                    )
            elif ntype in ("s", "del", "strike"):
                for child in getattr(n, "children", []):
                    _traverse(
                        child,
                        is_bold=is_bold,
                        is_italic=is_italic,
                        is_code=is_code,
                        is_strikethrough=True,
                        url=url,
                    )
            elif ntype == "image":
                attrs = getattr(n, "attrs", {})
                alt = getattr(n, "content", "")
                src = attrs.get("src", "")
                full_text_parts.append(alt)
                runs.append(InlineRun(text=f"[Image: {alt or src}]", url=src))
            elif ntype in ("softbreak", "hardbreak"):
                full_text_parts.append("\n")
                runs.append(InlineRun(text="\n"))
            elif ntype == "strong":
                for child in getattr(n, "children", []):
                    _traverse(
                        child,
                        is_bold=True,
                        is_italic=is_italic,
                        is_code=is_code,
                        is_strikethrough=is_strikethrough,
                        url=url,
                    )
            elif ntype == "em":
                for child in getattr(n, "children", []):
                    _traverse(
                        child,
                        is_bold=is_bold,
                        is_italic=True,
                        is_code=is_code,
                        is_strikethrough=is_strikethrough,
                        url=url,
                    )
            elif ntype == "link":
                attrs = getattr(n, "attrs", {})
                link_url = attrs.get("href", url)
                for child in getattr(n, "children", []):
                    _traverse(
                        child,
                        is_bold=is_bold,
                        is_italic=is_italic,
                        is_code=is_code,
                        is_strikethrough=is_strikethrough,
                        url=link_url,
                    )
            elif ntype == "inline":
                for child in getattr(n, "children", []):
                    _traverse(
                        child,
                        is_bold=is_bold,
                        is_italic=is_italic,
                        is_code=is_code,
                        is_strikethrough=is_strikethrough,
                        url=url,
                    )
            else:
                for child in getattr(n, "children", []):
                    _traverse(
                        child,
                        is_bold=is_bold,
                        is_italic=is_italic,
                        is_code=is_code,
                        is_strikethrough=is_strikethrough,
                        url=url,
                    )

        for child in node.children:
            _traverse(child, is_bold=False, is_italic=False, is_code=False, is_strikethrough=False, url=None)

        full_text = "".join(full_text_parts).strip()
        cleaned_text = clean_latex_and_inline_math(full_text)
        cleaned_runs = [
            InlineRun(
                text=clean_latex_and_inline_math(r.text),
                is_bold=r.is_bold,
                is_italic=r.is_italic,
                is_code=r.is_code,
                is_strikethrough=r.is_strikethrough,
                url=r.url,
            )
            for r in runs
            if r.text != "\n"
        ]
        return cleaned_text, cleaned_runs

    def _check_standalone_image(self, paragraph_node: SyntaxTreeNode) -> Optional[ImageNode]:
        """Check if a paragraph node contains exclusively a single image element.

        Args:
            paragraph_node: Paragraph syntax tree node.

        Returns:
            ImageNode if paragraph is an image container, else None.
        """
        children = paragraph_node.children
        if len(children) == 1 and children[0].type == "inline":
            inline_children = children[0].children
            if len(inline_children) == 1 and inline_children[0].type == "image":
                img = inline_children[0]
                attrs = getattr(img, "attrs", {})
                alt = getattr(img, "content", "")
                src = attrs.get("src", "")
                title = attrs.get("title", "")
                return ImageNode(src=src, alt=alt, title=title)
        return None

    def _extract_list_items(self, list_node: SyntaxTreeNode, level: int) -> List[ListItem]:
        """Recursively extract ListItems from a bullet_list or ordered_list node.

        Args:
            list_node: SyntaxTreeNode representing a list.
            level: Current nesting depth level.

        Returns:
            List of ListItem objects.
        """
        items: List[ListItem] = []
        is_node_ordered = (list_node.type == "ordered_list")

        for item_node in list_node.children:
            if item_node.type != "list_item":
                continue

            item_text = ""
            item_runs: List[InlineRun] = []
            nested_items: List[ListItem] = []

            for child in item_node.children:
                if child.type == "paragraph":
                    t, r = self._extract_inline_content(child)
                    if not item_text:
                        item_text = t
                        item_runs = r
                    else:
                        item_text += f"\n{t}"
                        item_runs.extend(r)
                elif child.type in ("bullet_list", "ordered_list"):
                    nested_items.extend(self._extract_list_items(child, level=level + 1))

            item_level = level
            if item_text.strip() not in (".", "...", "…"):
                items.append(ListItem(text=item_text, runs=item_runs, level=item_level, is_ordered=is_node_ordered))
                items.extend(nested_items)

        return items

    def _extract_table(self, table_node: SyntaxTreeNode) -> TableNode:
        """Extract headers and rows from a Markdown table node.

        Args:
            table_node: SyntaxTreeNode representing a table.

        Returns:
            Constructed TableNode.
        """
        headers: List[str] = []
        rows: List[List[str]] = []
        raw_rows: List[TableRow] = []

        for child in table_node.children:
            if child.type in ("thead", "tbody"):
                for tr in child.children:
                    if tr.type != "tr":
                        continue

                    is_header_row = (child.type == "thead")
                    cell_strings: List[str] = []
                    cell_objects: List[TableCell] = []

                    for cell in tr.children:
                        if cell.type in ("th", "td"):
                            text, runs = self._extract_inline_content(cell)
                            cell_strings.append(text)
                            cell_objects.append(TableCell(text=text, is_header=is_header_row, runs=runs))

                    if is_header_row:
                        headers = cell_strings
                        raw_rows.append(TableRow(cells=cell_objects, is_header=True))
                    else:
                        rows.append(cell_strings)
                        raw_rows.append(TableRow(cells=cell_objects, is_header=False))

        return TableNode(headers=headers, rows=rows, raw_rows=raw_rows)

    def _parse_fallback_table(self, text: str) -> Optional[TableNode]:
        """Attempt to parse raw table pipe text into a structured TableNode.

        Args:
            text: Unparsed paragraph text containing table pipes '|'.

        Returns:
            Constructed TableNode or None if parsing fails.
        """
        lines = [line.strip() for line in text.split("\n") if line.strip() and "|" in line]
        if not lines:
            lines = [row.strip() for row in re.split(r"(?<=\|)\s*(?=\|)", text) if "|" in row]

        rows_data: List[List[str]] = []
        for l in lines:
            clean = l.strip("|").strip()
            if "---" in clean or re.match(r"^[\s:\-\|]+$", clean):
                continue
            cells = [clean_latex_and_inline_math(c.strip()) for c in clean.split("|")]
            if cells:
                rows_data.append(cells)

        if not rows_data:
            return None

        headers = rows_data[0]
        data_rows = rows_data[1:] if len(rows_data) > 1 else []
        raw_table_rows: List[TableRow] = []

        header_cells = [TableCell(text=h, is_header=True) for h in headers]
        raw_table_rows.append(TableRow(cells=header_cells, is_header=True))

        for r in data_rows:
            row_cells = [TableCell(text=c, is_header=False) for c in r]
            raw_table_rows.append(TableRow(cells=row_cells, is_header=False))

        return TableNode(headers=headers, rows=data_rows, raw_rows=raw_table_rows)

    def _is_roman_header(self, text: str) -> bool:
        t = text.strip()
        import re
        pattern = r"^\s*(\*\*|\#\#\s*)?(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+[A-Z]"
        return bool(re.match(pattern, t))

    def _is_top_level_section_header(self, node: ASTNode, is_first_section: bool = False, has_numbered_sections: bool = False) -> bool:
        """Check if node is a top-level section boundary (e.g. main title, numbered/Roman header, or H1/H2 header)."""
        import re
        if isinstance(node, HeaderNode):
            text = node.text.strip()
            # 1. Numbered or Roman Section Headers ALWAYS start a new major section
            if re.match(r"^\s*(\#\#?\s*)?\d{1,3}\.\s+[A-Z]", text) or self._is_roman_header(text):
                return True

            # 2. First H1 in the document is the main title section
            if is_first_section and node.level == 1:
                return True

            # 3. If document uses numbered sections (## 1., ## 2., etc.), non-numbered H1s inside sections are internal demo content
            if has_numbered_sections:
                return False

            # 4. Otherwise, for unnumbered documents, H1 and H2 headers create section boundaries
            if node.level in (1, 2):
                return True

        elif isinstance(node, ParagraphNode) and self._is_roman_header(node.text):
            return True

        elif isinstance(node, HorizontalRuleNode):
            return True

        return False

    def _is_demo_section_title(self, text: str) -> bool:
        t = text.lower().strip()
        demo_keywords = ["heading", "edge case", "format", "nest", "list", "image", "short", "final", "validation"]
        return any(kw in t for kw in demo_keywords)

    def _build_sections(self, nodes: List[ASTNode]) -> List[Section]:
        """Group flat document nodes into logical sections bounded by major top-level headers.

        Args:
            nodes: Top-level AST nodes.

        Returns:
            List of Section instances.
        """
        import re
        has_numbered = any(
            isinstance(n, HeaderNode) and re.match(r"^\s*(\#\#?\s*)?\d{1,3}\.\s+[A-Z]", n.text.strip())
            for n in nodes
        )

        sections: List[Section] = []
        current_section = Section(header=None, nodes=[])
        in_report_section = False

        for node in nodes:
            is_first = (len(sections) == 0 and current_section.header is None)

            if isinstance(node, BulletListNode):
                current_items: List[ListItem] = []
                for item in node.items:
                    is_slide_header = bool(re.match(r"^\s*Slide\s*\d+\s*[:\.\-]", item.text, flags=re.IGNORECASE))
                    if self._is_roman_header(item.text) or is_slide_header:
                        if current_items:
                            sub_node = BulletListNode(items=current_items, is_ordered=node.is_ordered)
                            current_section.nodes.append(sub_node)
                            current_items = []
                        if current_section.header or current_section.nodes:
                            sections.append(current_section)
                        clean_t = re.sub(r"^\s*Slide\s*\d+\s*[:\.\-]\s*", "", item.text, flags=re.IGNORECASE).strip()
                        h_node = HeaderNode(level=2, text=clean_t or item.text, runs=getattr(item, "runs", []))
                        current_section = Section(header=h_node, nodes=[])
                        in_report_section = False
                    else:
                        current_items.append(item)

                if current_items:
                    sub_node = BulletListNode(items=current_items, is_ordered=node.is_ordered)
                    current_section.nodes.append(sub_node)

            elif isinstance(node, HorizontalRuleNode):
                # Only split section on horizontal rule if current section already contains 3+ content nodes
                if len([n for n in current_section.nodes if not isinstance(n, HeaderNode)]) >= 3:
                    if current_section.header or current_section.nodes:
                        sections.append(current_section)
                    prev_h = current_section.header
                    current_section = Section(header=prev_h, nodes=[])

            elif self._is_top_level_section_header(node, is_first_section=is_first, has_numbered_sections=has_numbered):
                if current_section.header or current_section.nodes:
                    sections.append(current_section)
                h_node = node if isinstance(node, HeaderNode) else HeaderNode(level=2, text=getattr(node, "text", ""), runs=getattr(node, "runs", []))
                current_section = Section(header=h_node, nodes=[node])
                in_report_section = ("18." in h_node.text or "report" in h_node.text.lower())

            elif isinstance(node, HeaderNode) and node.level == 3:
                # Split H3 into dedicated slides for report sections (e.g. Section 18: Executive Summary, Objectives, etc.)
                if in_report_section:
                    if current_section.header or current_section.nodes:
                        sections.append(current_section)
                    current_section = Section(header=node, nodes=[node])
                else:
                    current_section.nodes.append(node)

            else:
                current_section.nodes.append(node)

        if current_section.header or current_section.nodes:
            sections.append(current_section)

        # Propagate range-based start index (e.g. (11-20) -> 11) to ordered list nodes
        for sec in sections:
            if sec.header and sec.header.text:
                m_range = re.search(r"\((\d{1,3})-\d{1,3}\)", sec.header.text)
                if m_range:
                    range_start = int(m_range.group(1))
                    for node in sec.nodes:
                        if isinstance(node, BulletListNode) and node.is_ordered and node.start_index == 1:
                            node.start_index = range_start

        return sections

    def _extract_document_title(self, nodes: List[ASTNode]) -> str:
        """Extract document main title from the first H1 header or top node.

        Args:
            nodes: Top-level AST nodes.

        Returns:
            Extracted title string or default 'Untitled Presentation'.
        """
        for node in nodes:
            if isinstance(node, HeaderNode) and node.level == 1:
                return node.text
        for node in nodes:
            if isinstance(node, HeaderNode):
                return node.text
        return "Untitled Presentation"
