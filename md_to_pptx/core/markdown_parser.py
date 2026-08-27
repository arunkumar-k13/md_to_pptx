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


def restore_markdown_structure(raw_text: str) -> str:
    """Pre-process and restore line break structure for Markdown strings.

    Ensures multiline formatting, headings, bullet lists, code blocks, and tables
    are properly separated by line breaks even if flattened by HTML forms or APIs,
    while protecting LaTeX math blocks from line splitting.
    """
    if not raw_text or not raw_text.strip():
        return raw_text

    text = raw_text.replace("\r\n", "\n").replace("\r", "\n")

    # Unescape literal '\n' and '\t' string sequences if present
    if "\\n" in text and "\n" not in text:
        text = text.replace("\\n", "\n").replace("\\t", "\t")

    # Shield math blocks before applying newline/bullet/heading restoration
    shielded_text, math_placeholders = shield_math_blocks(text)

    # Restore newlines before headings (#, ##, ###)
    shielded_text = re.sub(r"(?<!\n)(\s*)(#{1,6}\s+)", r"\n\n\2", shielded_text)

    lines = shielded_text.split("\n")
    processed: List[str] = []
    for l in lines:
        if "|" in l or l.strip().startswith("#"):
            # Preserve headings (#) and table lines (|) intact without splitting numbers or symbols inside them
            processed.append(l)
        else:
            # Match bullet symbols ONLY if not part of ** bold or heading or table
            l_mod = re.sub(r"(?<![\*\w\#])(\s*)((?<!\*)\*(?!\*)\s+)", r"\n\2", l)
            l_mod = re.sub(r"(?<![\*\w\#])(\s*)(-(?=[^:\-\|])\s+)", r"\n\2", l_mod)
            l_mod = re.sub(r"(?<![\*\w\#])(\s*)(\+\s+)", r"\n\2", l_mod)
            # Restore line breaks before numbered items (1., 2., 11., 100.)
            l_mod = re.sub(r"([^\n])\s+(\d{1,3}\.\s+)", r"\1\n\2", l_mod)
            l_mod = re.sub(r"(?<!\n)(\s*)(```[\w]*\s*)", r"\n\n\2", l_mod)
            l_mod = re.sub(r"(?<!\n)(\s*)(>\s+)", r"\n\n\2", l_mod)
            processed.append(l_mod)

    text = "\n".join(processed)

    # Normalize multiple blank lines to a maximum of 2 blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Unshield math blocks back to original intact form
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
    """Convert LaTeX math notation and inline math delimiters into clean Unicode text.

    Examples:
        '$\\frac{1}{2}$' -> '(1/2)'
        '$\\frac{a}{\\frac{b}{c}}$' -> '(a/(b/c))'
        '$\\sqrt[3]{y}$' -> '∛(y)'
        '$x + y = z$' -> 'x + y = z'
        '~~strikethrough~~' -> 'strikethrough'
    """
    if not text:
        return text

    result = text

    # Remove display math $$ delimiters
    result = re.sub(r"\$\$\s*", "", result)
    result = re.sub(r"\s*\$\$", "", result)

    # Convert LaTeX matrices \begin{pmatrix} a & b \\ c & d \end{pmatrix} -> [a  b | c  d]
    def _clean_matrix(m):
        content = m.group(1).replace("\\\\", " | ").replace("&", "  ").replace("\n", " ")
        content = re.sub(r"\s+", " ", content).strip()
        return f"[{content}]"

    result = re.sub(r"\\begin\{[a-zA-Z]*matrix\}(.*?)\\end\{[a-zA-Z]*matrix\}", _clean_matrix, result, flags=re.DOTALL)

    # Convert LaTeX cases \begin{cases} x & x > 0 \\ -x & x \leq 0 \end{cases}
    def _clean_cases(m):
        content = m.group(1).replace("\\\\", " ; ").replace("&", " if ").replace("\n", " ")
        content = re.sub(r"\s+", " ", content).strip()
        return f"{{ {content} }}"

    result = re.sub(r"\\begin\{cases\}(.*?)\\end\{cases\}", _clean_cases, result, flags=re.DOTALL)

    # Convert LaTeX aligned equations \begin{aligned} a &= b \\ c &= d \end{aligned}
    def _clean_aligned(m):
        content = m.group(1).replace("\\\\", " | ").replace("&", "").replace("\n", " ")
        content = re.sub(r"\s+", " ", content).strip()
        return f"[{content}]"

    result = re.sub(r"\\begin\{aligned\}(.*?)\\end\{aligned\}", _clean_aligned, result, flags=re.DOTALL)

    # Convert recursive fractions \frac{num}{den}
    while "\\frac{" in result:
        idx = result.find("\\frac{")
        num, next_pos = parse_balanced_braces(result, idx + 5)
        if next_pos < len(result) and result[next_pos] == '{':
            den, end_pos = parse_balanced_braces(result, next_pos)
            clean_num = clean_latex_and_inline_math(num)
            clean_den = clean_latex_and_inline_math(den)
            result = result[:idx] + f"({clean_num}/{clean_den})" + result[end_pos:]
        else:
            break

    # Convert n-th roots \sqrt[n]{x} and \sqrt{x}
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

    # Set notation and mathbb
    result = result.replace("\\mathbb{R}", "ℝ").replace("\\mathbb{N}", "ℕ").replace("\\mathbb{Z}", "ℤ").replace("\\mathbb{Q}", "ℚ").replace("\\mathbb{C}", "ℂ")
    result = result.replace("\\{", "{").replace("\\}", "}").replace("\\mid", "|")

    # Delimiters and floor/ceiling
    result = re.sub(r"\\lfloor\s*(.*?)\s*\\rfloor", r"⌊\1⌋", result)
    result = re.sub(r"\\ceil\s*(.*?)\s*\\rceil", r"⌈\1⌉", result)
    result = re.sub(r"\\left[\(\[\{\|]?", "", result)
    result = re.sub(r"\\right[\)\]\}\|]?", "", result)

    # Strip \text{...}, \boxed{...}, \color{...}{...}, \overbrace{...}^{...}, \underbrace{...}_{...}
    result = re.sub(r"\\text\{([^}]+)\}", r"\1", result)
    result = re.sub(r"\\boxed\{([^}]+)\}", r"[\1]", result)
    result = re.sub(r"\\color\{[^}]+\}\{([^}]+)\}", r"\1", result)
    result = re.sub(r"\\overbrace\{([^}]*)\}\^\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\underbrace\{([^}]*)\}_\{([^}]*)\}", r"\1", result)

    # Accent functions
    result = re.sub(r"\\(hat|check|tilde|dot|ddot)\{([^}]+)\}", r"\2", result)

    # LaTeX spacing control sequences
    result = result.replace("\\quad", " ").replace("\\,", " ").replace("\\!", "").replace("\\:", " ").replace("\\;", " ")

    # Font styles & functions
    result = re.sub(r"\\(mathcal|normalfont|mathrm|mathit|mathbf|mathsf|mathtt|mathfrak)\{([^}]+)\}", r"\2", result)
    result = re.sub(r"\\binom\{([^}]+)\}\{([^}]+)\}", r"(\1 choose \2)", result)

    replacements = [
        (r"\\lim_\{([^}]+)\}", r"lim(\1)"),
        (r"\\lim\b", "lim"),
        (r"\\iint_\{([^}]+)\}", r"∬(\1)"),
        (r"\\iint\b", "∬"),
        (r"\\iiint_\{([^}]+)\}", r"∭(\1)"),
        (r"\\iiint\b", "∭"),
        (r"\\oint_\{([^}]+)\}", r"∮(\1)"),
        (r"\\oint\b", "∮"),
        (r"\\prod_\{([^}]+)\}\^\{([^}]+)\}", r"∏(\1 to \2)"),
        (r"\\prod\b", "∏"),
        (r"\\sum_\{([^}]+)\}\^\{([^}]+)\}", r"∑(\1 to \2)"),
        (r"\\sum\b", "∑"),
        (r"\\int_\{([^}]+)\}\^\{([^}]+)\}", r"∫(\1 to \2)"),
        (r"\\int\b", "∫"),
        (r"\\partial\b", "∂"),
        (r"\\infty\b", "∞"),
        (r"\\alpha\b", "α"),
        (r"\\beta\b", "β"),
        (r"\\gamma\b", "γ"),
        (r"\\delta\b", "δ"),
        (r"\\epsilon\b", "ε"),
        (r"\\pi\b", "π"),
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
        (r"\\Theta\b", "Θ"),
        (r"\\Lambda\b", "Λ"),
        (r"\\Omega\b", "Ω"),
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
    ]

    for pattern, repl in replacements:
        result = re.sub(pattern, repl, result)

    # Strip dollar wrappers $formula$ -> formula
    result = re.sub(r"\$([^\$]+)\$", r"\1", result)

    # Clean strikethrough syntax ~~text~~ -> text
    result = re.sub(r"~~([^~]+)~~", r"\1", result)
    result = result.replace("~~", "")

    # Ensure space after bold/italic inline colons (e.g. Primary Focus:Comprehensive -> Primary Focus: Comprehensive)
    result = re.sub(r"([a-zA-Z0-9\):]):([a-zA-Z0-9])", r"\1: \2", result)

    return result.strip()


class MarkdownParser:
    """Parser converting Markdown text or files into a structured Document AST."""

    def __init__(self) -> None:
        """Initialize the markdown-it-py parser engine with standard extensions."""
        self._md = MarkdownIt("commonmark").enable("table")

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

                # Table Fallback Recovery: Parse table pipes in paragraph text as structured TableNode
                if "|" in text and text.count("|") >= 4 and ("---" in text or text.count("\n") >= 1 or "| |" in text):
                    tbl = self._parse_fallback_table(text)
                    if tbl:
                        return tbl

                clean_txt = clean_latex_and_inline_math(text)
                return ParagraphNode(text=clean_txt, runs=runs)

            elif ntype in ("bullet_list", "ordered_list"):
                is_ordered = (ntype == "ordered_list")
                items = self._extract_list_items(node, level=0)
                return BulletListNode(items=items, is_ordered=is_ordered)

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
                return QuoteNode(children=child_nodes, text=aggregated_text)

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

        def _traverse(n: SyntaxTreeNode, is_bold: bool, is_italic: bool, is_code: bool, url: Optional[str]) -> None:
            ntype = n.type
            if ntype == "text":
                text = getattr(n, "content", "")
                if text:
                    full_text_parts.append(text)
                    runs.append(InlineRun(text=text, is_bold=is_bold, is_italic=is_italic, is_code=is_code, url=url))
            elif ntype == "code_inline":
                text = getattr(n, "content", "")
                if text:
                    full_text_parts.append(text)
                    runs.append(InlineRun(text=text, is_bold=is_bold, is_italic=is_italic, is_code=True, url=url))
            elif ntype == "image":
                attrs = getattr(n, "attrs", {})
                alt = getattr(n, "content", "")
                src = attrs.get("src", "")
                full_text_parts.append(alt)
                runs.append(InlineRun(text=f"[Image: {alt or src}]", url=src))
            elif ntype == "strong":
                for child in n.children:
                    _traverse(child, is_bold=True, is_italic=is_italic, is_code=is_code, url=url)
            elif ntype == "em":
                for child in n.children:
                    _traverse(child, is_bold=is_bold, is_italic=True, is_code=is_code, url=url)
            elif ntype == "link":
                attrs = getattr(n, "attrs", {})
                link_url = attrs.get("href", url)
                for child in n.children:
                    _traverse(child, is_bold=is_bold, is_italic=is_italic, is_code=is_code, url=link_url)
            elif ntype == "inline":
                for child in n.children:
                    _traverse(child, is_bold=is_bold, is_italic=is_italic, is_code=is_code, url=url)
            else:
                for child in getattr(n, "children", []):
                    _traverse(child, is_bold=is_bold, is_italic=is_italic, is_code=is_code, url=url)

        for child in node.children:
            _traverse(child, is_bold=False, is_italic=False, is_code=False, url=None)

        full_text = "".join(full_text_parts).strip()
        cleaned_text = clean_latex_and_inline_math(full_text)
        cleaned_runs = [
            InlineRun(
                text=clean_latex_and_inline_math(r.text),
                is_bold=r.is_bold,
                is_italic=r.is_italic,
                is_code=r.is_code,
                url=r.url,
            )
            for r in runs
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

        for item_node in list_node.children:
            if item_node.type != "list_item":
                continue

            item_text = ""
            item_runs: List[InlineRun] = []
            nested_nodes: List[ASTNode] = []

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
                    nested_items = self._extract_list_items(child, level=level + 1)
                    nested_nodes.append(
                        BulletListNode(items=nested_items, is_ordered=(child.type == "ordered_list"))
                    )
                else:
                    conv = self._convert_node(child)
                    if conv:
                        if isinstance(conv, list):
                            nested_nodes.extend(conv)
                        else:
                            nested_nodes.append(conv)

            items.append(ListItem(text=item_text, runs=item_runs, children=nested_nodes, level=level))

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
                            text, _ = self._extract_inline_content(cell)
                            cell_strings.append(text)
                            cell_objects.append(TableCell(text=text, is_header=is_header_row))

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

    def _build_sections(self, nodes: List[ASTNode]) -> List[Section]:
        """Group flat document nodes into logical sections bounded by major headers.

        Args:
            nodes: Top-level AST nodes.

        Returns:
            List of Section instances.
        """
        has_h1_or_h2 = any(isinstance(n, HeaderNode) and n.level in (1, 2) for n in nodes)
        split_levels = (1, 2) if has_h1_or_h2 else (1, 2, 3)

        sections: List[Section] = []
        current_section = Section(header=None, nodes=[])

        for node in nodes:
            if isinstance(node, HeaderNode) and node.level in split_levels:
                if current_section.header or current_section.nodes:
                    sections.append(current_section)
                current_section = Section(header=node, nodes=[node])
            else:
                current_section.nodes.append(node)

        if current_section.header or current_section.nodes:
            sections.append(current_section)

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
