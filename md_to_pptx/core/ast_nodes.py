"""Markdown Abstract Syntax Tree (AST) Domain Model.

This module defines pure, presentation-agnostic dataclasses representing the
parsed Markdown document structure. It adheres strictly to clean architecture
by decoupling raw document representation from presentation planning or rendering engines.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass(slots=True)
class InlineRun:
    """Represents a formatted inline text run within a paragraph or list item.

    Attributes:
        text: Raw text string.
        is_bold: True if text has bold formatting.
        is_italic: True if text has italic formatting.
        is_code: True if text is inline code.
        url: Target URL if this text run forms a hyperlink, else None.
    """

    text: str
    is_bold: bool = False
    is_italic: bool = False
    is_code: bool = False
    is_strikethrough: bool = False
    url: Optional[str] = None


@dataclass(slots=True)
class ASTNode:
    """Base abstract node for all Markdown Abstract Syntax Tree elements."""

    pass


@dataclass(slots=True)
class HeaderNode(ASTNode):
    """Represents a Markdown heading (H1 - H6).

    Attributes:
        level: Heading level (1 through 6).
        text: Heading text content.
        runs: Formatted inline runs within the heading.
    """

    level: int
    text: str
    runs: List[InlineRun] = field(default_factory=list)


@dataclass(slots=True)
class ParagraphNode(ASTNode):
    """Represents a standard text paragraph.

    Attributes:
        text: Full plain text content of the paragraph.
        runs: List of formatted inline text runs.
    """

    text: str
    runs: List[InlineRun] = field(default_factory=list)


@dataclass(slots=True)
class ListItem(ASTNode):
    """Represents an individual item in a list, supporting nested sub-nodes.

    Attributes:
        text: Plain text content of the list item.
        runs: Inline text runs.
        children: Nested AST nodes (e.g. sub-lists or nested paragraphs).
        level: Indentation/nesting depth (0-indexed).
    """

    text: str
    runs: List[InlineRun] = field(default_factory=list)
    children: List[ASTNode] = field(default_factory=list)
    level: int = 0
    is_ordered: bool = False


@dataclass(slots=True)
class BulletListNode(ASTNode):
    """Represents a bullet or numbered list.

    Attributes:
        items: List of item elements.
        is_ordered: True if ordered (numbered) list, False for bullet list.
    """

    items: List[ListItem] = field(default_factory=list)
    is_ordered: bool = False
    start_index: int = 1


@dataclass(slots=True)
class TableCell:
    """Represents a single cell within a table.

    Attributes:
        text: Plain text content.
        is_header: True if cell is part of header row.
        runs: Formatted inline runs.
    """

    text: str
    is_header: bool = False
    runs: List[InlineRun] = field(default_factory=list)


@dataclass(slots=True)
class TableRow:
    """Represents a horizontal row in a table.

    Attributes:
        cells: List of cells in this row.
        is_header: True if this row is the table header row.
    """

    cells: List[TableCell] = field(default_factory=list)
    is_header: bool = False


@dataclass(slots=True)
class TableNode(ASTNode):
    """Represents a tabular data block.

    Attributes:
        headers: List of column header title strings.
        rows: List of row data (each row is a list of cell strings).
        raw_rows: Structured TableRow instances.
    """

    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    raw_rows: List[TableRow] = field(default_factory=list)


@dataclass(slots=True)
class ImageNode(ASTNode):
    """Represents an embedded image block or inline image.

    Attributes:
        src: Image file path or URL.
        alt: Alternate description text.
        title: Optional title attribute.
    """

    src: str
    alt: str = ""
    title: str = ""


@dataclass(slots=True)
class CodeBlockNode(ASTNode):
    """Represents a fenced code block.

    Attributes:
        code: Source code text content.
        language: Programming or markup language identifier (e.g., 'python', 'json').
    """

    code: str
    language: str = ""


@dataclass(slots=True)
class QuoteNode(ASTNode):
    """Represents a blockquote container.

    Attributes:
        children: Nested nodes contained inside the quote.
        text: Plain text content aggregated from nested nodes.
    """

    children: List[ASTNode] = field(default_factory=list)
    text: str = ""
    runs: List[InlineRun] = field(default_factory=list)


@dataclass(slots=True)
class HorizontalRuleNode(ASTNode):
    """Represents a horizontal divider rule ('---', '***')."""

    pass


@dataclass(slots=True)
class Section(ASTNode):
    """Represents a logical section of the document bounded by a major heading.

    Attributes:
        header: Optional leading header for this section.
        nodes: Child AST nodes belonging to this section.
    """

    header: Optional[HeaderNode] = None
    nodes: List[ASTNode] = field(default_factory=list)


@dataclass(slots=True)
class Document(ASTNode):
    """Root container node representing a full parsed Markdown document.

    Attributes:
        title: Document main title (derived from first H1 or document header).
        nodes: Linear sequence of top-level AST nodes.
        sections: Logical sections derived from headers.
    """

    title: str = ""
    nodes: List[ASTNode] = field(default_factory=list)
    sections: List[Section] = field(default_factory=list)
