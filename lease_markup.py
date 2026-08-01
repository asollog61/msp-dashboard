"""Markdown-subset parser for lease clause text.

Turns the markup people type into the Lease Builder's clause editors into a
block/run tree the Word renderer can walk. Deliberately pure — no Streamlit, no
python-docx, no lease data — so it can be unit-tested on its own and reused by
the HTML preview.

Supported markup (LEASE_FORMAT_SPEC.md):

    **text**            bold
    _text_              italic
    - item              bullet
    1. item             numbered   (the app renumbers; source numbers ignored)
    A. item             lettered sub-clause, level 1
    (i)                 level-2 sub-clause — inline, NOT its own block
    leading tab / 4sp   one nesting level deeper
    blank line          new paragraph
    [KP:Name]           cross-reference to a key provision (left unresolved)
    ^                   line break inside a paragraph
    \\*                  a literal asterisk (backslash escapes * _ [ ^ \\)

Structural markers, each its own block:

    [RentTable:Base]        place the base rent schedule here
    [RentTable:Option 1]    place an option schedule here
    [PageBreak]             hard page break
    [Exhibit:A|Floor Plan]  exhibit heading

    Base Rent Table:  …  Table End      legacy form of [RentTable:Base]

[KP:Name] comes out as an unresolved run carrying the provision name. The
renderer substitutes the value and builds the green hyperlink, so parsing needs
no lease data.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
import re
from typing import Any, Iterator

# ---------------------------------------------------------------------------
# Model
# ---------------------------------------------------------------------------

RUN_TEXT = "text"
RUN_KP_REF = "kp_ref"
RUN_BREAK = "break"

BLOCK_PARAGRAPH = "paragraph"
BLOCK_BULLET = "bullet"
BLOCK_NUMBERED = "numbered"
BLOCK_LETTERED = "lettered"
BLOCK_RENT_TABLE = "rent_table"
BLOCK_PAGE_BREAK = "page_break"
BLOCK_EXHIBIT = "exhibit"

LIST_KINDS = (BLOCK_BULLET, BLOCK_NUMBERED, BLOCK_LETTERED)
STRUCTURAL_KINDS = (BLOCK_RENT_TABLE, BLOCK_PAGE_BREAK, BLOCK_EXHIBIT)

MAX_LEVEL = 4  # deeper indents flatten; Word has no use for them here


@dataclass(frozen=True)
class Run:
    """One stretch of text with uniform formatting.

    kind == RUN_KP_REF: `name` holds the provision name and `text` is empty
    until the renderer substitutes it.
    kind == RUN_BREAK: a line break; carries no text.
    """

    text: str = ""
    bold: bool = False
    italic: bool = False
    kind: str = RUN_TEXT
    name: str = ""

    @property
    def is_ref(self) -> bool:
        return self.kind == RUN_KP_REF

    @property
    def is_break(self) -> bool:
        return self.kind == RUN_BREAK


@dataclass
class Block:
    """One paragraph-level thing: a paragraph, a list item, or a marker."""

    kind: str = BLOCK_PARAGRAPH
    level: int = 0
    runs: list[Run] = field(default_factory=list)
    # Filled in by assign_numbers(); the source's own numbering is discarded.
    number: str = ""
    # Structural blocks only: {"schedule": "Base"} / {"letter": "A", "title": …}
    data: dict[str, Any] = field(default_factory=dict)

    @property
    def is_list(self) -> bool:
        return self.kind in LIST_KINDS

    @property
    def is_structural(self) -> bool:
        return self.kind in STRUCTURAL_KINDS

    @property
    def text(self) -> str:
        """Plain text with refs shown as [KP:Name]; formatting dropped."""
        return "".join(
            "\n" if run.is_break else (f"[KP:{run.name}]" if run.is_ref else run.text)
            for run in self.runs
        )

    def refs(self) -> list[str]:
        return [run.name for run in self.runs if run.is_ref]


# ---------------------------------------------------------------------------
# Inline parsing
# ---------------------------------------------------------------------------

LINE_BREAK = "^"
# The list delimiters are escapable too, so a paragraph that merely starts
# "a) …" can be written back out without turning into a list item.
_ESCAPABLE = set("*_[]^\\.)-•")

# Bold before italic: ** must win over the _ rule and over a single *.
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.DOTALL)
# Underscores only count as italic at a word boundary, so snake_case_names and
# file_name.docx survive untouched.
_ITALIC_RE = re.compile(r"(?<![A-Za-z0-9_])_(?!\s)(.+?)(?<!\s)_(?![A-Za-z0-9_])", re.DOTALL)
_KP_RE = re.compile(r"\[\s*KPS?\s*:\s*([^\]]+?)\s*\]", re.IGNORECASE)

# Distinct open and close bytes matter: with the same byte at both ends, the
# closing sentinel of one placeholder plus the opening of the next can be read
# as a third placeholder, and restoring corrupts the text. (Found by fuzzing on
# input like r"[\1*1[".)
_PLACEHOLDER = "\x00{}\x01"
_SENTINELS = "\x00\x01"


def _protect_escapes(text: str) -> tuple[str, list[str]]:
    """Replace \\x escapes with sentinels so the markup passes over them."""
    stash: list[str] = []
    out: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char == "\\" and index + 1 < len(text) and text[index + 1] in _ESCAPABLE:
            stash.append(text[index + 1])
            out.append(_PLACEHOLDER.format(len(stash) - 1))
            index += 2
            continue
        out.append(char)
        index += 1
    return "".join(out), stash


def _restore_escapes(text: str, stash: list[str]) -> str:
    for position, char in enumerate(stash):
        text = text.replace(_PLACEHOLDER.format(position), char)
    return text


def parse_inline(text: str) -> list[Run]:
    """Split one paragraph's text into formatted runs.

    Unbalanced markers are left as literal characters — a stray asterisk in a
    signed lease is better than silently swallowing half a sentence.
    """
    source = str(text or "")
    # The sentinels must not occur in real content, so pasted control bytes go.
    for sentinel in _SENTINELS:
        source = source.replace(sentinel, "")
    if not source:
        return []
    protected, stash = _protect_escapes(source)

    # Walk the string once, deepest-binding marker first, recursing into the
    # content so bold and italic can nest.
    def walk(chunk: str, bold: bool, italic: bool) -> list[Run]:
        runs: list[Run] = []
        cursor = 0
        while cursor < len(chunk):
            bold_match = _BOLD_RE.search(chunk, cursor) if not bold else None
            italic_match = _ITALIC_RE.search(chunk, cursor) if not italic else None
            kp_match = _KP_RE.search(chunk, cursor)
            candidates = [m for m in (bold_match, italic_match, kp_match) if m]
            if not candidates:
                runs.extend(_plain_runs(chunk[cursor:], bold, italic, stash))
                break
            match = min(candidates, key=lambda m: m.start())
            if match.start() > cursor:
                runs.extend(_plain_runs(chunk[cursor:match.start()], bold, italic, stash))
            if match is kp_match:
                runs.append(Run(kind=RUN_KP_REF, name=_restore_escapes(match.group(1), stash),
                                bold=bold, italic=italic))
            elif match is bold_match:
                runs.extend(walk(match.group(1), True, italic))
            else:
                runs.extend(walk(match.group(1), bold, True))
            cursor = match.end()
        return runs

    return _merge_runs(walk(protected, False, False))


def _plain_runs(chunk: str, bold: bool, italic: bool, stash: list[str]) -> Iterator[Run]:
    """Literal text, split on the caret line-break marker."""
    for position, piece in enumerate(chunk.split(LINE_BREAK)):
        if position:
            yield Run(kind=RUN_BREAK)
        restored = _restore_escapes(piece, stash)
        if restored:
            yield Run(text=restored, bold=bold, italic=italic)


def _merge_runs(runs: list[Run]) -> list[Run]:
    """Collapse adjacent runs that share formatting; drop empties."""
    merged: list[Run] = []
    for run in runs:
        if run.kind == RUN_TEXT and not run.text:
            continue
        if (
            merged
            and merged[-1].kind == RUN_TEXT
            and run.kind == RUN_TEXT
            and merged[-1].bold == run.bold
            and merged[-1].italic == run.italic
        ):
            merged[-1] = replace(merged[-1], text=merged[-1].text + run.text)
        else:
            merged.append(run)
    return merged


# ---------------------------------------------------------------------------
# Block parsing
# ---------------------------------------------------------------------------

_BULLET_RE = re.compile(r"^([-*•])\s+(.*)$")
_NUMBERED_RE = re.compile(r"^(\d{1,3})[.)]\s+(.*)$")
_LETTERED_RE = re.compile(r"^([A-Za-z])[.)]\s+(.*)$")
_RENT_TABLE_RE = re.compile(r"^\[\s*RentTable\s*:\s*([^\]]+?)\s*\]$", re.IGNORECASE)
_PAGE_BREAK_RE = re.compile(r"^\[\s*PageBreak\s*\]$", re.IGNORECASE)
_EXHIBIT_RE = re.compile(r"^\[\s*Exhibit\s*:\s*([^\]|]+?)\s*(?:\|\s*([^\]]*?)\s*)?\]$", re.IGNORECASE)
# Legacy markers left over from the Word-template era.
_LEGACY_TABLE_RE = re.compile(r"^(Base|Option\s*\d*)\s+Rent\s+Table\s*:?\s*$", re.IGNORECASE)
_LEGACY_TABLE_END_RE = re.compile(r"^(Table\s+End|DELETE\s+TABLE)\s*$", re.IGNORECASE)

_INDENT_RE = re.compile(r"^(?:\t| {4})+")


def _indent_level(line: str) -> tuple[int, str]:
    """Count leading tabs / 4-space groups and return the stripped remainder."""
    level = 0
    rest = line
    while True:
        if rest.startswith("\t"):
            rest = rest[1:]
        elif rest.startswith("    "):
            rest = rest[4:]
        else:
            break
        level += 1
    return min(level, MAX_LEVEL), rest.strip()


def _classify(line: str) -> tuple[str, str, str]:
    """(kind, source_marker, content) for one already-dedented line."""
    match = _BULLET_RE.match(line)
    if match:
        return BLOCK_BULLET, match.group(1), match.group(2)
    match = _NUMBERED_RE.match(line)
    if match:
        return BLOCK_NUMBERED, match.group(1), match.group(2)
    match = _LETTERED_RE.match(line)
    if match:
        # "I. " and "V. " are ambiguous with roman numerals but the spec's level
        # 1 is lettered, so a single letter is always a lettered item.
        return BLOCK_LETTERED, match.group(1), match.group(2)
    return BLOCK_PARAGRAPH, "", line


def parse_blocks(text: str, renumber: bool = True) -> list[Block]:
    """Parse clause text into blocks.

    Consecutive non-blank lines join into one block (Word wraps for us); a blank
    line starts a new one. List markers and structural markers always start a
    new block regardless of blank lines.
    """
    source = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    # Strip the escape sentinels here as well as in parse_inline: they are not
    # whitespace, so a line holding only one would otherwise read as content and
    # glue two paragraphs together before being dropped downstream.
    for sentinel in _SENTINELS:
        source = source.replace(sentinel, "")
    blocks: list[Block] = []
    pending: Block | None = None
    pending_lines: list[str] = []
    legacy_table_open = False

    def flush() -> None:
        nonlocal pending, pending_lines
        if pending is not None:
            pending.runs = parse_inline(" ".join(pending_lines).strip())
            # A marker with no content at all is noise, not a paragraph.
            if pending.runs or pending.is_structural:
                blocks.append(pending)
        pending = None
        pending_lines = []

    for raw_line in source.split("\n"):
        if not raw_line.strip():
            flush()
            continue

        level, line = _indent_level(raw_line)

        structural = _structural_block(line, level, legacy_table_open)
        if structural is not None:
            flush()
            if structural.kind == BLOCK_RENT_TABLE and structural.data.get("legacy"):
                legacy_table_open = True
                structural.data.pop("legacy")
            elif structural.kind == "":  # legacy Table End: swallow the line
                legacy_table_open = False
                continue
            blocks.append(structural)
            continue

        kind, _marker, content = _classify(line)
        if kind == BLOCK_PARAGRAPH and pending is not None and not pending.is_structural:
            # Continuation of the paragraph or list item already open.
            pending_lines.append(content)
            continue

        flush()
        pending = Block(kind=kind, level=level)
        pending_lines = [content]

    flush()
    if renumber:
        assign_numbers(blocks)
    return blocks


def _structural_block(line: str, level: int, legacy_table_open: bool = False) -> Block | None:
    """A marker line, or None if this is ordinary content."""
    match = _RENT_TABLE_RE.match(line)
    if match:
        return Block(kind=BLOCK_RENT_TABLE, level=level, data={"schedule": match.group(1).strip()})
    if _PAGE_BREAK_RE.match(line):
        return Block(kind=BLOCK_PAGE_BREAK, level=level)
    match = _EXHIBIT_RE.match(line)
    if match:
        return Block(
            kind=BLOCK_EXHIBIT,
            level=level,
            data={"letter": match.group(1).strip(), "title": (match.group(2) or "").strip()},
        )
    match = _LEGACY_TABLE_RE.match(line)
    if match:
        schedule = re.sub(r"\s+", " ", match.group(1).strip()).title()
        return Block(kind=BLOCK_RENT_TABLE, level=level,
                     data={"schedule": schedule, "legacy": True})
    # "Table End" closes a legacy table; on its own it is just a sentence, and
    # swallowing it would delete text that no marker ever opened.
    if legacy_table_open and _LEGACY_TABLE_END_RE.match(line):
        return Block(kind="")  # sentinel: swallow the line
    return None


# ---------------------------------------------------------------------------
# Numbering
# ---------------------------------------------------------------------------

def _letter(index: int, upper: bool = True) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters if upper else letters.lower()


def assign_numbers(blocks: list[Block]) -> list[Block]:
    """Assign each list item its displayed marker.

    Counters are per (level, kind). Returning to a shallower level resets every
    deeper counter, so a nested list restarts inside each parent item. Plain
    paragraphs do not reset anything — an explanatory paragraph between A. and
    B. is normal drafting and must not restart the letters. A page break or an
    exhibit does reset, since that is a new part of the document.
    """
    counters: dict[tuple[int, str], int] = {}
    for block in blocks:
        if block.kind in (BLOCK_PAGE_BREAK, BLOCK_EXHIBIT):
            counters.clear()
            continue
        if not block.is_list:
            continue
        for key in [k for k in counters if k[0] > block.level]:
            del counters[key]
        key = (block.level, block.kind)
        counters[key] = counters.get(key, 0) + 1
        index = counters[key]
        if block.kind == BLOCK_BULLET:
            block.number = "•"
        elif block.kind == BLOCK_NUMBERED:
            block.number = f"{index}."
        else:
            block.number = f"{_letter(index)}."
    return blocks


# ---------------------------------------------------------------------------
# Serializers
# ---------------------------------------------------------------------------

def to_plain_text(blocks: list[Block], resolve: dict[str, str] | None = None) -> str:
    """Markup-free text with numbering applied. Used by previews and diffs."""
    values = {str(k).strip().lower(): str(v) for k, v in (resolve or {}).items()}
    lines: list[str] = []
    for block in blocks:
        if block.kind == BLOCK_PAGE_BREAK:
            lines.append("\f")
            continue
        if block.kind == BLOCK_RENT_TABLE:
            lines.append(f"[{block.data.get('schedule', 'Base')} Rent Table]")
            continue
        if block.kind == BLOCK_EXHIBIT:
            title = block.data.get("title", "")
            lines.append(f"Exhibit {block.data.get('letter', '')}" + (f" — {title}" if title else ""))
            continue
        body = "".join(
            "\n" if run.is_break
            else (values.get(run.name.strip().lower(), f"[KP:{run.name}]") if run.is_ref else run.text)
            for run in block.runs
        )
        indent = "    " * block.level
        prefix = f"{block.number} " if block.number else ""
        lines.append(indent + prefix + body)
    return "\n\n".join(lines)


def to_markup(blocks: list[Block]) -> str:
    """Back to source form. Round-tripping this must be stable."""
    lines: list[str] = []
    for block in blocks:
        if block.kind == BLOCK_PAGE_BREAK:
            lines.append("[PageBreak]")
            continue
        if block.kind == BLOCK_RENT_TABLE:
            lines.append(f"[RentTable:{block.data.get('schedule', 'Base')}]")
            continue
        if block.kind == BLOCK_EXHIBIT:
            title = block.data.get("title", "")
            lines.append(f"[Exhibit:{block.data.get('letter', '')}" + (f"|{title}]" if title else "]"))
            continue
        # Edge whitespace cannot survive a re-parse — lines are stripped, and a
        # leading run of spaces would be counted as indentation — so dropping it
        # here is what makes to_markup a fixed point.
        body = "".join(_run_markup(run) for run in block.runs).strip()
        if not body:
            # Nothing survived — e.g. a run that was only whitespace. Emitting a
            # blank line here would vanish on the next parse.
            continue
        if block.kind == BLOCK_PARAGRAPH:
            body = _disambiguate_paragraph(body)
        indent = "\t" * block.level
        if block.kind == BLOCK_BULLET:
            prefix = "- "
        elif block.kind == BLOCK_NUMBERED:
            prefix = block.number + " " if block.number else "1. "
        elif block.kind == BLOCK_LETTERED:
            prefix = block.number + " " if block.number else "A. "
        else:
            prefix = ""
        lines.append(indent + prefix + body)
    return "\n\n".join(lines)


def _disambiguate_paragraph(body: str) -> str:
    """Stop a paragraph that happens to open like "a) …" or "- …" from being
    re-read as a list item. Escaping the delimiter is enough, and the escape
    disappears again on the next parse."""
    kind, marker, _content = _classify(body.strip())
    if kind == BLOCK_PARAGRAPH:
        return body
    lead = body[: len(body) - len(body.lstrip())]
    rest = body[len(lead):]
    if kind == BLOCK_BULLET:
        return f"{lead}\\{rest}"  # marker char is the first char
    # Numbered/lettered: the delimiter is the character after the marker text.
    return f"{lead}{marker}\\{rest[len(marker):]}"


def _escape(text: str) -> str:
    for char in ("\\", "*", "_", "[", LINE_BREAK):
        text = text.replace(char, "\\" + char)
    return text


def _escape_name(name: str) -> str:
    """Provision names are a controlled vocabulary, but a stray backslash or
    bracket in one must not silently re-cut the token on the next parse."""
    return str(name).replace("\\", "\\\\").replace("]", "\\]")


def _run_markup(run: Run) -> str:
    if run.is_break:
        return LINE_BREAK
    if run.is_ref:
        return f"[KP:{_escape_name(run.name)}]"
    if not (run.bold or run.italic):
        return _escape(run.text)
    # Emphasis markers must hug the words: "_ text _" is not italic by the
    # italic rule, so leaving the padding inside would not survive a re-parse.
    stripped = run.text.strip()
    if not stripped:
        return _escape(run.text)
    lead = run.text[: len(run.text) - len(run.text.lstrip())]
    trail = run.text[len(run.text.rstrip()):]
    body = _escape(stripped)
    if run.italic:
        body = f"_{body}_"
    if run.bold:
        body = f"**{body}**"
    return _escape(lead) + body + _escape(trail)


_HTML_ESCAPES = {"&": "&amp;", "<": "&lt;", ">": "&gt;"}


def _html_escape(text: str) -> str:
    for char, entity in _HTML_ESCAPES.items():
        text = text.replace(char, entity)
    return text


def to_html(blocks: list[Block], resolve: dict[str, str] | None = None,
            link_color: str = "1F7A33") -> str:
    """Preview HTML. Cross-references render the way the Word output will:
    green and underlined, or as a red unresolved token when the name is unknown."""
    values = {str(k).strip().lower(): str(v) for k, v in (resolve or {}).items()}
    out: list[str] = []
    for block in blocks:
        if block.kind == BLOCK_PAGE_BREAK:
            out.append('<hr style="border:0;border-top:1px dashed #999;margin:14px 0">')
            continue
        if block.kind == BLOCK_RENT_TABLE:
            out.append(
                '<div style="border:1px dashed #999;padding:6px;color:#555;font-style:italic">'
                f"{_html_escape(str(block.data.get('schedule', 'Base')))} rent table</div>"
            )
            continue
        if block.kind == BLOCK_EXHIBIT:
            title = _html_escape(str(block.data.get("title", "")))
            out.append(
                f'<p style="text-align:center;font-weight:700">Exhibit '
                f'{_html_escape(str(block.data.get("letter", "")))}'
                + (f"<br>{title}" if title else "")
                + "</p>"
            )
            continue

        pieces: list[str] = []
        for run in block.runs:
            if run.is_break:
                pieces.append("<br>")
                continue
            if run.is_ref:
                value = values.get(run.name.strip().lower())
                if value is None:
                    pieces.append(
                        '<span style="color:#b00020;font-weight:600">'
                        f"[KP:{_html_escape(run.name)}]</span>"
                    )
                else:
                    pieces.append(
                        f'<span style="color:#{link_color};text-decoration:underline">'
                        f"{_html_escape(value)}</span>"
                    )
                continue
            body = _html_escape(run.text)
            if run.bold:
                body = f"<strong>{body}</strong>"
            if run.italic:
                body = f"<em>{body}</em>"
            pieces.append(body)

        indent = block.level * 24 + (24 if block.is_list else 0)
        marker = f"{_html_escape(block.number)} " if block.number else ""
        out.append(
            f'<p style="margin:0 0 8px 0;text-align:justify;padding-left:{indent}px'
            f'{";text-indent:-24px" if block.number else ""}">{marker}{"".join(pieces)}</p>'
        )
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Convenience
# ---------------------------------------------------------------------------

def parse_clause(text: str) -> list[Block]:
    """Main entry point."""
    return parse_blocks(text)


def collect_refs(blocks: list[Block]) -> list[str]:
    """Provision names cited anywhere in these blocks, in order, deduplicated."""
    seen: list[str] = []
    for block in blocks:
        for name in block.refs():
            if name not in seen:
                seen.append(name)
    return seen


def rent_schedules_used(blocks: list[Block]) -> list[str]:
    """Which rent schedules the text asks for, in order."""
    seen: list[str] = []
    for block in blocks:
        if block.kind == BLOCK_RENT_TABLE:
            schedule = str(block.data.get("schedule", "Base"))
            if schedule not in seen:
                seen.append(schedule)
    return seen
