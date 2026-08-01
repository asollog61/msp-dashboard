"""The lease's own words, extracted out of Word and into JSON.

Until now the master `.docx` was the only live copy of the lease language: the
Lease Builder read its 60 sections and its key-provision bookmarks straight out
of the file on every load. That made the Word file load-bearing — deleting it
took the whole tab down — and it is exactly the dependency LEASE_FORMAT_SPEC.md
sets out to remove.

This module extracts that content once into `lease_content.json`, preserving
bold and italic as the markup subset `lease_markup.py` understands, and gives
the app a loader to read it back. After extraction the `.docx` is a historical
artifact, not a runtime dependency.

    python lease_content.py --extract     # rebuild the JSON from the master
    python lease_content.py --check       # verify JSON still matches the master
"""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import re
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.text.run import Run

import lease_builder as lb

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "data" / "Lease Builder"
CONTENT_FILE = TEMPLATE_DIR / "lease_content.json"
DEFAULT_SOURCE = TEMPLATE_DIR / "2026_07_30 MSP Master_Lease.v8 NNN Retail.docx"

CONTENT_VERSION = 1

# The heading is rebuilt by the renderer from the number and title, so it is
# stripped off the stored body. Matches "Section 5.<tab>Base and Additional
# Rent.  " including the run-in title.
_HEADING_RE = re.compile(
    r"^\s*Section\s+\d+(?:\.\d+)?\s*\.?\s*\t?\s*", re.IGNORECASE
)


# ---------------------------------------------------------------------------
# DOCX -> markup
# ---------------------------------------------------------------------------

def iter_runs(paragraph: Any) -> list[Any]:
    """Every run in the paragraph, in document order.

    `paragraph.runs` only returns direct w:r children, so runs nested inside a
    w:hyperlink are invisible to it — and cross-references are hyperlinks. Using
    it silently dropped text like "(#Landlord_Work_Itemized)" from section 9.
    Field-code runs are skipped: their instrText is Word plumbing, not content.
    """
    runs = []
    for element in paragraph._p.iter(qn("w:r")):
        if element.find(qn("w:instrText")) is not None:
            continue
        if element.find(qn("w:fldChar")) is not None:
            continue
        if _inside(element, "w:del"):
            continue  # tracked deletion: treat as already removed
        runs.append(Run(element, paragraph))
    return runs


def _inside(element: Any, tag: str) -> bool:
    parent = element.getparent()
    while parent is not None:
        if parent.tag == qn(tag):
            return True
        parent = parent.getparent()
    return False


def tracked_changes(document: Any) -> dict[str, int]:
    """How many unaccepted revisions the source still carries.

    This matters more than it looks. `paragraph.text` ignores runs nested in
    w:ins, so with revisions pending Word's own text and python-docx's disagree:
    section 47 read as "This ease contains the entire" and section 46 came back
    empty. Extraction resolves revisions the way Word's "Accept All" would —
    insertions kept, deletions dropped — and records the counts so the numbers
    can be sanity-checked against the file.
    """
    body = document.element.body
    return {
        "insertions": len(body.findall(".//" + qn("w:ins"))),
        "deletions": len(body.findall(".//" + qn("w:del"))),
    }


def paragraph_to_markup(paragraph: Any, skip_chars: int = 0) -> str:
    """One Word paragraph as markup, preserving bold and italic.

    Formatting is emitted per contiguous group of like-formatted runs, because
    Word splits a single bold phrase across many runs and per-run markers would
    produce "**Rent**** ****Commencement**".

    `skip_chars` drops that many leading visible characters — used to remove a
    run-in section heading whose bold may span the heading and the body both.
    """
    pieces: list[str] = []
    group_text = ""
    group_bold = False
    group_italic = False

    def close_group() -> None:
        nonlocal group_text, group_bold, group_italic
        if not group_text:
            return
        stripped = group_text.strip()
        if stripped and (group_bold or group_italic):
            lead = group_text[: len(group_text) - len(group_text.lstrip())]
            trail = group_text[len(group_text.rstrip()):]
            body = _escape_markup(stripped)
            if group_italic:
                body = f"_{body}_"
            if group_bold:
                body = f"**{body}**"
            pieces.append(lead + body + trail)
        else:
            pieces.append(_escape_markup(group_text))
        group_text = ""

    remaining_skip = max(0, int(skip_chars))
    for run in iter_runs(paragraph):
        text = run.text or ""
        if remaining_skip:
            if len(text) <= remaining_skip:
                remaining_skip -= len(text)
                continue
            text = text[remaining_skip:]
            remaining_skip = 0
        if not text:
            continue
        bold, italic = bool(run.bold), bool(run.italic)
        if bold == group_bold and italic == group_italic:
            group_text += text
        else:
            close_group()
            group_bold, group_italic = bold, italic
            group_text = text
    close_group()
    return "".join(pieces)


def _escape_markup(text: str) -> str:
    """Protect characters the parser would otherwise read as markup.

    [KP:Name] tokens are inserted by lease_builder before extraction and must
    survive, so brackets are only escaped when they are not part of a token.
    """
    placeholder_map: dict[str, str] = {}

    def stash(match: re.Match) -> str:
        key = f"\x02{len(placeholder_map)}\x03"
        placeholder_map[key] = match.group(0)
        return key

    protected = lb.KP_TOKEN_RE.sub(stash, text)
    for char in ("\\", "*", "_", "[", "]", "^"):
        protected = protected.replace(char, "\\" + char)
    for key, original in placeholder_map.items():
        protected = protected.replace(key, original)
    return protected


def heading_length(text: str, number: str, title: str) -> int:
    """How many leading characters of the first paragraph are the run-in heading.

    Measured on plain text and then applied to the runs, because the bold in
    "**Section 9.  Landlord's Work**." covers the heading but stops before the
    period — a regex over the markup cannot see where the heading ends.
    """
    pattern = (
        r"^\s*Section\s+" + re.escape(str(number)) + r"\s*\.?[ \t]*"
        + re.escape(str(title)) + r"\s*\.?[ \t]*"
    )
    match = re.match(pattern, str(text or ""), re.IGNORECASE)
    return match.end() if match else 0


def _section_body(paragraphs: list[Any], start: int, end: int,
                  number: str = "", title: str = "") -> str:
    """Markup for one section, run-in heading removed."""
    lines: list[str] = []
    for index in range(start, end):
        paragraph = paragraphs[index]
        skip = heading_length(paragraph.text, number, title) if index == start else 0
        markup = paragraph_to_markup(paragraph, skip_chars=skip).strip()
        if markup:
            lines.append(markup)
    return "\n\n".join(lines)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def extract_from_docx(source: str | Path = DEFAULT_SOURCE) -> dict[str, Any]:
    """Read the master template into a plain-data clause library."""
    path = Path(source)
    document = Document(str(path))

    bookmarks = []
    for name, (_, _, elements) in lb._bookmark_ranges(document).items():
        if name not in lb.BOOKMARK_LABELS:
            continue
        value = "".join(
            element.text or "" for element in elements if element.tag == qn("w:t")
        ).strip()
        bookmarks.append(
            {"bookmark": name, "field": lb.BOOKMARK_LABELS[name], "value": value}
        )
    bookmarks.sort(key=lambda item: list(lb.BOOKMARK_LABELS).index(item["bookmark"]))

    known = {item["bookmark"] for item in bookmarks}
    orphan_refs = [
        {"bookmark": target, "field": lb.BOOKMARK_LABELS.get(target, lb.humanize_bookmark(target))}
        for target in lb.collect_ref_targets(document)
        if target not in known
    ]

    # Turn Word REF fields into [KP:Name] tokens before reading any text, so the
    # stored clause carries the cross-reference rather than a stale field.
    names_by_id = {item["bookmark"]: item["field"] for item in bookmarks}
    names_by_id.update({item["bookmark"]: item["field"] for item in orphan_refs})
    lb.convert_ref_fields_to_tokens(document, names_by_id)

    paragraphs = document.paragraphs
    sections = []
    for section in lb.scan_sections(document):
        sections.append(
            {
                "number": str(section["number"]),
                "title": section["title"],
                "category": section["category"],
                "body": _section_body(
                    paragraphs, section["start"], section["end"],
                    number=str(section["number"]), title=section["title"],
                ),
            }
        )

    return {
        "version": CONTENT_VERSION,
        "source": path.name,
        "extracted_at": datetime.now().isoformat(timespec="seconds"),
        "tracked_changes_resolved": tracked_changes(document),
        "key_provisions": bookmarks,
        "orphan_refs": orphan_refs,
        "sections": sections,
    }


def write_content(library: dict[str, Any], destination: str | Path = CONTENT_FILE) -> Path:
    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(library, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def load_content(source: str | Path = CONTENT_FILE) -> dict[str, Any]:
    """The clause library, or an empty shell when it has not been extracted."""
    try:
        with Path(source).open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {"version": CONTENT_VERSION, "sections": [], "key_provisions": [], "orphan_refs": []}
    if not isinstance(data, dict):
        return {"version": CONTENT_VERSION, "sections": [], "key_provisions": [], "orphan_refs": []}
    data.setdefault("sections", [])
    data.setdefault("key_provisions", [])
    data.setdefault("orphan_refs", [])
    return data


def content_available(source: str | Path = CONTENT_FILE) -> bool:
    return bool(load_content(source)["sections"])


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def _words(text: str) -> list[str]:
    """Comparable words: markup, numbering and whitespace removed."""
    plain = re.sub(r"\\([\\*_\[\]^])", r"\1", str(text or ""))
    plain = plain.replace("**", "")
    plain = re.sub(r"(?<![A-Za-z0-9_])_|_(?![A-Za-z0-9_])", "", plain)
    return re.sub(r"\s+", " ", plain).strip().split()


def compare_to_docx(library: dict[str, Any], source: str | Path = DEFAULT_SOURCE) -> list[str]:
    """Differences between a stored library and the Word master. Empty is good."""
    problems: list[str] = []
    fresh = extract_from_docx(source)

    stored_sections = {s["number"]: s for s in library.get("sections", [])}
    fresh_sections = {s["number"]: s for s in fresh["sections"]}

    for number in sorted(set(fresh_sections) - set(stored_sections), key=float):
        problems.append(f"Section {number} is in the master but missing from the library.")
    for number in sorted(set(stored_sections) - set(fresh_sections), key=float):
        problems.append(f"Section {number} is in the library but no longer in the master.")

    for number in sorted(set(stored_sections) & set(fresh_sections), key=float):
        stored, current = stored_sections[number], fresh_sections[number]
        if stored.get("title") != current.get("title"):
            problems.append(
                f"Section {number} title differs: {stored.get('title')!r} vs {current.get('title')!r}"
            )
        missing = [w for w in _words(current["body"]) if w not in _words(stored["body"])]
        if missing:
            problems.append(f"Section {number} is missing {len(missing)} word(s): {missing[:5]}")

    stored_kp = {item["bookmark"] for item in library.get("key_provisions", [])}
    fresh_kp = {item["bookmark"] for item in fresh["key_provisions"]}
    for bookmark in sorted(fresh_kp - stored_kp):
        problems.append(f"Key provision {bookmark} is missing from the library.")

    return problems


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    argument = sys.argv[1] if len(sys.argv) > 1 else "--extract"
    if argument == "--extract":
        library = extract_from_docx()
        written = write_content(library)
        print(f"Wrote {written}")
        print(f"  {len(library['sections'])} sections, "
              f"{len(library['key_provisions'])} key provisions, "
              f"{len(library['orphan_refs'])} orphan refs")
    elif argument == "--check":
        issues = compare_to_docx(load_content())
        if issues:
            print(f"{len(issues)} problem(s):")
            for issue in issues:
                print("  -", issue)
            sys.exit(1)
        print("Library matches the master template.")
    else:
        print(__doc__)
        sys.exit(2)
