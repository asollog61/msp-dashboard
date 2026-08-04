"""Read the master lease's highlight convention into a menu of choices.

The master `.docx` is both a readable Word document and the machine-readable
menu the Lease Builder offers. Highlight colour carries the meaning:

    (none)      in the lease by default; still removable, just not surfaced
    green       lives in the master, OUT of the lease unless opted in
    cyan        one of an adjacent set — pick exactly one, or none
    yellow      an optional key provision (front matter)
    lightGray   scaffolding: drafting directions, stripped and never rendered
    red         scaffolding: notes to self, stripped and never rendered

Two structural rules do work that colour cannot:

**Indent groups a block.** A block begins at a highlighted paragraph and
absorbs the highlighted paragraphs that follow it at a deeper indent. So
"Conditional Reduction" carries "Application of Credit" and "Reinstatement
Rights" with it, and switching the parent off takes the children too.

**Adjacency groups a choice.** Consecutive cyan blocks at the same indent form
one pick-one set. Real content between them closes the set; a rule of asterisks
does not — that is a divider drawn inside a set, and it is stripped on the way
out.

Containers are numbered sections *and* the exhibits after them. Without that
second half every exhibit collapses into Section 56, which is where the
Landlord's Work phases and the guaranty variants live — reading them as
alternatives to each other would let you pick Gas or Water but not both.

Pure data: takes a path, returns dictionaries. No Streamlit, no app state.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn

# Meaning of each highlight, and whether it survives into a lease.
OPTIONAL_OUT = "green"      # present in the master, absent from the lease
PICK_ONE = "cyan"           # one of an adjacent set
KEY_PROVISION = "yellow"    # optional key provision, front matter
SCAFFOLDING = ("lightGray", "red")

SECTION_RE = re.compile(r"^\s*Section\s+(\d+(?:\.\d+)?)\s*\.?\s*(.*)$", re.IGNORECASE)
EXHIBIT_RE = re.compile(r'^\s*[“"\']?\s*(Exhibit\s+[A-Z0-9]+)\s*[”"\']?\s*$', re.IGNORECASE)
# Centred all-caps headings that open their own container.
NAMED_CONTAINER_RE = re.compile(
    r"^\s*(GUARANTY OF LEASE|SIGNATURES|TENANT FORMS\.?)\s*$", re.IGNORECASE
)
# A run-in heading: the bold lead-in that names a block.
RUN_IN_RE = re.compile(r"^\s*([^.:\n]{3,70}?)\s*[.:]\s")
# A rule of repeated symbols: a hard break between choice groups. Used where a
# section is flat and indentation cannot express that one group has ended.
SEPARATOR_RE = re.compile(r"^[\s*_=~.\-\u2013\u2014]{3,}$")
# The colour legend at the top of the master describes the convention; it is
# not part of any lease.
LEGEND_RE = re.compile(r"^\s*(YELLOW|GREEN|CYAN|TEAL|RED|GREY|GRAY|MAGENTA)\s*=", re.IGNORECASE)


def _highlight(run: Any) -> str | None:
    properties = run._element.find(qn("w:rPr"))
    if properties is None:
        return None
    element = properties.find(qn("w:highlight"))
    return element.get(qn("w:val")) if element is not None else None


def _indent(paragraph: Any) -> int:
    properties = paragraph._element.find(qn("w:pPr"))
    if properties is None:
        return 0
    element = properties.find(qn("w:ind"))
    if element is None:
        return 0
    try:
        return int(element.get(qn("w:left")) or 0)
    except (TypeError, ValueError):
        return 0


def _alignment(paragraph: Any) -> str:
    properties = paragraph._element.find(qn("w:pPr"))
    if properties is None:
        return ""
    element = properties.find(qn("w:jc"))
    return (element.get(qn("w:val")) or "") if element is not None else ""


def paragraph_colour(paragraph: Any) -> str | None:
    """The one highlight that governs a paragraph, or None.

    A paragraph part-highlighted in a colour still counts as that colour: the
    convention works on whole paragraphs, and a half-marked one is far more
    likely to be a marking slip than a deliberate inline alternative. Those
    are reported rather than silently reinterpreted.
    """
    colours = {
        colour for run in paragraph.runs
        if (colour := _highlight(run)) and run.text.strip()
    }
    if not colours:
        return None
    for preferred in (PICK_ONE, OPTIONAL_OUT, KEY_PROVISION):
        if preferred in colours:
            return preferred
    return sorted(colours)[0]


def is_partly_highlighted(paragraph: Any) -> bool:
    """True when only some of a paragraph's text carries a highlight."""
    marked = "".join(run.text for run in paragraph.runs if _highlight(run)).strip()
    whole = paragraph.text.strip()
    return bool(marked) and len(marked) < len(whole) * 0.9


def block_name(text: str) -> str:
    """A block's label, taken from its run-in heading where there is one."""
    match = RUN_IN_RE.match(str(text or ""))
    if match:
        return match.group(1).strip()
    words = str(text or "").split()
    return " ".join(words[:7]) + ("…" if len(words) > 7 else "")


# ---------------------------------------------------------------------------
# Containers
# ---------------------------------------------------------------------------

def _container_for(paragraph: Any, text: str) -> tuple[str, str] | None:
    """(kind, label) when this paragraph opens a new container."""
    match = SECTION_RE.match(text)
    if match:
        return ("section", match.group(1))
    centred = _alignment(paragraph) == "center"
    match = EXHIBIT_RE.match(text)
    if match and centred:
        return ("exhibit", match.group(1).title())
    if centred and NAMED_CONTAINER_RE.match(text):
        return ("appendix", text.strip().title())
    return None


def _iter_body(document: Any):
    """Paragraphs and tables in the order they appear.

    python-docx's `document.paragraphs` skips table cells entirely, which is
    where the Key Provisions Summary lives — 45 of the yellow provisions were
    invisible until this walked the body itself.
    """
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield "tbl", Table(child, document)


def _key_provisions_from_table(table: Any) -> list[dict[str, Any]]:
    """Optional key provisions read off a summary table.

    A row is a field/value pair. Yellow anywhere in the row marks the whole
    provision optional, and the first cell names it.
    """
    found = []
    for row in table.rows:
        cells = row.cells
        if not cells:
            continue
        marked = any(
            paragraph_colour(paragraph) == KEY_PROVISION
            for cell in cells for paragraph in cell.paragraphs
        )
        if not marked:
            continue
        label = cells[0].text.strip()
        value = " ".join(cell.text.strip() for cell in cells[1:]).strip()
        if not label and not value:
            continue
        found.append({
            "name": block_name(label or value),
            "colour": KEY_PROVISION,
            "indent": 0,
            "text": value or label,
            "children": [],
            "position": -1,
        })
    return found


def read_master(path: str | Path) -> dict[str, Any]:
    """Every container in the master, with its optional blocks and choices."""
    document = Document(str(path))

    containers: list[dict[str, Any]] = []
    current = {"kind": "front", "label": "Key Provisions Summary",
               "title": "", "paragraphs": []}
    warnings: list[str] = []
    pending_title = False

    for kind, item in _iter_body(document):
        if kind == "tbl":
            current.setdefault("table_blocks", []).extend(_key_provisions_from_table(item))
            continue
        paragraph = item
        text = paragraph.text.strip()
        if not text:
            continue
        if LEGEND_RE.match(text):
            continue  # the colour key, not lease language
        if SEPARATOR_RE.match(text):
            current["paragraphs"].append({"text": "", "colour": None,
                                          "indent": 0, "separator": True})
            continue

        opened = _container_for(paragraph, text)
        if opened:
            containers.append(current)
            kind, label = opened
            title = ""
            if kind == "section":
                title = SECTION_RE.match(text).group(2).strip()
            current = {"kind": kind, "label": label, "title": title, "paragraphs": []}
            pending_title = kind in ("exhibit", "appendix")
            continue

        # The line under an exhibit heading is its title, not body text.
        if pending_title and _alignment(paragraph) == "center" and len(text) < 80:
            current["title"] = text
            pending_title = False
            continue
        pending_title = False

        colour = paragraph_colour(paragraph)
        if colour in SCAFFOLDING:
            continue  # drafting directions never reach a lease
        if colour and is_partly_highlighted(paragraph):
            warnings.append(
                f"{current['label']}: “{text[:60]}…” is only part-highlighted "
                f"({colour}). The whole paragraph will be treated as {colour}."
            )
        current["paragraphs"].append(
            {"text": text, "colour": colour, "indent": _indent(paragraph)}
        )

    containers.append(current)
    for container in containers:
        container["blocks"] = (container.pop("table_blocks", [])
                               + _group_blocks(container["paragraphs"]))
        container.pop("paragraphs")

    return {"source": Path(path).name, "containers": containers, "warnings": warnings}


# ---------------------------------------------------------------------------
# Blocks and choice groups
# ---------------------------------------------------------------------------

def _group_blocks(paragraphs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Highlighted paragraphs collapsed into blocks, then into choice groups."""
    blocks: list[dict[str, Any]] = []
    index = 0
    while index < len(paragraphs):
        entry = paragraphs[index]
        if entry.get("separator") or not entry["colour"]:
            index += 1
            continue
        children: list[str] = []
        cursor = index + 1
        # Deeper-indented highlighted paragraphs belong to this block.
        while cursor < len(paragraphs):
            following = paragraphs[cursor]
            if not following["colour"] or following["indent"] <= entry["indent"]:
                break
            children.append(following["text"])
            cursor += 1
        blocks.append({
            "name": block_name(entry["text"]),
            "colour": entry["colour"],
            "indent": entry["indent"],
            "text": entry["text"],
            "children": children,
            "position": index,
        })
        index = cursor

    _assign_choice_groups(blocks, paragraphs)
    return blocks


def _assign_choice_groups(blocks: list[dict[str, Any]], paragraphs: list[dict]) -> None:
    """Consecutive cyan blocks at one indent become a single pick-one group."""
    blocks = [b for b in blocks if b["position"] >= 0]
    group = 0
    run: list[dict[str, Any]] = []

    def close():
        nonlocal group, run
        if len(run) >= 2:
            group += 1
            for member in run:
                member["choice_group"] = group
        run = []

    previous_end = None
    for block in blocks:
        if block["colour"] != PICK_ONE:
            close()
            previous_end = None
            continue
        # Real content between two cyan blocks separates them. A rule of
        # asterisks does not: it is a divider drawn *inside* a set to show
        # where one option ends and the next begins, which is how it reads in
        # Word. Treating it as a group break turned every marked pair into two
        # unrelated singles.
        gap_clean = previous_end is None or not any(
            not paragraphs[i]["colour"] and not paragraphs[i].get("separator")
            for i in range(previous_end, block["position"])
        )
        if run and (block["indent"] != run[-1]["indent"] or not gap_clean):
            close()
        run.append(block)
        previous_end = block["position"] + len(block["children"]) + 1
        while (previous_end < len(paragraphs)
               and paragraphs[previous_end].get("separator")):
            previous_end += 1
    close()


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def report(master: dict[str, Any]) -> str:
    """A human-readable reading of the master, to check before it is trusted."""
    lines = [f"Master: {master['source']}", ""]
    totals = {OPTIONAL_OUT: 0, PICK_ONE: 0, KEY_PROVISION: 0}

    for container in master["containers"]:
        blocks = container["blocks"]
        if not blocks:
            continue
        heading = container["label"]
        if container["kind"] == "section":
            heading = f"Section {container['label']}. {container['title']}"
        elif container["title"]:
            heading = f"{container['label']} — {container['title']}"
        lines.append(heading)

        groups: dict[int, list[dict]] = {}
        for block in blocks:
            totals[block["colour"]] = totals.get(block["colour"], 0) + 1
            if block.get("choice_group"):
                groups.setdefault(block["choice_group"], []).append(block)

        shown = set()
        for block in blocks:
            group = block.get("choice_group")
            if group:
                if group in shown:
                    continue
                shown.add(group)
                members = groups[group]
                lines.append(f"    PICK ONE of {len(members)}:")
                for member in members:
                    extra = f" (+{len(member['children'])})" if member["children"] else ""
                    lines.append(f"        · {member['name']}{extra}")
            else:
                mark = {OPTIONAL_OUT: "OFF by default",
                        PICK_ONE: "optional (single)",
                        KEY_PROVISION: "key provision"}.get(block["colour"], block["colour"])
                extra = f" (+{len(block['children'])} sub)" if block["children"] else ""
                lines.append(f"    [ ] {block['name']}{extra}  — {mark}")
        lines.append("")

    lines.append(f"TOTALS  off-by-default: {totals.get(OPTIONAL_OUT, 0)}   "
                 f"pick-one: {totals.get(PICK_ONE, 0)}   "
                 f"key provisions: {totals.get(KEY_PROVISION, 0)}")
    if master["warnings"]:
        lines.append("")
        lines.append("WARNINGS")
        for warning in master["warnings"]:
            lines.append(f"  ! {warning}")
    return "\n".join(lines)
