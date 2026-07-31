"""Lease Builder helpers for the MSP Property Dashboard."""

from __future__ import annotations

from io import BytesIO
import json
from pathlib import Path
import re
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "data" / "Lease Builder"
CLAUSE_LIBRARY_FILE = TEMPLATE_DIR / "lease_clause_library.json"
SECTION_RE = re.compile(
    r"^\s*Section\s+(\d+(?:\.\d+)?)(?:\.(?:\s+|$)|\s+)(.*)$",
    re.IGNORECASE,
)

CATEGORY_RANGES = [
    (1, 9, "Lease Economics & Delivery"),
    (10, 17, "Premises Operations"),
    (18, 26, "Transfers, Casualty & Property Rights"),
    (27, 40, "Defaults, Remedies & Enforcement"),
    (41, 46, "Insurance, Indemnity & Environmental"),
    (47, 56, "General & Boilerplate"),
]


def discover_templates() -> list[dict[str, str]]:
    """Return deployable DOCX templates; future templates are auto-discovered."""
    templates = []
    if not TEMPLATE_DIR.exists():
        return templates
    for path in sorted(TEMPLATE_DIR.glob("*.docx")):
        if path.name.startswith("~$") or " TEST" in path.stem.upper():
            continue
        stem = path.stem
        label = re.sub(r"^\d{4}_\d{2}_\d{2}\s+", "", stem)
        label = re.sub(r"\s*\.v(\d+)\s*", r" — v\1 ", label, flags=re.IGNORECASE)
        templates.append({"label": label.strip(), "path": str(path)})
    return templates


def load_clause_library() -> dict[str, Any]:
    try:
        with CLAUSE_LIBRARY_FILE.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {"sections": {}, "additional_provisions": {}}


def category_for(number: str) -> str:
    try:
        integer = int(float(number))
    except ValueError:
        return "Other"
    for start, end, category in CATEGORY_RANGES:
        if start <= integer <= end:
            return category
    return "Other"


def _section_title(remainder: str, number: str) -> str:
    remainder = " ".join(remainder.split()).strip()
    if not remainder:
        return f"Section {number}"
    title = remainder.split(".", 1)[0].strip(" .\t")
    return title[:120] or f"Section {number}"


def scan_sections(document: Document) -> list[dict[str, Any]]:
    """Locate the numbered lease body, excluding drafting notes and exhibits."""
    paragraphs = document.paragraphs
    starts: list[tuple[int, str, str]] = []
    body_started = False
    for index, paragraph in enumerate(paragraphs):
        match = SECTION_RE.match(paragraph.text.strip())
        if not match:
            continue
        number, remainder = match.groups()
        if number == "1" and index >= 50:
            body_started = True
        if body_started:
            starts.append((index, number, remainder))
        if body_started and number == "56":
            break

    sections = []
    for position, (start, number, remainder) in enumerate(starts):
        if position + 1 < len(starts):
            end = starts[position + 1][0]
        else:
            end = next(
                (i for i in range(start + 1, len(paragraphs))
                 if paragraphs[i].text.strip().upper() == "SIGNATURES"),
                len(paragraphs),
            )
        clause_text = "\n\n".join(
            paragraph.text.strip()
            for paragraph in paragraphs[start:end]
            if paragraph.text.strip()
        )
        sections.append({
            "number": number,
            "title": _section_title(remainder, number),
            "category": category_for(number),
            "start": start,
            "end": end,
            "text": clause_text,
            "paragraph_count": end - start,
        })
    return sections


BOOKMARK_LABELS = {
    "Tx_BuildingAddress": "Property",
    "Tx_LeaseEffeftiveDt": "Lease Execution Date",
    "Tx_Landlord": "Landlord",
    "Tx_Tenant": "Tenant",
    "Tx_Premises": "Premises",
    "Tx_Sqft": "Premises Sqft",
    "Tx_TenantShareProperty": "Tenant Share of Property",
    "Tx_TenantShareFloor": "Tenant Share of Floor",
    "Tx_TenantSharePremisis": "Tenant Share of Premises",
    "Tx_PermittedUse": "Permitted Use",
    "Tx_LandlordContact": "Landlord Contact",
    "Tx_TenantContact": "Tenant Contact",
    "Tx_LeaseType": "Lease Type",
    "Tx_LeaseCommenceDt": "Lease Commencement Date",
    "Tx_DeliveryCondition": "Lease Delivery Condition",
    "Tx_RentCommDt": "Rent Commencement Date",
    "Tx_PermitPer": "Permit Period",
    "Tx_LeaseExpDt": "Lease Expiration Date",
    "Tx_LeaseTerm": "Lease Term",
    "Tx_AdditionalRent": "Additional Rent",
    "Tx_OptionPeriod": "Option Period(s)",
    "Tx_OptiontoCancel": "Option to Cancel",
    "Tx_Utilities": "Utilities and Services",
    "Tx_SecurityDeposit": "Security Deposit",
    "Tx_Brokers": "Broker(s)",
    "Tx_LandlordContFitUp": "Landlord Fit-Up Contribution",
}


def _bookmark_ranges(document: Document) -> dict[str, tuple[Any, Any, list[Any]]]:
    elements = list(document.element.body.iter())
    end_positions = {
        element.get(qn("w:id")): index
        for index, element in enumerate(elements)
        if element.tag == qn("w:bookmarkEnd")
    }
    ranges = {}
    for index, element in enumerate(elements):
        if element.tag != qn("w:bookmarkStart"):
            continue
        name = element.get(qn("w:name"))
        end_index = end_positions.get(element.get(qn("w:id")))
        if not name or end_index is None:
            continue
        ranges[name] = (element, elements[end_index], elements[index + 1:end_index])
    return ranges


def inspect_template(template_path: str | Path) -> dict[str, Any]:
    document = Document(str(template_path))
    bookmarks = []
    for name, (_, _, elements) in _bookmark_ranges(document).items():
        if name not in BOOKMARK_LABELS:
            continue
        value = "".join(
            element.text or "" for element in elements if element.tag == qn("w:t")
        ).strip()
        bookmarks.append({"bookmark": name, "field": BOOKMARK_LABELS[name], "value": value})
    bookmarks.sort(key=lambda item: list(BOOKMARK_LABELS).index(item["bookmark"]))
    return {"sections": scan_sections(document), "bookmarks": bookmarks}


def _replace_bookmark_text(document: Document, name: str, value: str) -> bool:
    target = _bookmark_ranges(document).get(name)
    if not target:
        return False
    start, _, elements = target
    text_elements = [element for element in elements if element.tag == qn("w:t")]
    if text_elements:
        text_elements[0].text = value
        if value.startswith(" ") or value.endswith(" "):
            text_elements[0].set(qn("xml:space"), "preserve")
        for element in text_elements[1:]:
            element.text = ""
    else:
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = value
        run.append(text)
        start.addnext(run)
    return True


def _delete_paragraph(paragraph: Paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _ancestor(element: Any, tag: str) -> Any | None:
    current = element
    while current is not None:
        if current.tag == tag:
            return current
        current = current.getparent()
    return None


def _relocate_empty_bookmark(document: Document, start: Any, end: Any) -> None:
    """Keep an excluded bookmark valid while removing its visible source row."""
    for element in (start, end):
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    paragraph = OxmlElement("w:p")
    run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    run_properties.append(OxmlElement("w:vanish"))
    run.append(run_properties)
    text = OxmlElement("w:t")
    text.text = ""
    run.append(text)
    paragraph.extend([start, run, end])

    body = document.element.body
    section_properties = body.find(qn("w:sectPr"))
    if section_properties is not None:
        section_properties.addprevious(paragraph)
    else:
        body.append(paragraph)


def _apply_key_provision_inclusions(document: Document, included_bookmarks: set[str]) -> None:
    """Remove fully excluded KPS rows while preserving blank REF sources."""
    ranges = _bookmark_ranges(document)
    row_groups: dict[int, dict[str, Any]] = {}
    for name in BOOKMARK_LABELS:
        target = ranges.get(name)
        if not target:
            continue
        row = _ancestor(target[0], qn("w:tr"))
        if row is None:
            if name not in included_bookmarks:
                _replace_bookmark_text(document, name, "")
            continue
        group = row_groups.setdefault(id(row), {"row": row, "names": []})
        group["names"].append(name)

    for group in row_groups.values():
        names = group["names"]
        excluded = [name for name in names if name not in included_bookmarks]
        if not excluded:
            continue
        if len(excluded) == len(names):
            current_ranges = _bookmark_ranges(document)
            for name in names:
                target = current_ranges.get(name)
                if target:
                    _relocate_empty_bookmark(document, target[0], target[1])
            row = group["row"]
            parent = row.getparent()
            if parent is not None:
                parent.remove(row)
        else:
            for name in excluded:
                _replace_bookmark_text(document, name, "")


def _section_contains_ref(paragraphs: list[Paragraph], bookmark_name: str) -> bool:
    pattern = re.compile(rf"\bREF\s+{re.escape(bookmark_name)}\b", re.IGNORECASE)
    for paragraph in paragraphs:
        for instruction in paragraph._p.iter(qn("w:instrText")):
            if pattern.search(instruction.text or ""):
                return True
    return False


def _insert_after(paragraph: Paragraph, title: str, text: str) -> Paragraph:
    new_element = OxmlElement("w:p")
    paragraph._p.addnext(new_element)
    new_paragraph = Paragraph(new_element, paragraph._parent)
    if paragraph.style:
        new_paragraph.style = paragraph.style
    heading = new_paragraph.add_run(f"{title}.  ")
    heading.bold = True
    new_paragraph.add_run(text.strip())
    return new_paragraph


def _set_update_fields(document: Document) -> None:
    settings = document.settings._element
    existing = settings.find(qn("w:updateFields"))
    if existing is None:
        existing = OxmlElement("w:updateFields")
        settings.append(existing)
    existing.set(qn("w:val"), "true")


def _clean_drafting_formatting(document: Document) -> None:
    """Accept revisions and remove template drafting markup from every Word story."""
    roots = []
    seen = set()

    # Main document plus headers/footers, comments and other OOXML story parts.
    for part in document.part.package.parts:
        root = getattr(part, "_element", None)
        if root is None:
            root = getattr(part, "element", None)
        if root is not None and id(root) not in seen:
            roots.append(root)
            seen.add(id(root))

    for root in roots:
        # Accept tracked insertions by unwrapping their contents into the document.
        for insertion in list(root.iter(qn("w:ins"))):
            parent = insertion.getparent()
            if parent is None:
                continue
            position = parent.index(insertion)
            for child in list(insertion):
                insertion.remove(child)
                parent.insert(position, child)
                position += 1
            parent.remove(insertion)

        # Reject deletions and remove comment anchors/references.
        remove_tags = (qn("w:del"), qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference"))
        for tag in remove_tags:
            for element in list(root.iter(tag)):
                parent = element.getparent()
                if parent is not None:
                    parent.remove(element)

        # The source template uses red/blue/highlight/shading as drafting guidance.
        # A built lease is a clean document, so retain typography/layout but strip markup.
        for run_properties in root.iter(qn("w:rPr")):
            for tag in (qn("w:highlight"), qn("w:color"), qn("w:shd")):
                for element in list(run_properties.findall(tag)):
                    run_properties.remove(element)
        for tag in (qn("w:shd"), qn("w:highlight")):
            for element in list(root.iter(tag)):
                parent = element.getparent()
                if parent is not None:
                    parent.remove(element)

    settings = document.settings._element
    for tag in (qn("w:trackRevisions"), qn("w:showInsDel"), qn("w:showMarkup")):
        for element in list(settings.findall(tag)):
            settings.remove(element)


def _remove_template_markers(document: Document) -> None:
    """Remove Word-template bookmark labels such as (#Property) from output text."""
    marker_in_parentheses = re.compile(r"\s*\(\s*#[A-Za-z_][A-Za-z0-9_]*\s*\)")
    bare_marker = re.compile(r"\s+#[A-Za-z_][A-Za-z0-9_]*")
    roots = []
    seen = set()
    for part in document.part.package.parts:
        root = getattr(part, "_element", None)
        if root is None:
            root = getattr(part, "element", None)
        if root is not None and id(root) not in seen:
            roots.append(root)
            seen.add(id(root))
    for root in roots:
        # Hyperlinked display markers are split across separate runs: "(" + #Property + ")".
        for hyperlink in list(root.iter(qn("w:hyperlink"))):
            hyperlink_text = "".join(
                item.text or "" for item in hyperlink.iter(qn("w:t"))
            ).strip()
            if not re.fullmatch(r"#[A-Za-z_][A-Za-z0-9_]*", hyperlink_text):
                continue
            parent = hyperlink.getparent()
            if parent is None:
                continue
            index = parent.index(hyperlink)
            neighbors = []
            if index > 0:
                neighbors.append(parent[index - 1])
            if index + 1 < len(parent):
                neighbors.append(parent[index + 1])
            parent.remove(hyperlink)
            for neighbor in neighbors:
                neighbor_text = "".join(item.text or "" for item in neighbor.iter(qn("w:t"))).strip()
                if neighbor_text in {"(", ")"} and neighbor.getparent() is parent:
                    parent.remove(neighbor)
        for text_element in root.iter(qn("w:t")):
            original = text_element.text or ""
            cleaned = bare_marker.sub("", marker_in_parentheses.sub("", original))
            if re.fullmatch(r"#[A-Za-z_][A-Za-z0-9_]*", cleaned.strip()):
                cleaned = ""
            if cleaned != original:
                text_element.text = cleaned
                if cleaned.startswith(" ") or cleaned.endswith(" "):
                    text_element.set(qn("xml:space"), "preserve")


def _clean_drafting_notes(document: Document) -> None:
    """Remove template directions and scratch clauses before the formal lease body."""
    paragraphs = list(document.paragraphs)
    body_start = next(
        (i for i, paragraph in enumerate(paragraphs)
         if paragraph.text.strip().startswith("This LEASE AGREEMENT")),
        None,
    )
    if body_start is not None:
        for paragraph in reversed(paragraphs[2:body_start]):
            _delete_paragraph(paragraph)

    scratch_markers = ("DELETE ABOVE", "LEAVE FOR TEMPLATE", "Replace Sec 1 if full guarantee")
    for paragraph in list(document.paragraphs):
        if any(marker.lower() in paragraph.text.lower() for marker in scratch_markers):
            _delete_paragraph(paragraph)


def build_word_document(
    template_path: str | Path,
    section_choices: dict[str, dict[str, Any]],
    bookmark_values: dict[str, str] | None = None,
    included_bookmarks: set[str] | list[str] | None = None,
    linked_provisions: list[dict[str, Any]] | None = None,
    additional_choices: list[dict[str, Any]] | None = None,
    clean_drafting_notes: bool = True,
    document_title: str = "MSP Lease Draft",
) -> bytes:
    """Build a redline-ready DOCX and return its bytes."""
    document = Document(str(template_path))
    sections = scan_sections(document)
    original_paragraphs = list(document.paragraphs)

    for section in reversed(sections):
        choice = section_choices.get(section["number"], {})
        clause_paragraphs = original_paragraphs[section["start"]:section["end"]]
        if not bool(choice.get("include", True)):
            for paragraph in reversed(clause_paragraphs):
                _delete_paragraph(paragraph)
            continue
        replacement = str(choice.get("replacement_text") or "").strip()
        if replacement:
            first = clause_paragraphs[0]
            title = str(choice.get("title") or section["title"]).strip()
            first.text = f"Section {section['number']}.  {title}.  {replacement}"
            for paragraph in reversed(clause_paragraphs[1:]):
                _delete_paragraph(paragraph)

    if bookmark_values:
        values = dict(bookmark_values)
        if "Tx_Landlord" in values:
            values.setdefault("Tx_Landlord1", values["Tx_Landlord"])
        for bookmark_name, value in values.items():
            _replace_bookmark_text(document, bookmark_name, str(value))

    if included_bookmarks is not None:
        _apply_key_provision_inclusions(document, set(included_bookmarks))

    current_sections = {section["number"]: section for section in scan_sections(document)}
    current_paragraphs = list(document.paragraphs)

    if linked_provisions:
        grouped_links: dict[str, list[dict[str, Any]]] = {}
        for item in linked_provisions:
            section_number = str(item.get("section", ""))
            if item.get("include") and section_number and str(item.get("value", "")).strip():
                grouped_links.setdefault(section_number, []).append(item)
        for section_number, items in grouped_links.items():
            anchor_section = current_sections.get(section_number)
            if not anchor_section:
                continue
            section_paragraphs = current_paragraphs[anchor_section["start"]:anchor_section["end"]]
            anchor = section_paragraphs[-1]
            for item in items:
                bookmark_name = str(item.get("bookmark", ""))
                if bookmark_name and _section_contains_ref(section_paragraphs, bookmark_name):
                    continue
                anchor = _insert_after(anchor, str(item.get("field", "Key Provision")), str(item["value"]))

    if additional_choices:
        for item in additional_choices:
            if not item.get("include") or not str(item.get("text", "")).strip():
                continue
            anchor_section = current_sections.get(str(item.get("insert_after", "11")))
            if not anchor_section:
                continue
            anchor = current_paragraphs[anchor_section["end"] - 1]
            _insert_after(anchor, str(item.get("title", "Additional Provision")), str(item["text"]))

    _remove_template_markers(document)

    if clean_drafting_notes:
        _clean_drafting_notes(document)
        _clean_drafting_formatting(document)

    document.core_properties.title = document_title
    document.core_properties.subject = "Generated by MSP Lease Builder"
    _set_update_fields(document)
    output = BytesIO()
    document.save(output)
    return output.getvalue()
