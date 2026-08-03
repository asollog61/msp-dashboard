"""One kind of saved thing: a lease document.

The Lease Builder used to keep two stores. "Templates" authored a menu of
approved clause language; "leases" consumed that menu, storing only their
selections and re-hydrating the alternates from a parent template on load. The
split bought inheritance, but it made every visit to the tab start with a
question — am I editing the menu or a deal? — and the answer was rarely obvious.

Now there is one list. A document called "Triple Net Template" and one called
"ABC Bakery Lease" are the same kind of object; the only difference is what you
left checked. Each carries its own provisions, its own alternates, its own
section choices and its own format profile, so nothing depends on a parent
still existing. A new document is made by copying an existing one.

This module is pure data: migration, normalization, and copying. No Streamlit.
"""

from __future__ import annotations

import copy
import re
from datetime import datetime
from typing import Any

DOC_VERSION = 1

ALTERNATE_SLOTS = 10

# A provision's Value is what prints in the lease. Choice records where that
# text came from — one of the Alt slots, or nothing, meaning the value was
# typed directly. Keeping Value authoritative means nothing downstream has to
# understand choices; keeping Choice alongside it is what makes "which leases
# use this clause" answerable.
#
# The chooser offers the Alt slots only. Blank is how a provision says "no
# alternate — this value was typed", and is the state every provision starts
# in. Earlier documents wrote the literal string "Current Value" for that, so
# it is still accepted on read and normalised away.
NO_CHOICE = ""
CURRENT_VALUE_CHOICE = "Current Value"  # Legacy spelling of NO_CHOICE.
_ALT_CHOICE_RE = re.compile(r"^\s*Alt\s*(\d+)\s*$", re.IGNORECASE)


def choice_options(row: Any) -> list[str]:
    """The choices offered for one provision: blank, then each filled alternate.

    Blank slots are not offered. Selecting "Alt 7" when Alt 7 is empty would
    silently blank a key provision in the lease.
    """
    source = row if isinstance(row, dict) else {}
    alternates = _pad_alternates(source.get("Alternates"))
    return [NO_CHOICE] + [
        f"Alt {index}"
        for index, value in enumerate(alternates, start=1)
        if str(value).strip()
    ]


def normalize_choice(row: Any) -> str:
    """A choice that still points at real text, or the default.

    An alternate that gets blanked after being chosen falls back here rather
    than leaving the provision pointing at nothing.
    """
    source = row if isinstance(row, dict) else {}
    raw = str(source.get("Choice", "") or "").strip()
    # "alt 2", "ALT2" and "Alt 2" are the same choice. Excel is a free-text
    # field even with a dropdown on it, so a case difference must not quietly
    # demote the row back to no choice at all.
    match = _ALT_CHOICE_RE.match(raw)
    canonical = f"Alt {int(match.group(1))}" if match else raw
    if canonical.casefold() == CURRENT_VALUE_CHOICE.casefold():
        canonical = NO_CHOICE
    return canonical if canonical in choice_options(source) else NO_CHOICE


def apply_choice(row: Any) -> dict[str, Any]:
    """Materialise the chosen alternate into Value.

    Re-applied on every edit, so changing the text of a chosen alternate keeps
    Value in step instead of leaving the lease printing a stale copy.
    """
    resolved = dict(row) if isinstance(row, dict) else {}
    choice = normalize_choice(resolved)
    resolved["Choice"] = choice
    match = _ALT_CHOICE_RE.match(choice)
    if match:
        alternates = _pad_alternates(resolved.get("Alternates"))
        index = int(match.group(1)) - 1
        if 0 <= index < len(alternates) and str(alternates[index]).strip():
            resolved["Value"] = str(alternates[index])
    else:
        # No choice means no value. Every provision's text lives in an Alt
        # slot; Value only ever mirrors the chosen one. A provision showing
        # text with nothing selected would be text no alternate accounts for,
        # and it would survive into the lease unnoticed.
        resolved["Value"] = ""
    return resolved


def adopt_legacy_value(row: Any) -> dict[str, Any]:
    """Move a typed value into Alt 1 for rows written before Choice existed.

    Those rows carry their text in Value with no Choice and no alternates.
    Blanking them to match the chooser would delete real lease language, so
    the text is promoted into Alt 1 and selected — the same move as typing it
    into the Alt 1 column by hand.
    """
    source = dict(row) if isinstance(row, dict) else {}
    value = str(source.get("Value", "") or "").strip()
    alternates = _pad_alternates(source.get("Alternates"))
    already_chosen = bool(_ALT_CHOICE_RE.match(str(source.get("Choice", "") or "").strip()))
    if not value or already_chosen or any(str(slot).strip() for slot in alternates):
        return source
    alternates[0] = str(source.get("Value", ""))
    source["Alternates"] = alternates
    source["Choice"] = "Alt 1"
    return source


def _rows(value: Any) -> list[dict[str, Any]]:
    return [dict(row) for row in value if isinstance(row, dict)] if isinstance(value, list) else []


def _pad_alternates(value: Any) -> list[str]:
    slots = [str(item) for item in value[:ALTERNATE_SLOTS]] if isinstance(value, (list, tuple)) else []
    return slots + [""] * (ALTERNATE_SLOTS - len(slots))


def normalize_provision(row: dict[str, Any]) -> dict[str, Any]:
    """One key-provision row with every field present and the right type."""
    row = adopt_legacy_value(row)
    return {
        "Group": "Mandatory" if str(row.get("Group", "")).strip() == "Mandatory" else "Optional",
        "Include": bool(row.get("Include", True)),
        "Field": str(row.get("Field", "") or "Key Provision"),
        "Value": str(row.get("Value", "") or ""),
        "Alternates": _pad_alternates(row.get("Alternates")),
        "Choice": normalize_choice(row),
        "Link": bool(row.get("Link", False)),
        "Section": str(row.get("Section", "") or ""),
        "Bookmark": str(row.get("Bookmark", "") or ""),
    }


def normalize_section(config: Any) -> dict[str, Any]:
    source = config if isinstance(config, dict) else {}
    entry = {
        "include": bool(source.get("include", True)),
        "choice": str(source.get("choice", "") or "Template language"),
    }
    # Text is only stored when it departs from the source language; an empty
    # string and a missing key mean the same thing and must stay distinguishable
    # from "deliberately blank clause".
    if "text" in source and str(source.get("text", "")).strip():
        entry["text"] = str(source["text"])
    return entry


def normalize_document(raw: Any) -> dict[str, Any]:
    """Coerce anything off the sheet into a complete document."""
    source = raw if isinstance(raw, dict) else {}
    sections = source.get("sections")
    return {
        "version": DOC_VERSION,
        "key_provisions": [normalize_provision(row) for row in _rows(source.get("key_provisions"))],
        "sections": {
            str(number): normalize_section(config)
            for number, config in (sections.items() if isinstance(sections, dict) else [])
        },
        "rent_schedules": source.get("rent_schedules") if isinstance(source.get("rent_schedules"), dict) else {},
        "format_profile": str(source.get("format_profile", "") or ""),
        # Which space in the tenancy workbook this lease covers. Stored as a
        # key rather than resolved values so reopening the document re-reads
        # the workbook and picks up corrections.
        "space_key": str(source.get("space_key", "") or ""),
        "copied_from": str(source.get("copied_from", "") or ""),
        "saved_at": str(source.get("saved_at", "") or ""),
    }


def normalize_store(raw: Any) -> dict[str, dict[str, Any]]:
    source = raw if isinstance(raw, dict) else {}
    return {
        str(name).strip(): normalize_document(document)
        for name, document in source.items()
        if str(name).strip()
    }


# ---------------------------------------------------------------------------
# Migration
# ---------------------------------------------------------------------------

def _hydrate_lease(lease: dict[str, Any], parent: dict[str, Any]) -> dict[str, Any]:
    """Bake a lease's parent into it, so it can stand alone.

    A saved lease stored only Bookmark/Include/Value and leaned on its template
    for the alternates and the section language. Those have to be folded in now
    or the alternates would simply vanish at migration.
    """
    parent_rows = _rows(parent.get("key_provisions"))
    by_bookmark = {str(row.get("Bookmark", "")): row for row in _rows(lease.get("key_provisions"))}

    merged: list[dict[str, Any]] = []
    for source in parent_rows:
        row = dict(source)
        saved = by_bookmark.pop(str(row.get("Bookmark", "")), None)
        if saved is not None:
            row["Include"] = bool(saved.get("Include", row.get("Include", True)))
            row["Value"] = str(saved.get("Value", row.get("Value", "")))
            if saved.get("Choice"):
                row["Choice"] = saved["Choice"]
        merged.append(row)
    # A provision added on the lease itself has no parent row; keep it.
    for leftover in by_bookmark.values():
        merged.append(leftover)

    sections = dict(parent.get("sections") or {})
    for number, config in (lease.get("sections") or {}).items():
        entry = dict(sections.get(str(number), {}))
        entry.update(config if isinstance(config, dict) else {})
        sections[str(number)] = entry

    return {
        "key_provisions": merged,
        "sections": sections,
        "rent_schedules": lease.get("rent_schedules") or {},
        "format_profile": parent.get("format_profile", ""),
        "copied_from": str(lease.get("template_name", "") or ""),
        "saved_at": lease.get("saved_at", ""),
    }


def unique_name(name: str, taken: Any) -> str:
    """A name not already in use, suffixed rather than silently overwriting."""
    label = str(name).strip() or "Untitled"
    if label not in taken:
        return label
    index = 2
    while f"{label} ({index})" in taken:
        index += 1
    return f"{label} ({index})"


def migrate_stores(templates: Any, leases: Any) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Fold the old template and lease stores into one list of documents.

    Returns (documents, notes). Templates come across as they are; leases are
    hydrated from their parent first. Names collide only if a template and a
    lease shared one, in which case the lease is suffixed — losing a saved deal
    to a name clash would be the worst possible migration outcome.
    """
    template_source = templates if isinstance(templates, dict) else {}
    lease_source = leases if isinstance(leases, dict) else {}
    documents: dict[str, dict[str, Any]] = {}
    notes: list[str] = []

    for name, payload in template_source.items():
        label = str(name).strip()
        if not label:
            continue
        documents[label] = normalize_document(payload)

    for name, payload in lease_source.items():
        label = str(name).strip()
        if not label:
            continue
        parent_name = str((payload or {}).get("template_name", "") or "")
        parent = template_source.get(parent_name) or {}
        if parent_name and not parent:
            notes.append(
                f"“{label}” referenced the template “{parent_name}”, which no longer "
                "exists — its own selections were kept, but no alternates could be recovered."
            )
        hydrated = _hydrate_lease(payload or {}, parent)
        final = unique_name(label, documents)
        if final != label:
            notes.append(f"“{label}” was saved as “{final}” — a template already had that name.")
        documents[final] = normalize_document(hydrated)

    return documents, notes


# ---------------------------------------------------------------------------
# Creating and saving
# ---------------------------------------------------------------------------

def copy_document(document: Any, source_name: str = "") -> dict[str, Any]:
    """A deep, independent copy. Editing the copy must never touch the original."""
    duplicate = normalize_document(copy.deepcopy(document))
    duplicate["copied_from"] = str(source_name or "")
    duplicate["saved_at"] = ""
    return duplicate


def build_document(key_provisions: Any, sections: Any, rent_schedules: Any,
                   format_profile: str, copied_from: str = "") -> dict[str, Any]:
    """The payload written to the sheet when you press Save."""
    return normalize_document({
        "key_provisions": key_provisions,
        "sections": sections,
        "rent_schedules": rent_schedules,
        "format_profile": format_profile,
        "copied_from": copied_from,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    })


def describe_document(document: Any) -> str:
    """One-line summary for the picker and the header."""
    resolved = normalize_document(document)
    provisions = resolved["key_provisions"]
    sections = resolved["sections"]
    used_provisions = sum(1 for row in provisions if row["Include"])
    used_sections = sum(1 for config in sections.values() if config["include"])
    with_alternates = sum(1 for row in provisions if any(row["Alternates"]))
    parts = [
        f"{used_provisions}/{len(provisions)} provisions",
        f"{used_sections}/{len(sections)} sections",
    ]
    if with_alternates:
        parts.append(f"{with_alternates} with alternates")
    if resolved["saved_at"]:
        parts.append(f"saved {resolved['saved_at'][:10]}")
    return " · ".join(parts)
