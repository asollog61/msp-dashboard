"""Space data from the MSP Tenancy workbook, and the [Space:...] tokens that reference it.

A lease repeats the same handful of facts about the space it covers — building,
unit, tenant, square footage, the three shares, lease type. Those already live
in the tenancy workbook, and retyping them into every lease is how they drift.

So a key provision holds the sentence and a token holds the fact:

    Approximately [Space:Sqft] Total Gross Square Feet
    Tenant's Share of the Property being [Space:ShareOfProperty]

Picking a space in the Lease Builder resolves every token. The tokens stay in
the saved document — only the rendered output is substituted — so the same
template produces a correct lease for any space, and changing the space changes
every number at once.

This mirrors the [KP:Name] tokens already used inside clause text, with one
difference: KP tokens point at other provisions in the same document, Space
tokens point out at the tenancy workbook.

Pure data. No Streamlit, no file IO — the caller supplies the summary rows.
"""

from __future__ import annotations

import re
from typing import Any

# Deliberately the same bracket shape as KP_TOKEN_RE in lease_builder, so the
# two kinds of token look alike in a clause and neither can swallow the other.
SPACE_TOKEN_RE = re.compile(r"\[\s*Space\s*:\s*([^\]]+?)\s*\]", re.IGNORECASE)

# The workbook header really is spelled "Buidling". Correcting it there would
# break every other reader, so the typo is absorbed here instead.
_BUILDING_COLUMN = "Buidling"


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _number(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number


def _format_sqft(value: Any) -> str:
    number = _number(value)
    if number is None:
        return ""
    # The workbook carries placeholder rows as 1e-07 rather than blank, and a
    # lease that says "approximately 0 square feet" is worse than one that
    # says nothing.
    if number < 1:
        return ""
    return f"{round(number):,}"


def _format_percent(value: Any) -> str:
    number = _number(value)
    if number is None:
        return ""
    # The workbook stores shares as fractions (0.2172), except where someone
    # has typed a whole number. Anything above 1 is already a percentage.
    percent = number * 100 if number <= 1 else number
    if percent <= 0:
        return ""
    text = f"{percent:.2f}".rstrip("0").rstrip(".")
    return f"{text}%"


# Token name -> (label for the UI, how to read it off a summary row).
# Order is the order they appear in the reference panel.
FIELDS: dict[str, tuple[str, Any]] = {
    "Building":        ("Building",           lambda row: _text(row.get(_BUILDING_COLUMN))),
    "Unit":            ("Unit / Space",       lambda row: _text(row.get("Unit"))),
    "Tenant":          ("Tenant",             lambda row: _text(row.get("Tenant Name"))),
    "Sqft":            ("Square Feet",        lambda row: _format_sqft(row.get("New Sqft"))),
    "ShareOfProperty": ("Share of Property",  lambda row: _format_percent(row.get("_building_share"))),
    "ShareOfFloor":    ("Share of Floor",     lambda row: _format_percent(row.get("Floor %"))),
    "ShareOfPremises": ("Share of Premises",  lambda row: _format_percent(row.get("Category %"))),
    "LeaseType":       ("Lease Type",         lambda row: _text(row.get("Lease Type"))),
    "Floor":           ("Floor",              lambda row: _text(row.get("Floor"))),
    "UseType":         ("Use Type",           lambda row: _text(row.get("Type"))),
}

# Aliases, so a token typed the way the field reads still resolves.
ALIASES = {
    "space": "Unit",
    "squarefeet": "Sqft",
    "sf": "Sqft",
    "shareofbuilding": "ShareOfProperty",
    "property": "Building",
    "use": "UseType",
    "type": "UseType",
}


def token_names() -> list[str]:
    return list(FIELDS)


def field_label(name: str) -> str:
    entry = FIELDS.get(canonical_name(name))
    return entry[0] if entry else str(name)


def canonical_name(name: str) -> str:
    """Map a token's spelling onto a field name, ignoring case and spacing."""
    key = re.sub(r"[\s_-]+", "", str(name or "")).casefold()
    for field in FIELDS:
        if field.casefold() == key:
            return field
    return ALIASES.get(key, "")


def space_records(summaries: Any) -> list[dict[str, str]]:
    """One record per leasable space, with every token field resolved.

    Share of Property is computed here rather than read, because the workbook
    has no such column: it is this space's square footage over the total for
    its building, which is the same basis the dashboard uses elsewhere.
    """
    rows = [row for row in (summaries or []) if isinstance(row, dict)]

    building_totals: dict[str, float] = {}
    for row in rows:
        building = _text(row.get(_BUILDING_COLUMN)).casefold()
        sqft = _number(row.get("New Sqft")) or 0.0
        if building and sqft >= 1:
            building_totals[building] = building_totals.get(building, 0.0) + sqft

    records = []
    for row in rows:
        building = _text(row.get(_BUILDING_COLUMN))
        if not building:
            continue
        total = building_totals.get(building.casefold(), 0.0)
        sqft = _number(row.get("New Sqft")) or 0.0
        enriched = dict(row)
        enriched["_building_share"] = (sqft / total) if (total > 0 and sqft >= 1) else None

        record = {name: reader(enriched) for name, (_, reader) in FIELDS.items()}
        unit = record["Unit"] or "—"
        tenant = record["Tenant"]
        # The picker is chosen by eye, so the label carries enough to tell two
        # units in the same building apart without opening anything.
        record["_label"] = f"{building} · {unit}" + (f" · {tenant}" if tenant else "")
        record["_key"] = f"{building}|{unit}|{tenant}".casefold()
        records.append(record)

    records.sort(key=lambda record: (record["Building"], record["Unit"]))
    return records


def find_space(records: Any, key: str) -> dict[str, str] | None:
    for record in records or []:
        if isinstance(record, dict) and record.get("_key") == key:
            return record
    return None


def resolve(text: Any, space: Any) -> str:
    """Substitute every [Space:Field] token in one string.

    An unknown field, or a field the workbook leaves blank, is left as the
    literal token. Silently emitting an empty string would put "Approximately
    Total Gross Square Feet" into a signed lease.
    """
    source = str(text or "")
    if not source or not isinstance(space, dict):
        return source

    def substitute(match: re.Match) -> str:
        field = canonical_name(match.group(1))
        if not field:
            return match.group(0)
        value = str(space.get(field, "") or "")
        return value if value else match.group(0)

    return SPACE_TOKEN_RE.sub(substitute, source)


def resolve_provisions(rows: Any, space: Any) -> list[dict[str, Any]]:
    """Every provision with its Value resolved. The stored rows are untouched."""
    resolved = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        copy = dict(row)
        copy["Value"] = resolve(copy.get("Value", ""), space)
        resolved.append(copy)
    return resolved


def unresolved(text: Any, space: Any = None) -> list[str]:
    """Token names in this text that would not resolve. Drives the warning."""
    missing = []
    for match in SPACE_TOKEN_RE.finditer(str(text or "")):
        field = canonical_name(match.group(1))
        if not field:
            missing.append(match.group(1))
        elif isinstance(space, dict) and not str(space.get(field, "") or ""):
            missing.append(match.group(1))
        elif space is None:
            missing.append(match.group(1))
    return missing
