"""Formatting settings for template-free lease generation.

The renderer builds a lease from rules rather than from a base .docx, so every
page-setup and typography decision has to live somewhere explicit. That is this
module: a flat, JSON-serializable settings dict, its defaults (taken from
LEASE_FORMAT_SPEC.md, itself derived from the executed Chez Alice lease), and a
normalizer that turns anything loaded off a Google Sheet back into a usable
dict.

Deliberately pure — no Streamlit, no python-docx — so the settings model can be
unit-tested on its own and imported by both the form and the renderer.
"""

from __future__ import annotations

from typing import Any

SETTINGS_VERSION = 1

# Choices offered in the UI. Kept here so the form and the renderer cannot drift.
PAGE_SIZES = {
    # name: (width_in, height_in)
    "Letter": (8.5, 11.0),
    "Legal": (8.5, 14.0),
    "A4": (8.27, 11.69),
}
BODY_FONTS = [
    "Times New Roman",
    "Georgia",
    "Garamond",
    "Cambria",
    "Book Antiqua",
    "Calibri",
    "Arial",
]
ALIGNMENTS = ["justify", "left"]
PAGE_NUMBER_POSITIONS = ["right", "center", "left", "none"]
SUBCLAUSE_LEVEL1_STYLES = ["A.", "(A)", "(a)", "1."]
SUBCLAUSE_LEVEL2_STYLES = ["(i)", "(a)", "(1)"]


DEFAULTS: dict[str, Any] = {
    "version": SETTINGS_VERSION,

    # ---- Page setup ------------------------------------------------------
    # Spec: US Letter 8.5 x 11, 1 in margins all sides.
    "page_size": "Letter",
    "margin_top_in": 1.0,
    "margin_bottom_in": 1.0,
    "margin_left_in": 1.0,
    "margin_right_in": 1.0,

    # ---- Body typography -------------------------------------------------
    # Spec: serif body, 11 pt, justified.
    "body_font": "Times New Roman",
    "body_size_pt": 11.0,
    "body_alignment": "justify",
    "body_line_spacing": 1.0,
    "space_after_pt": 6.0,
    "first_line_indent_in": 0.0,

    # ---- Footer ----------------------------------------------------------
    # Spec: counsel document ID left, page number right. Doc ID is per-template.
    "footer_doc_id": "",
    "footer_font_size_pt": 9.0,
    "page_number_position": "right",

    # ---- Title block -----------------------------------------------------
    # Spec: LEASE AGREEMENT then KEY PROVISIONS SUMMARY, centered/bold/underlined.
    "title_text": "LEASE AGREEMENT",
    "key_provisions_title": "KEY PROVISIONS SUMMARY",
    "title_size_pt": 12.0,
    "title_bold": True,
    "title_underline": True,

    # ---- Key Provisions table -------------------------------------------
    # Bordered; Notice Addresses splits the value into Landlord | Tenant, which
    # is why the underlying table is 3 columns with most rows merged across
    # cols 2-3.
    #
    # The spec's "~1.9 / ~5.1" is approximate and sums to 7.0 in, which overruns
    # the 6.5 in printable area. These are the widths measured off the master
    # template (1.706 / 2.346 / 2.442 = 6.494 in), rounded to hundredths.
    "kp_label_width_in": 1.71,
    "kp_value_width_in": 4.79,
    "kp_split_left_width_in": 2.35,   # Landlord half of a split row
    "kp_split_right_width_in": 2.44,  # Tenant half
    "kp_borders": True,
    "kp_label_bold": True,
    "kp_cell_padding_pt": 3.0,
    "kp_split_fields": ["Notice Addresses"],
    "kp_link_color": "1F7A33",       # green, underlined cross-reference
    "kp_link_underline": True,

    # ---- Section headings ------------------------------------------------
    # Spec: run-in bold heading, body continues on the same line, tab after the
    # number. Decimal sections are peers, not children.
    "section_word": "Section",
    "section_heading_bold": True,
    "section_tab_stop_in": 0.5,
    "section_space_before_pt": 10.0,

    # ---- Sub-clauses -----------------------------------------------------
    # Spec: two levels only. Level 1 lettered with a 0.25 in hanging indent;
    # level 2 roman, inline within the paragraph.
    "subclause_level1_style": "A.",
    "subclause_level1_indent_in": 0.25,
    "subclause_level1_hanging_in": 0.25,
    "subclause_level2_style": "(i)",
    "subclause_lead_in_bold": True,

    # ---- Rent table ------------------------------------------------------
    # Spec: bold "Base Rent Table:" label, 3 bordered columns, bold header row,
    # right-aligned currency. Appears in the KP value and in the rent section.
    "rent_table_label": "Base Rent Table:",
    "rent_col_term_width_in": 2.5,
    "rent_col_monthly_width_in": 2.0,
    "rent_col_annual_width_in": 2.0,
    "rent_header_bold": True,
    "rent_borders": True,
    "rent_amount_alignment": "right",

    # ---- Signatures ------------------------------------------------------
    "signature_title": "SIGNATURES",
    "signature_landlord_label": "(LANDLORD)",
    "signature_tenant_label": "(TENANT)",
    "signature_keep_together": True,

    # ---- Exhibits --------------------------------------------------------
    # Spec: page break before each; centered bold "Exhibit A" then a centered
    # bold subtitle.
    "exhibit_page_break": True,
    "exhibit_label_format": '"Exhibit {letter}"',
    "exhibit_title_size_pt": 12.0,
    "exhibit_image_max_width_in": 6.5,
}


# Bounds that keep a hand-typed or sheet-round-tripped value from producing a
# document Word refuses to open. (min, max)
_NUMERIC_BOUNDS: dict[str, tuple[float, float]] = {
    "margin_top_in": (0.25, 3.0),
    "margin_bottom_in": (0.25, 3.0),
    "margin_left_in": (0.25, 3.0),
    "margin_right_in": (0.25, 3.0),
    "body_size_pt": (6.0, 24.0),
    "body_line_spacing": (0.8, 3.0),
    "space_after_pt": (0.0, 36.0),
    "first_line_indent_in": (0.0, 1.5),
    "footer_font_size_pt": (6.0, 14.0),
    "title_size_pt": (8.0, 28.0),
    "kp_label_width_in": (0.75, 4.0),
    "kp_value_width_in": (1.5, 8.0),
    "kp_split_left_width_in": (0.75, 6.0),
    "kp_split_right_width_in": (0.75, 6.0),
    "kp_cell_padding_pt": (0.0, 12.0),
    "section_tab_stop_in": (0.0, 2.0),
    "section_space_before_pt": (0.0, 36.0),
    "subclause_level1_indent_in": (0.0, 2.0),
    "subclause_level1_hanging_in": (0.0, 2.0),
    "rent_col_term_width_in": (0.75, 4.0),
    "rent_col_monthly_width_in": (0.75, 4.0),
    "rent_col_annual_width_in": (0.75, 4.0),
    "exhibit_title_size_pt": (8.0, 28.0),
    "exhibit_image_max_width_in": (1.0, 10.0),
}

_CHOICES: dict[str, list[str]] = {
    "page_size": list(PAGE_SIZES),
    "body_font": BODY_FONTS,
    "body_alignment": ALIGNMENTS,
    "page_number_position": PAGE_NUMBER_POSITIONS,
    "subclause_level1_style": SUBCLAUSE_LEVEL1_STYLES,
    "subclause_level2_style": SUBCLAUSE_LEVEL2_STYLES,
    "rent_amount_alignment": ["right", "left", "center"],
}

_BOOL_KEYS = [key for key, value in DEFAULTS.items() if isinstance(value, bool)]
_LIST_KEYS = ["kp_split_fields"]


def _as_float(value: Any, fallback: float) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return fallback
    if number != number or number in (float("inf"), float("-inf")):
        return fallback
    return number


def _as_bool(value: Any, fallback: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        token = value.strip().lower()
        if token in {"true", "yes", "1", "on"}:
            return True
        if token in {"false", "no", "0", "off", ""}:
            return False
    return fallback


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def normalize_hex_color(value: Any, fallback: str = "1F7A33") -> str:
    """Accept '#1f7a33', '1F7A33', or junk; always return six upper-case hex digits."""
    token = str(value or "").strip().lstrip("#").upper()
    if len(token) == 3 and all(char in "0123456789ABCDEF" for char in token):
        token = "".join(char * 2 for char in token)
    if len(token) == 6 and all(char in "0123456789ABCDEF" for char in token):
        return token
    return fallback


def default_settings() -> dict[str, Any]:
    """A fresh copy of the spec defaults; safe for the caller to mutate."""
    settings = dict(DEFAULTS)
    settings["kp_split_fields"] = list(DEFAULTS["kp_split_fields"])
    return settings


def normalize_settings(raw: Any) -> dict[str, Any]:
    """Coerce anything (None, partial dict, sheet round-trip with string values)
    into a complete, in-range settings dict. Unknown keys are dropped so a stale
    saved template cannot smuggle junk into the renderer."""
    source = raw if isinstance(raw, dict) else {}
    settings = default_settings()

    for key, fallback in DEFAULTS.items():
        if key not in source:
            continue
        value = source[key]
        if key in _BOOL_KEYS:
            settings[key] = _as_bool(value, bool(fallback))
        elif key in _LIST_KEYS:
            if isinstance(value, str):
                items = [part.strip() for part in value.split(",")]
            elif isinstance(value, (list, tuple)):
                items = [str(part).strip() for part in value]
            else:
                items = list(fallback)
            settings[key] = [item for item in items if item]
        elif key in _CHOICES:
            settings[key] = value if value in _CHOICES[key] else fallback
        elif isinstance(fallback, float):
            low, high = _NUMERIC_BOUNDS.get(key, (float("-inf"), float("inf")))
            settings[key] = _clamp(_as_float(value, fallback), low, high)
        elif isinstance(fallback, int):
            settings[key] = int(_as_float(value, fallback))
        else:
            settings[key] = str(value)

    settings["kp_link_color"] = normalize_hex_color(
        source.get("kp_link_color", DEFAULTS["kp_link_color"]), DEFAULTS["kp_link_color"]
    )
    settings["version"] = SETTINGS_VERSION
    return settings


def page_dimensions(settings: dict[str, Any]) -> tuple[float, float]:
    """(width_in, height_in) for the configured page size."""
    return PAGE_SIZES.get(str(settings.get("page_size")), PAGE_SIZES["Letter"])


def content_width_in(settings: dict[str, Any]) -> float:
    """Printable width — the budget the KP and rent tables have to fit inside."""
    width, _ = page_dimensions(settings)
    return round(
        width
        - _as_float(settings.get("margin_left_in"), 1.0)
        - _as_float(settings.get("margin_right_in"), 1.0),
        4,
    )


def validate_settings(settings: dict[str, Any]) -> list[str]:
    """Human-readable warnings for a normalized dict.

    Normalization already guarantees the document will open; these are the
    remaining problems a person should see and decide about, chiefly tables
    that overrun the margins.
    """
    warnings: list[str] = []
    available = content_width_in(settings)

    kp_total = _as_float(settings.get("kp_label_width_in"), 0) + _as_float(
        settings.get("kp_value_width_in"), 0
    )
    if kp_total > available + 0.01:
        warnings.append(
            f"Key Provisions table is {kp_total:.2f} in wide but only "
            f"{available:.2f} in fits between the margins."
        )

    split_total = _as_float(settings.get("kp_split_left_width_in"), 0) + _as_float(
        settings.get("kp_split_right_width_in"), 0
    )
    if abs(split_total - _as_float(settings.get("kp_value_width_in"), 0)) > 0.02:
        warnings.append(
            f"Split Landlord|Tenant halves total {split_total:.2f} in but the value "
            f"column is {_as_float(settings.get('kp_value_width_in'), 0):.2f} in; "
            "the Notice Addresses row will not line up with the rows above it."
        )

    rent_total = sum(
        _as_float(settings.get(key), 0)
        for key in ("rent_col_term_width_in", "rent_col_monthly_width_in", "rent_col_annual_width_in")
    )
    if rent_total > available + 0.01:
        warnings.append(
            f"Rent table is {rent_total:.2f} in wide but only {available:.2f} in "
            "fits between the margins."
        )

    if _as_float(settings.get("exhibit_image_max_width_in"), 0) > available + 0.01:
        warnings.append(
            f"Exhibit image width exceeds the {available:.2f} in printable area; "
            "images will be scaled down."
        )

    if settings.get("subclause_level1_hanging_in", 0) > settings.get(
        "subclause_level1_indent_in", 0
    ):
        warnings.append(
            "Sub-clause hanging indent is larger than its left indent, so the "
            "letter will sit in the left margin."
        )

    if not str(settings.get("footer_doc_id", "")).strip():
        warnings.append("No counsel document ID set — the footer will show only the page number.")

    return warnings


def settings_diff(settings: dict[str, Any]) -> dict[str, Any]:
    """Only the values that depart from the spec defaults. Useful for a compact
    save payload and for showing 'what did I change'."""
    normalized = normalize_settings(settings)
    return {
        key: value
        for key, value in normalized.items()
        if key != "version" and value != DEFAULTS[key]
    }
