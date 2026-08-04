"""Render the master as a document you can read, not a form you have to decode.

The Lease Builder used to present the lease as a grid of fields and a list of
section numbers. That is fine for editing one clause and useless for answering
"what does this lease actually say", which is the question being asked most of
the time.

This renders the extracted blocks back into something shaped like the lease —
numbered headings, run-in bold lead-ins, indented sub-clauses, the same reading
order as Word — with the decisions marked inline where they occur:

    a pick-one set   shows its options stacked, the selected one solid and the
                     rest dimmed, so the alternatives stay visible in context
    an optional block shows whether it is in or out
    an unresolved token stays visible rather than rendering as a blank

Every section gets an anchor and the index above links to it, so the document
is navigable rather than scrolled.

Produces HTML. No Streamlit, so it can be tested by reading the string.
"""

from __future__ import annotations

import html
import re
from typing import Any

import lease_blocks as lbk

# Word's own proportions, roughly: a serif face, generous line height, and a
# measure that does not run the full width of a dashboard.
STYLE = """
<style>
.lv-wrap { font-family: 'Times New Roman', Georgia, serif; font-size: 15px;
           line-height: 1.55; max-width: 52rem; }
.lv-index { font-family: system-ui, sans-serif; font-size: 12.5px;
            line-height: 1.9; margin-bottom: 1.5rem;
            border-left: 3px solid rgba(128,128,128,.35); padding-left: .9rem; }
.lv-index a { text-decoration: none; opacity: .85; }
.lv-index a:hover { text-decoration: underline; opacity: 1; }
.lv-index .lv-count { opacity: .55; font-size: 11.5px; }
.lv-sec { margin: 0 0 1.6rem 0; scroll-margin-top: 1rem; }
.lv-h { font-weight: 700; margin: 0 0 .45rem 0; }
.lv-p { margin: 0 0 .55rem 0; text-align: justify; }
.lv-lead { font-weight: 700; }
.lv-i1 { margin-left: 2.2rem; }
.lv-i2 { margin-left: 4.4rem; }
.lv-opt { border-left: 3px solid #2e9e5b; padding-left: .8rem; margin: .6rem 0; }
.lv-opt-off { border-left-color: rgba(128,128,128,.4); opacity: .45; }
.lv-set { border-left: 3px solid #1f8fa8; padding-left: .8rem; margin: .6rem 0; }
.lv-choice { margin: 0 0 .5rem 0; }
.lv-choice-off { opacity: .38; }
.lv-tag { font-family: system-ui, sans-serif; font-size: 10.5px;
          letter-spacing: .04em; text-transform: uppercase; opacity: .7;
          display: block; margin-bottom: .2rem; }
.lv-token { color: #d13438; font-weight: 600; }
.lv-kp { width: 100%; border-collapse: collapse; margin-bottom: 1.6rem;
         font-size: 14px; }
.lv-kp td { border: 1px solid rgba(128,128,128,.35); padding: .35rem .55rem;
            vertical-align: top; }
.lv-kp td:first-child { font-weight: 700; width: 30%; }
</style>
"""

LEAD_IN_RE = re.compile(r"^([^.:\n]{3,70}?)([.:]\s)(.*)$", re.DOTALL)


def _para(text: str, depth: int = 0, extra: str = "") -> str:
    """One paragraph, with its run-in heading bolded the way Word shows it."""
    escaped = html.escape(str(text or ""))
    escaped = lbk.SPACE_TOKEN_RE.sub(
        lambda m: f'<span class="lv-token">{html.escape(m.group(0))}</span>',
        escaped,
    ) if hasattr(lbk, "SPACE_TOKEN_RE") else escaped
    match = LEAD_IN_RE.match(escaped)
    if match and len(match.group(1)) < 70:
        escaped = (f'<span class="lv-lead">{match.group(1)}{match.group(2).rstrip()}</span> '
                   f'{match.group(3)}')
    indent = f" lv-i{min(depth, 2)}" if depth else ""
    return f'<p class="lv-p{indent}{extra}">{escaped}</p>'


def _token_marked(text: str) -> str:
    escaped = html.escape(str(text or ""))
    return re.sub(r"\[\s*(Space|KP)\s*:[^\]]*\]",
                  lambda m: f'<span class="lv-token">{m.group(0)}</span>', escaped)


def render_block(block: dict[str, Any], selection: Any = None) -> str:
    """One optional block or pick-one set, with its state shown."""
    chosen = (selection or {}).get(block.get("name"))
    parts = []

    if block.get("choice_group"):
        # Handled by render_group; a lone member should never reach here.
        return ""

    on = bool(chosen) if chosen is not None else False
    classes = "lv-opt" if on else "lv-opt lv-opt-off"
    parts.append(f'<div class="{classes}">')
    parts.append(f'<span class="lv-tag">optional · {"included" if on else "not included"}</span>')
    parts.append(_para(block["text"], block.get("indent", 0)))
    for child in block.get("children", []):
        parts.append(_para(child, block.get("indent", 0) + 1))
    parts.append("</div>")
    return "".join(parts)


def render_group(members: list[dict[str, Any]], selected: str | None = None) -> str:
    """A pick-one set: every option visible, the chosen one solid."""
    parts = [f'<div class="lv-set">',
             f'<span class="lv-tag">choose one of {len(members)}</span>']
    for member in members:
        active = (selected == member["name"]) or (selected is None and False)
        klass = "lv-choice" if active else "lv-choice lv-choice-off"
        parts.append(f'<div class="{klass}">')
        parts.append(_para(member["text"], member.get("indent", 0)))
        for child in member.get("children", []):
            parts.append(_para(child, member.get("indent", 0) + 1))
        parts.append("</div>")
    parts.append("</div>")
    return "".join(parts)


def _anchor(container: dict[str, Any]) -> str:
    label = str(container.get("label", "")).replace(" ", "-").replace(".", "_")
    return f"lv-{container.get('kind','x')}-{label}"


def _heading(container: dict[str, Any]) -> str:
    if container["kind"] == "section":
        title = container.get("title", "")
        return f"Section {container['label']}. {title}".strip().rstrip(".")
    if container.get("title"):
        return f"{container['label']} — {container['title']}"
    return str(container["label"])


def decisions_in(container: dict[str, Any]) -> int:
    """How many choices this container asks of you."""
    blocks = container.get("blocks", [])
    groups = {b["choice_group"] for b in blocks if b.get("choice_group")}
    singles = sum(1 for b in blocks
                  if not b.get("choice_group") and b.get("field") is None)
    return len(groups) + singles


def render_index(master: dict[str, Any]) -> str:
    """The clickable table of contents, with a count of decisions per section."""
    rows = []
    for container in master["containers"]:
        if container["kind"] == "front" and not container.get("blocks"):
            continue
        count = decisions_in(container)
        note = (f' <span class="lv-count">· {count} '
                f'{"decision" if count == 1 else "decisions"}</span>') if count else ""
        rows.append(f'<div><a href="#{_anchor(container)}">{html.escape(_heading(container))}</a>{note}</div>')
    return f'<div class="lv-index">{"".join(rows)}</div>'


def render_master(master: dict[str, Any], selection: Any = None,
                  only: str | None = None) -> str:
    """The whole document, or one container when `only` names its label."""
    selection = selection or {}
    parts = [STYLE, '<div class="lv-wrap">']
    if only is None:
        parts.append(render_index(master))

    for container in master["containers"]:
        if only is not None and str(container.get("label")) != only:
            continue
        blocks = container.get("blocks", [])
        provisions = [b for b in blocks if b.get("field") is not None]
        options = [b for b in blocks if b.get("field") is None]
        if not blocks and container["kind"] == "front":
            continue

        parts.append(f'<div class="lv-sec" id="{_anchor(container)}">')
        parts.append(f'<p class="lv-h">{html.escape(_heading(container))}</p>')

        if provisions:
            parts.append('<table class="lv-kp">')
            for provision in provisions:
                value = provision.get("text", "")
                if provision.get("alternatives"):
                    value = provision["alternatives"][0]
                cell = _token_marked(value).replace("\n", "<br>")
                tag = ""
                if provision.get("alternatives"):
                    tag = (f'<span class="lv-tag">choose one of '
                           f'{len(provision["alternatives"])}</span>')
                elif provision.get("optional"):
                    tag = '<span class="lv-tag">optional</span>'
                parts.append(
                    f'<tr><td>{html.escape(provision.get("field",""))}</td>'
                    f'<td>{tag}{cell}</td></tr>'
                )
            parts.append("</table>")

        groups: dict[int, list[dict]] = {}
        for block in options:
            if block.get("choice_group"):
                groups.setdefault(block["choice_group"], []).append(block)

        rendered = set()
        for block in options:
            group = block.get("choice_group")
            if group:
                if group in rendered:
                    continue
                rendered.add(group)
                parts.append(render_group(groups[group], selection.get(f"group:{group}")))
            else:
                parts.append(render_block(block, selection))
        parts.append("</div>")

    parts.append("</div>")
    return "".join(parts)
