"""Render the master as a document on the left and its decisions on the right.

Two different questions were being answered by one view, badly:

    "what does this lease say?"      needs the text as it will print, with
                                     nothing else in the way
    "what could it say instead?"     needs every alternative, including the
                                     ones not chosen

So there are two renderers. `render_document` shows only what is in — fixed
text, the selected option from each set, the optional blocks switched on. It
reads like the lease because it is the lease. `render_options` shows one
section's decisions: every competing version, every block that can be dropped.

The index links carry a `?lb_section=` query parameter rather than a plain
anchor, because Streamlit reruns on a query change and a bare `#anchor` only
scrolls. That is what makes clicking a heading drive the right-hand pane.

Produces HTML. No Streamlit import, so it can be tested by reading the string.
"""

from __future__ import annotations

import html
import re
from typing import Any

import lease_blocks as lbk

STYLE = """
<style>
.lv-doc { font-family: 'Times New Roman', Georgia, serif; font-size: 14.5px;
          line-height: 1.5; }
.lv-doc p { margin: 0 0 .5rem 0; text-align: justify; }
.lv-h { font-weight: 700; margin: 1.1rem 0 .4rem 0 !important; }
.lv-lead { font-weight: 700; }
.lv-i1 { margin-left: 1.9rem !important; }
.lv-i2 { margin-left: 3.8rem !important; }
.lv-token { color: #d13438; font-weight: 600; }

.lv-index { font-family: system-ui, sans-serif; font-size: 12.5px;
            line-height: 1.85; border-left: 3px solid rgba(128,128,128,.3);
            padding-left: .8rem; margin-bottom: 1.1rem; }
.lv-index a { text-decoration: none; opacity: .8; }
.lv-index a:hover { text-decoration: underline; opacity: 1; }
.lv-index .on { font-weight: 700; opacity: 1; }
.lv-badge { font-size: 11px; opacity: .75; }
.lv-badge.set { color: #46b8d0; }
.lv-badge.opt { color: #4bbd7a; }

.lv-opts { font-family: system-ui, sans-serif; font-size: 13px; }
.lv-card { border: 1px solid rgba(128,128,128,.35); border-radius: .5rem;
           padding: .7rem .8rem; margin-bottom: .8rem; }
.lv-card.set { border-left: 3px solid #1f8fa8; }
.lv-card.opt { border-left: 3px solid #2e9e5b; }
.lv-tag { font-size: 10.5px; letter-spacing: .05em; text-transform: uppercase;
          opacity: .7; display: block; margin-bottom: .45rem; }
.lv-choice { border-top: 1px solid rgba(128,128,128,.22); padding: .45rem 0 .1rem 0;
             font-family: 'Times New Roman', Georgia, serif; font-size: 13.5px;
             line-height: 1.45; }
.lv-choice:first-of-type { border-top: none; }
.lv-name { font-family: system-ui, sans-serif; font-size: 11.5px;
           font-weight: 600; opacity: .85; display: block; margin-bottom: .15rem; }
.lv-in  { color: #4bbd7a; }
.lv-out { opacity: .5; }
</style>
"""

LEAD_IN_RE = re.compile(r"^([^.:\n]{3,70}?)([.:]\s)(.*)$", re.DOTALL)
TOKEN_RE = re.compile(r"\[\s*(?:Space|KP)\s*:[^\]]*\]", re.IGNORECASE)


def _escape(text: str) -> str:
    """Escaped, with unresolved tokens made obvious rather than invisible."""
    escaped = html.escape(str(text or ""))
    return TOKEN_RE.sub(lambda m: f'<span class="lv-token">{m.group(0)}</span>', escaped)


def _para(text: str, depth: int = 0) -> str:
    body = _escape(text)
    match = LEAD_IN_RE.match(body)
    if match:
        body = (f'<span class="lv-lead">{match.group(1)}{match.group(2).rstrip()}</span> '
                f'{match.group(3)}')
    klass = f' class="lv-i{min(depth, 2)}"' if depth else ""
    return f"<p{klass}>{body}</p>"


def _heading(container: dict[str, Any]) -> str:
    if container["kind"] == "section":
        return f"Section {container['label']}. {container.get('title','')}".strip().rstrip(".")
    if container.get("title"):
        return f"{container['label']} — {container['title']}"
    return str(container["label"])


def _anchor(container: dict[str, Any]) -> str:
    label = str(container.get("label", "")).replace(" ", "-").replace(".", "_")
    return f"lv-{container.get('kind','x')}-{label}"


# ---------------------------------------------------------------------------
# Decisions
# ---------------------------------------------------------------------------

def groups_in(container: dict[str, Any]) -> dict[int, list[dict]]:
    groups: dict[int, list[dict]] = {}
    for block in container.get("blocks", []):
        if block.get("choice_group"):
            groups.setdefault(block["choice_group"], []).append(block)
    return groups


def optional_in(container: dict[str, Any]) -> list[dict]:
    return [b for b in container.get("blocks", [])
            if not b.get("choice_group") and b.get("field") is None]


def provisions_in(container: dict[str, Any]) -> list[dict]:
    return [b for b in container.get("blocks", []) if b.get("field") is not None]


def decisions_in(container: dict[str, Any]) -> int:
    return len(groups_in(container)) + len(optional_in(container)) + sum(
        1 for p in provisions_in(container)
        if p.get("alternatives") or p.get("optional")
    )


def describe_decisions(container: dict[str, Any]) -> str:
    """"2 options · 3 optional" — what this section is asking of you."""
    parts = []
    sets_ = groups_in(container)
    if sets_:
        total = sum(len(members) for members in sets_.values())
        parts.append(f'<span class="lv-badge set">{total} options</span>')
    optional = optional_in(container)
    if optional:
        parts.append(f'<span class="lv-badge opt">{len(optional)} optional</span>')
    provisions = [p for p in provisions_in(container)
                  if p.get("alternatives") or p.get("optional")]
    if provisions:
        parts.append(f'<span class="lv-badge set">{len(provisions)} provisions</span>')
    return " · ".join(parts)


# ---------------------------------------------------------------------------
# Selection
# ---------------------------------------------------------------------------

def default_selection(master: dict[str, Any]) -> dict[str, Any]:
    """What the lease says before you change anything.

    A pick-one set takes its first option, because a lease has to say
    something. A green block stays out, because that is what green means.
    """
    selection: dict[str, Any] = {}
    for container in master["containers"]:
        label = str(container["label"])
        for group, members in groups_in(container).items():
            selection[f"{label}|set{group}"] = members[0]["name"]
        for block in optional_in(container):
            selection[f"{label}|opt|{block['name']}"] = False
    return selection


def _chosen(selection: dict, label: str, group: int, members: list[dict]) -> dict:
    name = (selection or {}).get(f"{label}|set{group}")
    for member in members:
        if member["name"] == name:
            return member
    return members[0]


def _is_on(selection: dict, label: str, block: dict) -> bool:
    return bool((selection or {}).get(f"{label}|opt|{block['name']}", False))


# ---------------------------------------------------------------------------
# Left pane: the lease as it reads
# ---------------------------------------------------------------------------

def render_document(master: dict[str, Any], selection: dict | None = None) -> str:
    """Only what is in the lease. No alternatives, no tags, no dimming."""
    selection = selection if selection is not None else default_selection(master)
    parts = [STYLE, '<div class="lv-doc">']

    for container in master["containers"]:
        label = str(container["label"])
        blocks = container.get("blocks", [])
        provisions = provisions_in(container)
        if not blocks and container["kind"] == "front":
            continue

        parts.append(f'<p class="lv-h" id="{_anchor(container)}">'
                     f'{html.escape(_heading(container))}</p>')

        for provision in provisions:
            value = provision.get("text", "")
            if provision.get("alternatives"):
                value = provision["alternatives"][0]
            if value:
                parts.append(
                    f'<p><span class="lv-lead">{html.escape(provision.get("field",""))}</span> '
                    f'{_escape(value)}</p>'
                )

        groups = groups_in(container)
        rendered = set()
        for block in container.get("blocks", []):
            if block.get("field") is not None:
                continue
            group = block.get("choice_group")
            if group:
                if group in rendered:
                    continue
                rendered.add(group)
                winner = _chosen(selection, label, group, groups[group])
                parts.append(_para(winner["text"], winner.get("indent", 0)))
                for child in winner.get("children", []):
                    parts.append(_para(child, winner.get("indent", 0) + 1))
            elif _is_on(selection, label, block):
                parts.append(_para(block["text"], block.get("indent", 0)))
                for child in block.get("children", []):
                    parts.append(_para(child, block.get("indent", 0) + 1))

    parts.append("</div>")
    return "".join(parts)


# ---------------------------------------------------------------------------
# Right pane: the index, and one section's decisions
# ---------------------------------------------------------------------------

def render_index(master: dict[str, Any], active: str | None = None) -> str:
    """Sections that ask something of you, each saying what and how many."""
    rows = []
    for container in master["containers"]:
        if not decisions_in(container):
            continue
        label = str(container["label"])
        on = " on" if active == label else ""
        rows.append(
            f'<div><a class="lv-link{on}" href="?lb_section={html.escape(label)}">'
            f'{html.escape(_heading(container))}</a> '
            f'{describe_decisions(container)}</div>'
        )
    if not rows:
        return '<div class="lv-index">No options marked in this master.</div>'
    return f'<div class="lv-index">{"".join(rows)}</div>'


def render_options(container: dict[str, Any], selection: dict | None = None) -> str:
    """Everything this section could say instead of what it says now."""
    label = str(container["label"])
    selection = selection or {}
    parts = [STYLE, '<div class="lv-opts">']

    for provision in provisions_in(container):
        if not (provision.get("alternatives") or provision.get("optional")):
            continue
        parts.append('<div class="lv-card set">')
        head = html.escape(provision.get("field", ""))
        if provision.get("alternatives"):
            parts.append(f'<span class="lv-tag">{head} · choose one of '
                         f'{len(provision["alternatives"])}</span>')
            for alternative in provision["alternatives"]:
                parts.append(f'<div class="lv-choice">{_escape(alternative)}</div>')
        else:
            parts.append(f'<span class="lv-tag">{head} · optional</span>')
            parts.append(f'<div class="lv-choice">{_escape(provision.get("text",""))}</div>')
        parts.append("</div>")

    groups = groups_in(container)
    for group, members in sorted(groups.items()):
        winner = _chosen(selection, label, group, members)
        parts.append('<div class="lv-card set">')
        parts.append(f'<span class="lv-tag">choose one of {len(members)}</span>')
        for member in members:
            mark = ('<span class="lv-in">● in the lease</span>'
                    if member["name"] == winner["name"] else
                    '<span class="lv-out">○ not used</span>')
            parts.append(
                f'<div class="lv-choice"><span class="lv-name">'
                f'{html.escape(member["name"])} — {mark}</span>{_escape(member["text"])}'
                + "".join(f'<div class="lv-i1">{_escape(c)}</div>'
                          for c in member.get("children", []))
                + "</div>"
            )
        parts.append("</div>")

    for block in optional_in(container):
        on = _is_on(selection, label, block)
        parts.append('<div class="lv-card opt">')
        parts.append(
            f'<span class="lv-tag">optional · '
            + ('<span class="lv-in">included</span>' if on
               else '<span class="lv-out">not included</span>')
            + "</span>"
        )
        parts.append(
            f'<div class="lv-choice"><span class="lv-name">'
            f'{html.escape(block["name"])}</span>{_escape(block["text"])}'
            + "".join(f'<div class="lv-i1">{_escape(c)}</div>'
                      for c in block.get("children", []))
            + "</div>"
        )
        parts.append("</div>")

    if len(parts) == 2:
        parts.append('<div class="lv-card">Nothing to decide in this section.</div>')
    parts.append("</div>")
    return "".join(parts)


def find_container(master: dict[str, Any], label: str | None):
    if label is None:
        return None
    for container in master["containers"]:
        if str(container["label"]) == str(label):
            return container
    return None


def first_with_decisions(master: dict[str, Any]):
    for container in master["containers"]:
        if decisions_in(container):
            return container
    return None
