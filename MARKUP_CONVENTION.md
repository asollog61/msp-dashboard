# Master lease markup convention

How the master `.docx` tells the Lease Builder what is fixed, what is optional,
and what is a choice between competing versions.

Everything here is expressed with **Word highlight colours**, so the master
stays a readable Word document that a lawyer can open, and the same file is
also the machine-readable menu the builder extracts.

Source of truth: `2026_08_01 MSP Master_Lease.v9B NNN Retail.docx`.

---

## The three colours

| Highlight | Meaning | Behaviour in the builder |
|---|---|---|
| **Green** | Independent optional block | A checkbox. On or off, no effect on anything else. |
| **Cyan** | One of a set of competing versions | A radio group. Exactly one may be selected, or none. |
| **Magenta** | Optional key provision (front matter only) | Offered in the Key Provisions table rather than the section body. |

**No highlight** means the text is fixed. It is always in the lease and cannot
be switched off. That is the default, and most of the document should stay that
way.

Magenta is used because it is currently unused as a highlight anywhere in the
master. Yellow is deliberately *not* used — it reads as "unfinished" and three
stray yellow runs already exist.

### Red text is not part of this

Red font colour (`EE0000`, 232 runs today) is drafting guidance for a human —
notes, reminders, things to decide. It is stripped from generated leases and
carries no meaning for the extractor. Do not use red to mark options.

---

## Where a block starts and stops

**You do not need to mark this. Indentation already says it.**

A block begins at a highlighted paragraph at the shallowest highlighted indent
level. Highlighted paragraphs that follow it at a *deeper* indent belong to it
and travel with it.

Worked example, Section 8 as it stands today:

```
        Section 8. Security Deposit.                    no highlight  -> fixed
  720   Initial Deposit.                                no highlight  -> fixed
  720   Maintenance of Deposit.                         GREEN  -> block 1
  720   Return of Balance.                              no highlight  -> fixed
 1440     The expiration of the Lease Term, or          no highlight  -> fixed
 1440     Tenant vacating of the Premises...            no highlight  -> fixed
  720   Conditional Reduction.                          GREEN  -> block 2
 1440     Application of Credit:                        GREEN  -> block 2 (child)
 1440     Reinstatement Rights:                         GREEN  -> block 2 (child)
  720   Application to Final Month.                     GREEN  -> block 3
```

Three optional blocks, not five. Turning off *Conditional Reduction* takes its
two children with it, which is the only sane reading — "Application of Credit"
means nothing on its own.

**Consequence to be aware of:** a child paragraph must be indented deeper than
its parent. If you flatten the indent, it becomes a separate block.

---

## Cyan: competing versions

Consecutive cyan blocks at the same indent level form **one** radio group.
A block of any other kind, or an unhighlighted paragraph, closes the group.

```
  720   Termination Right. Tenant may...                CYAN  ┐
  720   Rent Credit. Tenant may elect...                CYAN  ┘ pick one
```

The literal `OR` paragraph you have been writing between alternatives is no
longer needed — colour and adjacency carry it. **Delete those `OR` paragraphs**
when you re-colour, or they will be extracted as a one-word alternative.

To offer two *separate* radio groups inside one section, put an unhighlighted
paragraph between them, or make one of them green.

### The case this fixes

Section 2 currently holds three near-identical copies of the Option to Cancel
language, adjacent, with nothing marking them as alternatives to each other.
Colour them all cyan and they become one pick-one group. Leave them green and
you will be able to select all three into the same lease.

---

## Naming

A block's name comes from its **run-in heading** — the bold lead-in ending in a
full stop or colon:

> **Conditional Reduction.** Notwithstanding the foregoing, provided that…

becomes a checkbox labelled *Conditional Reduction*.

A block with no run-in heading is named from its first few words, which is
usually worse. Give every optional block a bold run-in heading.

---

## Front matter (before Section 1)

Twenty highlighted paragraphs currently sit above Section 1. These are optional
key provisions — items that belong in the Key Provisions Summary table and,
when included, need corresponding language somewhere in the section body.

Mark these **magenta**. They are routed to the Key Provisions table, which
already has its own chooser, rather than being treated as section text.

An unhighlighted front-matter provision is mandatory and always present.

---

## Checklist before re-extracting

- [ ] Every optional block is green **or** cyan, never both
- [ ] Competing versions are cyan and adjacent, with nothing between them
- [ ] Leftover `OR` separator paragraphs deleted
- [ ] Child paragraphs indented deeper than their parent
- [ ] Every optional block has a bold run-in heading
- [ ] Front-matter optional provisions are magenta
- [ ] Nothing is highlighted that should always appear in every lease

Then run **Re-extract** and check the report, which lists every block found,
its colour, its section and its name. Read it before saving any lease — it is
the only place a mis-colour becomes visible.

---

## What this does not cover yet

**Inline alternatives.** Two paragraphs in the current master highlight a phrase
*within* a sentence rather than the whole paragraph. The convention above works
on whole paragraphs only. Either promote those to full alternative paragraphs,
or leave them and accept that the whole paragraph becomes the option.

**Cross-section dependencies.** Choosing "no parking" does not yet adjust the
Common Areas or Rules and Regulations clauses that reference parking. Those
remain a manual check.
