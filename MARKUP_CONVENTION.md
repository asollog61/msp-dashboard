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
| **Bright green** | Independent optional block, **out** by default | A checkbox. Off unless you opt in. Two adjacent green blocks can both be in. |
| **Turquoise (cyan)** | One of a set of competing versions | A radio group. Exactly one may be selected, or none. |
| **Yellow** | Optional key provision (front matter only) | Offered in the Key Provisions table rather than the section body. |
| **Grey / red** | Scaffolding — directions and notes to self | Stripped on extract. Never reaches a lease. |

**No highlight** means the text is in the lease by default. It can still be
switched off deal by deal — it simply is not surfaced as a decision. Most of
the document should stay this way.

The three active colours are all light ones, so text stays readable behind
them. Highlights never survive into a generated lease, so they exist purely for
you in Word.

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

A run of cyan is **one** pick-one set. Inside it, a rule of asterisks ends one
choice and begins the next:

```
        AS IS/WHERE IS                                  CYAN  ┐
        The premises to be delivered AS-IS...           CYAN  │ choice 1
        ***********************************                   │
        PHASE 1 –                                       CYAN  │
        DEMO THE ENTIRE PREMISES...                     CYAN  │ choice 2
        Gas: ...  Water: ...  Sewer: ...  HVAC: ...     CYAN  │
        PHASE 2 –   Electric: ...                       CYAN  ┘
```

**One separator means two choices. Two separators mean three.** A choice is
everything between two separators, however many paragraphs that takes — the
eight paragraphs of the phased build-out above are a single option, not eight.

Separators appear **only inside cyan**. They are stripped and never reach a
lease.

A cyan run with no separator in it is a single optional block, not a choice —
there is nothing to choose between.

### Green is not a choice

Two adjacent green blocks are two independent things that can both be in the
lease. Section 2's *Termination Right* and *Rent Credit* are green on purpose:
when that provision is included, the lease describes both remedies and the
tenant elects between them later, in the real world. Making them cyan would
force one out of the document at drafting time, which is wrong.

Use cyan only when including one version means the other must not appear at
all.

## Naming

A block's name comes from its **run-in heading** — the bold lead-in ending in a
full stop or colon:

> **Conditional Reduction.** Notwithstanding the foregoing, provided that…

becomes a checkbox labelled *Conditional Reduction*.

A block with no run-in heading is named from its first few words, which is
usually worse. Give every optional block a bold run-in heading.

---

## Front matter (before Section 1)

These live in the Key Provisions Summary **table**, and are read from the table
rows rather than from body paragraphs. Thirteen are marked today.

Mark them **yellow**. They are routed to the Key Provisions table, which has
its own chooser, rather than being treated as section text.

An unhighlighted front-matter provision is mandatory and always present.

**Every key provision needs a matching section.** If *Exclusivity* appears in
the summary, there must be an exclusivity clause in the body. The reverse is
fine — a section may exist with no summary row.

---

## Checklist before re-extracting

- [ ] Every optional block is green **or** cyan, never both
- [ ] Cyan is used only where including one version means the other must not appear
- [ ] Separators appear only inside cyan, one fewer than the number of choices
- [ ] Child paragraphs indented deeper than their parent
- [ ] Every optional block has a bold run-in heading
- [ ] No two blocks in one set share a name
- [ ] Front-matter optional provisions are yellow
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
