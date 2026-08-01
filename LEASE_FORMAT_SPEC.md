# Lease Generation Spec

Derived from the executed lease `1_Executed_Chez Alice_MSP Lease_6.16.26.pdf` (25 pp).
This is the target output for generating a lease **without a base Word template**.

## Decisions taken

| Question | Decision |
|---|---|
| Base `.docx` | **None.** Everything generated from rules below. |
| Clause markup | **Markdown subset** (see below). |
| Numbering | **App assigns all numbering.** |
| Exhibit floor plans | **Per-lease image upload** — they are deal-specific, not template content. |
| `(#Address)` style markers | **Dropped.** Superseded by `[KP:Name]` tokens. |
| Key provisions list | App-owned (already done). |
| Cross-references | `[KP:Name]` → green underlined hyperlink to the summary row (already done). |
| Line break in a value | `^` caret (already done). |

## Page setup

- US Letter 8.5 × 11 in, margins 1 in all sides.
- Serif body (Times New Roman), 11 pt, **justified**.
- Footer: counsel document ID left (e.g. `4928-4211-2690, v. 2`), page number right.
  Doc ID is a per-template setting; page number is a `PAGE` field.

## Document order

1. Title block — `LEASE AGREEMENT`, then `KEY PROVISIONS SUMMARY`, each centered, bold, underlined.
2. Key Provisions table.
3. Preamble (`This LEASE AGREEMENT ... BETWEEN / AND`).
4. Sections 1..N.
5. `SIGNATURES` block.
6. Exhibits A, B, C ...

## Key Provisions table

- Bordered, 2 columns: label ≈ 1.9 in, value ≈ 5.1 in.
- **Exception:** the *Notice Addresses* row splits the value into two columns
  (Landlord | Tenant). This is why the source table is 3 columns with most rows
  merged across cols 2–3. Any provision may opt into the split.
- A value may embed the rent table (see below).
- Each row is bookmarked `_MSP_KP_<Field>` as the target of body cross-references.

## Section headings

Run-in bold, body continues on the same line:

```
Section 5.<tab>Base and Additional Rent.  Tenant acknowledges and agrees that ...
```

- `Section` + number + `.` bold; tab; title + `.` bold; then body in normal weight.
- Decimal sections (`4.1`, `5.2`, `17.1`, `46.1`) are **peers**, not children —
  they render identically, just with a decimal number.
- Each section start is bookmarked `_MSP_Sec_<number>`.

## Sub-clauses

Two levels only:

- Level 1: `A.` `B.` `C.` — lettered, hanging indent ≈ 0.25 in, justified.
  Frequently opens with a bold lead-in phrase ending in a period:
  `A. **Rent Commencement Date.** The Rent Commencement Date shall be ...`
- Level 2: `(i)` `(ii)` — inline within the paragraph, not a separate list.

## Clause markup (markdown subset)

| Input | Output |
|---|---|
| `**text**` | bold |
| `_text_` | italic |
| `- item` | bullet |
| `1. item` | numbered (app renumbers) |
| `A. item` | lettered sub-clause |
| leading tab / 4 spaces | one nesting level deeper |
| blank line | new paragraph |
| `[KP:Name]` | provision value, green underlined hyperlink to summary row |
| `^` inside a value | line break within the cell/run |

## Rent table

- Bold label `Base Rent Table:` above.
- 3 columns: `Term` | `Monthly Amount` | `Annual Amount`, bordered, header row bold.
- Amounts right-aligned currency.
- Closed by a `Table End` line.
- Appears twice: inside the Key Provisions *Base Rent* value, and in the rent section.
- Source data is the caret/Excel rent schedule already implemented
  (`Term` block in A:C, `Option N` blocks in F:H; blank or `0` = does not exist).

## Signatures

- Centered bold underlined `SIGNATURES`.
- `In Witness Whereof, ...` paragraph.
- Two-column block, landlord left / tenant right:
  entity name, `(LANDLORD)` / `(TENANT)`, `By:` rule, name, title, date.

## Exhibits

- Page break before each.
- Centered bold `"Exhibit A"`, then centered bold subtitle
  (`Floor Plan of the Premises`).
- Body is either uploaded image(s) — per lease — or editable text
  (Exhibit B *Landlord's Work Itemized*, Exhibit C *Tenant's Work*).

## Build order (proposed)

1. Formatting-settings model + form; persist with the template.
2. Markdown-subset parser → paragraph/run model (pure, unit-testable).
3. Renderer: page setup, styles, footer.
4. Key Provisions table incl. the Notice-Addresses split.
5. Sections with auto-numbering and run-in headings.
6. Rent tables.
7. Signature block.
8. Exhibits incl. per-lease image upload.
9. Compare generated output against the executed PDF page by page.

## Verification

Every step must be run, not just written. Minimum bar before shipping:
- `python-docx` opens the result and re-reads the structure.
- No duplicate bookmark IDs (Word reports these as a damaged file).
- LibreOffice converts it to PDF without error.
- Page-by-page visual diff against `1_Executed_Chez Alice_MSP Lease_6.16.26.pdf`.
