"""Unit tests for the clause markup parser.

Run: python -m unittest test_lease_markup -v
"""

from __future__ import annotations

import unittest

import lease_markup as lm


def kinds(blocks):
    return [b.kind for b in blocks]


def texts(blocks):
    return [b.text for b in blocks]


class TestInline(unittest.TestCase):
    def test_plain_text_is_one_run(self):
        runs = lm.parse_inline("The Tenant shall pay rent.")
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0].text, "The Tenant shall pay rent.")
        self.assertFalse(runs[0].bold or runs[0].italic)

    def test_bold(self):
        runs = lm.parse_inline("Pay **all rent** promptly.")
        self.assertEqual([r.text for r in runs], ["Pay ", "all rent", " promptly."])
        self.assertEqual([r.bold for r in runs], [False, True, False])

    def test_italic(self):
        runs = lm.parse_inline("the _pro rata_ share")
        self.assertEqual([r.italic for r in runs], [False, True, False])

    def test_bold_inside_italic_and_vice_versa(self):
        runs = lm.parse_inline("**bold _both_ bold**")
        both = [r for r in runs if r.text == "both"][0]
        self.assertTrue(both.bold and both.italic)

    def test_unbalanced_markers_stay_literal(self):
        self.assertEqual(lm.parse_inline("2 ** 3 asterisks")[0].text, "2 ** 3 asterisks")
        self.assertEqual(lm.parse_inline("a * b")[0].text, "a * b")

    def test_snake_case_is_not_italic(self):
        runs = lm.parse_inline("see lease_builder_notes for detail")
        self.assertEqual(len(runs), 1)
        self.assertFalse(runs[0].italic)

    def test_underscore_needs_a_word_boundary(self):
        self.assertEqual(len(lm.parse_inline("a_b_c")), 1)

    def test_kp_token_becomes_a_ref_run(self):
        runs = lm.parse_inline("located at [KP:Property] today")
        self.assertEqual([r.kind for r in runs], ["text", "kp_ref", "text"])
        self.assertEqual(runs[1].name, "Property")
        self.assertEqual(runs[1].text, "")

    def test_kp_token_variants(self):
        for token in ("[KP:Property]", "[kp: Property ]", "[KPS:Property]"):
            runs = lm.parse_inline(f"x {token} y")
            self.assertEqual(runs[1].name, "Property", token)

    def test_multiword_provision_name(self):
        runs = lm.parse_inline("[KP:Tenant Share of Property]")
        self.assertEqual(runs[0].name, "Tenant Share of Property")

    def test_ref_inherits_surrounding_bold(self):
        runs = lm.parse_inline("**due on [KP:Rent Commencement Date] hereof**")
        ref = [r for r in runs if r.is_ref][0]
        self.assertTrue(ref.bold)

    def test_caret_becomes_a_break_run(self):
        runs = lm.parse_inline("123 Main St^Westfield, NJ")
        self.assertEqual([r.kind for r in runs], ["text", "break", "text"])

    def test_escapes(self):
        self.assertEqual(lm.parse_inline(r"5 \* 3")[0].text, "5 * 3")
        self.assertEqual(lm.parse_inline(r"\[KP:Property]")[0].text, "[KP:Property]")
        self.assertEqual(lm.parse_inline(r"a \^ b")[0].text, "a ^ b")
        self.assertEqual(lm.parse_inline(r"a \\ b")[0].text, "a \\ b")

    def test_empty_input(self):
        self.assertEqual(lm.parse_inline(""), [])
        self.assertEqual(lm.parse_inline(None), [])

    def test_adjacent_same_format_runs_merge(self):
        runs = lm.parse_inline("a[KP:X]b")
        self.assertEqual(len(runs), 3)
        self.assertEqual(len(lm.parse_inline("plain text here")), 1)


class TestBlocks(unittest.TestCase):
    def test_blank_line_splits_paragraphs(self):
        blocks = lm.parse_blocks("First para.\n\nSecond para.")
        self.assertEqual(kinds(blocks), ["paragraph", "paragraph"])

    def test_wrapped_lines_join(self):
        blocks = lm.parse_blocks("The Tenant\nshall pay\nrent.")
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0].text, "The Tenant shall pay rent.")

    def test_bullets(self):
        blocks = lm.parse_blocks("- one\n- two")
        self.assertEqual(kinds(blocks), ["bullet", "bullet"])
        self.assertEqual(texts(blocks), ["one", "two"])
        self.assertEqual([b.number for b in blocks], ["•", "•"])

    def test_numbered_items_are_renumbered(self):
        blocks = lm.parse_blocks("7. first\n9. second\n1. third")
        self.assertEqual([b.number for b in blocks], ["1.", "2.", "3."])

    def test_lettered_subclauses(self):
        blocks = lm.parse_blocks("A. first\nB. second\nZ. third")
        self.assertEqual(kinds(blocks), ["lettered"] * 3)
        self.assertEqual([b.number for b in blocks], ["A.", "B.", "C."])

    def test_paren_marker_accepted(self):
        self.assertEqual(kinds(lm.parse_blocks("1) one\n2) two")), ["numbered", "numbered"])
        self.assertEqual(kinds(lm.parse_blocks("A) one")), ["lettered"])

    def test_level_2_roman_stays_inline(self):
        blocks = lm.parse_blocks("A. Tenant shall (i) pay rent and (ii) insure.")
        self.assertEqual(len(blocks), 1)
        self.assertIn("(i) pay rent", blocks[0].text)

    def test_tab_indent_nests(self):
        blocks = lm.parse_blocks("A. outer\n\n\t- inner")
        self.assertEqual([b.level for b in blocks], [0, 1])

    def test_four_spaces_nest(self):
        blocks = lm.parse_blocks("A. outer\n\n    - inner")
        self.assertEqual([b.level for b in blocks], [0, 1])

    def test_deep_indent_is_capped(self):
        blocks = lm.parse_blocks("\t" * 10 + "- deep")
        self.assertEqual(blocks[0].level, lm.MAX_LEVEL)

    def test_list_marker_starts_a_block_without_a_blank_line(self):
        blocks = lm.parse_blocks("Intro sentence.\n- one\n- two")
        self.assertEqual(kinds(blocks), ["paragraph", "bullet", "bullet"])

    def test_continuation_line_stays_with_its_list_item(self):
        blocks = lm.parse_blocks("A. first line\ncontinues here\n\nB. second")
        self.assertEqual(len(blocks), 2)
        self.assertEqual(blocks[0].text, "first line continues here")

    def test_bold_lead_in_on_a_subclause(self):
        blocks = lm.parse_blocks("A. **Rent Commencement Date.** The date shall be set.")
        self.assertEqual(blocks[0].kind, "lettered")
        self.assertTrue(blocks[0].runs[0].bold)
        self.assertEqual(blocks[0].runs[0].text, "Rent Commencement Date.")

    def test_empty_and_whitespace_input(self):
        self.assertEqual(lm.parse_blocks(""), [])
        self.assertEqual(lm.parse_blocks("   \n\n  \t \n"), [])
        self.assertEqual(lm.parse_blocks(None), [])

    def test_crlf_is_handled(self):
        self.assertEqual(len(lm.parse_blocks("a\r\n\r\nb")), 2)

    def test_bare_marker_stays_literal_rather_than_vanishing(self):
        """A lone '-' is more likely a typed separator than an empty bullet, and
        silently dropping text from a lease is the worse failure."""
        blocks = lm.parse_blocks("-  \n")
        self.assertEqual(kinds(blocks), ["paragraph"])
        self.assertEqual(blocks[0].text, "-")


class TestNumbering(unittest.TestCase):
    def test_paragraph_between_items_does_not_restart(self):
        blocks = lm.parse_blocks("A. first\n\nAn explanatory paragraph.\n\nB. second")
        lettered = [b.number for b in blocks if b.kind == "lettered"]
        self.assertEqual(lettered, ["A.", "B."])

    def test_nested_list_restarts_inside_each_parent(self):
        blocks = lm.parse_blocks(
            "A. first\n\n\t1. alpha\n\t2. beta\n\nB. second\n\n\t1. gamma"
        )
        numbers = [(b.level, b.number) for b in blocks]
        self.assertEqual(
            numbers, [(0, "A."), (1, "1."), (1, "2."), (0, "B."), (1, "1.")]
        )

    def test_numbered_and_lettered_count_separately(self):
        blocks = lm.parse_blocks("A. a\n1. one\nB. b\n2. two")
        self.assertEqual([b.number for b in blocks], ["A.", "1.", "B.", "2."])

    def test_page_break_resets_numbering(self):
        blocks = lm.parse_blocks("A. first\n\n[PageBreak]\n\nA. fresh")
        lettered = [b.number for b in blocks if b.kind == "lettered"]
        self.assertEqual(lettered, ["A.", "A."])

    def test_letters_roll_over_past_z(self):
        source = "\n".join(f"A. item {i}" for i in range(27))
        blocks = lm.parse_blocks(source)
        self.assertEqual(blocks[25].number, "Z.")
        self.assertEqual(blocks[26].number, "AA.")

    def test_renumber_can_be_switched_off(self):
        blocks = lm.parse_blocks("7. seven", renumber=False)
        self.assertEqual(blocks[0].number, "")


class TestStructuralMarkers(unittest.TestCase):
    def test_rent_table_marker(self):
        blocks = lm.parse_blocks("Rent is payable monthly.\n\n[RentTable:Base]")
        self.assertEqual(kinds(blocks), ["paragraph", "rent_table"])
        self.assertEqual(blocks[1].data["schedule"], "Base")

    def test_option_rent_table(self):
        blocks = lm.parse_blocks("[RentTable:Option 1]")
        self.assertEqual(blocks[0].data["schedule"], "Option 1")

    def test_legacy_table_markers(self):
        blocks = lm.parse_blocks("Base Rent Table:\n\nTable End\n\nAfter.")
        self.assertEqual(kinds(blocks), ["rent_table", "paragraph"])
        self.assertEqual(blocks[0].data["schedule"], "Base")

    def test_legacy_end_marker_only_closes_a_table_it_opened(self):
        """'Table End' on its own is a sentence, not a marker; swallowing it
        would delete text no marker ever opened."""
        self.assertEqual(kinds(lm.parse_blocks("DELETE TABLE")), ["paragraph"])
        self.assertEqual(kinds(lm.parse_blocks("Table End")), ["paragraph"])
        self.assertEqual(
            kinds(lm.parse_blocks("Base Rent Table:\n\nDELETE TABLE")), ["rent_table"]
        )

    def test_paragraph_that_looks_like_a_list_survives_a_round_trip(self):
        # "a)" alone is not a marker (nothing follows it), so the wrapped lines
        # join into one paragraph whose text now *starts* like a lettered item.
        blocks = lm.parse_blocks("a)\nnot a list")
        self.assertEqual(kinds(blocks), ["paragraph"])
        self.assertEqual(blocks[0].text, "a) not a list")
        again = lm.parse_blocks(lm.to_markup(blocks))
        self.assertEqual(kinds(again), ["paragraph"])
        self.assertEqual(again[0].text, blocks[0].text)

    def test_page_break(self):
        self.assertEqual(kinds(lm.parse_blocks("[PageBreak]")), ["page_break"])

    def test_exhibit_with_title(self):
        blocks = lm.parse_blocks("[Exhibit:A|Floor Plan of the Premises]")
        self.assertEqual(blocks[0].kind, "exhibit")
        self.assertEqual(blocks[0].data, {"letter": "A", "title": "Floor Plan of the Premises"})

    def test_exhibit_without_title(self):
        blocks = lm.parse_blocks("[Exhibit:B]")
        self.assertEqual(blocks[0].data, {"letter": "B", "title": ""})

    def test_marker_must_be_alone_on_its_line(self):
        blocks = lm.parse_blocks("See [RentTable:Base] below.")
        self.assertEqual(kinds(blocks), ["paragraph"])

    def test_schedules_used(self):
        blocks = lm.parse_blocks("[RentTable:Base]\n\nx\n\n[RentTable:Option 1]\n\n[RentTable:Base]")
        self.assertEqual(lm.rent_schedules_used(blocks), ["Base", "Option 1"])


class TestRefs(unittest.TestCase):
    def test_collect_refs_in_order_deduplicated(self):
        blocks = lm.parse_blocks("[KP:Premises] and [KP:Property]\n\nagain [KP:Premises]")
        self.assertEqual(lm.collect_refs(blocks), ["Premises", "Property"])

    def test_no_refs(self):
        self.assertEqual(lm.collect_refs(lm.parse_blocks("plain")), [])


class TestSerializers(unittest.TestCase):
    def test_plain_text_applies_numbering(self):
        out = lm.to_plain_text(lm.parse_blocks("9. first\n9. second"))
        self.assertIn("1. first", out)
        self.assertIn("2. second", out)

    def test_plain_text_keeps_unresolved_tokens_visible(self):
        out = lm.to_plain_text(lm.parse_blocks("at [KP:Property]"))
        self.assertIn("[KP:Property]", out)

    def test_plain_text_substitutes_when_given_values(self):
        out = lm.to_plain_text(lm.parse_blocks("at [KP:Property]"), {"Property": "114 Central Ave"})
        self.assertEqual(out, "at 114 Central Ave")

    def test_plain_text_drops_formatting(self):
        self.assertEqual(lm.to_plain_text(lm.parse_blocks("**bold** and _italic_")), "bold and italic")

    def test_caret_becomes_a_newline_in_plain_text(self):
        self.assertEqual(lm.to_plain_text(lm.parse_blocks("a^b")), "a\nb")

    def test_html_escapes_and_formats(self):
        html = lm.to_html(lm.parse_blocks("**A & B** <tag>"))
        self.assertIn("<strong>A &amp; B</strong>", html)
        self.assertIn("&lt;tag&gt;", html)

    def test_html_marks_unresolved_refs_red(self):
        html = lm.to_html(lm.parse_blocks("[KP:Nope]"))
        self.assertIn("#b00020", html)

    def test_html_renders_resolved_refs_green_and_underlined(self):
        html = lm.to_html(lm.parse_blocks("[KP:Property]"), {"Property": "114 Central"})
        self.assertIn("#1F7A33", html)
        self.assertIn("underline", html)
        self.assertIn("114 Central", html)


class TestRoundTrip(unittest.TestCase):
    def assert_stable(self, source):
        once = lm.parse_blocks(source)
        twice = lm.parse_blocks(lm.to_markup(once))
        self.assertEqual(lm.to_markup(once), lm.to_markup(twice), source)
        self.assertEqual(lm.to_plain_text(once), lm.to_plain_text(twice), source)

    def test_round_trip_is_stable(self):
        for source in [
            "Plain paragraph.",
            "**Bold** and _italic_ text.",
            "A. first\n\nB. second",
            "- one\n- two",
            "1. one\n2. two",
            "at [KP:Tenant Share of Property] hereof",
            "123 Main St^Westfield, NJ^07090",
            "A. outer\n\n\t1. inner\n\t2. inner two",
            "[RentTable:Base]",
            "[Exhibit:A|Floor Plan of the Premises]",
            "[PageBreak]",
            "Intro.\n\n[RentTable:Option 1]\n\nOutro.",
            r"literal \* asterisk and \[brackets]",
        ]:
            self.assert_stable(source)

    def test_escaped_markup_survives_a_round_trip(self):
        blocks = lm.parse_blocks(r"5 \* 3 = 15")
        self.assertEqual(lm.to_plain_text(lm.parse_blocks(lm.to_markup(blocks))), "5 * 3 = 15")

    def test_markup_output_reparses_to_the_same_text(self):
        source = "A. **Lead-in.** Body with [KP:Property] and a break^here."
        blocks = lm.parse_blocks(source)
        self.assertEqual(
            [(b.kind, b.text) for b in lm.parse_blocks(lm.to_markup(blocks))],
            [(b.kind, b.text) for b in blocks],
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
