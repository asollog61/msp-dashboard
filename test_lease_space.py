"""Unit tests for space records and [Space:...] token resolution.

Run: python -m unittest test_lease_space -v
"""

from __future__ import annotations

import unittest

import lease_space as lsp


def summary(building="114 Central", unit="102", tenant="Nabig Sakr", sqft=1063,
            lease_type="NNN", floor_pct=0.2172, category_pct=0.2172, floor="1",
            use_type="Retail"):
    return {
        "Buidling": building,          # The workbook's spelling.
        "Unit": unit,
        "Tenant Name": tenant,
        "New Sqft": sqft,
        "Lease Type": lease_type,
        "Floor %": floor_pct,
        "Category %": category_pct,
        "Floor": floor,
        "Type": use_type,
    }


class TestSpaceRecords(unittest.TestCase):
    def test_fields_are_read_off_the_summary_row(self):
        record = lsp.space_records([summary()])[0]
        self.assertEqual(record["Building"], "114 Central")
        self.assertEqual(record["Unit"], "102")
        self.assertEqual(record["Tenant"], "Nabig Sakr")
        self.assertEqual(record["LeaseType"], "NNN")

    def test_square_feet_are_formatted_with_a_thousands_separator(self):
        self.assertEqual(lsp.space_records([summary(sqft=12345)])[0]["Sqft"], "12,345")

    def test_shares_render_as_percentages(self):
        record = lsp.space_records([summary(floor_pct=0.2172)])[0]
        self.assertEqual(record["ShareOfFloor"], "21.72%")

    def test_a_share_already_stored_as_a_whole_number_is_not_multiplied(self):
        self.assertEqual(lsp.space_records([summary(floor_pct=45)])[0]["ShareOfFloor"], "45%")

    def test_share_of_property_is_computed_across_the_building(self):
        rows = [summary(unit="102", sqft=1000), summary(unit="104", sqft=3000, tenant="Other")]
        records = lsp.space_records(rows)
        self.assertEqual(records[0]["ShareOfProperty"], "25%")
        self.assertEqual(records[1]["ShareOfProperty"], "75%")

    def test_buildings_are_kept_separate_when_computing_share(self):
        rows = [summary(building="A", sqft=1000), summary(building="B", sqft=9000, tenant="X")]
        records = lsp.space_records(rows)
        self.assertEqual({r["Building"]: r["ShareOfProperty"] for r in records},
                         {"A": "100%", "B": "100%"})

    def test_placeholder_sqft_is_treated_as_missing(self):
        # The workbook carries combined units as 1e-07 rather than blank. A
        # lease saying "approximately 0 square feet" would be worse than one
        # leaving the token visible.
        record = lsp.space_records([summary(sqft=1e-07)])[0]
        self.assertEqual(record["Sqft"], "")
        self.assertEqual(record["ShareOfProperty"], "")

    def test_placeholder_rows_do_not_dilute_other_shares(self):
        rows = [summary(unit="102", sqft=1000), summary(unit="102-104", sqft=1e-07, tenant="Combo")]
        records = lsp.space_records(rows)
        real = next(r for r in records if r["Unit"] == "102")
        self.assertEqual(real["ShareOfProperty"], "100%")

    def test_rows_without_a_building_are_skipped(self):
        self.assertEqual(lsp.space_records([{"Unit": "102"}]), [])

    def test_junk_input_is_survivable(self):
        self.assertEqual(lsp.space_records(None), [])
        self.assertEqual(lsp.space_records(["not a dict", 7]), [])

    def test_label_distinguishes_units_in_one_building(self):
        records = lsp.space_records([summary(unit="102"), summary(unit="104", tenant="Other")])
        self.assertNotEqual(records[0]["_label"], records[1]["_label"])
        self.assertIn("102", records[0]["_label"])

    def test_records_are_sorted_by_building_then_unit(self):
        rows = [summary(building="B", unit="1"), summary(building="A", unit="2"),
                summary(building="A", unit="1")]
        got = [(r["Building"], r["Unit"]) for r in lsp.space_records(rows)]
        self.assertEqual(got, [("A", "1"), ("A", "2"), ("B", "1")])


class TestCanonicalName(unittest.TestCase):
    def test_exact_name(self):
        self.assertEqual(lsp.canonical_name("Sqft"), "Sqft")

    def test_case_and_spacing_are_ignored(self):
        for spelling in ("sqft", "SQFT", " Sq ft ", "sq_ft"):
            self.assertEqual(lsp.canonical_name(spelling), "Sqft", spelling)

    def test_aliases_resolve(self):
        self.assertEqual(lsp.canonical_name("SquareFeet"), "Sqft")
        self.assertEqual(lsp.canonical_name("Space"), "Unit")

    def test_unknown_name_is_empty(self):
        self.assertEqual(lsp.canonical_name("Bogus"), "")


class TestResolve(unittest.TestCase):
    def setUp(self):
        self.space = lsp.space_records([summary()])[0]

    def test_a_token_inside_prose_is_substituted(self):
        self.assertEqual(
            lsp.resolve("Approximately [Space:Sqft] Total Gross Square Feet", self.space),
            "Approximately 1,063 Total Gross Square Feet",
        )

    def test_several_tokens_in_one_string(self):
        self.assertEqual(
            lsp.resolve("[Space:Tenant] at [Space:Building] Unit [Space:Unit]", self.space),
            "Nabig Sakr at 114 Central Unit 102",
        )

    def test_tokens_are_case_insensitive(self):
        self.assertEqual(lsp.resolve("[space:sqft]", self.space), "1,063")

    def test_spacing_inside_the_brackets_is_tolerated(self):
        self.assertEqual(lsp.resolve("[ Space : Sqft ]", self.space), "1,063")

    def test_an_unknown_field_is_left_alone(self):
        # Leaving it visible is the point: a lease reading "[Space:Bogus]" is
        # obviously unfinished, one reading "" is not.
        self.assertEqual(lsp.resolve("x [Space:Bogus] y", self.space), "x [Space:Bogus] y")

    def test_a_blank_field_is_left_as_the_token(self):
        blank = lsp.space_records([summary(sqft=1e-07)])[0]
        self.assertEqual(lsp.resolve("[Space:Sqft]", blank), "[Space:Sqft]")

    def test_no_space_selected_changes_nothing(self):
        self.assertEqual(lsp.resolve("[Space:Sqft]", None), "[Space:Sqft]")

    def test_text_without_tokens_is_untouched(self):
        self.assertEqual(lsp.resolve("plain text", self.space), "plain text")

    def test_kp_tokens_are_not_disturbed(self):
        # The two token types share a bracket shape and must not eat each other.
        self.assertEqual(lsp.resolve("[KP:Landlord] leases [Space:Sqft]", self.space),
                         "[KP:Landlord] leases 1,063")

    def test_empty_and_none_input(self):
        self.assertEqual(lsp.resolve("", self.space), "")
        self.assertEqual(lsp.resolve(None, self.space), "")


class TestResolveProvisions(unittest.TestCase):
    def setUp(self):
        self.space = lsp.space_records([summary()])[0]

    def test_values_are_resolved(self):
        rows = [{"Field": "Premises Sqft", "Value": "Approximately [Space:Sqft] SF"}]
        self.assertEqual(lsp.resolve_provisions(rows, self.space)[0]["Value"],
                         "Approximately 1,063 SF")

    def test_the_stored_rows_are_not_mutated(self):
        # Resolution happens on the way to Word. Writing it back would bake one
        # unit's numbers into the template permanently.
        rows = [{"Field": "Premises Sqft", "Value": "[Space:Sqft]"}]
        lsp.resolve_provisions(rows, self.space)
        self.assertEqual(rows[0]["Value"], "[Space:Sqft]")

    def test_other_keys_survive(self):
        rows = [{"Field": "X", "Value": "[Space:Sqft]", "Bookmark": "Tx_X", "Include": True}]
        resolved = lsp.resolve_provisions(rows, self.space)[0]
        self.assertEqual(resolved["Bookmark"], "Tx_X")
        self.assertTrue(resolved["Include"])


class TestUnresolved(unittest.TestCase):
    def test_reports_unknown_fields(self):
        space = lsp.space_records([summary()])[0]
        self.assertEqual(lsp.unresolved("[Space:Bogus]", space), ["Bogus"])

    def test_reports_blank_fields(self):
        blank = lsp.space_records([summary(sqft=1e-07)])[0]
        self.assertEqual(lsp.unresolved("[Space:Sqft]", blank), ["Sqft"])

    def test_resolvable_tokens_are_not_reported(self):
        space = lsp.space_records([summary()])[0]
        self.assertEqual(lsp.unresolved("[Space:Sqft]", space), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
