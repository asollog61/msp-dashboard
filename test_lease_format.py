"""Unit tests for the lease formatting settings model.

Run: python -m unittest test_lease_format -v
"""

from __future__ import annotations

import json
import unittest

import lease_format as lf


class TestDefaults(unittest.TestCase):
    def test_defaults_match_spec(self):
        settings = lf.default_settings()
        self.assertEqual(settings["page_size"], "Letter")
        self.assertEqual(settings["body_font"], "Times New Roman")
        self.assertEqual(settings["body_size_pt"], 11.0)
        self.assertEqual(settings["body_alignment"], "justify")
        # Measured off the master template, not the spec's approximations.
        self.assertEqual(settings["kp_label_width_in"], 1.71)
        self.assertEqual(settings["kp_value_width_in"], 4.79)
        self.assertEqual(settings["kp_link_color"], "1F7A33")
        self.assertEqual(settings["subclause_level1_indent_in"], 0.25)
        for margin in ("top", "bottom", "left", "right"):
            self.assertEqual(settings[f"margin_{margin}_in"], 1.0)

    def test_default_settings_is_a_copy(self):
        first = lf.default_settings()
        first["kp_split_fields"].append("Injected")
        first["body_size_pt"] = 99
        second = lf.default_settings()
        self.assertEqual(second["kp_split_fields"], ["Notice Addresses"])
        self.assertEqual(second["body_size_pt"], 11.0)

    def test_defaults_are_json_serializable(self):
        json.dumps(lf.default_settings())

    def test_spec_defaults_produce_only_the_doc_id_warning(self):
        self.assertEqual(
            lf.validate_settings(lf.default_settings()),
            ["No counsel document ID set — the footer will show only the page number."],
        )


class TestNormalize(unittest.TestCase):
    def test_none_and_junk_give_defaults(self):
        for junk in (None, "", 0, [], "not a dict", {"unknown_key": 1}):
            self.assertEqual(lf.normalize_settings(junk), lf.default_settings())

    def test_unknown_keys_are_dropped(self):
        self.assertNotIn("evil", lf.normalize_settings({"evil": "payload"}))

    def test_partial_dict_keeps_the_rest(self):
        settings = lf.normalize_settings({"body_size_pt": 12})
        self.assertEqual(settings["body_size_pt"], 12.0)
        self.assertEqual(settings["body_font"], "Times New Roman")

    def test_sheet_round_trip_strings_are_coerced(self):
        settings = lf.normalize_settings(
            {"body_size_pt": "12.5", "kp_borders": "false", "title_bold": "TRUE"}
        )
        self.assertEqual(settings["body_size_pt"], 12.5)
        self.assertIs(settings["kp_borders"], False)
        self.assertIs(settings["title_bold"], True)

    def test_out_of_range_numbers_are_clamped(self):
        self.assertEqual(lf.normalize_settings({"body_size_pt": 900})["body_size_pt"], 24.0)
        self.assertEqual(lf.normalize_settings({"margin_left_in": -5})["margin_left_in"], 0.25)

    def test_unparseable_number_falls_back(self):
        self.assertEqual(lf.normalize_settings({"body_size_pt": "eleven"})["body_size_pt"], 11.0)
        self.assertEqual(lf.normalize_settings({"margin_top_in": None})["margin_top_in"], 1.0)

    def test_invalid_choice_falls_back(self):
        self.assertEqual(lf.normalize_settings({"page_size": "Tabloid"})["page_size"], "Letter")
        self.assertEqual(
            lf.normalize_settings({"body_alignment": "centered"})["body_alignment"], "justify"
        )

    def test_split_fields_accept_a_comma_string(self):
        settings = lf.normalize_settings({"kp_split_fields": "Notice Addresses, Guarantor , "})
        self.assertEqual(settings["kp_split_fields"], ["Notice Addresses", "Guarantor"])

    def test_normalize_is_idempotent(self):
        once = lf.normalize_settings({"body_size_pt": "13", "page_size": "A4"})
        self.assertEqual(lf.normalize_settings(once), once)

    def test_version_is_always_stamped(self):
        self.assertEqual(lf.normalize_settings({"version": 99})["version"], lf.SETTINGS_VERSION)


class TestColor(unittest.TestCase):
    def test_accepts_common_forms(self):
        for value in ("#1f7a33", "1F7A33", " #1F7A33 "):
            self.assertEqual(lf.normalize_hex_color(value), "1F7A33")

    def test_expands_shorthand(self):
        self.assertEqual(lf.normalize_hex_color("#0f0"), "00FF00")

    def test_junk_falls_back(self):
        self.assertEqual(lf.normalize_hex_color("rebeccapurple"), "1F7A33")
        self.assertEqual(lf.normalize_hex_color(None), "1F7A33")

    def test_color_picker_output_survives_normalize(self):
        self.assertEqual(
            lf.normalize_settings({"kp_link_color": "#aabbcc"})["kp_link_color"], "AABBCC"
        )


class TestGeometry(unittest.TestCase):
    def test_letter_content_width(self):
        self.assertEqual(lf.content_width_in(lf.default_settings()), 6.5)

    def test_margins_reduce_the_content_width(self):
        settings = lf.normalize_settings({"margin_left_in": 1.5, "margin_right_in": 1.5})
        self.assertEqual(lf.content_width_in(settings), 5.5)

    def test_a4_is_narrower_than_letter(self):
        a4 = lf.normalize_settings({"page_size": "A4"})
        self.assertLess(lf.content_width_in(a4), lf.content_width_in(lf.default_settings()))

    def test_default_kp_table_fits_the_page(self):
        settings = lf.default_settings()
        total = settings["kp_label_width_in"] + settings["kp_value_width_in"]
        self.assertLessEqual(total, lf.content_width_in(settings) + 1e-9)

    def test_default_split_halves_equal_the_value_column(self):
        settings = lf.default_settings()
        halves = settings["kp_split_left_width_in"] + settings["kp_split_right_width_in"]
        self.assertAlmostEqual(halves, settings["kp_value_width_in"], places=2)

    def test_default_rent_table_fits_the_page(self):
        settings = lf.default_settings()
        total = (
            settings["rent_col_term_width_in"]
            + settings["rent_col_monthly_width_in"]
            + settings["rent_col_annual_width_in"]
        )
        self.assertLessEqual(total, lf.content_width_in(settings) + 1e-9)


class TestValidate(unittest.TestCase):
    def test_overwide_kp_table_warns(self):
        settings = lf.normalize_settings({"kp_value_width_in": 7.0, "footer_doc_id": "X"})
        self.assertTrue(any("Key Provisions table" in w for w in lf.validate_settings(settings)))

    def test_mismatched_split_halves_warn(self):
        settings = lf.normalize_settings(
            {"kp_split_left_width_in": 1.0, "kp_split_right_width_in": 1.0, "footer_doc_id": "X"}
        )
        self.assertTrue(any("Split Landlord|Tenant" in w for w in lf.validate_settings(settings)))

    def test_overwide_rent_table_warns(self):
        settings = lf.normalize_settings(
            {
                "rent_col_term_width_in": 3.5,
                "rent_col_monthly_width_in": 3.5,
                "rent_col_annual_width_in": 3.5,
                "footer_doc_id": "X",
            }
        )
        self.assertTrue(any("Rent table" in w for w in lf.validate_settings(settings)))

    def test_hanging_indent_larger_than_indent_warns(self):
        settings = lf.normalize_settings(
            {"subclause_level1_indent_in": 0.1, "subclause_level1_hanging_in": 0.5, "footer_doc_id": "X"}
        )
        self.assertTrue(any("hanging indent" in w for w in lf.validate_settings(settings)))

    def test_doc_id_silences_its_warning(self):
        settings = lf.normalize_settings({"footer_doc_id": "4928-4211-2690, v. 2"})
        self.assertEqual(lf.validate_settings(settings), [])


class TestDiffAndPersistence(unittest.TestCase):
    def test_defaults_diff_to_nothing(self):
        self.assertEqual(lf.settings_diff(lf.default_settings()), {})

    def test_diff_reports_only_changes(self):
        diff = lf.settings_diff(lf.normalize_settings({"body_size_pt": 12, "page_size": "A4"}))
        self.assertEqual(diff, {"body_size_pt": 12.0, "page_size": "A4"})

    def test_save_load_round_trip(self):
        """What the template save actually does: diff -> JSON -> normalize."""
        edited = lf.normalize_settings(
            {
                "body_font": "Georgia",
                "margin_left_in": 1.25,
                "footer_doc_id": "4928-4211-2690, v. 2",
                "kp_link_color": "#AA0000",
                "kp_split_fields": ["Notice Addresses", "Guarantor"],
                "kp_borders": False,
            }
        )
        stored = json.loads(json.dumps(lf.settings_diff(edited)))
        self.assertEqual(lf.normalize_settings(stored), edited)

    def test_diff_payload_stays_small(self):
        diff = lf.settings_diff(lf.normalize_settings({"body_size_pt": 12}))
        self.assertLess(len(json.dumps(diff)), 100)


if __name__ == "__main__":
    unittest.main(verbosity=2)
