"""Tests for the merged lease-document store.

Run: python -m unittest test_lease_docs -v
"""

from __future__ import annotations

import json
import unittest

import lease_docs as ld


def template(name_rows, sections=None, profile="MSP House Style"):
    return {
        "base_template": "MSP NNN Retail",
        "key_provisions": name_rows,
        "sections": sections or {"1": {"include": True, "choice": "Template language"}},
        "format_profile": profile,
        "saved_at": "2026-07-30T10:00:00",
    }


def provision(field, value="", alternates=None, include=True, bookmark=None):
    return {
        "Group": "Optional",
        "Include": include,
        "Field": field,
        "Value": value,
        "Alternates": alternates or [],
        "Link": False,
        "Section": "Section 1 — Premises",
        "Bookmark": bookmark or f"Tx_{field.replace(' ', '')}",
    }


class TestNormalize(unittest.TestCase):
    def test_junk_becomes_an_empty_document(self):
        for junk in (None, "", 0, [], "nonsense"):
            doc = ld.normalize_document(junk)
            self.assertEqual(doc["key_provisions"], [])
            self.assertEqual(doc["sections"], {})

    def test_alternates_are_padded_to_ten(self):
        doc = ld.normalize_document({"key_provisions": [provision("Rent", alternates=["a", "b"])]})
        self.assertEqual(len(doc["key_provisions"][0]["Alternates"]), 10)
        self.assertEqual(doc["key_provisions"][0]["Alternates"][:2], ["a", "b"])

    def test_extra_alternates_are_truncated_not_dropped_silently(self):
        doc = ld.normalize_document({"key_provisions": [provision("Rent", alternates=list("abcdefghijklmno"))]})
        self.assertEqual(len(doc["key_provisions"][0]["Alternates"]), 10)

    def test_section_text_only_kept_when_it_departs(self):
        doc = ld.normalize_document({"sections": {"1": {"include": True, "text": "   "}}})
        self.assertNotIn("text", doc["sections"]["1"])
        doc = ld.normalize_document({"sections": {"1": {"include": True, "text": "Custom."}}})
        self.assertEqual(doc["sections"]["1"]["text"], "Custom.")

    def test_group_falls_back_to_optional(self):
        doc = ld.normalize_document({"key_provisions": [{"Field": "X", "Group": "Weird"}]})
        self.assertEqual(doc["key_provisions"][0]["Group"], "Optional")

    def test_document_round_trips_through_json(self):
        doc = ld.build_document([provision("Rent", "5,000", ["4,500"])],
                                {"1": {"include": True, "choice": "Template language"}},
                                {}, "MSP House Style")
        self.assertEqual(ld.normalize_document(json.loads(json.dumps(doc))), doc)

    def test_store_drops_blank_names(self):
        self.assertEqual(list(ld.normalize_store({"  ": {}, "Real": {}})), ["Real"])


class TestMigration(unittest.TestCase):
    def test_templates_come_across_whole(self):
        docs, notes = ld.migrate_stores({"Triple Net": template([provision("Rent", "5,000", ["4,500", "6,000"])])}, {})
        self.assertEqual(list(docs), ["Triple Net"])
        self.assertEqual(docs["Triple Net"]["key_provisions"][0]["Alternates"][:2], ["4,500", "6,000"])
        self.assertEqual(notes, [])

    def test_lease_gets_its_parents_alternates_baked_in(self):
        """The whole point of the migration: a lease stored only its selections."""
        templates = {"Triple Net": template([provision("Rent", "5,000", ["4,500", "6,000"])])}
        leases = {"ABC Bakery": {
            "template_name": "Triple Net",
            "key_provisions": [{"Bookmark": "Tx_Rent", "Include": True, "Value": "6,000"}],
            "sections": {"1": {"include": False, "choice": "Template language"}},
            "saved_at": "2026-07-31T09:00:00",
        }}
        docs, notes = ld.migrate_stores(templates, leases)
        bakery = docs["ABC Bakery"]
        self.assertEqual(bakery["key_provisions"][0]["Value"], "6,000")
        self.assertEqual(bakery["key_provisions"][0]["Alternates"][:2], ["4,500", "6,000"])
        self.assertFalse(bakery["sections"]["1"]["include"])
        self.assertEqual(bakery["copied_from"], "Triple Net")
        self.assertEqual(notes, [])

    def test_lease_keeps_provisions_its_parent_never_had(self):
        templates = {"Triple Net": template([provision("Rent")])}
        leases = {"ABC": {"template_name": "Triple Net",
                          "key_provisions": [{"Bookmark": "Custom_9", "Field": "Signage", "Value": "yes"}]}}
        docs, _ = ld.migrate_stores(templates, leases)
        fields = [row["Field"] for row in docs["ABC"]["key_provisions"]]
        self.assertIn("Signage", fields)
        self.assertIn("Rent", fields)

    def test_missing_parent_is_reported_not_swallowed(self):
        leases = {"Orphan": {"template_name": "Deleted Template",
                             "key_provisions": [{"Bookmark": "Tx_Rent", "Value": "5,000"}]}}
        docs, notes = ld.migrate_stores({}, leases)
        self.assertIn("Orphan", docs)
        self.assertEqual(len(notes), 1)
        self.assertIn("Deleted Template", notes[0])
        # The lease's own selections still survive.
        self.assertEqual(docs["Orphan"]["key_provisions"][0]["Value"], "5,000")

    def test_name_clash_suffixes_the_lease_rather_than_overwriting(self):
        templates = {"Retail": template([provision("Rent", "1")])}
        leases = {"Retail": {"template_name": "Retail",
                             "key_provisions": [{"Bookmark": "Tx_Rent", "Value": "2"}]}}
        docs, notes = ld.migrate_stores(templates, leases)
        self.assertEqual(sorted(docs), ["Retail", "Retail (2)"])
        self.assertEqual(docs["Retail"]["key_provisions"][0]["Value"], "1")
        self.assertEqual(docs["Retail (2)"]["key_provisions"][0]["Value"], "2")
        self.assertTrue(notes)

    def test_lease_with_no_parent_named(self):
        docs, notes = ld.migrate_stores({}, {"Solo": {"key_provisions": [{"Field": "Rent"}]}})
        self.assertIn("Solo", docs)
        self.assertEqual(notes, [])

    def test_empty_stores(self):
        self.assertEqual(ld.migrate_stores({}, {}), ({}, []))
        self.assertEqual(ld.migrate_stores(None, None), ({}, []))

    def test_migration_is_idempotent(self):
        templates = {"Triple Net": template([provision("Rent", "5,000", ["4,500"])])}
        leases = {"ABC": {"template_name": "Triple Net",
                          "key_provisions": [{"Bookmark": "Tx_Rent", "Value": "6,000"}]}}
        once, _ = ld.migrate_stores(templates, leases)
        twice, _ = ld.migrate_stores(once, {})
        self.assertEqual(sorted(once), sorted(twice))
        self.assertEqual(once["ABC"]["key_provisions"], twice["ABC"]["key_provisions"])

    def test_migrated_lease_no_longer_needs_a_parent(self):
        """After migration the parent can be deleted with no loss."""
        templates = {"Triple Net": template([provision("Rent", "5,000", ["4,500", "6,000"])])}
        leases = {"ABC": {"template_name": "Triple Net",
                          "key_provisions": [{"Bookmark": "Tx_Rent", "Value": "6,000"}]}}
        docs, _ = ld.migrate_stores(templates, leases)
        standalone = json.loads(json.dumps(docs["ABC"]))   # parent now gone
        self.assertEqual(ld.normalize_document(standalone)["key_provisions"][0]["Alternates"][:2],
                         ["4,500", "6,000"])


class TestCopy(unittest.TestCase):
    def test_copy_is_deep(self):
        original = ld.normalize_document(template([provision("Rent", "5,000", ["4,500"])]))
        duplicate = ld.copy_document(original, "Triple Net")
        duplicate["key_provisions"][0]["Value"] = "9,999"
        duplicate["key_provisions"][0]["Alternates"][0] = "changed"
        duplicate["sections"]["1"]["include"] = False
        self.assertEqual(original["key_provisions"][0]["Value"], "5,000")
        self.assertEqual(original["key_provisions"][0]["Alternates"][0], "4,500")
        self.assertTrue(original["sections"]["1"]["include"])

    def test_copy_records_its_origin_and_clears_the_timestamp(self):
        duplicate = ld.copy_document(template([provision("Rent")]), "Triple Net")
        self.assertEqual(duplicate["copied_from"], "Triple Net")
        self.assertEqual(duplicate["saved_at"], "")

    def test_copy_keeps_the_format_profile(self):
        duplicate = ld.copy_document(template([provision("Rent")], profile="Compact"), "X")
        self.assertEqual(duplicate["format_profile"], "Compact")

    def test_unique_name(self):
        taken = {"Retail", "Retail (2)"}
        self.assertEqual(ld.unique_name("Retail", taken), "Retail (3)")
        self.assertEqual(ld.unique_name("New", taken), "New")
        self.assertEqual(ld.unique_name("  ", taken), "Untitled")


class TestChoiceOptions(unittest.TestCase):
    def test_only_filled_slots_are_offered(self):
        row = {"Alternates": ["first", "", "third", "", "", "", "", "", "", ""]}
        self.assertEqual(ld.choice_options(row), ["Current Value", "Alt 1", "Alt 3"])

    def test_no_alternates_means_only_the_default(self):
        self.assertEqual(ld.choice_options({"Alternates": []}), ["Current Value"])

    def test_whitespace_is_not_a_choice(self):
        self.assertEqual(ld.choice_options({"Alternates": ["   ", "real"]}),
                         ["Current Value", "Alt 2"])


class TestNormalizeChoice(unittest.TestCase):
    def test_a_missing_choice_defaults(self):
        self.assertEqual(ld.normalize_choice({"Alternates": ["a"]}), "Current Value")

    def test_a_valid_choice_is_kept(self):
        self.assertEqual(ld.normalize_choice({"Choice": "Alt 1", "Alternates": ["a"]}), "Alt 1")

    def test_a_choice_pointing_at_a_blank_slot_falls_back(self):
        # The alternate was emptied after being chosen. Printing nothing in a
        # key provision would be worse than printing the default.
        self.assertEqual(ld.normalize_choice({"Choice": "Alt 2", "Alternates": ["a", ""]}),
                         "Current Value")

    def test_a_nonsense_choice_falls_back(self):
        self.assertEqual(ld.normalize_choice({"Choice": "banana", "Alternates": ["a"]}),
                         "Current Value")

    def test_out_of_range_choice_falls_back(self):
        self.assertEqual(ld.normalize_choice({"Choice": "Alt 11", "Alternates": ["a"]}),
                         "Current Value")


class TestApplyChoice(unittest.TestCase):
    def test_choosing_an_alternate_sets_the_value(self):
        row = ld.apply_choice({"Value": "default", "Choice": "Alt 2",
                               "Alternates": ["one", "two", "three"]})
        self.assertEqual(row["Value"], "two")
        self.assertEqual(row["Choice"], "Alt 2")

    def test_current_value_leaves_the_typed_value_alone(self):
        row = ld.apply_choice({"Value": "typed", "Choice": "Current Value",
                               "Alternates": ["one"]})
        self.assertEqual(row["Value"], "typed")

    def test_reapplying_picks_up_an_edited_alternate(self):
        # Retyping a chosen alternate must move the value with it, or the lease
        # prints a copy of text that no longer exists anywhere in the row.
        row = {"Value": "default", "Choice": "Alt 1", "Alternates": ["first"]}
        row = ld.apply_choice(row)
        self.assertEqual(row["Value"], "first")
        row["Alternates"] = ["first, amended"]
        self.assertEqual(ld.apply_choice(row)["Value"], "first, amended")

    def test_a_blanked_alternate_restores_the_default(self):
        row = {"Value": "default", "Choice": "Alt 1", "Alternates": [""]}
        resolved = ld.apply_choice(row)
        self.assertEqual(resolved["Value"], "default")
        self.assertEqual(resolved["Choice"], "Current Value")

    def test_apply_is_idempotent(self):
        row = ld.apply_choice({"Value": "d", "Choice": "Alt 1", "Alternates": ["one"]})
        self.assertEqual(ld.apply_choice(ld.apply_choice(row)), row)

    def test_the_original_row_is_not_mutated(self):
        original = {"Value": "default", "Choice": "Alt 1", "Alternates": ["one"]}
        ld.apply_choice(original)
        self.assertEqual(original["Value"], "default")

    def test_case_and_spacing_in_a_choice_still_match(self):
        row = ld.apply_choice({"Value": "d", "Choice": "alt 1", "Alternates": ["one"]})
        self.assertEqual(row["Value"], "one")


class TestChoiceSurvivesNormalisation(unittest.TestCase):
    def test_normalize_provision_keeps_a_valid_choice(self):
        row = ld.normalize_provision({"Field": "Rent", "Value": "d", "Choice": "Alt 1",
                                      "Alternates": ["one"]})
        self.assertEqual(row["Choice"], "Alt 1")

    def test_normalize_provision_supplies_a_default(self):
        self.assertEqual(ld.normalize_provision({"Field": "Rent"})["Choice"], "Current Value")

    def test_a_document_round_trips_its_choices(self):
        doc = ld.normalize_document({
            "key_provisions": [{"Field": "Rent", "Value": "d", "Choice": "Alt 2",
                                "Alternates": ["one", "two"]}],
            "sections": {},
        })
        self.assertEqual(ld.normalize_document(doc)["key_provisions"][0]["Choice"], "Alt 2")

    def test_a_copy_keeps_the_choice(self):
        doc = ld.normalize_document({
            "key_provisions": [{"Field": "Rent", "Value": "d", "Choice": "Alt 1",
                                "Alternates": ["one"]}],
            "sections": {},
        })
        self.assertEqual(ld.copy_document(doc)["key_provisions"][0]["Choice"], "Alt 1")

    def test_documents_saved_before_choice_existed_still_load(self):
        legacy = {"key_provisions": [{"Field": "Rent", "Value": "d", "Alternates": ["one"]}],
                  "sections": {}}
        self.assertEqual(ld.normalize_document(legacy)["key_provisions"][0]["Choice"],
                         "Current Value")


class TestDescribe(unittest.TestCase):
    def test_counts_what_is_used(self):
        doc = ld.normalize_document(template(
            [provision("Rent", include=True, alternates=["a"]), provision("Signage", include=False)],
            {"1": {"include": True}, "2": {"include": False}},
        ))
        text = ld.describe_document(doc)
        self.assertIn("1/2 provisions", text)
        self.assertIn("1/2 sections", text)
        self.assertIn("1 with alternates", text)

    def test_empty_document(self):
        self.assertIn("0/0 provisions", ld.describe_document({}))


if __name__ == "__main__":
    unittest.main(verbosity=2)
