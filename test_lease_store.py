"""Unit tests for the lease document store.

Everything here runs against LocalBackend in a temporary directory, so the
tests need no GitHub token and no network. The GitHub backend implements the
same four-method interface, so what is proven here about the store's logic
holds for both.

Run: python -m unittest test_lease_store -v
"""

from __future__ import annotations

import base64
import json
import shutil
import tempfile
import unittest
from pathlib import Path

import lease_store as ls


def sample(**overrides):
    """A document shaped like the ones lease_docs.build_document produces."""
    document = {
        "version": 1,
        "key_provisions": [
            {"Field": "Premises", "Value": "1,200 SF", "Include": True, "Alternates": [""] * 10},
        ],
        "sections": {"12": {"include": True, "choice": "Template language"}},
        "rent_schedules": {},
        "format_profile": "",
        "copied_from": "",
        "saved_at": "2026-08-02T10:00:00",
    }
    document.update(overrides)
    return document


class StoreTestCase(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="lease-store-test-"))
        self.store = ls.LeaseStore(ls.LocalBackend(self.root))

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)


# ---------------------------------------------------------------------------
# Naming
# ---------------------------------------------------------------------------

class TestSlugify(unittest.TestCase):
    def test_spaces_become_hyphens(self):
        self.assertEqual(ls.slugify("MSP NNN test 3"), "MSP-NNN-test-3")

    def test_em_dash_and_punctuation_are_dropped(self):
        self.assertEqual(ls.slugify("MSP NNN Retail — Restaurant"), "MSP-NNN-Retail-Restaurant")

    def test_path_separators_cannot_survive(self):
        # A name like this must never produce a nested or escaping path.
        self.assertNotIn("/", ls.slugify("../../etc/passwd"))
        self.assertNotIn("\\", ls.slugify(r"..\..\windows"))

    def test_empty_name_still_yields_a_filename(self):
        self.assertEqual(ls.slugify("!!!"), "untitled")

    def test_long_names_are_capped(self):
        self.assertLessEqual(len(ls.slugify("x" * 500)), 100)


# ---------------------------------------------------------------------------
# Round trip
# ---------------------------------------------------------------------------

class TestSaveAndLoad(StoreTestCase):
    def test_save_then_load_returns_the_same_document(self):
        self.store.save_document("MSP NNN test 3", sample())
        loaded = self.store.load_document("MSP NNN test 3")
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded["sections"], {"12": {"include": True, "choice": "Template language"}})
        self.assertEqual(loaded["key_provisions"][0]["Value"], "1,200 SF")

    def test_display_name_is_stored_inside_the_file(self):
        self.store.save_document("MSP NNN Retail — Restaurant", sample())
        # The slug drops the em dash; the file must still carry the real name,
        # or rebuilding the index would silently rename the document.
        loaded = self.store.load_document("MSP NNN Retail — Restaurant")
        self.assertEqual(loaded["name"], "MSP NNN Retail — Restaurant")

    def test_unknown_document_is_none_not_an_error(self):
        self.assertIsNone(self.store.load_document("Never Saved"))

    def test_saving_twice_overwrites_rather_than_duplicating(self):
        self.store.save_document("Lease A", sample())
        self.store.save_document("Lease A", sample(format_profile="Compact"))
        self.assertEqual(len(self.store.list_documents()), 1)
        self.assertEqual(self.store.load_document("Lease A")["format_profile"], "Compact")

    def test_blank_name_is_rejected(self):
        with self.assertRaises(ls.StoreError):
            self.store.save_document("   ", sample())

    def test_unicode_and_long_clause_text_survive(self):
        # The Sheet backend gzipped to fit a 50k cell; nothing should be lossy now.
        long_text = "The Tenant shall pay — without offset — “Base Rent”. " * 2000
        self.store.save_document("Big", sample(sections={"1": {"include": True, "text": long_text}}))
        loaded = self.store.load_document("Big")
        self.assertEqual(loaded["sections"]["1"]["text"], long_text)
        self.assertGreater(len(long_text), 50_000)


# ---------------------------------------------------------------------------
# Save As
# ---------------------------------------------------------------------------

class TestSaveAs(StoreTestCase):
    def test_two_documents_are_independent_files(self):
        self.store.save_document("Lease A", sample())
        self.store.save_document("Lease B", sample(format_profile="Compact"))
        self.assertEqual(sorted(self.store.list_documents()), ["Lease A", "Lease B"])
        self.assertEqual(self.store.load_document("Lease A")["format_profile"], "")

    def test_names_that_slug_identically_get_separate_files(self):
        # "Lease A" and "Lease/A" both slug to "Lease-A"; one must not clobber
        # the other, which is the failure mode that would lose a saved deal.
        self.store.save_document("Lease A", sample(format_profile="first"))
        self.store.save_document("Lease/A", sample(format_profile="second"))
        entries = self.store.list_documents()
        self.assertEqual(len(entries), 2)
        self.assertNotEqual(entries["Lease A"]["file"], entries["Lease/A"]["file"])
        self.assertEqual(self.store.load_document("Lease A")["format_profile"], "first")
        self.assertEqual(self.store.load_document("Lease/A")["format_profile"], "second")

    def test_editing_a_copy_does_not_touch_the_original(self):
        self.store.save_document("Original", sample())
        copied = self.store.load_document("Original")
        copied["sections"]["12"]["include"] = False
        self.store.save_document("Copy", copied)
        self.assertTrue(self.store.load_document("Original")["sections"]["12"]["include"])


# ---------------------------------------------------------------------------
# Delete
# ---------------------------------------------------------------------------

class TestDelete(StoreTestCase):
    def test_delete_removes_the_document_and_its_index_entry(self):
        self.store.save_document("Doomed", sample())
        self.assertTrue(self.store.delete_document("Doomed"))
        self.assertNotIn("Doomed", self.store.list_documents())
        self.assertIsNone(self.store.load_document("Doomed"))

    def test_deleting_something_absent_reports_false(self):
        self.assertFalse(self.store.delete_document("Never Existed"))

    def test_delete_leaves_other_documents_alone(self):
        self.store.save_document("Keep", sample())
        self.store.save_document("Drop", sample())
        self.store.delete_document("Drop")
        self.assertEqual(list(self.store.list_documents()), ["Keep"])


# ---------------------------------------------------------------------------
# The index is a cache, not the truth
# ---------------------------------------------------------------------------

class TestIndexRecovery(StoreTestCase):
    def test_a_missing_index_is_rebuilt_from_the_files(self):
        self.store.save_document("Lease A", sample())
        self.store.save_document("Lease B", sample())
        (self.root / ls.INDEX_PATH).unlink()
        self.assertEqual(sorted(self.store.list_documents()), ["Lease A", "Lease B"])

    def test_a_corrupt_index_is_rebuilt_rather_than_hiding_documents(self):
        self.store.save_document("Lease A", sample())
        (self.root / ls.INDEX_PATH).write_text("{ this is not json", encoding="utf-8")
        self.assertIn("Lease A", self.store.list_documents())

    def test_rebuild_recovers_the_real_display_name(self):
        self.store.save_document("MSP NNN Retail — Restaurant", sample())
        (self.root / ls.INDEX_PATH).unlink()
        self.assertIn("MSP NNN Retail — Restaurant", self.store.rebuild_index())

    def test_an_unreadable_document_does_not_break_the_whole_list(self):
        self.store.save_document("Good", sample())
        (self.root / ls.DOCUMENTS_PREFIX / "broken.json").write_text("{ nope", encoding="utf-8")
        (self.root / ls.INDEX_PATH).unlink()
        self.assertIn("Good", self.store.list_documents())

    def test_corrupt_document_raises_rather_than_returning_empty(self):
        # Silently returning a blank document would look like a wiped lease.
        self.store.save_document("Lease A", sample())
        entry = self.store.list_documents()["Lease A"]
        (self.root / ls.DOCUMENTS_PREFIX / entry["file"]).write_text("{ broken", encoding="utf-8")
        with self.assertRaises(ls.StoreError):
            self.store.load_document("Lease A")


class TestLoadAll(StoreTestCase):
    def test_load_all_matches_what_was_saved(self):
        self.store.save_document("Lease A", sample(format_profile="A"))
        self.store.save_document("Lease B", sample(format_profile="B"))
        everything = self.store.load_all()
        self.assertEqual(sorted(everything), ["Lease A", "Lease B"])
        self.assertEqual(everything["Lease B"]["format_profile"], "B")

    def test_empty_store_is_an_empty_dict(self):
        self.assertEqual(self.store.load_all(), {})


# ---------------------------------------------------------------------------
# Published Word files
# ---------------------------------------------------------------------------

class TestPublished(StoreTestCase):
    def test_publish_then_read_returns_the_same_bytes(self):
        payload = b"PK\x03\x04 pretend this is a docx"
        self.store.publish("MSP NNN Retail 2026.docx", payload)
        self.assertEqual(self.store.read_published("MSP-NNN-Retail-2026.docx"), payload)

    def test_published_files_are_listed(self):
        self.store.publish("One.docx", b"a")
        self.store.publish("Two.docx", b"b")
        self.assertEqual(sorted(self.store.list_published()), ["One.docx", "Two.docx"])

    def test_extension_is_preserved(self):
        self.store.publish("Some Lease.docx", b"a")
        self.assertTrue(self.store.list_published()[0].endswith(".docx"))


# ---------------------------------------------------------------------------
# Safety
# ---------------------------------------------------------------------------

class TestPathSafety(StoreTestCase):
    def test_a_document_cannot_be_written_outside_the_store(self):
        backend = ls.LocalBackend(self.root)
        with self.assertRaises(ls.StoreError):
            backend.write("../escaped.json", b"{}", "nope")

    def test_a_traversal_name_stays_inside_documents(self):
        self.store.save_document("../../escape", sample())
        written = list((self.root / ls.DOCUMENTS_PREFIX).glob("*.json"))
        self.assertEqual(len(written), 1)
        self.assertFalse((self.root.parent / "escape.json").exists())


# ---------------------------------------------------------------------------
# GitHub backend
#
# Exercised against a stub rather than the network, so these run offline and in
# CI. What is being checked is the protocol: that an update sends the blob sha
# GitHub requires, that a create does not, and that each failure code turns
# into a message naming the actual fix.
# ---------------------------------------------------------------------------

class FakeResponse:
    def __init__(self, status_code, payload=None, text=""):
        self.status_code = status_code
        self._payload = payload
        self.text = text
        self.content = b"x" if payload is not None else b""

    @property
    def ok(self):
        return 200 <= self.status_code < 300

    def json(self):
        return self._payload


class FakeRequests:
    """Stands in for the requests module, recording what it was asked to do."""

    def __init__(self, responses):
        self.responses = responses          # (method, path) -> FakeResponse or list
        self.calls = []

    def request(self, method, url, **kwargs):
        path = url.split("/contents/", 1)[1]
        self.calls.append((method, path, kwargs.get("json")))
        entry = self.responses.get((method, path), FakeResponse(404))
        if isinstance(entry, list):
            return entry.pop(0) if entry else FakeResponse(404)
        return entry


class GitHubTestCase(unittest.TestCase):
    def make(self, responses):
        import sys
        fake = FakeRequests(responses)
        self._saved = sys.modules.get("requests")
        sys.modules["requests"] = fake
        self.addCleanup(self._restore)
        return ls.GitHubBackend("tok", "owner/repo", "main"), fake

    def _restore(self):
        import sys
        if self._saved is None:
            sys.modules.pop("requests", None)
        else:
            sys.modules["requests"] = self._saved


class TestGitHubBackend(GitHubTestCase):
    def test_read_decodes_base64_content(self):
        payload = {"content": base64.b64encode(b'{"name": "A"}').decode(), "sha": "abc"}
        backend, _ = self.make({("GET", "documents/a.json"): FakeResponse(200, payload)})
        self.assertEqual(backend.read("documents/a.json"), b'{"name": "A"}')

    def test_missing_file_reads_as_none(self):
        backend, _ = self.make({})
        self.assertIsNone(backend.read("documents/nope.json"))

    def test_creating_a_new_file_sends_no_sha(self):
        backend, fake = self.make({
            ("GET", "documents/new.json"): FakeResponse(404),
            ("PUT", "documents/new.json"): FakeResponse(201, {"content": {"sha": "s1"}}),
        })
        backend.write("documents/new.json", b"{}", "create")
        put = [c for c in fake.calls if c[0] == "PUT"][0]
        self.assertNotIn("sha", put[2])
        self.assertEqual(put[2]["branch"], "main")

    def test_updating_an_existing_file_sends_its_sha(self):
        # GitHub rejects an update that omits the current sha, so this is the
        # difference between saving and a 409 on every save after the first.
        backend, fake = self.make({
            ("GET", "documents/a.json"): FakeResponse(200, {"content": "", "sha": "old"}),
            ("PUT", "documents/a.json"): FakeResponse(200, {"content": {"sha": "new"}}),
        })
        backend.write("documents/a.json", b"{}", "update")
        self.assertEqual([c for c in fake.calls if c[0] == "PUT"][0][2]["sha"], "old")

    def test_sha_from_a_write_is_reused_by_the_next_write(self):
        backend, fake = self.make({
            ("GET", "documents/a.json"): FakeResponse(404),
            ("PUT", "documents/a.json"): [
                FakeResponse(201, {"content": {"sha": "s1"}}),
                FakeResponse(200, {"content": {"sha": "s2"}}),
            ],
        })
        backend.write("documents/a.json", b"{}", "first")
        backend.write("documents/a.json", b"{}", "second")
        puts = [c for c in fake.calls if c[0] == "PUT"]
        self.assertEqual(puts[1][2]["sha"], "s1")
        # One GET only: the second write must not re-fetch what it already knows.
        self.assertEqual(len([c for c in fake.calls if c[0] == "GET"]), 1)

    def test_list_dir_returns_only_files(self):
        backend, _ = self.make({("GET", "documents"): FakeResponse(200, [
            {"type": "file", "path": "documents/a.json", "sha": "s1"},
            {"type": "dir", "path": "documents/sub", "sha": "s2"},
            {"type": "file", "path": "documents/b.json", "sha": "s3"},
        ])})
        self.assertEqual(backend.list_dir("documents"), ["documents/a.json", "documents/b.json"])

    def test_delete_sends_the_sha(self):
        backend, fake = self.make({
            ("GET", "documents/a.json"): FakeResponse(200, {"content": "", "sha": "old"}),
            ("DELETE", "documents/a.json"): FakeResponse(200, {}),
        })
        backend.delete("documents/a.json", "remove")
        self.assertEqual([c for c in fake.calls if c[0] == "DELETE"][0][2]["sha"], "old")

    def test_deleting_a_missing_file_is_not_an_error(self):
        backend, fake = self.make({})
        backend.delete("documents/gone.json", "remove")
        self.assertEqual([c for c in fake.calls if c[0] == "DELETE"], [])

    def test_a_404_on_write_raises_instead_of_pretending_to_succeed(self):
        # The branch in secrets pointing at "master" on a "main" repo lands
        # here. Returning quietly made a failed save look like a saved lease.
        backend, _ = self.make({
            ("GET", "documents/a.json"): FakeResponse(404),
            ("PUT", "documents/a.json"): FakeResponse(404, {}, "Not Found"),
        })
        with self.assertRaises(ls.StoreError) as caught:
            backend.write("documents/a.json", b"{}", "write")
        self.assertIn("branch", str(caught.exception).lower())

    def test_a_404_on_read_is_still_just_missing(self):
        backend, _ = self.make({})
        self.assertIsNone(backend.read("documents/a.json"))

    def test_a_save_through_the_store_surfaces_the_404(self):
        backend, _ = self.make({
            ("GET", ls.INDEX_PATH): FakeResponse(404),
            ("GET", "documents"): FakeResponse(404),
            ("GET", "documents/Lease-A.json"): FakeResponse(404),
            ("PUT", "documents/Lease-A.json"): FakeResponse(404, {}, "Not Found"),
        })
        with self.assertRaises(ls.StoreError):
            ls.LeaseStore(backend).save_document("Lease A", {"sections": {}})

    def test_oversized_write_is_refused_before_the_request(self):
        backend, fake = self.make({})
        with self.assertRaises(ls.StoreError):
            backend.write("documents/a.json", b"x" * (ls.MAX_DOCUMENT_BYTES + 1), "too big")
        self.assertEqual(fake.calls, [])


class TestGitHubErrors(GitHubTestCase):
    def assert_message(self, status, needle):
        backend, _ = self.make({("GET", "documents/a.json"): FakeResponse(status, {}, "detail")})
        with self.assertRaises(ls.StoreError) as caught:
            backend.read("documents/a.json")
        self.assertIn(needle, str(caught.exception).lower())

    def test_401_names_the_token(self):
        self.assert_message(401, "token")

    def test_403_names_the_permission_to_fix(self):
        self.assert_message(403, "contents")

    def test_409_tells_you_to_reload(self):
        self.assert_message(409, "reload")

    def test_500_surfaces_the_status(self):
        self.assert_message(500, "500")


class TestStoreOverGitHub(GitHubTestCase):
    def test_a_save_writes_the_document_then_the_index(self):
        backend, fake = self.make({
            ("GET", ls.INDEX_PATH): FakeResponse(404),
            ("GET", "documents"): FakeResponse(200, []),
            ("PUT", ls.INDEX_PATH): FakeResponse(201, {"content": {"sha": "i1"}}),
            ("GET", "documents/Lease-A.json"): FakeResponse(404),
            ("PUT", "documents/Lease-A.json"): FakeResponse(201, {"content": {"sha": "d1"}}),
        })
        store = ls.LeaseStore(backend)
        store.save_document("Lease A", {"sections": {}})
        written = [c[1] for c in fake.calls if c[0] == "PUT"]
        # Document first: an index naming a file that failed to write would be
        # worse than an index that is briefly missing an entry.
        self.assertEqual(written, ["documents/Lease-A.json", ls.INDEX_PATH])

    def test_the_saved_body_carries_the_display_name(self):
        backend, fake = self.make({
            ("GET", ls.INDEX_PATH): FakeResponse(404),
            ("GET", "documents"): FakeResponse(200, []),
            ("PUT", ls.INDEX_PATH): FakeResponse(201, {"content": {"sha": "i1"}}),
            ("GET", "documents/Lease-A.json"): FakeResponse(404),
            ("PUT", "documents/Lease-A.json"): FakeResponse(201, {"content": {"sha": "d1"}}),
        })
        ls.LeaseStore(backend).save_document("Lease A", {"sections": {}})
        put = [c for c in fake.calls if c[0] == "PUT" and c[1].startswith("documents/")][0]
        body = json.loads(base64.b64decode(put[2]["content"]).decode())
        self.assertEqual(body["name"], "Lease A")
        self.assertTrue(body["saved_at"])


class TestBuildStore(unittest.TestCase):
    def test_local_root_builds_a_local_store(self):
        root = Path(tempfile.mkdtemp(prefix="lease-store-build-"))
        try:
            store = ls.build_store(local_root=root)
            self.assertIsInstance(store.backend, ls.LocalBackend)
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_missing_config_returns_none_so_the_app_can_fall_back(self):
        self.assertIsNone(ls.build_store(secrets={}))
        self.assertIsNone(ls.build_store(secrets=None))

    def test_partial_config_returns_none(self):
        self.assertIsNone(ls.build_store(secrets={"lease_data": {"repo": "a/b"}}))
        self.assertIsNone(ls.build_store(secrets={"lease_data": {"token": "x"}}))

    def test_a_malformed_repo_is_rejected_loudly(self):
        with self.assertRaises(ls.StoreError):
            ls.GitHubBackend("token", "not-a-repo")


if __name__ == "__main__":
    unittest.main(verbosity=2)
