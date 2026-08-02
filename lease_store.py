"""Durable storage for Lease Builder documents.

Streamlit Cloud rebuilds its container from GitHub on every push, every reboot
and every wake from idle, so anything the app writes to its own filesystem is
gone by the next visit. That is why saved documents used to live in cell A1 of
a Google Sheet: it was the only durable store the deployed app could write to.

A single cell is a poor home for a growing set of legal documents. Google caps
a cell at 50,000 characters, which the store already fought by gzipping and by
dropping clause text that matched the template. It also collapses every
document into one opaque blob, so there is no way to see what changed between
two saves.

Documents now live as one JSON file each in a separate private GitHub
repository, reached through the Contents API. Each save is a commit, so a
document has real history and a real diff.

The data repo must be separate from msp-dashboard. A push to the app repo
triggers a Streamlit redeploy, so writing a lease into it would restart the app
in the middle of an edit.

Layout in the data repo:

    index.json               name -> file, saved_at (one call to fill the picker)
    documents/<slug>.json    one saved lease or template
    published/<name>.docx    generated Word files

Two backends implement the same interface. GitHubBackend is what the deployed
app uses. LocalBackend is a plain directory, used by the tests and by anyone
running the app on their own machine, where the filesystem is durable.

This module deliberately does not import Streamlit. Credentials are passed in
by the caller and caching is the caller's job, which keeps the whole thing
testable without a running app.
"""

from __future__ import annotations

import base64
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Protocol

DOCUMENTS_PREFIX = "documents"
PUBLISHED_PREFIX = "published"
INDEX_PATH = "index.json"

# GitHub rejects a Contents API write above 100 MB and gets slow well before
# that. A lease document is tens of kilobytes; anything near this ceiling means
# something has gone wrong upstream, and failing loudly beats a silent truncate.
MAX_DOCUMENT_BYTES = 5 * 1024 * 1024


class StoreError(Exception):
    """A storage operation failed in a way the caller should surface."""


# ---------------------------------------------------------------------------
# Naming
# ---------------------------------------------------------------------------

def slugify(name: str) -> str:
    """A filename that survives git, Windows and a URL, for a display name.

    The slug is not required to round-trip. The authoritative display name is
    stored inside the document under "name", so an em dash or a slash in a
    lease name is never lost — it simply does not appear in the filename.
    """
    text = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9._-]+", "-", text).strip("-._")
    return (text or "untitled")[:100]


def _unique_slug(name: str, taken: set[str]) -> str:
    """A slug not already used by a different document."""
    base = slugify(name)
    if base not in taken:
        return base
    index = 2
    while f"{base}-{index}" in taken:
        index += 1
    return f"{base}-{index}"


# ---------------------------------------------------------------------------
# Backends
# ---------------------------------------------------------------------------

class Backend(Protocol):
    """The minimum a store needs: read, write, delete and list a path."""

    def read(self, path: str) -> bytes | None: ...
    def write(self, path: str, data: bytes, message: str) -> None: ...
    def delete(self, path: str, message: str) -> None: ...
    def list_dir(self, prefix: str) -> list[str]: ...


class LocalBackend:
    """A directory on disk. Durable when the app runs locally; used by tests."""

    def __init__(self, root: str | Path):
        self.root = Path(root)

    def _full(self, path: str) -> Path:
        target = (self.root / path).resolve()
        # A document name should never be able to escape the store.
        if not str(target).startswith(str(self.root.resolve())):
            raise StoreError(f"Refusing to touch a path outside the store: {path}")
        return target

    def read(self, path: str) -> bytes | None:
        target = self._full(path)
        if not target.exists():
            return None
        return target.read_bytes()

    def write(self, path: str, data: bytes, message: str) -> None:
        target = self._full(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)

    def delete(self, path: str, message: str) -> None:
        target = self._full(path)
        if target.exists():
            target.unlink()

    def list_dir(self, prefix: str) -> list[str]:
        folder = self._full(prefix)
        if not folder.is_dir():
            return []
        return sorted(f"{prefix}/{item.name}" for item in folder.iterdir() if item.is_file())


class GitHubBackend:
    """A private GitHub repository, via the Contents API.

    Every write is a commit. The blob sha needed to update a file is cached per
    path, because GitHub requires the current sha on update and fetching it
    again before each write would double the request count.
    """

    API = "https://api.github.com"

    def __init__(self, token: str, repo: str, branch: str = "main", timeout: int = 20):
        if not token:
            raise StoreError("No GitHub token was supplied.")
        if not repo or "/" not in repo:
            raise StoreError(f'Repo must look like "owner/name", got: {repo!r}')
        self.token = token
        self.repo = repo
        self.branch = branch
        self.timeout = timeout
        self._sha: dict[str, str] = {}

    # -- plumbing ----------------------------------------------------------

    def _request(self, method: str, path: str, **kwargs: Any) -> Any:
        import requests

        url = f"{self.API}/repos/{self.repo}/contents/{path}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }
        response = requests.request(method, url, headers=headers, timeout=self.timeout, **kwargs)

        if response.status_code == 404:
            # A missing file is a normal answer to a read. A missing file is
            # never a normal answer to a write: GitHub also returns 404 when the
            # repo or the branch does not exist, and treating that as "nothing
            # there" made a failed save look like a successful one.
            if method == "GET":
                return None
            raise StoreError(
                f"GitHub returned 404 for {method} {path}. Nothing was written. "
                f"Check that “{self.repo}” exists and has a branch named "
                f"“{self.branch}” — a branch set to “master” on a repo whose "
                "default is “main” fails exactly this way."
            )
        if response.status_code == 401:
            raise StoreError("GitHub rejected the token (401). It may be expired or revoked.")
        if response.status_code == 403:
            raise StoreError(
                "GitHub refused the request (403). The token most likely lacks "
                "Contents: Read and write on this repository."
            )
        if response.status_code == 409:
            raise StoreError(
                "The file changed on GitHub since this session loaded it (409). "
                "Reload the page and save again."
            )
        if not response.ok:
            raise StoreError(f"GitHub returned HTTP {response.status_code}: {response.text[:200]}")
        return response.json() if response.content else {}

    # -- interface ---------------------------------------------------------

    def read(self, path: str) -> bytes | None:
        payload = self._request("GET", path, params={"ref": self.branch})
        if not payload or "content" not in payload:
            return None
        self._sha[path] = payload.get("sha", "")
        return base64.b64decode(payload["content"])

    def write(self, path: str, data: bytes, message: str) -> None:
        if len(data) > MAX_DOCUMENT_BYTES:
            raise StoreError(
                f"{path} is {len(data):,} bytes, past the {MAX_DOCUMENT_BYTES:,} byte "
                "ceiling. Nothing was written."
            )
        body: dict[str, Any] = {
            "message": message,
            "content": base64.b64encode(data).decode("ascii"),
            "branch": self.branch,
        }
        sha = self._sha.get(path)
        if sha is None:
            # Not seen this session; ask GitHub whether it exists so an update
            # is not mistaken for a create, which GitHub rejects.
            existing = self._request("GET", path, params={"ref": self.branch})
            sha = (existing or {}).get("sha")
        if sha:
            body["sha"] = sha
        result = self._request("PUT", path, json=body)
        new_sha = ((result or {}).get("content") or {}).get("sha")
        if new_sha:
            self._sha[path] = new_sha

    def delete(self, path: str, message: str) -> None:
        sha = self._sha.get(path)
        if not sha:
            existing = self._request("GET", path, params={"ref": self.branch})
            if not existing:
                return
            sha = existing.get("sha")
        self._request("DELETE", path, json={"message": message, "sha": sha, "branch": self.branch})
        self._sha.pop(path, None)

    def list_dir(self, prefix: str) -> list[str]:
        payload = self._request("GET", prefix, params={"ref": self.branch})
        if not isinstance(payload, list):
            return []
        names = []
        for entry in payload:
            if entry.get("type") == "file":
                names.append(entry["path"])
                if entry.get("sha"):
                    self._sha[entry["path"]] = entry["sha"]
        return sorted(names)


# ---------------------------------------------------------------------------
# The store
# ---------------------------------------------------------------------------

class LeaseStore:
    """Documents by display name, backed by one JSON file each.

    An index.json maps display names to files so the picker can be filled with
    a single request. The index is a cache, not the truth: rebuild_index()
    reconstructs it from the documents themselves, so a corrupt or stale index
    can never strand a saved lease.
    """

    def __init__(self, backend: Backend):
        self.backend = backend

    # -- index -------------------------------------------------------------

    def _read_index(self) -> dict[str, dict[str, Any]]:
        raw = self.backend.read(INDEX_PATH)
        if not raw:
            return {}
        try:
            payload = json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}
        entries = payload.get("documents") if isinstance(payload, dict) else None
        return entries if isinstance(entries, dict) else {}

    def _write_index(self, entries: dict[str, dict[str, Any]], message: str) -> None:
        payload = {
            "version": 1,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "documents": entries,
        }
        self.backend.write(
            INDEX_PATH,
            json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8"),
            message,
        )

    def _discover(self) -> dict[str, dict[str, Any]]:
        """Read every document file and derive the index from it. Writes nothing."""
        entries: dict[str, dict[str, Any]] = {}
        for path in self.backend.list_dir(DOCUMENTS_PREFIX):
            if not path.endswith(".json"):
                continue
            raw = self.backend.read(path)
            if not raw:
                continue
            try:
                document = json.loads(raw.decode("utf-8"))
            except (ValueError, UnicodeDecodeError):
                # One unreadable file must not hide every other document.
                continue
            name = str(document.get("name") or Path(path).stem)
            entries[name] = {
                "file": Path(path).name,
                "saved_at": str(document.get("saved_at", "")),
            }
        return entries

    def rebuild_index(self) -> dict[str, dict[str, Any]]:
        """Reconstruct the index from the document files and persist it."""
        entries = self._discover()
        self._write_index(entries, "Rebuild lease document index")
        return entries

    # -- reading -----------------------------------------------------------

    def list_documents(self) -> dict[str, dict[str, Any]]:
        """Display name -> {file, saved_at}. One request in the common case."""
        entries = self._read_index()
        if entries:
            return entries
        # No index yet, or it was unreadable. Falling back to the files is the
        # safe answer: an empty picker looks exactly like "no saved leases".
        discovered = self._discover()
        if discovered:
            # Documents exist but the index did not describe them. Persist the
            # repair so the next page load is one request again.
            self._write_index(discovered, "Rebuild lease document index")
        # An empty store writes nothing. Otherwise the first save of a brand new
        # repo would commit an empty index, then the document, then the index
        # again — three commits to save one lease.
        return discovered

    def load_document(self, name: str) -> dict[str, Any] | None:
        entries = self.list_documents()
        entry = entries.get(str(name).strip())
        if not entry:
            return None
        raw = self.backend.read(f"{DOCUMENTS_PREFIX}/{entry['file']}")
        if not raw:
            return None
        try:
            return json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as exc:
            raise StoreError(f"“{name}” is not readable JSON: {exc}") from exc

    def load_all(self) -> dict[str, dict[str, Any]]:
        """Every document, keyed by display name.

        One request per document, so this is for migration and verification
        rather than for filling the picker on each page load.
        """
        documents: dict[str, dict[str, Any]] = {}
        for name in self.list_documents():
            document = self.load_document(name)
            if document is not None:
                documents[name] = document
        return documents

    # -- writing -----------------------------------------------------------

    def save_document(self, name: str, payload: dict[str, Any]) -> str:
        """Write one document and update the index. Returns the file name."""
        label = str(name).strip()
        if not label:
            raise StoreError("A document needs a name.")

        entries = self.list_documents()
        existing = entries.get(label)
        if existing:
            filename = existing["file"]
        else:
            taken = {Path(entry["file"]).stem for entry in entries.values()}
            filename = f"{_unique_slug(label, taken)}.json"

        document = dict(payload)
        # The display name lives in the file so the index can always be
        # rebuilt, and so a name with characters the slug drops is never lost.
        document["name"] = label
        document.setdefault("saved_at", datetime.now().isoformat(timespec="seconds"))

        self.backend.write(
            f"{DOCUMENTS_PREFIX}/{filename}",
            json.dumps(document, indent=2, ensure_ascii=False).encode("utf-8"),
            f"Save lease document: {label}",
        )
        entries[label] = {"file": filename, "saved_at": document["saved_at"]}
        self._write_index(entries, f"Index: {label}")
        return filename

    def delete_document(self, name: str) -> bool:
        label = str(name).strip()
        entries = self.list_documents()
        entry = entries.pop(label, None)
        if not entry:
            return False
        self.backend.delete(f"{DOCUMENTS_PREFIX}/{entry['file']}", f"Delete lease document: {label}")
        self._write_index(entries, f"Index: removed {label}")
        return True

    # -- published Word files ---------------------------------------------

    def publish(self, filename: str, data: bytes) -> str:
        """Store a generated .docx so it survives the next redeploy."""
        safe = slugify(Path(filename).stem) + Path(filename).suffix
        path = f"{PUBLISHED_PREFIX}/{safe}"
        self.backend.write(path, data, f"Publish: {filename}")
        return path

    def list_published(self) -> list[str]:
        return [Path(p).name for p in self.backend.list_dir(PUBLISHED_PREFIX)]

    def read_published(self, filename: str) -> bytes | None:
        return self.backend.read(f"{PUBLISHED_PREFIX}/{Path(filename).name}")


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------

def build_store(secrets: Any = None, local_root: str | Path | None = None) -> LeaseStore | None:
    """A store from Streamlit secrets, or a local directory.

    Pass st.secrets from the app. Expects:

        [lease_data]
        token  = "github_pat_..."
        repo   = "asollog61/msp-lease-data"
        branch = "main"

    Returns None when no configuration is present, so the caller can fall back
    to the Google Sheet rather than crash.
    """
    if local_root is not None:
        return LeaseStore(LocalBackend(local_root))

    try:
        config = dict(secrets["lease_data"]) if secrets else {}
    except (KeyError, TypeError):
        return None

    token = str(config.get("token", "") or "")
    repo = str(config.get("repo", "") or "")
    if not token or not repo:
        return None
    return LeaseStore(GitHubBackend(token, repo, str(config.get("branch", "main") or "main")))
