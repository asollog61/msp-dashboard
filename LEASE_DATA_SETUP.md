# Lease data repo — setup

Saved lease documents used to live as one gzipped JSON blob in cell A1 of the
"Lease Documents" tab of the dashboard spreadsheet. They now live as one JSON
file each in a **separate private GitHub repository**, so every save is a commit
with real history and a real diff.

Nothing is deleted from the Google Sheet. It stays readable as a fallback and
as a known-good copy you can always walk back to.

**This takes about ten minutes and all of it is on your side — the code is
already in place and will simply say "not configured" until you finish.**

---

## Why a second repo

Streamlit Cloud rebuilds the app's filesystem from GitHub on every push, every
reboot and every wake from idle. Anything the app writes to its own disk is
gone by the next visit. That is why documents could not just be files.

The data repo has to be **separate from `msp-dashboard`**. Pushing to the app
repo triggers a redeploy, so saving a lease into it would restart the app in
the middle of your edit.

---

## 1. Create the repo

On GitHub, create a new repository:

- **Name:** `msp-lease-data`
- **Owner:** `asollog61`
- **Visibility: Private.** These are real lease terms.
- **Tick "Add a README file"** — the repo must not be empty, or the first save
  fails with a confusing 404.

Nothing else. No `.gitignore`, no licence.

---

## 2. Create a token

GitHub → your avatar → **Settings** → **Developer settings** →
**Personal access tokens** → **Fine-grained tokens** → **Generate new token**.

| Field | Value |
|---|---|
| Token name | `msp-lease-data` |
| Expiration | 1 year (put a reminder in your calendar) |
| Repository access | **Only select repositories** → `msp-lease-data` |
| Permissions → Repository → **Contents** | **Read and write** |

Contents is the only permission needed. Leave everything else alone.

Copy the token when it appears — `github_pat_…`. GitHub shows it once.

> **Do not paste the token into a chat, an issue, or any file inside
> `msp-dashboard-src`.** It goes in exactly the two places below, both of which
> are already gitignored.

---

## 3. Tell the live app

Streamlit Cloud → your app → **Manage app** → **⋮** → **Settings** → **Secrets**.

Add, keeping whatever is already there:

```toml
[lease_data]
token  = "github_pat_paste_yours_here"
repo   = "asollog61/msp-lease-data"
branch = "main"
```

Save. The app restarts on its own.

> If your new repo's default branch is `master` rather than `main`, change
> `branch` to match. GitHub's default is `main` for new repos.

---

## 4. Tell your local copy (optional)

Only needed if you run the dashboard on your own machine. Add the same block to:

```
msp-dashboard-src\.streamlit\secrets.toml
```

That file is gitignored, so it will not be committed.

---

## 5. Migrate

Open the **🧱 Lease Builder** tab. Because the repo is empty, it reads your
existing documents from the Google Sheet and shows a banner:

> *N document(s) are still in the old Google Sheet.*

Click **📦 Move to data repo**.

Migration is all-or-nothing on purpose. If any single document fails to write,
nothing is migrated and the Sheet is left as it was — a half-finished migration
that looked complete is the one failure here that could lose a lease.

---

## 6. Check it worked

1. Open `github.com/asollog61/msp-lease-data`. You should see `index.json` and a
   `documents/` folder with one `.json` per lease.
2. Open any document file. It should be readable JSON with a `"name"` field
   matching what the picker shows.
3. Back in the app, open a document, change something small, press **💾 Save**.
4. On GitHub, the repo's **Commits** should show `Save lease document: <name>`.
   Click it to see exactly what changed.

That last step is the real payoff: you can now see what changed between two
versions of a lease.

---

## How it is laid out

```
msp-lease-data/
  index.json               name -> file, saved_at
  documents/
    MSP-NNN-test-3.json    one saved lease or template
    ...
  published/
    MSP-NNN-Retail-2026.docx
```

**`index.json` is a cache, not the truth.** It exists so the document picker can
be filled with one API call instead of one per document. If it is ever deleted
or corrupted, the app rebuilds it by reading the document files themselves, so a
bad index can never strand a saved lease.

**Filenames are slugs; the real name lives inside the file.** A lease called
`MSP NNN Retail — Restaurant` is stored as `MSP-NNN-Retail-Restaurant.json`,
with the em dash preserved in the file's `"name"` field. Two documents whose
names slug identically get separate files.

---

## Want the files on your computer

Clone the data repo into Dropbox and you get a real browsable folder — the list
of lease JSONs and the folder of generated Word files — that you can open in
Explorer:

```
cd "C:\Dropbox\ASRA Investments\Marion St Properties"
git clone https://github.com/asollog61/msp-lease-data.git
```

Then `git pull` in that folder whenever you want the latest. The app writes to
GitHub; your clone is a local mirror.

---

## If something goes wrong

| Symptom | Cause | Fix |
|---|---|---|
| "The lease data repo is not configured" | No `[lease_data]` in secrets, or `token`/`repo` blank | Redo step 3 |
| "GitHub rejected the token (401)" | Token expired, revoked, or mistyped | Generate a new one, update secrets |
| "GitHub refused the request (403)" | Token lacks **Contents: Read and write**, or the repo is not in its "selected repositories" | Edit the token's permissions |
| Save fails with a 404 | Repo is empty, or `repo`/`branch` is wrong | Confirm the repo has a README and the branch name matches |
| Picker is empty but GitHub has files | Stale index | It self-heals on load; if not, delete `index.json` on GitHub and reload |
| "Reading from the old Google Sheet" won't go away | Repo unreachable or still empty | Check the token, then click **Move to data repo** |

**Nothing here can lose data.** The Google Sheet is never written to again, so
the pre-migration copy of every document stays exactly where it was.
