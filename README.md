# MSP Property Dashboard — Quick Guide

## What Is This?

A web-based property management dashboard for Marion Street Properties. It shows tenant data, vacancy tracking, insurance status, security deposits, and SOPs — all in one place, accessible from any browser.

**Live URL:** https://asollog61-msp-dashboard.streamlit.app

Share that link with anyone (Jason, Richie, etc.) — no login or software needed.

---

## How It All Connects

```
Your Computer (Dropbox)          GitHub                 Streamlit Cloud
┌──────────────────────┐    git push    ┌──────────┐    auto-deploy    ┌──────────────┐
│ msp-dashboard-src/   │ ───────────▶  │  GitHub   │ ──────────────▶  │  Live Web App │
│  app.py              │               │  Repo     │                  │  (public URL) │
│  data/MSP Tenancy.xlsx│              └──────────┘                  └──────────────┘
│  data/Marion_St_SOP… │
│  requirements.txt    │
└──────────────────────┘
```

- **Dropbox folder** = your working copy of all files
- **GitHub** = cloud backup that Streamlit reads from
- **Streamlit Cloud** = free hosting, runs the app, gives you the URL
- **Google Sheets** = stores shared vacancy activity & marketing entries (auto-saves from the app)

---

## Updating Data

### Tenant Data (rents, leases, expirations)
1. Update `data\MSP Tenancy.xlsx` in this folder
2. Push to GitHub (see below)
3. Dashboard refreshes automatically within ~1 minute

### SOP Manual
1. Replace `data\Marion_St_SOP_Manual.pdf` in this folder
2. Push to GitHub
3. Dashboard refreshes automatically

### Vacancy Activity & Marketing Notes
- Just type them directly in the dashboard — they save automatically to Google Sheets
- No git push needed for these

### Dashboard Code / Layout Changes
- Sis updates `app.py` → you push to GitHub

---

## How to Push Updates to GitHub

Open **Command Prompt** and run:

```
cd "C:\Dropbox\OpenClaw\Share Jason\General Procedures\msp-dashboard-src"
git add -A
git commit -m "Updated tenancy data"
git push
```

Or just double-click **`Push_Updates.bat`** in this folder (does the same thing).

### What Those Commands Do
| Command | What It Does |
|---|---|
| `git add -A` | Stages all changed files |
| `git commit -m "message"` | Saves a snapshot with a description |
| `git push` | Uploads to GitHub → Streamlit auto-redeploys |

---

## Quick Push Batch File

Double-click **`Push_Updates.bat`** to push all changes in one click. It will:
1. Stage all changes
2. Commit with a timestamped message
3. Push to GitHub
4. Pause so you can see the result

---

## For a Different App

Same process, different folder:
1. Create a new folder (e.g., `deal-tracker-src`)
2. `git init` inside it
3. Create a new GitHub repo
4. Connect them: `git remote add origin https://github.com/asollog61/new-repo.git`
5. Deploy on Streamlit Cloud pointing to the new repo
6. Same `git add / commit / push` workflow from that folder

**Key rule:** Always `cd` into the correct folder before running git commands.

---

## Files in This Folder

| File | Purpose |
|---|---|
| `app.py` | Dashboard application code |
| `requirements.txt` | Python packages needed to run the app |
| `data/MSP Tenancy.xlsx` | Tenant database (update this to refresh tenant data) |
| `data/Marion_St_SOP_Manual.pdf` | SOP manual (update this to refresh SOPs) |
| `building_expenses.json` | Per-building annual expense config |
| `column_widths.json` | Column width preferences (auto-saved from dashboard) |
| `.streamlit/config.toml` | Streamlit theme/display settings |
| `.gitignore` | Files excluded from GitHub (secrets, etc.) |
| `Push_Updates.bat` | One-click push to GitHub |
| `README.md` | This file |

---

## Troubleshooting

**Dashboard shows old data after pushing?**
→ Wait 1-2 minutes for Streamlit to redeploy. Check the app's "Manage app" menu (bottom-right corner) for deploy status.

**`git push` rejected?**
→ Try `git pull --rebase` then `git push` again. Or `git push --force` if you're the only one making changes.

**Google Sheets data not showing?**
→ The service account credentials are stored in Streamlit Cloud secrets (not in this folder). If they expire, update them in Streamlit Cloud → Manage app → Settings → Secrets.

**Need to change the app URL?**
→ Go to share.streamlit.io → your app → Settings → change the custom subdomain.
