# KNAemailsender

Email sender project — synced with Google Apps Script via GitHub.

## GitHub (done)

- Repo: https://github.com/Invisibl5/KNAemailsender
- Branch: `main`

## Connect Google Apps Script to this repo

Google Apps Script doesn’t “pull from GitHub” in the browser. Use **clasp** (Google’s CLI) so that:

1. **GitHub** = source of truth (you push code here).
2. **clasp** = syncs that code to your Apps Script project (pull from GitHub locally, then `clasp push`).

### 1. Install clasp

```bash
npm install -g @google/clasp
```

### 2. Log in to Google

```bash
clasp login
```

Use the Google account that owns the Apps Script project.

### 3. Link this folder to an Apps Script project

**Option A – You already have an Apps Script project**

```bash
cd "/path/to/Kumon Email Sender"
clasp clone <SCRIPT_ID>
```

Get `SCRIPT_ID`: Apps Script editor → Project settings (gear) → “Script ID”.

**Option B – Create a new Apps Script project from this repo**

```bash
cd "/path/to/Kumon Email Sender"
clasp create --type standalone --title "KNAemailsender"
```

That creates a new project and writes `.clasp.json` here. Then:

```bash
git add .clasp.json
git commit -m "Add clasp project config"
git push
```

(Only do this if you want the script ID in the repo; otherwise add `.clasp.json` to `.gitignore`.)

### 4. Workflow: GitHub → Apps Script

1. **Edit code** (here or after pulling from GitHub).
2. **Push to GitHub:**
   ```bash
   git add .
   git commit -m "Your message"
   git push origin main
   ```
3. **Deploy to Apps Script:**
   ```bash
   clasp push
   ```

So: **push to GitHub** for version control; **clasp push** to update the live Apps Script project.

### 5. Optional: Pull from Apps Script into this repo

If you (or someone) edits in the Apps Script editor and you want those changes here and on GitHub:

```bash
clasp pull
git add .
git commit -m "Sync from Apps Script"
git push origin main
```

### Files clasp expects

- Root-level `.gs` files (e.g. `Code.gs`, `Main.gs`).
- `appsscript.json` (manifest) in the project root.

If your script is only in the Apps Script editor, run `clasp pull` once from this folder (after `clasp clone` or `clasp create`) to get the files, then commit and push to GitHub.

---

**Summary:** Push code to **GitHub**; use **clasp push** to update **Google Apps Script** from this repo.

---

## Auto-update (GitHub Actions)

You can have **Apps Script update automatically** whenever you push to `main`. A workflow in `.github/workflows/sync-to-apps-script.yml` runs `clasp push` on every push to `main`.

### One-time setup

1. **Enable the Google Apps Script API**
   - Go to [script.google.com](https://script.google.com/) → **Settings** (gear in the left sidebar).
   - Turn **ON** the **Google Apps Script API**.

2. **Get your clasp credentials**
   - On your machine: `clasp login` (use the same Google account as your Apps Script project).
   - Copy the contents of your clasp config:
     ```bash
     cat ~/.clasprc.json
     ```
   - Copy the full JSON output (one line is fine).

3. **Add the secret in GitHub**
   - Repo → **Settings** → **Secrets and variables** → **Actions**.
   - **New repository secret** → Name: `CLASPRC_JSON` → Value: paste the `~/.clasprc.json` contents → **Add secret**.

4. **Ensure the repo has `.clasp.json`**
   - The workflow needs `.clasp.json` in the repo so clasp knows which Apps Script project to push to.
   - If you haven’t already: run `clasp clone <SCRIPT_ID>` or `clasp create --type standalone --title "KNAemailsender"` in this folder, then commit and push `.clasp.json`.

After this, every **push to `main`** will trigger the workflow and update your Apps Script project automatically.
