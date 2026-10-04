# High School Registry

A shared queue for linking the way each recruiting service writes a high school's name to its
College Board (ETS/CEEB) code. Part of a Virginia Tech capstone project.

This site holds **school-level data only**: school names as each source writes them, the source's
school ids, state, and (for ESPN) the school's address. It also has a count of how many athletes list
each school, which sets the queue order, and the public College Board high school list. No athlete
names, profiles or other personal data are in this repo.

Team decisions are saved to a private Google Sheet through a Google Apps Script web app. Saving
needs a team passcode, which is not stored here.

## One-time setup (owner)

1. **Create the Sheet.** In Google Drive, create a blank Google Sheet, e.g. "HS Registry decisions".
2. **Add the script.** In the Sheet, go to **Extensions → Apps Script**. Replace the contents of
   `Code.gs` with [`apps_script/Code.gs`](apps_script/Code.gs), then save.
3. **Set the passcode.** In the Apps Script editor, go to **Project Settings** (gear icon) → **Script Properties**
   → **Add script property**. Set the name to `PASSCODE` and the value to a passphrase you'll give your teammates.
4. **Deploy.** Click **Deploy → New deployment** → type **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**

   Then authorize when Google asks. It warns that the app is unverified because it's your own script: choose
   **Advanced → Go to (project)**. Copy the **Web app URL** (ends in `/exec`).
5. **Connect the page.** Paste that URL into `config.js` as `endpoint`, then commit and push.
6. **Turn on Pages.** On GitHub, go to **Settings → Pages**. Set Source to *Deploy from a branch*, branch `main`, folder `/ (root)`.

Send teammates the Pages link and the passcode (privately). Each person enters their name and the passcode
once. Their name is recorded with every decision.

If you change `Code.gs` later, use **Deploy → Manage deployments → Edit → New version** so the URL stays the same.

## How it works

- **Same school** links the name to a CEEB code. **Needs research** sends it to a separate research list.
  **Not a high school** excludes it. **Skip** is personal and temporary.
- Every save writes one row per school name to the `decisions` tab (latest decision wins) and appends to the
  `log` tab, so the full history is kept. Pages check for teammates' decisions every 15 seconds.
- **Export → Download decisions (JSON)** produces the file `export_bigquery.py` (in the main project) turns into
  the BigQuery tables `hs_master` and `hs_alias`, plus the separate research file.

## Updating the queue

New schools come from the main project's `hs_registry/build_queue.py`, which rewrites `data/`. Commit and push
`data/` to publish them. Decisions already made are kept, because they're keyed by a stable `alias_id`.
