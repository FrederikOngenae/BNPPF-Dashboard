# Publishing the MyGuide dashboard via GitHub Pages → Webflow

Written for your repo: **`FrederikOngenae/BNPPF-Dashboard`**. Every URL below is
the real one — no placeholders to swap out.

The dashboard can't be pasted into Webflow directly: Webflow's Code Embed caps
at 50,000 characters (the file is ~45,000) and the dashboard's CSS would leak
into the rest of your page. So GitHub Pages serves the file and Webflow shows it
in an iframe that resizes itself to fit.

One thing to know before you start: **`github.com` is not a web host.** It shows
you source code and refuses to be put in an iframe — that's the "heeft de
verbinding geweigerd" error. GitHub Pages is a different service on a different
domain (`github.io`) that serves the same files as an actual website. Turning it
on is step 1.

---

## Step 1 · Turn on GitHub Pages

1. Go to <https://github.com/FrederikOngenae/BNPPF-Dashboard/settings/pages>
2. Under **Build and deployment → Source**, choose **Deploy from a branch**.
3. Under **Branch**, set the dropdowns to **`main`** and **`/ (root)`**, then
   click **Save**.
4. Wait. First build is usually 30–60 seconds but can take a few minutes. The
   Pages settings page will show a green banner with your live URL when it's
   ready. You can watch progress under the repo's **Actions** tab.

Your site root becomes:

```
https://frederikongenae.github.io/BNPPF-Dashboard/
```

## Step 2 · Check the file loads

You uploaded the folder rather than the loose files, so the dashboard sits one
level down in `MyGuide-webflow/`. Open this in a normal browser tab:

```
https://frederikongenae.github.io/BNPPF-Dashboard/MyGuide-webflow/MyGuide-dashboard.html
```

You should see the dashboard render, dark green header and all. **Do not
continue until this works** — if it doesn't render here it won't render in
Webflow either.

- *404?* The build hasn't finished, or the path is wrong. Check capitalisation:
  `MyGuide-webflow` and `MyGuide-dashboard.html` are case-sensitive.
- *Source code instead of a page?* You're on a `github.com` or
  `raw.githubusercontent.com` URL, not `github.io`. Look at the domain.

**Optional — shorten the URL.** If you'd rather not have `MyGuide-webflow/` in
the path, move the four files up to the repo root (in GitHub: open each file →
pencil icon → change the filename to strip the folder prefix → commit). Then
drop `MyGuide-webflow/` from every URL below. Purely cosmetic.

## Step 3 · Paste the embed into Webflow

Drag a **Code embed** element onto your page and paste the contents of
`webflow-embed-snippet.html`. It already contains your URLs — nothing to
replace.

Two things that trip people up:

- The embed shows a grey placeholder on the Designer canvas. That's normal —
  use **Preview**, or publish, to see it.
- Don't wrap the snippet in `<html>`, `<head>` or `<body>` tags. Webflow
  rejects those and it breaks the page layout.

## Step 4 · Password-protect the Webflow page

Page settings → **General** → *Password protection* → set a password, then
publish. Also set the page to `noindex` under the SEO tab.

Page-level password protection needs a paid **Site plan** on that Webflow site.
Site-wide protection is available on all paid Site plans.

---

## Visibility — read this one

Your repo is **public**, which means the dashboard file is publicly reachable at
its `github.io` URL by anyone who has it, regardless of the Webflow password.
The Webflow password protects the Webflow *page*, not the underlying file.

For internal production percentages on a client project this is probably
acceptable, but decide deliberately rather than by accident. If it isn't:

- **Make the repo private.** GitHub Pages on a private repository requires a
  paid GitHub plan (Pro or above) — check yours before relying on it. Note that
  on most plans the *published site* stays publicly accessible even when the
  repo is private; the repo going private mainly stops people finding the file
  by browsing your GitHub. Genuinely access-controlled Pages sites are an
  Enterprise Cloud feature.
- **Or move to a host with password protection**, e.g. Netlify or Cloudflare
  Pages on their paid tiers.
- **Or keep it public and strip anything sensitive** — the dashboard contains
  chapter titles, slide counts and percentages, no client data.

---

## The Excel workbook is the master record

`MyGuide-progress.xlsx` is where progress actually gets recorded. Put it in
SharePoint and the team can co-edit it — that's the point of using Excel rather
than clicking numbers into the dashboard one at a time. The dashboard reads the
workbook; it never writes back to it.

**Three tabs:**

| Tab | What's in it |
|---|---|
| **Read me** | Which cells to edit, allowed values, how the weighting works. |
| **Progress** | 60 rows — one per deck × chapter. Columns: Segment, Audience, Language, Chapter, Chapter title, Progress %, Feedback 1, Feedback 2, Validated. |
| **Slides** | 20 rows — one per segment × audience × chapter. Slide counts don't change between languages, so they're entered once rather than 60 times. |

Blue cells are inputs; grey columns A–E identify the deck and must not be
edited, because that's how the importer finds each row. Progress %, the two
feedback columns and Validated all have dropdowns, so there's nothing to type
by hand and nothing to get wrong.

One rule worth passing on to whoever else edits it: **don't type notes in the
rows underneath either table.** Anything below the data gets read as if it were
a row. There's space on the Read me tab for remarks.

The importer matches columns by *header name*, not position, so adding your own
columns (owner, due date, comments) or reordering the existing ones is fine —
the extra columns are simply ignored.

## Publishing an update — automatic

Once the sync in the next section is set up, there is no publishing step. Edit
the workbook in SharePoint, save, and the published page catches up within about
fifteen minutes. Everything below this line is one-time setup, plus the manual
fallback for when you want a change live *now*.

## Setting up the automatic sync

A scheduled GitHub Action downloads the workbook every 15 minutes, converts it,
and commits `progress.json` only when the figures actually changed.

### 1 · Get a direct link to the workbook

In SharePoint, select `MyGuide-progress.xlsx` → **Share** → change the audience
to **Anyone with the link** → set it to **Can view** → **Copy link**.

It has to be "Anyone with the link". A "People in your organisation" link needs
a sign-in, and GitHub's runner has no way to sign in — it'll get an HTML login
page instead of a file. The script detects exactly this and tells you so.

Be aware of the trade-off: that link makes the workbook readable to anyone who
has it. It holds chapter names, slide counts and percentages — no client data —
but decide rather than drift into it. If it's not acceptable, the Microsoft
Graph route with an Entra app registration keeps the file private; ask and I'll
write that version instead.

### 2 · Store the link as a repository secret

Repo → **Settings** → *Secrets and variables* → **Actions** → **New repository
secret**. Name it exactly `XLSX_URL`, paste the link as the value, save.

Secrets are never printed in logs and aren't readable from a fork, so this is
safe even though the repo is public. Don't paste the link into a file instead —
that would publish it.

### 3 · Add the two files

The repo needs to end up like this. Note that **`.github` and `scripts` sit at
the repo root**, *not* inside `MyGuide-webflow`:

```
BNPPF-Dashboard/
├── .github/
│   └── workflows/
│       └── update-progress.yml
├── scripts/
│   └── xlsx_to_json.py
└── MyGuide-webflow/
    ├── MyGuide-dashboard.html
    ├── progress.json
    └── webflow-embed-snippet.html
```

Dragging a folder called `.github` into the GitHub web uploader usually fails,
because browsers hide dot-folders in the file picker. Create them by typing the
path instead:

1. **Add file → Create new file**.
2. In the filename box type `.github/workflows/update-progress.yml` — typing
   each `/` turns the preceding part into a folder as you go.
3. Paste the file contents, then **Commit changes**.
4. Repeat with `scripts/xlsx_to_json.py`.

### 4 · Test it

Repo → **Actions** tab → *Update progress.json from SharePoint* → **Run
workflow**. Don't wait for the schedule; force the first run so you see the
result immediately.

A green tick means it worked — open the run log and you should see
`Matched 60/60 progress cells and 20/20 slide values.` A red cross means it
tells you why, in plain language. The commonest cause by far is a share link
that still requires sign-in.

### What the sync will and won't do

- **It refuses to publish incomplete data.** If fewer than 60 progress rows
  match — someone renamed a language, or is mid-restructure — the run fails
  rather than quietly writing those chapters as 0%. The log names the offending
  rows. GitHub emails you when a scheduled run fails.
- **It won't commit when nothing changed**, so your history stays readable and
  the date stamp alone doesn't trigger a commit every quarter of an hour.
- **Extra columns are ignored**, and columns are matched by header name, so the
  team can add owner or due-date columns without breaking the sync.
- **Timing is approximate.** GitHub queues scheduled runs and delays them under
  load. Fifteen minutes is typical; occasionally it's longer.

### Two things to watch

**Scheduled workflows switch off after 60 days of repository inactivity.** If
the workbook goes untouched that long there are no commits, so nothing counts as
activity. GitHub emails the repo owner before disabling it, and re-enabling is
one click in the Actions tab.

**If you ever make this repo private, change the schedule.** Actions minutes are
free and unlimited on public repos. On a private repo the free allowance is 2,000
minutes a month, and a run every 15 minutes would consume most of it. Change the
`cron` line in the workflow to `'0 * * * *'` for hourly or `'0 7 * * 1-5'` for
weekday mornings.

## Publishing an update — manual fallback

Useful when you want a change live immediately rather than waiting for the next
scheduled run, or if the sync is broken and you need to ship anyway.

1. Update `MyGuide-progress.xlsx` and save.
2. Open your local `MyGuide-dashboard.html`, click **Import from Excel (.xlsx)**
   and pick the workbook. It'll tell you how many rows it read.
3. Click **Export figures (.json)** → you get `MyGuide-progress-YYYY-MM-DD.json`.
4. Rename it to `progress.json`.
5. In GitHub, go to
   <https://github.com/FrederikOngenae/BNPPF-Dashboard/tree/main/MyGuide-webflow>,
   click **Add file → Upload files**, drag the new `progress.json` in, and
   commit. Same filename replaces the old one.
6. Wait ~1 minute for Pages to rebuild.

Steps 2–4 take about thirty seconds once you've done them once. The dashboard
file itself never changes, so after setup you never touch Webflow again. The published page reads the new JSON on the next load, and its footer
shows which file and which date it's reading — a quick way to confirm the update
landed.

> If the page still shows old numbers, it's browser cache, not GitHub. Hard-refresh
> (Ctrl+Shift+R). The dashboard requests the JSON with `cache: no-store`, but
> GitHub's CDN can hold a copy for a short while.

## How the embed behaves

`?embed=1` puts the dashboard in embed mode:

- **Read-only.** Progress strips, slide boxes and feedback pips are all inert —
  no one can click your numbers around.
- Export/import buttons, the lock toggle and the "how to update" notes are
  hidden. The **weighting explanation stays visible**, so the client can see why
  the percentages are what they are.
- The *Slide-weighted* and *Dark* toggles stay live — they only change the view.
- The page posts its height to Webflow whenever it changes, so the iframe grows
  to fit. No scrollbar inside a scrollbar, and your footer sits flush underneath.

`?data=progress.json` tells it where to read figures, relative to the dashboard's
own folder. If that file is missing or unreachable, the dashboard silently falls
back to the figures baked into the HTML — a bad upload shows stale numbers
rather than an empty page.

Embed mode also switches on automatically whenever the page is inside an iframe,
so `?embed=1` is belt and braces. Add `?embed=0` if you ever want the editable
version inside a frame.

---

## Troubleshooting

**"github.com heeft de verbinding geweigerd" / "refused to connect"**
The iframe is pointing at `github.com`. None of these work as an iframe source:

```
github.com/FrederikOngenae/BNPPF-Dashboard/blob/main/...   ← repo UI, refuses framing
raw.githubusercontent.com/FrederikOngenae/...              ← served as plain text
github.com/FrederikOngenae/.../raw/main/...                ← redirects to the above
```

The only URL that works is the `frederikongenae.github.io` one from step 2.

**404 on the github.io URL**
Either Pages hasn't finished its first build (check the Actions tab), or the
path is wrong. The folder and filename are case-sensitive.

**Dashboard loads but the frame stays 2600px tall**
The auto-height messages aren't getting through. Check that `ORIGIN` in the
snippet is exactly `https://frederikongenae.github.io` — scheme and host only,
no repo name, no path, no trailing slash.

**Blank iframe on a published Webflow page**
Check the Webflow page is served over `https`. An `http` iframe inside an
`https` page is blocked as mixed content.

**Nothing shows in the Webflow Designer**
Expected. Code embeds only render in Preview and on the published site.

**"Could not open that file" when importing from Excel**
The workbook is a cloud-only SharePoint placeholder. Right-click it in File
Explorer → *Always keep on this device*, wait for the sync tick, then try again.

**"Could not read that workbook"**
It has to be a real `.xlsx`. A `.csv`, an old `.xls`, or a `.url` shortcut to
the online copy won't work. In Excel: *File → Save a Copy → Excel Workbook
(\*.xlsx)*.

**The scheduled sync failed**
Open the failed run under the Actions tab and read the last step. The script
explains itself: a sign-in-required share link, an expired link, or fewer than
60 matching rows. Fix the cause and use **Run workflow** to retry rather than
waiting fifteen minutes.

**The sync succeeds but the page shows old numbers**
Check the commit actually landed (repo home, "X minutes ago" on progress.json),
then hard-refresh the Webflow page. GitHub's CDN holds a copy briefly.

**Import says "no row matched a known deck"**
Something in columns A–E got edited. Segment must read Private or Priority,
Audience must read Individual or Professional, Language must read NL, FR or EN,
and Chapter must be 1–5. Capitalisation doesn't matter, spelling does.

---

## Files

| File | What it is |
|---|---|
| `MyGuide-progress.xlsx` | **The master record.** Put this in SharePoint. Keep it out of the GitHub upload. |
| `MyGuide-dashboard.html` | The dashboard. Same file works locally (import/export) and hosted (read-only in an iframe). |
| `progress.json` | The figures, generated from the workbook. Replace this to publish an update. |
| `webflow-embed-snippet.html` | Paste into the Webflow Code embed element. Contains your real URLs. |
| `.github/workflows/update-progress.yml` | The scheduled sync. Goes at the **repo root**. |
| `scripts/xlsx_to_json.py` | Converts the workbook to progress.json. Goes at the **repo root**. |
| `README-webflow-setup.md` | This file. |
