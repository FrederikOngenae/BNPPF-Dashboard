# Embedding the MyGuide dashboard in Webflow

Webflow can't host a raw `.html` file, and its Code Embed element caps out at
50,000 characters — the dashboard is ~45,000, which would fit but leaves almost
nothing for the rest of the site and would leak its CSS into your page styles.
So the dashboard lives on a free static host and Webflow shows it in an iframe
that resizes itself. Three files, about ten minutes.

---

## 1 · Put the two files on a static host

Upload **`MyGuide-dashboard.html`** and **`progress.json`** side by side, in the
same folder. Any of these work and all have a free tier:

| Host | How |
|---|---|
| **Netlify Drop** (fastest) | Go to `app.netlify.com/drop`, drag the folder in. You get a URL immediately. Rename the site in Site settings to something tidy. |
| **Cloudflare Pages** | Create a project → *Direct Upload* → drag the folder. |
| **GitHub Pages** | Push the folder to a repo, Settings → Pages → deploy from branch. |

Confirm both of these load in a browser before continuing:

```
https://YOUR-SITE.example.com/MyGuide-dashboard.html
https://YOUR-SITE.example.com/progress.json
```

## 2 · Paste the embed into Webflow

Drag a **Code embed** element onto the page and paste the contents of
`webflow-embed-snippet.html` into it. Replace **both** occurrences of
`https://YOUR-SITE.example.com` with your real origin — one in the iframe `src`,
one in the `ORIGIN` constant (that one has no trailing slash and no path; it's
what verifies the resize messages really come from your dashboard).

Two things that trip people up:

- The embed shows a grey placeholder on the Designer canvas. That's normal —
  use **Preview** or publish to see it.
- Don't wrap the snippet in `<html>`, `<head>` or `<body>` tags. Webflow
  rejects those and it breaks the page layout.

## 3 · Password-protect the page

Page settings → **General** → *Password protection* → set a password, then
publish. Set the page to `noindex` in the SEO tab as well if you'd rather it
stayed out of search results entirely.

Note that page- and folder-level password protection needs a paid **Site plan**
on that site — Webflow's own docs are explicit that you have to add one to
unlock it. Site-*wide* password protection is available on all paid Site plans
(Basic, CMS, Business, Ecommerce). If this site isn't on a paid plan yet, that's
the one prerequisite to sort out before step 3.

One caveat worth knowing: the Webflow password guards the *page*, not the
dashboard file. Anyone with the direct host URL can still open
`MyGuide-dashboard.html`. It's progress percentages rather than anything
sensitive, but if that matters, put the host behind its own access control —
Netlify and Cloudflare Pages both offer password protection on their paid tiers.

---

## Publishing an update

Your local copy of `MyGuide-dashboard.html` stays your working file. When you
want the client to see new figures:

1. Open your local copy, click the numbers to where they should be.
2. **Export figures (.json)** → you get `MyGuide-progress-YYYY-MM-DD.json`.
3. Rename it to `progress.json` and re-upload it to the host, replacing the old
   one. On Netlify Drop, drag the folder in again.

The dashboard itself never changes, so you never touch Webflow again after
setup. The published page reads the new file on the next load and its footer
shows which file and date it's reading.

## How the embed behaves

Adding `?embed=1` to the URL puts the dashboard in embed mode:

- Read-only — no one can click your numbers around. Progress strips, slide
  boxes and feedback pips are all inert.
- Export/import buttons and the lock toggle are hidden, as are the
  "how to update" notes. The **weighting explanation stays visible**, so the
  client can see why the percentages are what they are.
- The *Slide-weighted* and *Dark* toggles stay live — they only change the view.
- The page posts its height to Webflow whenever it changes, so the iframe grows
  to fit. No scrollbar inside a scrollbar, and your footer sits flush underneath.

`?data=progress.json` tells it where to read the figures. If that file is
missing or unreachable the dashboard silently falls back to the figures baked
into the HTML, so a bad upload shows stale numbers rather than an empty page.

Embed mode also switches on automatically whenever the page is inside an iframe,
so `?embed=1` is belt and braces. Add `?embed=0` if you ever want the editable
version inside a frame.

## Files

| File | What it is |
|---|---|
| `MyGuide-dashboard.html` | The dashboard. Same file works locally (editable) and hosted (read-only in an iframe). |
| `progress.json` | Figures, zeroed out. Replace with your export. |
| `webflow-embed-snippet.html` | Paste into the Webflow Code embed element. |
