# AGENTS.md — Graphite

Handoff notes for anyone (human or agent) picking up this project.

## What this is

Graphite turns a spreadsheet into an interactive chart, entirely in the browser. It replaced an
old Flask backend: there is no server for rendering, and the user's data never leaves their device
unless they choose **Save to Cloud**.

Three pages, each a single self-contained `index.html` (no build step, no framework):

| Page | Folder | Chart | Live at |
|---|---|---|---|
| Graphite | `index.html` | Line chart, searchable series | `zhengjia.dev/graphite` |
| Graphite Combo | `combo/index.html` | Bars (left axis) + one line (right axis) | under `zhengjia.dev/graphite/…` |
| Graphite Histogram | `histogram/index.html` | Bars per bin, value labels | under `zhengjia.dev/graphite/…` |

Look and feel follows zhengjia.dev: random background photo from `cdn.zhengjia.dev/static/N.jpg`,
Adobe Clean font, dark inputs with a gradient border that glows only on focus.

## Files

```
index.html                 Graphite (line chart)
combo/index.html           Graphite Combo
combo/static/              Local mirror of the combo/histogram chart files (hosted on the CDN, see below)
combo/samples/*.xlsx       Combo sample datasets
combo/test.csv             Example data for the combo chart (also the built-in ?id=sample)
histogram/index.html       Graphite Histogram
histogram/samples/*.xlsx   Histogram sample datasets
samples/*.xlsx             Graphite (line chart) sample datasets
tools/update_samples.py    Downloads official data and rebuilds all sample files + the pages' sample lists
update_samples.bat         Double-click wrapper for the above (installs openpyxl if missing)
test.xlsx                  Ad-hoc test spreadsheet
```

**The three pages are near-copies with no shared code.** Combo and Histogram were generated from
the Graphite page and then diverged. A fix to shared behaviour (parsing, validation, cloud save,
sandboxing…) must be applied to **all three files** by hand.

## How a page works

1. **Input**: file upload (drag & drop / browse) or pasted cells, parsed with SheetJS
   (`xlsx@0.18.5` from cdnjs). Plus title fields: Graphite = graph title + search-bar title;
   Combo = graph, x, y (bars), optional line-axis title (defaults to the line column's name);
   Histogram = graph, x, y.
2. **Validation** (`readFile` → `readXlsx` / `readCsv`): only `.xlsx` / `.csv`; contents must match
   the extension (xlsx = zip containing `xl/workbook.xml`, csv = text, UTF-8 or Windows-1252);
   10 MB, 10,000 rows, 100 series; table checked on upload. Combo also needs ≥ 2 number columns.
3. **Parsing** (`sheetRows` → `toDataset`): works on SheetJS cell objects. Labels keep Excel's
   displayed text; values use the raw number (`$1,234.50` → 1234.5, `25%` cell → 25, `(500)` → −500,
   text numbers incl. `1.234,56`). Skips title rows above the table, empty/formatted-only columns,
   blank rows, `Total`/`Sum` rows and note rows with no numbers. Excel's default date format is
   shown as e.g. "1 May 2024" (SheetJS would print US-style 5/1/24). Pasted text and CSV are read
   with `raw: true` so labels are kept as typed.
4. **Rendering** (`buildGraphHtml`): builds a complete standalone HTML page for the chart and shows
   it in `<iframe id="frame" sandbox="allow-scripts allow-popups">` via `srcdoc`.
   **Download** saves exactly that HTML.
5. **Share links**: `?id=<uuid>` loads a cloud-saved graph; `?id=sample` shows a built-in sample.
   Shared graphs open with **Download** / **Create your own copy** instead of Save to Cloud / New.

### Chart files (on the CDN)

| Page | Chart library |
|---|---|
| Graphite | `cdn.zhengjia.dev/graph/static/supp_1.js`, `supp_2.css`, `supp_3a.css`, `embedded.js` (old Flourish line template) |
| Combo + Histogram | `cdn.zhengjia.dev/graphite/combo/template.js`, `bundle.css`, `choices.min.css`, `icons/` (Flourish "Line, bar and pie charts" template 118910 v4) + Graphite's `graph/static/embedded.js` |

Histogram deliberately **reuses** the combo chart files (`CHART_STATIC` points at `/graphite/combo/`);
the two Flourish exports were identical apart from a hidden credit string.

### Chart settings worth knowing

- Combo: the template draws the **first** value column as the line (`secondary_series: 1`). Users put
  the line in the **last** column, so `buildGraphHtml` moves it to the front.
- Combo/Histogram set `layout.body_font` to Source Sans Pro so they match Graphite's charts
  (the template's default is Canva Sans).
- Histogram: `legend_mode` is `off` for one series (as in the export), `auto` for several.

## Gotchas (each of these cost real debugging time)

- **Short charts on phones.** Flourish's `embedded.js` decides `fixed_height` by reading the parent
  page's URL. In a sandboxed `srcdoc` frame that throws (console shows "Script error"), and the
  chart falls back to 400 px on phones / 650 px on desktop. Every page sets
  `Flourish.fixed_height = true; Flourish.checkHeight = function () {};` **after** the scripts
  load and **before** `template.draw()`. Setting it earlier is useless — `embedded.js` overwrites it.
- **Never write a literal `</body>`, `<body`, `<head>` or `</head>` inside the page's `<script>`.**
  VS Code Live Server (`127.0.0.1:5500`, used for local testing) injects its reload script before the
  *first* `</body>` in the file; if that's inside a JS string the whole page script breaks (blank
  page, no greeting). Build tags as `"</" + "body>"`. Comments count too.
- **Flourish renders titles as HTML.** Shared graphs come from other people, hence the sandbox
  (opaque origin) and `fromSaved()` stripping `<` `>` and keeping only the expected JSON shape.
- **Testing the sandbox locally:** Chrome blocks a sandboxed frame from loading scripts from
  `localhost`, so a local copy of `template.js` won't load in the real frame (CDN files are fine).
  To test, inline the script into a test `srcdoc`, or test against the CDN copy. When inlining via
  `String.replace`, use a function replacer — `$` sequences in minified code corrupt it otherwise.
- **Screenshots of the chart** often catch the 1.5 s entrance animation (empty or half-drawn bars).
  Wait or re-take before concluding something is broken.
- **Flourish label types:** with `"auto"`, labels like "May 1" were read as years (2001…). Graphite
  only lets plain 4-digit years be `auto`; everything else is `string`.

## Save to Cloud

- Flow: Cloudflare Turnstile → `POST https://graphite-upload.zhengjia.workers.dev/graphite/sign`
  `{token, size}` → Worker returns a presigned S3 PUT URL (2-minute expiry, exact `content-length`
  and `content-type` signed) → browser PUTs the JSON straight to the Backblaze B2 bucket behind the CDN, at
  `graphite/shared/<uuid>.json` → share link `?id=<uuid>`.
- Loading: `GET …/graphite/load?id=<uuid>` (Worker proxies `cdn.zhengjia.dev`, adds CORS).
- Saved JSON: `{ v: 1, kind, title, …, columns, data, labelType }`. `kind` is absent for Graphite
  (line), `"combo"` or `"histogram"` for the others; each page refuses the other kinds with a message.
- Worker config: vars `B2_BUCKET`, `B2_ENDPOINT`, `B2_REGION`, `PUBLIC_BASE`, `ALLOWED_ORIGINS`;
  secrets `B2_KEY_ID`, `B2_APP_KEY`, `TURNSTILE_SECRET`. The public Turnstile site key is in each page
  (`TURNSTILE_SITEKEY`) and must list the domain the pages are served from.
  B2 CORS rule `graphite-upload` allows `s3_put` from `https://zhengjia.dev`.
- **The Worker source (`graphite-worker.js`) is not in this folder.** It was written in an earlier
  working folder that has since moved; the deployed copy lives in the Cloudflare dashboard.

## Sample datasets

All samples come from official Singapore open data, downloaded by `tools/update_samples.py`
(run `update_samples.bat`). It never uploads anything.

- `GROUPS` in the script = Graphite → `samples/`, Combo → `combo/samples/`,
  Histogram → `histogram/samples/`, each with its page and CDN folder. Every run processes all
  groups (18 files today) and ends with "Checked all N sample files across M pages".
- Before downloading anything it checks coverage and stops if something would be missed: any
  `index.html` in the project with a `// SAMPLES:BEGIN` list that isn't in `GROUPS`, or any `.xlsx`
  in a sample folder that no dataset builds. **Adding a new chart type = add its datasets and a
  `GROUPS` entry**, otherwise the script refuses to run.
- Each `.xlsx`: A1 title, A2 summary (shown in the picker), A3 source citation, blank row,
  header on row 5, data from row 6. A data fingerprint is stored in the workbook properties
  (`keywords`); unchanged data is not rewritten.
- Each page's sample list sits between `// SAMPLES:BEGIN` and `// SAMPLES:END` and is rewritten by
  the script: file, `v` (the fingerprint, used as `?v=` so Cloudflare's 7-day cache can't serve a
  stale file), title, summary, and for Combo/Histogram the axis titles.
- Sources: SingStat Table Builder API and data.gov.sg datastore API.

Data-source quirks the script already handles:
- SingStat rejects Python's default client — requests send a browser `User-Agent`.
- SingStat's `limit`/`offset` count **data points**, not rows — page until empty.
- Some SingStat tables repeat row names across groups (age table: total/male/female) — use
  `singstat_rows()` with `seriesNo`, not `singstat()`.
- openpyxl read-only workbooks must be closed, or Windows can't replace the file.
- HDB resale applications (`d_02aa…`) sometimes repeat a quarter with revised figures and publish
  new quarters flat type by flat type: keep the latest record, skip incomplete quarters.
- Census tables are rounded per cell, so bins can differ from the stated total by 1.
- The HDB resale-price histogram always uses the last 3 *complete* months.

## Deploying

Everything is uploaded manually (do **not** upload to Backblaze/the CDN from scripts or agents):

- Pages → wherever zhengjia.dev serves them.
- Sample files → `cdn.zhengjia.dev/graphite/samples/`, `/graphite/combo/samples/`,
  `/graphite/histogram/samples/`. Upload the files **before** the page whose `v=` tags point at them.
- Chart files → `cdn.zhengjia.dev/graphite/combo/` (from `combo/static/`; `embedded.js` there is unused).

Last checked (19 Oct 2026): chart files `template.js`, `bundle.css`, `choices.min.css` and the
combo + histogram sample files are on the CDN. Still missing (404): `graphite/combo/icons/cross.svg`
and `cross-inverse.svg` — only the "×" icons in the chart's search box. Whether the latest page
HTML has been uploaded isn't verifiable from here.

## Testing

No automated tests. Serve the folder and use a browser:

```
python -m http.server 8765
```

(Not Live Server for sandbox checks — see the localhost note above.) Useful checks: upload
`combo/test.csv`; open `?id=sample`; run every sample in the picker; phone width 412 px; confirm the
chart reaches the bottom of the frame.

## Workflow

- Uploading/deploying is done by hand; don't push to Backblaze or the CDN automatically.
- Local testing uses VS Code Live Server; mobile checks use Chrome device emulation on the live site.
- No secrets belong in this repo: Worker keys live only in Cloudflare's secret settings.
