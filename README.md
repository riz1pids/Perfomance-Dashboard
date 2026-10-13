# Shwapno Overall Performance Dashboard

POS sales, GPV, footfall and basket against SPLY (same day last year), by
category, brand, person and outlet, with chain targets. One page, no server,
no internet needed to view.

```
index.html                  the dashboard: open it, or let GitHub Pages serve it
build_dashboard_data.py     turns the month's input files into data/
ci/fetch_published.py       keeps earlier months when GitHub rebuilds the site
.github/workflows/build.yml  Drive -> build -> GitHub Pages
tests/                      synthetic inputs and a headless page check
```

`input/` and `data/` sit beside the script and are never committed. They are
large, they change every month, and they hold company figures.

> **This repository is public.** If you switch on GitHub Pages, the dashboard
> and its outlet-level figures are readable by anyone with the address.
> Before the first publish, decide whether that is acceptable. To keep it
> internal, make the repository private (Pages on a private repository needs a
> paid GitHub plan), or skip Pages and use the dashboard locally.

## Use it on your computer

```
pip install -r requirements.txt
python build_dashboard_data.py
```

Put the month's files in `input/` first (subfolders are fine). The builder
writes `data/` beside `index.html`; double-click `index.html` to open it. It
reads the data folder directly from disk, so it works offline.

## Publish it from Google Drive

The workflow downloads the shared Drive folder, builds the data and publishes
the site to GitHub Pages. Nothing is committed back to the repository.

**Once:** Settings > Pages > Build and deployment > Source: **GitHub Actions**.

**Each time:** Actions > *Build and publish dashboard* > **Run workflow**.
Untick *Publish* for a dry run: it downloads from Drive and runs the build,
which proves the files are readable, and publishes nothing.

The site then lives at `https://riz1pids.github.io/Perfomance-Dashboard/`.

To rebuild automatically every morning, uncomment the `schedule` block in
`.github/workflows/build.yml`. To read from a different Drive folder, set a
repository variable `DRIVE_FOLDER_ID`.

The Drive folder must be shared as *Anyone with the link can view*. Native
Google Sheets are downloaded as `.xlsx`, which the builder reads. Drive
sometimes refuses a large file that has been downloaded many times in a day
("too many users have viewed"); the download retries three times, and if it
still fails, wait an hour and run again.

## Each new month

Upload the new files to the Drive folder, then edit the `CONFIG` block at the
top of `build_dashboard_data.py` (on GitHub: open the file, pencil icon):

| Key | What changes |
|---|---|
| `FILE_SALES_TY`, `FILE_SALES_LY` | this year's and last year's POS export |
| `FILE_SAME_STORE`, `FILE_SAME_DAY`, `FILE_TARGET`, `FILE_ZONE` | the month's lists |
| `PERIOD_LABEL`, `SAME_DAY_SHEET` | e.g. `October 2026`, `Oct'26` |

File names are matched in any letter case, and a sales file may be `.csv` or
an Excel workbook. If a file is missing, the build stops and lists every file
it did find, so a renamed file is easy to spot.

The site keeps the latest three months (`KEEP_MONTHS`). Rebuilding a month
replaces that month only.

## What the numbers mean

- **SPLY** pairs each day with its weekday-matched day last year, from the
  same-day workbook. It is never "minus 365 days".
- **Growth %** is `(this − last) ÷ |last| × 100`. **GP %** is GPV ÷ sales.
  **Basket** is sales ÷ footfall. **ACH %** is actual ÷ target × 100.
- Every ratio is recomputed from whatever is on screen, so a filtered view and
  the grand total are always derived the same way.
- **Targets** are published chain-wide at Category 3. They show for All store
  only, and disappear under an outlet, brand or article filter rather than
  showing a number that could not be split that way.

## What follows the date range, and what doesn't

| View | Day-wise? |
|---|---|
| Cards, chart, category, brand, person and outlet views | yes |
| Article tab and article search | whole period only |

Article detail is stored as period totals. The Article tab asks for the full
period, and an article search holds the range to the full period and replaces
the chart with a note. Brand and article are not stored by outlet, so they
switch off while an outlet filter is on.

Tables draw the first 2,000 rows after sorting. Totals, subtotals and the CSV
export always cover every row.

## Checks

`tests/make_inputs.py` writes a full synthetic set of input files with the
real file names, sheets and columns. `tests/test_page.js` boots the page
headlessly and checks every view against totals taken from the raw files.

```
python tests/make_inputs.py /tmp/in
SHWAPNO_INPUT_DIR=/tmp/in python build_dashboard_data.py
npm install jsdom
node tests/test_page.js index.html '{"all":"2.68","same":"2.04","frozen":"0.37","minArticles":50,"minFrozenOutlets":20,"q":"FROZEN"}'
```

`make_inputs.py` prints the expected totals for the arguments.
