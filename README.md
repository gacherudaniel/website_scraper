# KeNADA / KNBS Indicator Extraction & Table Inventory Generator

## Setup
```bash
pip install pandas requests beautifulsoup4 pdfplumber tqdm openpyxl
# optional, for faster PDF parsing on large reports:
pip install pymupdf
```

## Run
Place `Reports_and_Datasets_from_Kenada_and_Website.xlsx` in the same folder
as the script (or pass `--input` with a path), then:

```bash
python kenada_indicator_extractor.py
```

Useful flags:
- `--limit 20` — only process the first 20 rows (quick smoke test)
- `--input "path/to/file.xlsx"` — use a different input workbook
- `--sheet "Datasets"` — use a different sheet name (default: Datasets)
- `--start-fresh` — wipe `./downloaded_reports/` and re-download everything
- `--output-csv`, `--output-xlsx` — customize output paths

## What it does
1. **Filter to KNBS-website rows only** — rows whose `Source` column isn't
   exactly `Website` (i.e. KeNADA rows) are skipped entirely, with a
   `[SKIP]` log line and a count in the run summary. Nothing is downloaded
   from KeNADA.
2. **Resolve & download** — for each remaining row, if `Excel_File_URL`
   isn't already a direct `.pdf` link, it fetches the landing page and
   scores every link on the page to find the real report PDF (prioritizing
   "Download PDF" links and `wp-content/uploads/...pdf` URLs,
   de-prioritizing questionnaires/manuals). PDFs are cached in
   `./downloaded_reports/`, so re-runs skip files already downloaded.
   - **Fallback if the link doesn't work** — if the URL in the sheet can't
     be resolved to a PDF, *or* the resolved link fails to actually
     download (dead file, 404, etc.), the script searches knbs.or.ke's own
     site search (`https://www.knbs.or.ke/?s=...`) for the publication name,
     ranks the results by title similarity (with a small bonus if the row's
     Year also appears in the match), and tries to resolve/download a PDF
     from the best-matching result page instead. Rows resolved this way are
     tagged `[FALLBACK]` in the logs and counted separately in the summary.
3. **Extract** — each PDF is scanned two ways: (a) its "Table of
   Contents"/"List of Tables" pages (using the printed page reference), and
   (b) a full page-by-page body scan that tracks the current chapter heading
   and detects every `Table X.Y: Title` caption with its real PDF page
   number. The two passes are merged, with the body-scan page number
   preferred and the TOC entry kept only as a fallback for tables the body
   scan didn't independently confirm (e.g. captions rendered as an image).
4. **Output** — a single `Publication Name / Chapter Name / Indicator Name /
   Page Found` inventory, written to both `indicator_inventory.csv` and
   `indicator_inventory.xlsx`.

Progress is shown with a `tqdm` bar plus live `[DOWNLOAD]/[PARSE]/[EXTRACT]/
[SUCCESS]/[ERROR]` status lines, a full log is written to `./logs/run.log`,
and a summary (publications processed, indicators found, failures) prints at
the end.

## Notes
- This was validated against the actual `Datasets` sheet structure (1,143
  rows, 1,014 of them `Source=Website`) and tested with a synthetic
  multi-chapter PDF (TOC + body tables) to confirm the extraction and merge
  logic. The KeNADA-skip filter, the site-search fallback's HTML parsing/
  ranking, and the full fallback chain (search → candidate page → resolved
  PDF) were each verified with mocked responses. It was **not** run
  against the live knbs.or.ke site end-to-end in this environment (that
  network isn't reachable from here) — run it in your own environment with
  normal internet access.
- KNBS report pages are WordPress-based and generally expose a direct
  `wp-content/uploads/.../<report>.pdf` link, which the resolver and the
  fallback search are both tuned for. If a publication still can't be found
  after both the direct link and the site-search fallback are tried, it's
  logged as a failure in the run summary — check `logs/run.log` for the
  specific search query used and adjust `SITE_SEARCH_MIN_SCORE` or
  `search_knbs_website()` if it's a titling/matching issue.
- If you ever do want KeNADA rows included again, change `ALLOWED_SOURCE` at
  the top of the script (or ask for a `--include-kenada` flag to be added).