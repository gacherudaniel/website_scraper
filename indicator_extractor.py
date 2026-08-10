#!/usr/bin/env python3
"""
KNBS Indicator Extraction & Table Inventory Generator
================================================================

Reads a list of publication URLs from an Excel workbook (sheet "Datasets"),
downloads the underlying PDF report for each publication whose Source is
"Website" (rows sourced from KeNADA are skipped entirely - see ALLOWED_SOURCE),
resolving landing pages -> direct PDF links where necessary. If the URL given
in the sheet doesn't resolve to a PDF, or the resolved link is dead, the
script falls back to searching knbs.or.ke directly for the publication and
tries to download the best-matching result instead.

Each downloaded PDF is scanned for its Table of Contents / List of Tables as
well as its full body text, and every table / indicator found is compiled
into a single "Table Inventory":

    Publication Name | Chapter Name | Indicator Name | Page Found

Results are written to both indicator_inventory.csv and indicator_inventory.xlsx.

Usage
-----
    python kenada_indicator_extractor.py
    python kenada_indicator_extractor.py --input "Reports_and_Datasets_from_Kenada_and_Website.xlsx"
    python kenada_indicator_extractor.py --limit 20          # quick test run
    python kenada_indicator_extractor.py --workers 4         # parallel downloads

Dependencies
------------
    pip install pandas requests beautifulsoup4 pdfplumber tqdm openpyxl
    (PyMuPDF is optional and used automatically if installed, for faster parsing:
     pip install pymupdf)
"""

from __future__ import annotations

import argparse
import difflib
import logging
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional
from urllib.parse import quote_plus, urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

try:
    import fitz  # PyMuPDF - optional, faster PDF parsing if available
    HAVE_FITZ = True
except ImportError:
    HAVE_FITZ = False

import pdfplumber

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

DEFAULT_INPUT_XLSX = "Reports_and_Datasets_from_Kenada_and_Website.xlsx"
SHEET_NAME = "Datasets"
DOWNLOAD_DIR = Path("./downloaded_reports")
LOG_DIR = Path("./logs")
OUTPUT_CSV = "indicator_inventory.csv"
OUTPUT_XLSX = "indicator_inventory.xlsx"

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}

REQUEST_TIMEOUT = 30          # seconds
MAX_RETRIES = 3
RETRY_BACKOFF = 2.0           # seconds, multiplied by attempt number
TOC_SEARCH_PAGES = 15         # how many leading pages to scan for a TOC

# Only rows whose Source column matches this (case-insensitive) are
# downloaded - KeNADA rows are skipped entirely, per requirement.
ALLOWED_SOURCE = "website"

# KNBS runs on WordPress, which exposes a native search endpoint at
# https://www.knbs.or.ke/?s=<query>. Used as a fallback when the URL given
# in the input sheet is dead, blocked, or doesn't resolve to a PDF.
KNBS_SEARCH_URL_TEMPLATE = "https://www.knbs.or.ke/?s={query}"
SITE_SEARCH_MAX_CANDIDATES = 5
SITE_SEARCH_MIN_SCORE = 0.30  # min title-similarity to accept a search result

# Words that mark a landing-page link as "this is probably the report PDF"
PDF_LINK_TEXT_HINTS = [
    "download pdf", "download report", "download full report", "download",
    "get microdata", "full report", "view pdf", "pdf report", "read more",
    "view report", "report pdf",
]

# Words that indicate a link should be de-prioritised (questionnaires, forms,
# metadata, etc. that sometimes sit next to the real report on the same page)
PDF_LINK_TEXT_PENALTIES = [
    "questionnaire", "consent form", "manual", "concept note", "study",
    "metadata", "citation",
]

# --------------------------------------------------------------------------- #
# Logging
# --------------------------------------------------------------------------- #

LOG_DIR.mkdir(parents=True, exist_ok=True)
logger = logging.getLogger("indicator_extractor")
logger.setLevel(logging.INFO)
logger.propagate = False

_fh = logging.FileHandler(LOG_DIR / "run.log", encoding="utf-8")
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(_fh)


def log_and_print(message: str, level: str = "info") -> None:
    """Print a live status line to the console AND write it to the log file."""
    print(message, flush=True)
    getattr(logger, level, logger.info)(message)


# --------------------------------------------------------------------------- #
# Data classes
# --------------------------------------------------------------------------- #

@dataclass
class Indicator:
    publication_name: str
    chapter_name: str
    indicator_name: str
    page_found: Optional[str]
    source: str = "body"  # "body" (actual PDF page scan) or "toc" (Table of Contents listing)


@dataclass
class RunStats:
    total_rows: int = 0
    processed_ok: int = 0
    failed_downloads: int = 0
    failed_extractions: int = 0
    skipped_existing: int = 0
    skipped_source: int = 0
    resolved_via_fallback: int = 0
    total_indicators: int = 0
    failures: list = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Helpers: text cleaning / filenames
# --------------------------------------------------------------------------- #

def clean_text(text: str) -> str:
    """Strip leader dots, excess whitespace, and stray line breaks."""
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\n", " ").replace("\r", " ")
    # Remove TOC leader dots / dot-fill runs: "Table 1.1 Title ..... 12"
    text = re.sub(r"\.{2,}", " ", text)
    text = re.sub(r"[\u2026]+", " ", text)  # unicode ellipsis
    # Collapse whitespace
    text = re.sub(r"\s+", " ", text).strip()
    # Trim stray trailing punctuation left over after removing leader dots
    text = text.strip(" .-\u2013\u2014")
    return text


def safe_filename(name: str, year: str = "", max_len: int = 150) -> str:
    """Build a filesystem-safe file name from a publication name and year."""
    base = f"{name}_{year}" if year else name
    base = unicodedata.normalize("NFKD", str(base)).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^\w\s\-]", "", base)
    base = re.sub(r"\s+", "_", base.strip())
    base = re.sub(r"_+", "_", base).strip("_")
    return (base[:max_len] or "report") + ".pdf"


# --------------------------------------------------------------------------- #
# Phase A: URL resolution & download
# --------------------------------------------------------------------------- #

def _score_pdf_link(href: str, link_text: str) -> int:
    """Heuristic score for how likely an <a> tag is to be the main report PDF."""
    href_l = href.lower()
    text_l = (link_text or "").lower()
    score = 0

    if href_l.endswith(".pdf"):
        score += 5
    if "wp-content/uploads" in href_l:
        score += 3
    if "/download" in href_l or "get_file" in href_l or "fulltext" in href_l:
        score += 2

    for hint in PDF_LINK_TEXT_HINTS:
        if hint in text_l:
            score += 3
            break

    for penalty in PDF_LINK_TEXT_PENALTIES:
        if penalty in text_l or penalty in href_l:
            score -= 4

    return score


def find_pdf_links_in_html(html: str, base_url: str) -> list[tuple[str, int]]:
    """
    Pure parsing helper (no network calls) so it can be unit-tested in isolation.
    Returns a list of (absolute_url, score) tuples, sorted best-first.
    """
    soup = BeautifulSoup(html, "html.parser")
    candidates: dict[str, int] = {}

    # 1. Anchor tags
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith("javascript:") or href.startswith("#"):
            continue
        abs_url = urljoin(base_url, href)
        parsed = urlparse(abs_url)
        if parsed.scheme not in ("http", "https"):
            continue
        link_text = a.get_text(" ", strip=True)
        score = _score_pdf_link(abs_url, link_text)
        if score <= 0 and not abs_url.lower().endswith(".pdf"):
            continue
        candidates[abs_url] = max(candidates.get(abs_url, -999), score)

    # 2. <embed>/<iframe> that point straight at a PDF (common on NADA "Downloads" tabs)
    for tag in soup.find_all(["embed", "iframe"], src=True):
        src = tag["src"].strip()
        abs_url = urljoin(base_url, src)
        if abs_url.lower().endswith(".pdf"):
            candidates[abs_url] = max(candidates.get(abs_url, -999), 6)

    # 3. <meta property="og:..."> pointing at a pdf (rare, but cheap to check)
    for meta in soup.find_all("meta", content=True):
        content = meta["content"].strip()
        if content.lower().endswith(".pdf"):
            abs_url = urljoin(base_url, content)
            candidates[abs_url] = max(candidates.get(abs_url, -999), 4)

    ranked = sorted(candidates.items(), key=lambda kv: kv[1], reverse=True)
    return ranked


def _get_with_retries(url: str, *, stream: bool = False) -> requests.Response:
    last_exc: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url,
                headers=REQUEST_HEADERS,
                timeout=REQUEST_TIMEOUT,
                stream=stream,
                allow_redirects=True,
            )
            if resp.status_code == 200:
                return resp
            if resp.status_code in (403, 429) and attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF * attempt)
                continue
            resp.raise_for_status()
            return resp
        except requests.RequestException as exc:
            last_exc = exc
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF * attempt)
    assert last_exc is not None
    raise last_exc


def resolve_pdf_link(landing_url: str) -> Optional[str]:
    """
    Given a URL that may already be a direct PDF, or may be an HTML landing
    page (KNBS report page or KeNADA/NADA catalog page), return the best
    direct PDF URL, or None if no PDF link could be located.
    """
    landing_url = (landing_url or "").strip()
    if not landing_url:
        return None

    if landing_url.lower().split("?")[0].endswith(".pdf"):
        return landing_url

    try:
        resp = _get_with_retries(landing_url)
    except requests.RequestException as exc:
        logger.warning(f"[RESOLVE] Could not fetch landing page {landing_url}: {exc}")
        return None

    content_type = resp.headers.get("Content-Type", "").lower()
    if "application/pdf" in content_type:
        return landing_url

    try:
        ranked = find_pdf_links_in_html(resp.text, landing_url)
    except Exception as exc:  # malformed HTML, etc.
        logger.warning(f"[RESOLVE] Could not parse landing page {landing_url}: {exc}")
        return None

    if not ranked:
        # NADA catalog pages sometimes hide the download behind a secondary
        # "Downloads" or "get_microdata" tab URL - try one common pattern.
        if "/catalog/" in landing_url:
            guess = landing_url.rstrip("/") + "/download"
            try:
                probe = _get_with_retries(guess)
                if "application/pdf" in probe.headers.get("Content-Type", "").lower():
                    return guess
            except requests.RequestException:
                pass
        return None

    return ranked[0][0]


def search_knbs_website(pub_name: str, year: str = "") -> list[str]:
    """
    Fall back to KNBS's own site search (https://www.knbs.or.ke/?s=...) to
    locate a publication's report page when the URL given in the input sheet
    is dead, blocked, or doesn't resolve to a PDF.

    Returns candidate report-page URLs on knbs.or.ke, ranked best-match first
    by text similarity between the search-result link text and the
    publication name (with a small bonus if the row's Year also appears in
    the link).
    """
    pub_name = (pub_name or "").strip()
    if not pub_name:
        return []

    search_url = KNBS_SEARCH_URL_TEMPLATE.format(query=quote_plus(pub_name))

    try:
        resp = _get_with_retries(search_url)
    except requests.RequestException as exc:
        logger.warning(f"[FALLBACK] KNBS site search request failed for '{pub_name}': {exc}")
        return []

    try:
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as exc:
        logger.warning(f"[FALLBACK] Could not parse KNBS search results for '{pub_name}': {exc}")
        return []

    scored: dict[str, float] = {}
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        link_text = a.get_text(" ", strip=True)
        if not href or not link_text:
            continue

        abs_url = urljoin(search_url, href)
        parsed = urlparse(abs_url)
        if "knbs.or.ke" not in parsed.netloc:
            continue  # ignore off-site links (social icons, ads, nav, etc.)
        if not (abs_url.lower().split("?")[0].endswith(".pdf") or "/reports/" in abs_url.lower()):
            continue  # only consider report pages or direct PDF links

        score = difflib.SequenceMatcher(None, link_text.lower(), pub_name.lower()).ratio()
        if year and (year in href or year in link_text):
            score += 0.15

        scored[abs_url] = max(scored.get(abs_url, 0.0), score)

    ranked = sorted(scored.items(), key=lambda kv: kv[1], reverse=True)
    return [url for url, score in ranked if score >= SITE_SEARCH_MIN_SCORE][:SITE_SEARCH_MAX_CANDIDATES]


def resolve_via_site_search(pub_name: str, year: str = "") -> Optional[str]:
    """
    Search knbs.or.ke directly for a publication and try to resolve a real
    PDF from the best-matching result(s). This is the fallback path used
    when the URL supplied in the input sheet doesn't work.
    """
    candidates = search_knbs_website(pub_name, year)
    for candidate_url in candidates:
        pdf_url = resolve_pdf_link(candidate_url)
        if pdf_url:
            return pdf_url
    return None


def download_pdf(url: str, output_path: Path) -> bool:
    """
    Download a PDF from `url` to `output_path`. Returns True on success.
    Skips the download (returns True) if the file already exists locally.
    """
    if output_path.exists() and output_path.stat().st_size > 0:
        log_and_print(f"[SKIP] Already downloaded: {output_path.name}")
        return True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_suffix(output_path.suffix + ".part")

    try:
        resp = _get_with_retries(url, stream=True)
        content_type = resp.headers.get("Content-Type", "").lower()
        if "pdf" not in content_type and not url.lower().split("?")[0].endswith(".pdf"):
            logger.warning(
                f"[DOWNLOAD] URL did not return a PDF content-type ({content_type}): {url}"
            )

        with open(tmp_path, "wb") as fh:
            for chunk in resp.iter_content(chunk_size=1024 * 64):
                if chunk:
                    fh.write(chunk)

        if tmp_path.stat().st_size == 0:
            raise IOError("Downloaded file is empty")

        tmp_path.rename(output_path)
        return True

    except Exception as exc:
        log_and_print(f"[ERROR] Failed to download {url}: {exc}", level="error")
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        return False


# --------------------------------------------------------------------------- #
# Phase B: PDF extraction (TOC + in-body scanning)
# --------------------------------------------------------------------------- #

# "Table 3.1: Gross Domestic Product by Activity ......... 45"
TOC_TABLE_LINE_RE = re.compile(
    r"^(?P<label>(?:Table|TABLE)\s*[:\-]?\s*\d+[A-Za-z]?(?:[.\-]\d+)*)"
    r"\s*[:.\-]?\s*(?P<title>.+?)\s*(?:\.{2,}|\s{2,}|\u2026)\s*(?P<page>[ivxlcIVXLC\d]+)\s*$"
)
# "Table 3.1: Gross Domestic Product by Activity" (no trailing page ref)
TOC_TABLE_LINE_NO_PAGE_RE = re.compile(
    r"^(?P<label>(?:Table|TABLE)\s*[:\-]?\s*\d+[A-Za-z]?(?:[.\-]\d+)*)"
    r"\s*[:.\-]?\s*(?P<title>.+)$"
)
TOC_HEADER_RE = re.compile(
    r"(table\s+of\s+contents|list\s+of\s+tables|list\s+of\s+figures)", re.IGNORECASE
)
# Stop treating a section as "the TOC" once we clearly hit body content
TOC_END_HINTS_RE = re.compile(
    r"^(chapter\s+one|chapter\s+1\b|1\.0\s|1\.1\s|introduction\b)", re.IGNORECASE
)

# In-body table caption, e.g. "Table 4.2: Population by County"
BODY_TABLE_RE = re.compile(
    r"^Table\s+(\d+[A-Za-z]?(?:[.\-]\d+)*)\s*[:.\-]?\s*(.+)$"
)

# Chapter headings, e.g. "CHAPTER ONE", "Chapter 1: Introduction", "1.0 INTRODUCTION"
CHAPTER_PATTERNS = [
    re.compile(r"^CHAPTER\s+[A-Z]+\b.*$", re.IGNORECASE),
    re.compile(r"^Chapter\s+\d+\s*[:.\-]?\s*.*$", re.IGNORECASE),
    re.compile(r"^\d{1,2}\.0\s+[A-Z][A-Za-z \-,&/]+$"),
    re.compile(r"^\d{1,2}\.0\s+[A-Z\s]{4,}$"),
    re.compile(r"^[A-Z][A-Z \-,&/]{6,}$"),  # long ALL-CAPS heading line
]

ROMAN_NUMERAL_RE = re.compile(r"^[ivxlcdmIVXLCDM]+$")


def _looks_like_chapter_heading(line: str) -> bool:
    line = line.strip()
    if not line or len(line) > 100:
        return False
    if BODY_TABLE_RE.match(line):
        return False
    for pat in CHAPTER_PATTERNS:
        if pat.match(line):
            return True
    return False


def _extract_toc_indicators(pdf, pub_name: str) -> list[Indicator]:
    """Scan the first TOC_SEARCH_PAGES pages for a 'List of Tables'/TOC block."""
    found: list[Indicator] = []
    in_toc_section = False
    current_chapter = "Table of Contents"

    n_pages = min(TOC_SEARCH_PAGES, len(pdf.pages))
    for page_idx in range(n_pages):
        try:
            text = pdf.pages[page_idx].extract_text() or ""
        except Exception:
            continue

        for raw_line in text.split("\n"):
            line = raw_line.strip()
            if not line:
                continue

            if TOC_HEADER_RE.search(line):
                in_toc_section = True
                current_chapter = clean_text(line)
                continue

            if not in_toc_section:
                continue

            if _looks_like_chapter_heading(line) and not BODY_TABLE_RE.match(line):
                current_chapter = clean_text(line)
                continue

            # A genuine "List of Tables" entry always terminates in a leader
            # of dots (or wide spacing) followed by a page number - that is
            # what distinguishes it from a real in-body table caption. Only
            # matching this stricter pattern (rather than also accepting a
            # bare "Table X: Title" line) keeps this pass from bleeding into
            # ordinary body text if the TOC section boundary isn't crisp.
            m = TOC_TABLE_LINE_RE.match(line)
            if not m:
                continue

            title = clean_text(m.group("title"))
            label = clean_text(m.group("label"))
            page_ref = m.group("page")

            if not title:
                continue
            indicator_name = clean_text(f"{label}: {title}")
            found.append(
                Indicator(
                    publication_name=pub_name,
                    chapter_name=current_chapter,
                    indicator_name=indicator_name,
                    page_found=page_ref,
                    source="toc",
                )
            )

    return found


def _extract_body_indicators(pdf, pub_name: str) -> list[Indicator]:
    """Scan every page of the document, tracking chapter headings and table captions."""
    found: list[Indicator] = []
    current_chapter = "Unknown / Front Matter"

    for page_number, page in enumerate(pdf.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            logger.warning(f"[EXTRACT] Could not read page {page_number}: {exc}")
            continue

        for raw_line in text.split("\n"):
            line = raw_line.strip()
            if not line:
                continue

            if _looks_like_chapter_heading(line):
                current_chapter = clean_text(line)
                continue

            # Lines that still carry TOC-style dot-leaders / page references
            # (e.g. "Table 2.1: GDP by Activity ......... 12") belong to a
            # Table of Contents block, not real in-body captions - the TOC
            # pass already captures these, so skip them here to avoid
            # double-counting with a bogus (wrong) page number.
            if TOC_TABLE_LINE_RE.match(line):
                continue

            m = BODY_TABLE_RE.match(line)
            if m:
                label_num = m.group(1)
                title = clean_text(m.group(2))
                if not title:
                    continue
                indicator_name = clean_text(f"Table {label_num}: {title}")
                found.append(
                    Indicator(
                        publication_name=pub_name,
                        chapter_name=current_chapter,
                        indicator_name=indicator_name,
                        page_found=str(page_number),
                        source="body",
                    )
                )

    return found


def _dedupe_indicators(indicators: Iterable[Indicator]) -> list[Indicator]:
    """
    Merge the TOC pass and the in-body pass into one inventory per table.

    Body-scanned entries carry the actual PDF page the caption was found on,
    so they are the preferred/authoritative record. TOC-only entries (tables
    listed in the "List of Tables" but never matched in the body scan - e.g.
    because the caption spans two lines, or the table is rendered as an
    image) are kept as a fallback so nothing is silently dropped.
    """
    by_key: dict[str, Indicator] = {}
    order: list[str] = []

    def key_for(ind: Indicator) -> str:
        return re.sub(r"[^a-z0-9]+", "", ind.indicator_name.lower())[:120]

    for ind in indicators:
        k = key_for(ind)
        if k not in by_key:
            by_key[k] = ind
            order.append(k)
            continue

        existing = by_key[k]
        # A body-sourced record always wins over a TOC-sourced one for the
        # same table; between two records of the same source, keep the first
        # one encountered.
        if existing.source == "toc" and ind.source == "body":
            by_key[k] = ind

    return [by_key[k] for k in order]


def extract_indicators_from_pdf(pdf_path: Path, pub_name: str) -> list[Indicator]:
    """
    Open a downloaded PDF and extract every table/indicator it contains,
    combining a Table-of-Contents pass with a full in-body scan.
    """
    indicators: list[Indicator] = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        log_and_print(f"[PARSE] Processing PDF: {pdf_path.name} ({total_pages} pages)...")

        toc_indicators = _extract_toc_indicators(pdf, pub_name)
        if toc_indicators:
            log_and_print(f"[EXTRACT] Found TOC / Extracted {len(toc_indicators)} table indicators...")
        else:
            log_and_print("[EXTRACT] No explicit Table of Contents found; relying on body scan.")

        body_indicators = _extract_body_indicators(pdf, pub_name)
        indicators = _dedupe_indicators(toc_indicators + body_indicators)

    return indicators


# --------------------------------------------------------------------------- #
# Phase D: Input loading
# --------------------------------------------------------------------------- #

REQUIRED_COLUMNS = ["Name", "Year", "Source", "Has_XLSX", "Excel_File_URL"]


def load_dataset_rows(input_path: Path, sheet_name: str = SHEET_NAME) -> pd.DataFrame:
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Input sheet '{sheet_name}' is missing required column(s): {missing}. "
            f"Found columns: {list(df.columns)}"
        )
    df = df.dropna(subset=["Excel_File_URL"]).reset_index(drop=True)
    return df


# --------------------------------------------------------------------------- #
# Main pipeline
# --------------------------------------------------------------------------- #

def process_row(row: pd.Series, stats: RunStats) -> list[Indicator]:
    pub_name = str(row["Name"]).strip()
    year = str(row.get("Year", "")).strip()
    source = str(row.get("Source", "")).strip()
    landing_url = str(row["Excel_File_URL"]).strip()

    # Only download from the KNBS website - KeNADA rows are skipped entirely.
    if source.strip().lower() != ALLOWED_SOURCE:
        stats.skipped_source += 1
        log_and_print(f"[SKIP] Not a KNBS website source (Source={source or 'unknown'}): {pub_name}")
        return []

    log_and_print(f"[DOWNLOAD] Downloading: {pub_name}...")

    try:
        pdf_url = resolve_pdf_link(landing_url)
        used_fallback = False

        if not pdf_url:
            log_and_print(
                f"[FALLBACK] Provided link didn't resolve to a PDF; "
                f"searching knbs.or.ke for \"{pub_name}\"..."
            )
            pdf_url = resolve_via_site_search(pub_name, year)
            used_fallback = pdf_url is not None

        if not pdf_url:
            raise RuntimeError(
                f"No PDF could be resolved from {landing_url}, and a knbs.or.ke "
                f"site search turned up no match either"
            )

        out_path = DOWNLOAD_DIR / safe_filename(pub_name, year)
        already_had_file = out_path.exists() and out_path.stat().st_size > 0

        ok = download_pdf(pdf_url, out_path)

        if not ok and not used_fallback:
            # The provided/resolved link is broken (dead file, 404, etc.) -
            # search the KNBS website for the report before giving up.
            log_and_print(
                f"[FALLBACK] Resolved link failed to download; "
                f"searching knbs.or.ke for \"{pub_name}\"..."
            )
            alt_pdf_url = resolve_via_site_search(pub_name, year)
            if alt_pdf_url and alt_pdf_url != pdf_url:
                ok = download_pdf(alt_pdf_url, out_path)
                if ok:
                    pdf_url = alt_pdf_url
                    used_fallback = True

        if not ok:
            raise RuntimeError(f"Download failed for {pdf_url}")

        if used_fallback:
            stats.resolved_via_fallback += 1
            log_and_print(f"[FALLBACK] Resolved via knbs.or.ke search: {pdf_url}")

        if already_had_file:
            stats.skipped_existing += 1

        indicators = extract_indicators_from_pdf(out_path, pub_name)
        stats.total_indicators += len(indicators)
        stats.processed_ok += 1
        log_and_print(f"[SUCCESS] Finished {pub_name} \u2014 Total Indicators: {len(indicators)}")
        return indicators

    except Exception as exc:
        stats.failed_downloads += 1
        stats.failures.append({"Publication": pub_name, "URL": landing_url, "Error": str(exc)})
        log_and_print(f"[ERROR] Failed to process {pub_name}: {exc}", level="error")
        return []


def write_outputs(all_indicators: list[Indicator], csv_path: str, xlsx_path: str) -> pd.DataFrame:
    records = [
        {
            "Publication Name": ind.publication_name,
            "Chapter Name": ind.chapter_name,
            "Indicator Name": ind.indicator_name,
            "Page Found": ind.page_found,
        }
        for ind in all_indicators
    ]
    out_df = pd.DataFrame(records, columns=["Publication Name", "Chapter Name", "Indicator Name", "Page Found"])
    out_df.to_csv(csv_path, index=False)
    out_df.to_excel(xlsx_path, index=False, sheet_name="Table Inventory")
    return out_df


def print_summary(stats: RunStats) -> None:
    print("\n" + "=" * 60)
    print("RUN SUMMARY")
    print("=" * 60)
    print(f"Total rows in input:              {stats.total_rows}")
    print(f"Skipped (non-KNBS-website src):   {stats.skipped_source}")
    print(f"Successfully processed:           {stats.processed_ok}")
    print(f"  (of which, already on disk):    {stats.skipped_existing}")
    print(f"  (of which, via site-search):    {stats.resolved_via_fallback}")
    print(f"Failed downloads/extractions:     {stats.failed_downloads}")
    print(f"Total indicators extracted:       {stats.total_indicators}")
    if stats.failures:
        print("\nFailures:")
        for f in stats.failures[:25]:
            print(f"  - {f['Publication']}: {f['Error']}")
        if len(stats.failures) > 25:
            print(f"  ... and {len(stats.failures) - 25} more (see logs/run.log)")
    print("=" * 60)


def main() -> None:
    parser = argparse.ArgumentParser(description="KeNADA/KNBS Indicator Extraction & Table Inventory Generator")
    parser.add_argument("--input", default=DEFAULT_INPUT_XLSX, help="Path to the input .xlsx workbook")
    parser.add_argument("--sheet", default=SHEET_NAME, help="Sheet name containing the dataset list")
    parser.add_argument("--output-csv", default=OUTPUT_CSV, help="Path to write the CSV inventory")
    parser.add_argument("--output-xlsx", default=OUTPUT_XLSX, help="Path to write the XLSX inventory")
    parser.add_argument("--limit", type=int, default=None, help="Only process the first N rows (useful for testing)")
    parser.add_argument(
        "--start-fresh", action="store_true",
        help="Ignore/overwrite any PDFs already present in ./downloaded_reports/"
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        log_and_print(f"[ERROR] Input file not found: {input_path}", level="error")
        sys.exit(1)

    df = load_dataset_rows(input_path, args.sheet)
    if args.limit:
        df = df.head(args.limit)

    stats = RunStats(total_rows=len(df))
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    if args.start_fresh:
        for f in DOWNLOAD_DIR.glob("*.pdf"):
            f.unlink()

    all_indicators: list[Indicator] = []

    for _, row in tqdm(df.iterrows(), total=len(df), desc="Processing publications", unit="pub"):
        indicators = process_row(row, stats)
        all_indicators.extend(indicators)

    out_df = write_outputs(all_indicators, args.output_csv, args.output_xlsx)
    log_and_print(f"[SUCCESS] Wrote {len(out_df)} rows to {args.output_csv} and {args.output_xlsx}")

    print_summary(stats)


if __name__ == "__main__":
    main()