#!/usr/bin/env python3
"""
knbs_scraper.py
================
Scrapes report/dataset metadata from two Kenya National Bureau of Statistics
(KNBS) sources and exports a combined, deduplicated dataset to Excel.

Sources
-------
1. KNBS main website  -> https://www.knbs.or.ke/all-reports/          (Source = "Website")
2. KeNADA microdata    -> https://statistics.knbs.or.ke/nada/index.php/catalog/central  (Source = "KeNADA")

Output
------
knbs_scraped_reports.xlsx with columns:
    Name | Year | Source | Has_XLSX | Excel_File_URL

Install requirements
---------------------
    pip install requests beautifulsoup4 pandas openpyxl cryptography

(cryptography is optional but recommended — it lets this script
auto-repair TLS handshakes where a server sends an incomplete
certificate chain; see the "SSL chain repair" section below.)

Run
---
    python knbs_scraper.py

Notes
-----
- Public website markup changes over time. This script uses flexible,
  pattern-based extraction (regex + heuristic link matching) rather than
  brittle, hard-coded CSS classes, so it keeps working even if KNBS/KeNADA
  tweak their templates slightly. If a site redesign changes the URL
  patterns used below (see FIND_* constants), update those patterns.
- Be a good citizen: requests are throttled with time.sleep(REQUEST_DELAY)
  and pagination loops have hard safety caps (MAX_PAGES_*) to avoid
  hammering the servers or looping forever if a stop condition is missed.
"""

import re
import time
import ssl
import socket
import tempfile
import logging
from urllib.parse import urljoin, urlparse, parse_qs

import requests
import certifi
from bs4 import BeautifulSoup
import pandas as pd

try:
    from cryptography import x509
    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives import serialization
    from cryptography.x509.oid import ExtensionOID, AuthorityInformationAccessOID
    CRYPTOGRAPHY_AVAILABLE = True
except ImportError:
    CRYPTOGRAPHY_AVAILABLE = False

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

WEBSITE_BASE = "https://www.knbs.or.ke"
WEBSITE_REPORTS_URL = "https://www.knbs.or.ke/all-reports/"

KENADA_BASE = "https://statistics.knbs.or.ke"
KENADA_CATALOG_URL = "https://statistics.knbs.or.ke/nada/index.php/catalog/central"
KENADA_PAGE_SIZE = 100  # 'ps' query param -> larger page size, fewer requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

REQUEST_TIMEOUT = 15      # seconds, per request
REQUEST_DELAY = 1         # seconds, between requests (be respectful)
MAX_RETRIES = 2           # simple retry count for transient failures

# Safety caps so a missed "last page" condition can't loop forever.
MAX_PAGES_WEBSITE = 50
MAX_PAGES_KENADA = 50

# Only fetch each report's/study's own detail page to look for a direct
# .xlsx/.xls download if True. Turns off to speed up big/testing runs.
FOLLOW_DETAIL_PAGES = True

# SSL verification. Keep this True. Only flip to False as a temporary,
# last-resort diagnostic if certifi's bundle (see get_session()) still
# doesn't resolve a persistent "unable to get local issuer certificate"
# error on your machine — e.g. behind a corporate SSL-inspecting proxy.
# Disabling verification removes protection against man-in-the-middle
# attacks, so treat "it worked with this off" as a sign to fix your local
# trust store or add your organization's root CA, not as a permanent fix.
VERIFY_SSL = True

YEAR_PATTERN = re.compile(r"\b(19|20)\d{2}\b")
EXCEL_LINK_PATTERN = re.compile(r"\.(xlsx|xls)(\?.*)?$", re.IGNORECASE)

OUTPUT_FILE = "knbs_scraped_reports.xlsx"

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger("knbs_scraper")


# --------------------------------------------------------------------------
# HTTP helpers
# --------------------------------------------------------------------------

def get_session(verify_path=None) -> requests.Session:
    """Build a requests Session with standard headers."""
    session = requests.Session()
    session.headers.update(HEADERS)
    if not VERIFY_SSL:
        session.verify = False
        # Silence the noisy "InsecureRequestWarning" that requests/urllib3
        # prints on every call when verification is off.
        requests.packages.urllib3.disable_warnings(
            requests.packages.urllib3.exceptions.InsecureRequestWarning
        )
        log.warning(
            "VERIFY_SSL is False — HTTPS certificate verification is disabled. "
            "This is insecure; only use it to confirm the CA bundle is the "
            "problem, then re-enable it."
        )
    else:
        # Prefer a bundle that's had any server-side missing intermediate
        # certificates repaired (see build_repaired_verify_bundle). Falls
        # back to certifi's plain root bundle if that wasn't run/available.
        session.verify = verify_path or certifi.where()
    return session


# --------------------------------------------------------------------------
# SSL chain repair ("AIA chasing")
#
# Some servers are misconfigured to send only their leaf (site) certificate
# during the TLS handshake, omitting the intermediate certificate(s) needed
# to build a chain up to a trusted root. This produces
# "SSLCertVerificationError: unable to get local issuer certificate" in
# requests/urllib3/openssl — even with a perfectly up-to-date CA bundle,
# and even after reinstalling the OS's ca-certificates package — because
# the missing piece isn't your trust store, it's what the server sent you.
#
# Web browsers silently work around this by fetching the missing
# intermediate from the URL published in the leaf certificate's
# "Authority Information Access" (AIA) extension ("CA Issuers"), then
# caching it. curl/openssl/requests don't do this automatically. The
# functions below replicate that behaviour: walk the AIA chain, download
# whatever intermediate(s) are missing, and write a combined PEM (certifi's
# trusted roots + the fetched intermediates) that `requests` can verify
# against. This still performs full cryptographic chain validation up to a
# real trusted root — it does not weaken security the way verify=False does.
# --------------------------------------------------------------------------

def _get_leaf_certificate(hostname: str, port: int = 443, timeout: int = 10):
    """Connect to hostname:port and return its leaf cert as a cryptography x509 object."""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE  # we only want to *inspect* it here
    with socket.create_connection((hostname, port), timeout=timeout) as sock:
        with ctx.wrap_socket(sock, server_hostname=hostname) as ssock:
            der = ssock.getpeercert(binary_form=True)
    return x509.load_der_x509_certificate(der, default_backend())


def _aia_ca_issuer_urls(cert):
    """Return any 'CA Issuers' URLs from a cert's Authority Information Access extension."""
    try:
        aia = cert.extensions.get_extension_for_oid(
            ExtensionOID.AUTHORITY_INFORMATION_ACCESS
        ).value
    except x509.ExtensionNotFound:
        return []
    return [
        desc.access_location.value
        for desc in aia
        if desc.access_method == AuthorityInformationAccessOID.CA_ISSUERS
    ]


def _fetch_missing_intermediates(hostname: str, max_hops: int = 5):
    """
    Walk the AIA chain for `hostname`'s certificate and return a list of
    intermediate certs (as PEM strings) that were fetched. Empty list if
    none were needed or any step failed (network issues here just mean we
    fall back to the plain certifi bundle, so failures are non-fatal).
    """
    pem_certs = []
    try:
        cert = _get_leaf_certificate(hostname)
    except Exception as exc:
        log.warning("Could not inspect %s's TLS certificate: %s", hostname, exc)
        return pem_certs

    seen_urls = set()
    for _ in range(max_hops):
        if cert.issuer == cert.subject:
            break  # reached a self-signed root; nothing more to fetch
        urls = _aia_ca_issuer_urls(cert)
        if not urls:
            break
        url = urls[0]
        if url in seen_urls:
            break
        seen_urls.add(url)
        try:
            resp = requests.get(url, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            issuer_cert = x509.load_der_x509_certificate(resp.content, default_backend())
        except Exception as exc:
            log.warning("Could not fetch intermediate certificate from %s: %s", url, exc)
            break
        pem_certs.append(issuer_cert.public_bytes(serialization.Encoding.PEM).decode("ascii"))
        cert = issuer_cert  # walk up towards the root
    return pem_certs


def build_repaired_verify_bundle(hostnames):
    """
    For each hostname, check whether the server's TLS handshake is missing
    intermediate certificates and, if so, fetch them via AIA chasing.
    Returns a filesystem path suitable for `requests`' `verify=` argument:
    a temp file containing certifi's roots plus any fetched intermediates,
    or plain certifi.where() if nothing extra was found/needed.
    """
    if not CRYPTOGRAPHY_AVAILABLE:
        log.info(
            "The 'cryptography' package isn't installed, so automatic SSL "
            "chain repair is skipped (pip install cryptography to enable it). "
            "Falling back to certifi's standard bundle."
        )
        return certifi.where()

    extra_pem_blocks = []
    for host in hostnames:
        found = _fetch_missing_intermediates(host)
        if found:
            log.info("Fetched %d missing intermediate certificate(s) for %s.", len(found), host)
            extra_pem_blocks.extend(found)

    if not extra_pem_blocks:
        return certifi.where()

    with open(certifi.where(), "r", encoding="ascii") as f:
        base_bundle = f.read()

    combined = base_bundle + "\n" + "\n".join(extra_pem_blocks)
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".pem", prefix="knbs_ca_bundle_", delete=False
    )
    tmp.write(combined)
    tmp.close()
    log.info("Built repaired CA bundle at %s", tmp.name)
    return tmp.name


def fetch(session: requests.Session, url: str, params: dict = None):
    """
    GET a URL with retries, a timeout, and error handling.
    Returns a BeautifulSoup object, or None on failure.
    """
    for attempt in range(1, MAX_RETRIES + 2):
        try:
            resp = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if resp.status_code == 200:
                return BeautifulSoup(resp.text, "html.parser")
            elif resp.status_code == 404:
                log.info("404 (likely past the last page): %s", resp.url)
                return None
            else:
                log.warning(
                    "Unexpected status %s for %s (attempt %d/%d)",
                    resp.status_code, url, attempt, MAX_RETRIES + 1,
                )
        except requests.exceptions.SSLError as exc:
            log.warning(
                "SSL error for %s (attempt %d/%d): %s",
                url, attempt, MAX_RETRIES + 1, exc,
            )
            log.warning(
                "If this keeps happening, either your machine's CA trust "
                "store is missing/broken, or the server itself is failing "
                "to send its full certificate chain. This script already "
                "tries to auto-repair the latter (see build_repaired_verify_bundle) "
                "if 'cryptography' is installed: pip install cryptography. "
                "As a last-resort diagnostic only, you can set VERIFY_SSL = False "
                "at the top of this script — do not leave that on for real use."
            )
        except requests.exceptions.RequestException as exc:
            log.warning(
                "Request error for %s (attempt %d/%d): %s",
                url, attempt, MAX_RETRIES + 1, exc,
            )
        time.sleep(REQUEST_DELAY)
    log.error("Giving up on %s after %d attempts", url, MAX_RETRIES + 1)
    return None


# --------------------------------------------------------------------------
# Shared helpers
# --------------------------------------------------------------------------

def clean_text(text: str) -> str:
    """Collapse whitespace/newlines and strip a string."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def extract_year(*texts: str) -> str:
    """
    Regex-extract a 4-digit year (19xx/20xx) from the first candidate
    string that contains one. Returns "" if none found.
    """
    for text in texts:
        if not text:
            continue
        match = YEAR_PATTERN.search(text)
        if match:
            return match.group(0)
    return ""


def find_excel_links(soup: BeautifulSoup, base_url: str):
    """Return all absolute .xlsx/.xls links found in a page's <a> tags."""
    links = []
    if soup is None:
        return links
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if EXCEL_LINK_PATTERN.search(href):
            links.append(urljoin(base_url, href))
    # de-duplicate while preserving order
    seen = set()
    unique_links = []
    for link in links:
        if link not in seen:
            seen.add(link)
            unique_links.append(link)
    return unique_links


def get_excel_link_from_detail_page(session: requests.Session, detail_url: str):
    """Visit a detail page and return the first .xlsx/.xls link found, if any."""
    if not FOLLOW_DETAIL_PAGES or not detail_url:
        return None
    soup = fetch(session, detail_url)
    time.sleep(REQUEST_DELAY)
    if soup is None:
        return None
    excel_links = find_excel_links(soup, detail_url)
    return excel_links[0] if excel_links else None


# --------------------------------------------------------------------------
# 1) KNBS main website — /all-reports/
# --------------------------------------------------------------------------

def scrape_website_reports(session: requests.Session):
    """
    Crawl all paginated pages of https://www.knbs.or.ke/all-reports/,
    extracting each report's title, year, and (if available) a direct
    .xlsx/.xls download link.
    """
    records = []
    seen_links = set()

    for page in range(1, MAX_PAGES_WEBSITE + 1):
        page_url = WEBSITE_REPORTS_URL if page == 1 else f"{WEBSITE_REPORTS_URL}page/{page}/"
        log.info("[Website] Fetching page %d: %s", page, page_url)
        soup = fetch(session, page_url)
        time.sleep(REQUEST_DELAY)

        if soup is None:
            log.info("[Website] No more pages (stopped at page %d).", page)
            break

        # Report entries link out to individual report detail pages, which on
        # this site live under /reports/<slug>/. We collect every such link
        # on the listing page rather than relying on a specific card/div
        # class, since that markup can change with template updates.
        page_links = []
        for a in soup.find_all("a", href=True):
            href = urljoin(WEBSITE_BASE, a["href"].strip())
            path = urlparse(href).path
            if re.match(r"^/reports/[^/]+/?$", path):
                title = clean_text(a.get_text())
                if title:  # skip empty/icon-only anchors pointing to the same card
                    page_links.append((href, title))

        if not page_links:
            log.info("[Website] No report links found on page %d — assuming end of pagination.", page)
            break

        new_items_found = False
        for href, title in page_links:
            if href in seen_links:
                continue
            seen_links.add(href)
            new_items_found = True

            # Direct xlsx/xls link straight from the listing card, if present
            # (rare, but some "download" buttons point straight at a file).
            direct_excel = None
            # look for an excel link within the same anchor's parent block
            parent = None
            for a in soup.find_all("a", href=True):
                if urljoin(WEBSITE_BASE, a["href"].strip()) == href:
                    parent = a.find_parent(["article", "div", "li"])
                    break
            if parent:
                nearby_excel = find_excel_links(parent, WEBSITE_BASE)
                if nearby_excel:
                    direct_excel = nearby_excel[0]

            # Otherwise, follow the detail page to look for .xlsx/.xls links
            # (e.g. "Chapter-1-International-Scene.xlsx" style data tables).
            excel_url = direct_excel or get_excel_link_from_detail_page(session, href)

            year = extract_year(title, href)
            if not year and FOLLOW_DETAIL_PAGES:
                # Fall back to scanning detail page text for a year if the
                # title/slug didn't contain one.
                detail_soup = fetch(session, href)
                time.sleep(REQUEST_DELAY)
                if detail_soup is not None:
                    year = extract_year(clean_text(detail_soup.get_text(" ")))

            records.append({
                "Name": title,
                "Year": year,
                "Source": "Website",
                "Has_XLSX": "Yes" if excel_url else "No",
                "Excel_File_URL": excel_url or href,
            })

        if not new_items_found:
            log.info("[Website] Page %d had no new items — stopping pagination.", page)
            break

    log.info("[Website] Collected %d report records.", len(records))
    return records


# --------------------------------------------------------------------------
# 2) KeNADA microdata catalog
# --------------------------------------------------------------------------

def scrape_kenada_catalog(session: requests.Session):
    """
    Crawl the KeNADA central catalog, paginating with `page` and using a
    larger `ps` (page size) to reduce the number of requests. Extracts each
    study's title, year, and (if available) a direct .xlsx/.xls link found
    either on the listing or the study's own detail page.
    """
    records = []
    seen_links = set()

    for page in range(1, MAX_PAGES_KENADA + 1):
        params = {"page": page, "ps": KENADA_PAGE_SIZE}
        log.info("[KeNADA] Fetching page %d (ps=%d)", page, KENADA_PAGE_SIZE)
        soup = fetch(session, KENADA_CATALOG_URL, params=params)
        time.sleep(REQUEST_DELAY)

        if soup is None:
            log.info("[KeNADA] No more pages (stopped at page %d).", page)
            break

        # Catalog entries link to individual studies at /catalog/<id> (the
        # NADA/DDI toolkit used by KeNADA). We match on that URL pattern
        # rather than a specific CSS class for the same resilience reason
        # as above.
        page_entries = []
        for a in soup.find_all("a", href=True):
            href = urljoin(KENADA_BASE, a["href"].strip())
            path = urlparse(href).path
            if re.match(r"^/nada/index\.php/catalog/\d+(/.*)?$", path):
                title = clean_text(a.get_text())
                if title:
                    # normalize to the study's canonical page (strip any
                    # sub-tab like /related-materials, /export)
                    match = re.match(r"^(/nada/index\.php/catalog/\d+)", path)
                    canonical_path = match.group(1) if match else path
                    canonical_url = f"{KENADA_BASE}{canonical_path}"
                    page_entries.append((canonical_url, title))

        if not page_entries:
            log.info("[KeNADA] No catalog entries found on page %d — assuming end of pagination.", page)
            break

        new_items_found = False
        for href, title in page_entries:
            if href in seen_links:
                continue
            seen_links.add(href)
            new_items_found = True

            year = extract_year(title)

            excel_url = None
            if FOLLOW_DETAIL_PAGES:
                detail_soup = fetch(session, href)
                time.sleep(REQUEST_DELAY)
                if detail_soup is not None:
                    excel_links = find_excel_links(detail_soup, href)
                    if excel_links:
                        excel_url = excel_links[0]
                    if not year:
                        year = extract_year(clean_text(detail_soup.get_text(" ")))

            records.append({
                "Name": title,
                "Year": year,
                "Source": "KeNADA",
                "Has_XLSX": "Yes" if excel_url else "No",
                "Excel_File_URL": excel_url or href,
            })

        if not new_items_found:
            log.info("[KeNADA] Page %d had no new items — stopping pagination.", page)
            break

        # Stop early if the site returned fewer entries than the requested
        # page size — that means we're on the last page.
        if len(page_entries) < KENADA_PAGE_SIZE:
            log.info("[KeNADA] Page %d returned fewer than %d entries — last page reached.",
                      page, KENADA_PAGE_SIZE)
            break

    log.info("[KeNADA] Collected %d study records.", len(records))
    return records


# --------------------------------------------------------------------------
# Data assembly & export
# --------------------------------------------------------------------------

def build_dataframe(website_records, kenada_records) -> pd.DataFrame:
    """Combine, clean, and deduplicate records from both sources."""
    all_records = website_records + kenada_records
    df = pd.DataFrame(all_records, columns=["Name", "Year", "Source", "Has_XLSX", "Excel_File_URL"])

    if df.empty:
        log.warning("No records were collected from either source.")
        return df

    df["Name"] = df["Name"].apply(clean_text)
    df["Year"] = df["Year"].apply(lambda y: clean_text(y) if y else "")

    # Deduplicate identical records (same name + source + link)
    before = len(df)
    df = df.drop_duplicates(subset=["Name", "Source", "Excel_File_URL"]).reset_index(drop=True)
    log.info("Deduplicated %d -> %d records.", before, len(df))

    df = df.sort_values(by=["Source", "Year", "Name"], ascending=[True, False, True]).reset_index(drop=True)
    return df


def export_to_excel(df: pd.DataFrame, filename: str = OUTPUT_FILE):
    """Write the final DataFrame to an .xlsx file with tidy column widths."""
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="KNBS_Reports")
        worksheet = writer.sheets["KNBS_Reports"]
        widths = {"A": 70, "B": 10, "C": 12, "D": 12, "E": 70}
        for col, width in widths.items():
            worksheet.column_dimensions[col].width = width
    log.info("Saved %d records to %s", len(df), filename)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    verify_path = None
    if VERIFY_SSL:
        log.info("Checking TLS certificate chains for the target sites...")
        verify_path = build_repaired_verify_bundle([
            urlparse(WEBSITE_BASE).hostname,
            urlparse(KENADA_BASE).hostname,
        ])

    session = get_session(verify_path)

    log.info("=== Scraping KNBS website reports ===")
    website_records = scrape_website_reports(session)

    log.info("=== Scraping KeNADA microdata catalog ===")
    kenada_records = scrape_kenada_catalog(session)

    log.info("=== Building combined dataset ===")
    df = build_dataframe(website_records, kenada_records)

    if df.empty:
        log.error("Nothing to export — check network access / site structure changes.")
        return

    export_to_excel(df, OUTPUT_FILE)
    print(df.head(20).to_string(index=False))


if __name__ == "__main__":
    main()