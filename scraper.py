#!/usr/bin/env python3
"""
RENAR scraper - accredited management-system certification bodies (OEC).

Source:   https://www.renar.ro/index.php/oec
List API: https://www.renar.ro/index.php/oec/get_oecs        (DataTables server-side JSON)
Details:  https://www.renar.ro/index.php/oec/get_oec_details/<id>

Collects, for every company in the selected RENAR categories (default: 5.2.1 - 5.2.8):
  name, address, phone, email, website, certificate standard, expiration date,
  plus the "Certificat" and "Anexa" PDF files (when published).

Outputs (in ./output):
  companies.csv, companies.xlsx, companies.json
  pdfs/<category>/<company>__<cert no>/Certificat.pdf, Anexa.pdf
  scraper.log

Usage:
  python scraper.py                    # all 5.2.x categories, all statuses
  python scraper.py --acts 8           # only 5.2.1 Calitate
  python scraper.py --status ACC       # only currently accredited
  python scraper.py --no-pdfs          # data only
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://www.renar.ro"
LIST_URL = f"{BASE}/index.php/oec/get_oecs"
DETAIL_URL = f"{BASE}/index.php/oec/get_oec_details/{{id}}"

# value of the "domain" <select> on the site  ->  RENAR code / label
CATEGORIES: dict[int, str] = {
    8: "5.2.1 Calitate",
    9: "5.2.2 Mediu",
    10: "5.2.3 Calitate pentru dispozitive medicale",
    11: "5.2.4 Sanatate si securitate in munca",
    12: "5.2.5 Energie",
    13: "5.2.6 Securitatea informatiei",
    7: "5.2.7 Siguranta alimentelor",
    22: "5.2.8 Anti-mita",
}
STATUSES = {"ACC": "Acreditate", "SUSP": "Suspendata", "RETR_P": "Restransa", "RETR": "Retrasa"}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
    "Accept-Language": "ro-RO,ro;q=0.9,en;q=0.8",
    "Referer": f"{BASE}/index.php/oec",
}

log = logging.getLogger("renar")


# --------------------------------------------------------------------------- model
@dataclass
class Company:
    renar_id: str
    category_code: str
    category: str
    name: str = ""
    laboratory: str = ""
    city: str = ""
    address: str = ""
    phone: str = ""
    fax: str = ""
    email: str = ""
    website: str = ""
    activity: str = ""
    status: str = ""
    standard: str = ""
    certificate_no: str = ""
    issue_date: str = ""
    expiration_date: str = ""
    certificate_pdf_url: str = ""
    annex_pdf_url: str = ""
    certificate_pdf_file: str = ""
    annex_pdf_file: str = ""
    detail_url: str = ""
    extra: dict = field(default_factory=dict)


# --------------------------------------------------------------------------- helpers
def norm(text: str) -> str:
    """lower-case, strip diacritics (ş/ș, ţ/ț, ă, â, î) and extra spaces."""
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text).strip().lower()


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def safe_name(text: str, max_len: int = 80) -> str:
    """Filename safe on Windows / macOS / Linux."""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:max_len] or "unnamed"


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(total=5, backoff_factor=1.5, status_forcelist=(429, 500, 502, 503, 504),
                  allowed_methods=("GET",), raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retry, pool_maxsize=10)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


# --------------------------------------------------------------------------- list
def fetch_list(session: requests.Session, act_id: int, status: str | None,
               page_size: int = 100) -> list[list[str]]:
    """Page through the DataTables JSON endpoint for one category."""
    rows: list[list[str]] = []
    start, draw = 0, 1
    while True:
        params = {
            "draw": draw, "start": start, "length": page_size,
            "order[0][column]": 2, "order[0][dir]": "asc",
            "search[value]": "", "search[regex]": "false",
            "act_id": act_id,
        }
        for i in range(8):  # the endpoint expects the column definitions
            params[f"columns[{i}][data]"] = i
        if status:
            params["status"] = status
        r = session.get(LIST_URL, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        batch = data.get("data", [])
        rows.extend(batch)
        total = int(data.get("recordsFiltered", 0))
        log.info("  list act_id=%s: %d/%d", act_id, len(rows), total)
        if not batch or len(rows) >= total:
            break
        start += page_size
        draw += 1
    return rows


# --------------------------------------------------------------------------- detail
LABELS = {  # normalised label on page -> Company attribute
    "persoana juridica": "name",
    "laborator": "laboratory",
    "localitate": "city",
    "adresa": "address",
    "telefon": "phone",
    "fax": "fax",
    "email": "email",
    "e-mail": "email",
    "website": "website",
    "activitate": "activity",
    "stare": "status",
    "referential": "standard",
    "numar certificat": "certificate_no",
    "data emiterii": "issue_date",
    "data expirarii": "expiration_date",
    "certificat": "certificate_pdf_url",
    "anexa": "annex_pdf_url",
}
LINK_FIELDS = {"email", "website", "certificate_pdf_url", "annex_pdf_url"}


def parse_table(table) -> dict[str, str]:
    out: dict[str, str] = {}
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 3:
            continue
        label, cell = norm(tds[1].get_text()), tds[2]
        attr = LABELS.get(label, label)
        if attr in LINK_FIELDS:
            a = cell.find("a", href=True)
            href = a["href"].strip() if a else ""
            if attr == "email":
                val = href[7:] if href.lower().startswith("mailto:") else clean(cell.get_text())
            elif attr == "website":
                val = clean(cell.get_text()) or href
                if val.startswith("//"):
                    val = "http:" + val
            else:
                val = urljoin(BASE, href) if href else ""
        else:
            val = clean(cell.get_text(" "))
        out[attr] = val
    return out


def fetch_detail(session: requests.Session, c: Company) -> Company:
    r = session.get(c.detail_url, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    tables = [parse_table(t) for t in soup.find_all("table")]
    tables = [t for t in tables if t]
    if not tables:
        log.warning("No data tables on %s", c.detail_url)
        return c

    merged: dict[str, str] = {}
    company_tbl = next((t for t in tables if "name" in t), tables[0])
    merged.update(company_tbl)
    # an accreditation block; if several exist pick the one with our certificate number
    acc_tbls = [t for t in tables if "certificate_no" in t or "standard" in t]
    if acc_tbls:
        acc = next((t for t in acc_tbls if t.get("certificate_no") == c.certificate_no), acc_tbls[0])
        merged.update(acc)

    known = set(Company.__dataclass_fields__)
    for k, v in merged.items():
        if k in known and k not in ("extra",):
            if v or not getattr(c, k):
                setattr(c, k, v)
        else:
            c.extra[k] = v
    return c


# --------------------------------------------------------------------------- pdfs
def download_file(session: requests.Session, url: str, dest_no_ext: Path) -> str:
    """Download url -> dest(.pdf/.ext). Skips if already present. Returns relative path or ''."""
    if not url:
        return ""
    for existing in dest_no_ext.parent.glob(dest_no_ext.name + ".*"):
        if existing.stat().st_size > 0:
            return str(existing)
    dest_no_ext.parent.mkdir(parents=True, exist_ok=True)
    try:
        with session.get(url, timeout=120, stream=True) as r:
            r.raise_for_status()
            ctype = r.headers.get("Content-Type", "").lower()
            ext = ".pdf"
            cd = r.headers.get("Content-Disposition", "")
            m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', cd, re.I)
            if m and "." in m.group(1):
                ext = "." + m.group(1).rsplit(".", 1)[-1].lower()[:5]
            elif "pdf" not in ctype and "octet" not in ctype and ctype:
                ext = "." + ctype.split("/")[-1].split(";")[0][:5]
            dest = dest_no_ext.with_suffix(ext)
            tmp = dest.with_suffix(ext + ".part")
            with open(tmp, "wb") as fh:
                for chunk in r.iter_content(65536):
                    fh.write(chunk)
            if ext == ".pdf":
                with open(tmp, "rb") as fh:
                    if fh.read(4) != b"%PDF":
                        log.warning("Not a PDF (%s): %s", ctype, url)
            tmp.replace(dest)
            return str(dest)
    except Exception as e:  # noqa: BLE001
        log.error("Download failed %s -> %s", url, e)
        return ""


# --------------------------------------------------------------------------- export
COLUMNS = [
    "category_code", "category", "name", "address", "city", "phone", "fax", "email", "website",
    "standard", "certificate_no", "status", "issue_date", "expiration_date", "activity",
    "laboratory", "certificate_pdf_url", "annex_pdf_url", "certificate_pdf_file",
    "annex_pdf_file", "detail_url", "renar_id",
]


def export(companies: list[Company], out_dir: Path) -> None:
    rows = [{k: asdict(c)[k] for k in COLUMNS} for c in companies]

    with open(out_dir / "companies.json", "w", encoding="utf-8") as fh:
        json.dump([asdict(c) for c in companies], fh, ensure_ascii=False, indent=2)

    # utf-8-sig so Excel shows Romanian diacritics correctly
    with open(out_dir / "companies.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        ws = wb.active
        ws.title = "Companies"
        ws.append(COLUMNS)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for r in rows:
            ws.append([r[k] for k in COLUMNS])
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("http"):
                    cell.hyperlink = v
                    cell.font = Font(color="0563C1", underline="single")
        widths = {"name": 40, "address": 50, "activity": 50, "category": 32}
        for i, col in enumerate(COLUMNS, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 20)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        wb.save(out_dir / "companies.xlsx")
    except ImportError:
        log.warning("openpyxl not installed - skipped companies.xlsx")


# --------------------------------------------------------------------------- main
def run(args: argparse.Namespace) -> int:
    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = out_dir / "pdfs"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        handlers=[logging.StreamHandler(sys.stdout),
                  logging.FileHandler(out_dir / "scraper.log", encoding="utf-8")],
    )
    session = make_session()

    # 1) collect list rows
    companies: list[Company] = []
    for act_id in args.acts:
        label = CATEGORIES.get(act_id, str(act_id))
        log.info("Category %s (act_id=%s)", label, act_id)
        for row in fetch_list(session, act_id, args.status):
            # row: [id, cert no, name, laboratory, city, issue, expiry, standard]
            rid = str(row[0])
            companies.append(Company(
                renar_id=rid,
                category_code=label.split(" ")[0],
                category=label,
                certificate_no=clean(str(row[1])),
                name=clean(BeautifulSoup(str(row[2]), "html.parser").get_text()),
                laboratory=clean(str(row[3])),
                city=clean(str(row[4])),
                standard=clean(str(row[7])) if len(row) > 7 else "",
                detail_url=DETAIL_URL.format(id=rid),
            ))
        time.sleep(args.delay)
    log.info("Total records: %d", len(companies))

    # 2) details (+ pdfs), polite concurrency
    def work(c: Company) -> Company:
        try:
            fetch_detail(session, c)
        except Exception as e:  # noqa: BLE001
            log.error("Detail failed %s (%s): %s", c.name, c.detail_url, e)
        if not args.no_pdfs:
            folder = pdf_dir / safe_name(c.category, 60) / safe_name(
                f"{c.name}__{c.certificate_no}".replace(" ", "_"), 100)
            c.certificate_pdf_file = download_file(session, c.certificate_pdf_url, folder / "Certificat")
            c.annex_pdf_file = download_file(session, c.annex_pdf_url, folder / "Anexa")
            for attr in ("certificate_pdf_file", "annex_pdf_file"):
                p = getattr(c, attr)
                if p:
                    setattr(c, attr, str(Path(p).relative_to(out_dir)))
        time.sleep(args.delay)
        return c

    done = 0
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = [pool.submit(work, c) for c in companies]
        for f in as_completed(futures):
            c = f.result()
            done += 1
            log.info("[%d/%d] %s | %s | exp %s | pdf: %s/%s", done, len(companies), c.name,
                     c.certificate_no, c.expiration_date,
                     "Y" if c.certificate_pdf_file else "-", "Y" if c.annex_pdf_file else "-")

    companies.sort(key=lambda c: (c.category_code, c.name.lower()))
    export(companies, out_dir)

    n_cert = sum(bool(c.certificate_pdf_file) for c in companies)
    n_anx = sum(bool(c.annex_pdf_file) for c in companies)
    log.info("Done. %d records | %d certificate PDFs | %d annex PDFs | output: %s",
             len(companies), n_cert, n_anx, out_dir)
    return 0


def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Scrape RENAR accredited certification bodies (5.2.x).")
    p.add_argument("--acts", type=int, nargs="+", default=list(CATEGORIES),
                   help="RENAR category ids (8=5.2.1 ... 22=5.2.8). Default: all 5.2.x")
    p.add_argument("--status", choices=list(STATUSES), default=None,
                   help="Filter by status (ACC = accredited). Default: all")
    p.add_argument("--out", default="output", help="Output folder (default: ./output)")
    p.add_argument("--workers", type=int, default=4, help="Parallel requests (default 4)")
    p.add_argument("--delay", type=float, default=0.5, help="Pause between requests per worker (s)")
    p.add_argument("--no-pdfs", action="store_true", help="Skip PDF downloads")
    return p.parse_args(argv)


if __name__ == "__main__":
    sys.exit(run(parse_args()))
