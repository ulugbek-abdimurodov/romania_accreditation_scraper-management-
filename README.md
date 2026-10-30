# romania_scraper

Scrapes RENAR (https://www.renar.ro/index.php/oec) for accredited management-system
certification bodies, categories **5.2.1 – 5.2.8** (~92 records as of Oct 2026).

For each company: name, address, city, phone, fax, email, website, standard
(Referențial), certificate no., status, issue date, expiration date, and the
**Certificat** and **Anexa** PDFs when published.

## Run
Windows: double-click `run.bat` · macOS/Linux: `./run.sh`
(or `pip install -r requirements.txt` then `python scraper.py`)

## Options
| flag | meaning |
|---|---|
| `--acts 8 9` | only some categories (8=5.2.1 Calitate, 9=5.2.2 Mediu, 10=5.2.3 Dispozitive medicale, 11=5.2.4 SSM, 12=5.2.5 Energie, 13=5.2.6 Securitatea informației, 7=5.2.7 Siguranța alimentelor, 22=5.2.8 Anti-mită) |
| `--status ACC` | only accredited (`SUSP`, `RETR_P`, `RETR` also available) |
| `--no-pdfs` | data only |
| `--workers 4 --delay 0.5` | politeness / speed |
| `--out output` | output folder |

## Output (`output/`)
- `companies.xlsx`, `companies.csv` (UTF‑8, Excel friendly), `companies.json`
- `pdfs/<category>/<Company>__<cert no>/Certificat.pdf`, `Anexa.pdf`
- `scraper.log`

Re-running is safe: already downloaded PDFs are skipped.
A company accredited in several categories appears once per category.

## How it works
The site's table is a DataTables server-side grid backed by
`/index.php/oec/get_oecs?act_id=<id>` (JSON). The scraper pages through it, then opens
`/index.php/oec/get_oec_details/<id>` for every row and parses the two detail tables
(labels matched diacritic-insensitively). PDFs are Alfresco shared links on cms.renar.ro.
