# Pulse Rx — Product Description Scraper

Matches every product in `pulse rx/Products/<ItemId>/` against two pharmacy
sites, fetches a rich product description, and writes it into that product's
`descriptions/product_details.xlsx` (cell **A2**, under the **Details** header).

It is built to be **perfect-by-verification**: scrape in batches, eyeball each
match in the UI, then commit only what you approve — or let it write directly.

---

## How matching works

| Site | Search | Description source | Notes |
|------|--------|--------------------|-------|
| **dwatson.pk** | `/catalogsearch/result/?q=<name>` (server-rendered) | `og:description` meta tag — the clean, structured *"Product Overview / Key Benefits / How It Works / ..."* format | **Preferred output format.** A match is only accepted when the name score ≥ *Min match score* **and** the dosage/strength tokens agree (so "Daplyza 10mg" never matches "Daplyza 5mg"). |
| **healthwire.pk** | `/searches` JSON API (`model=items`) | Product-page body sections: Description, Uses, Dosage, How To Use, Side Effects, Precautions, Drug Interactions, Storage | Healthwire's `item_id` **equals your Pulse Rx Item Id**, so matches are authoritative (100% reliable). Great fallback. |

For each product the engine tries the sources **in priority order** and uses the
first one that returns a usable description. If a site has no confident match,
it falls back to the other. (`Find in both — use whichever matches.`)

---

## Setup

```bash
cd "pulse rx/scraper"
python -m pip install -r requirements.txt
```

Python 3.10+ recommended (tested on 3.13).

## Run the UI

```bash
python app.py
```
or just double-click **`run.bat`** (installs deps then launches).

### UI walkthrough

**Configuration**
- **Products folder** / **Item List CSV** — auto-filled to the Pulse Rx paths; change with *Browse*.
- **Batch size** — how many products per run (default **50**, fully editable — set it to anything).
- **Start offset** — index into the product list to begin at. Auto-advances by the batch size after each run, so clicking *Run batch* repeatedly walks the whole catalog.
- **Min match score** — strictness for dwatson fuzzy matching (default 80).
- **Delay (s)** — polite pause between products (default 0.8). Lower = faster.
- **Source priority** — `dwatson, healthwire` (default) / `healthwire, dwatson` / single-site.
- **xlsx header** — kept in A1. The engine preserves the file's existing header if present.
- **Skip products already filled** — ignore products whose A2 already has text.
- **Save settings** — persists everything to `config.json`.

**Mode**
- **Preview then commit** — scrape a batch, review each row, then **Write selected to xlsx**.
- **Write directly while scraping** — writes each found description immediately.

**Results table**
- Click the **☑ / ☐** in the first column to include/exclude a row.
- **Double-click** any row to read the full description, open the source URL, or write just that one.
- *Select all* / *Deselect all* / *Clear results* as needed.

**Committing**
- *Write selected to xlsx* writes every checked, found row into its
  `product_details.xlsx` (A1 = header, A2 = description, wrapped).

---

## Command-line (optional, for automation)

```bash
# Preview 10 products starting at offset 0 (no writing)
python scraper_core.py --offset 0 --limit 10

# Scrape and WRITE 50 products, healthwire first
python scraper_core.py --offset 0 --limit 50 --order healthwire,dwatson --write
```

Flags: `--root`, `--csv`, `--offset`, `--limit`, `--order`, `--min-score`,
`--delay`, `--write`.

---

## Files

- `scraper_core.py` — UI-agnostic engine (site clients, matching, xlsx writer). Importable + CLI.
- `app.py` — Tkinter desktop UI.
- `requirements.txt` — dependencies.
- `run.bat` — Windows launcher.
- `config.json` — saved settings (created on first *Save settings*).

## Notes / tuning

- The scraper is polite (single-threaded with a configurable delay). For ~4,500
  products at 0.8 s that is roughly an hour; lower the delay to speed it up.
- If a site changes its HTML, only the small `DwatsonClient` / `HealthwireClient`
  classes in `scraper_core.py` need adjusting.
- Writing preserves the existing single-column structure used across all
  product files (header in A1, description in A2).
