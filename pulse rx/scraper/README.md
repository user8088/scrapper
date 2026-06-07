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
- **Delay (s)** — polite pause before each product (default 0.8). Lower = faster.
- **Threads** — how many products to scrape **concurrently** (default 4, max 16). The
  scraper is network-bound, so this is the biggest speed lever. Higher = faster, but
  you hit dwatson/healthwire harder: too many threads risks rate-limiting, timeouts,
  or temporary IP blocks (which can look like "not found"). 4–8 is a safe range; set
  it to **1** for the original gentle, sequential behaviour.
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

**Download report**
- *Download report (xlsx)* exports a polished, self-contained workbook of the
  current results:
  - **Summary** sheet — run statistics (total / found / not found / errors,
    success rate, and a breakdown of descriptions by source).
  - **Results** sheet — every product colour-marked by status (green = found,
    amber = not found, red = error, grey = no CSV row), with the **matched name**,
    **score**, a clickable **source URL** (where the description came from), and the
    full description. Header row is frozen with auto-filters for easy sorting.

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

## Distribution & auto-updates

The app ships as a **single `PulseRxScraper.exe`** — users download one file and
run it, no Python needed. On launch it quietly checks GitHub Releases and, if a
newer version exists, offers to download and restart into it. So you publish
once and every install updates itself.

### For users
Download the latest `PulseRxScraper.exe` from
**https://github.com/user8088/scrapper/releases/latest** and run it. Place it
wherever is convenient — on first run, use **Browse** to point *Products folder*
and *Item List CSV* at your data (it remembers them in `config.json` next to the
exe). When an update is published, click **Yes** on the prompt at startup (or
**Help → Check for updates…**) and it self-updates.

### For you — publishing an update
1. Make your changes.
2. Run the release helper with the new version number:
   ```powershell
   cd "pulse rx/scraper"
   .\release.ps1 1.1.0
   ```
   This bumps `version.py`, commits, tags `v1.1.0`, and pushes. **GitHub Actions**
   (`.github/workflows/release.yml`) builds the exe and publishes the release.
3. Existing installs detect the newer tag on their next launch and update.

> The version in `version.py` **must** match the tag — `release.ps1` keeps them
> in sync, and CI fails the build if they ever diverge.

### Building the exe locally (optional)
```powershell
cd "pulse rx/scraper"
.\build.ps1          # -> dist\PulseRxScraper.exe
```

## Files

- `scraper_core.py` — UI-agnostic engine (site clients, matching, xlsx writer). Importable + CLI.
- `app.py` — Tkinter desktop UI (+ auto-update on launch).
- `updater.py` — checks GitHub Releases, downloads & swaps in the new exe.
- `version.py` — the app's version (source of truth; bumped by `release.ps1`).
- `build.ps1` — builds the standalone exe with PyInstaller.
- `release.ps1` — bumps the version, tags, and pushes to trigger a release.
- `requirements.txt` — dependencies.
- `run.bat` — Windows launcher (run from source).
- `config.json` — saved settings (created on first *Save settings*).

## Notes / tuning

- The scraper is polite (single-threaded with a configurable delay). For ~4,500
  products at 0.8 s that is roughly an hour; lower the delay to speed it up.
- If a site changes its HTML, only the small `DwatsonClient` / `HealthwireClient`
  classes in `scraper_core.py` need adjusting.
- Writing preserves the existing single-column structure used across all
  product files (header in A1, description in A2).
