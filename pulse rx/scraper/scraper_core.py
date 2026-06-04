"""
High-end product description scraper for Pulse Rx.

Matches each product folder (named by Item Id) against two pharmacy sites and
fetches a rich product description:

  * dwatson.pk   -> Magento store. Search at /catalogsearch/result/?q=...
                    The full structured marketing description lives in the
                    `og:description` meta tag (the preferred output format).
  * healthwire.pk-> Rails store. JSON search at /searches. Its `item_id`
                    matches the Pulse Rx Item Id exactly, giving an
                    authoritative match. Description is built from the product
                    page body sections (Description / Uses / Dosage / ...).

The module is UI-agnostic: `process_product` returns a structured result and
`write_description` commits a description into the product's xlsx file.
"""

from __future__ import annotations

import csv
import os
import re
import time
import json
import threading
from dataclasses import dataclass, field, asdict
from typing import Callable, Optional

import requests
import urllib3
from bs4 import BeautifulSoup
from rapidfuzz import fuzz
from openpyxl import load_workbook, Workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# Sources
SOURCE_DWATSON = "dwatson"
SOURCE_HEALTHWIRE = "healthwire"

# Statuses
STATUS_FOUND = "found"
STATUS_NOT_FOUND = "not_found"
STATUS_NO_CSV = "no_csv_row"
STATUS_ERROR = "error"
STATUS_LOW_CONFIDENCE = "low_confidence"


# --------------------------------------------------------------------------- #
# Text helpers
# --------------------------------------------------------------------------- #
_STRENGTH_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(mcg|mg|ml|gm|g|iu|%|mg/ml)", re.I)
_PACK_RE = re.compile(r"(\d+)\s*(?:s\b|tablets?|capsules?|caps?|tabs?|sachets?|pcs?|pieces?|ampoules?|vials?)", re.I)


def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def strength_tokens(name: str) -> set:
    """Extract normalized strength tokens like {'18mcg', '500mg', '60ml'}."""
    out = set()
    for num, unit in _STRENGTH_RE.findall(name or ""):
        num = num.rstrip("0").rstrip(".") if "." in num else num
        out.add(f"{num}{unit.lower()}")
    return out


def pack_tokens(name: str) -> set:
    return {m for m in _PACK_RE.findall(name or "")}


def name_score(a: str, b: str) -> float:
    return fuzz.token_sort_ratio((a or "").lower(), (b or "").lower())


def is_strong_match(query_name: str, candidate_name: str, min_score: float) -> tuple[bool, float, str]:
    """
    Decide whether candidate_name is the same product as query_name.
    Requires a good fuzzy score AND consistent strength tokens (no conflicts).
    Returns (ok, score, reason).
    """
    score = name_score(query_name, candidate_name)
    q_str = strength_tokens(query_name)
    c_str = strength_tokens(candidate_name)

    # If both have strength tokens, they must overlap (no conflicting dosage).
    if q_str and c_str and not (q_str & c_str):
        return False, score, f"strength mismatch {sorted(q_str)} vs {sorted(c_str)}"

    if score < min_score:
        return False, score, f"score {score:.0f} < {min_score:.0f}"

    return True, score, "ok"


# --------------------------------------------------------------------------- #
# Catalog (CSV)
# --------------------------------------------------------------------------- #
@dataclass
class CatalogItem:
    item_id: str
    name: str
    generic: str = ""
    manufacturer: str = ""
    category: str = ""


def load_catalog(csv_path: str) -> dict[str, CatalogItem]:
    catalog: dict[str, CatalogItem] = {}
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            iid = (row.get("Item Id") or "").strip()
            if not iid:
                continue
            catalog[iid] = CatalogItem(
                item_id=iid,
                name=(row.get("Item Name") or "").strip(),
                generic=(row.get("Genaric Name") or row.get("Generic Name") or "").strip(),
                manufacturer=(row.get("Manufacturer") or "").strip(),
                category=(row.get("B2B Category") or row.get("Category") or "").strip(),
            )
    return catalog


def list_product_folders(products_root: str) -> list[str]:
    out = []
    for d in os.listdir(products_root):
        if os.path.isdir(os.path.join(products_root, d)) and d.isdigit():
            out.append(d)
    out.sort(key=lambda x: int(x))
    return out


# --------------------------------------------------------------------------- #
# Site clients
# --------------------------------------------------------------------------- #
class DwatsonClient:
    SEARCH_URL = "https://dwatson.pk/catalogsearch/result/"

    def __init__(self, session: requests.Session, timeout: int = 40):
        self.s = session
        self.timeout = timeout

    def search(self, query: str) -> list[tuple[str, str]]:
        r = self.s.get(self.SEARCH_URL, params={"q": query}, timeout=self.timeout, verify=False)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        out = []
        for a in soup.select("a.product-item-link"):
            title = a.get_text(strip=True)
            href = a.get("href")
            if title and href:
                out.append((title, href))
        return out

    def fetch_description(self, url: str) -> str:
        r = self.s.get(url, timeout=self.timeout, verify=False)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        og = soup.find("meta", attrs={"property": "og:description"})
        if og and og.get("content"):
            return normalize_text(og["content"])
        # Fallback: description block in the page body.
        node = soup.select_one("#description, .product.attribute.description .value, .description .value")
        if node:
            return normalize_text(node.get_text("\n", strip=True))
        return ""


class HealthwireClient:
    SEARCH_URL = "https://healthwire.pk/searches"
    BASE = "https://healthwire.pk"
    ATTRS = ["id", "name", "chemical_name", "image", "category", "actual_price",
             "discounted_price", "url", "manufacturer_name", "item_variant_id", "item_id"]
    # Sections to assemble into the description, in display order.
    SECTIONS = ["Description", "Uses", "How To Use", "Dosage", "Side Effects",
                "Precautions & Warnings", "Drug Interactions", "Storage/Disposal"]

    def __init__(self, session: requests.Session, timeout: int = 40):
        self.s = session
        self.timeout = timeout

    def search(self, query: str) -> list[dict]:
        params = [("searches[0][q]", query), ("searches[0][model]", "items")]
        for a in self.ATTRS:
            params.append(("searches[0][attributes][]", a))
        r = self.s.get(self.SEARCH_URL, params=params, timeout=self.timeout, verify=False,
                       headers={"X-Requested-With": "XMLHttpRequest",
                                "Accept": "application/json, text/javascript, */*; q=0.01"})
        r.raise_for_status()
        data = r.json()
        if isinstance(data, list) and data:
            return [h.get("_source", {}) for h in data[0].get("hits", [])]
        return []

    def fetch_description(self, url: str) -> str:
        full = url if url.startswith("http") else self.BASE + url
        r = self.s.get(full, timeout=self.timeout, verify=False)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        parts: list[str] = []
        for h2 in soup.find_all("h2"):
            title = h2.get_text(strip=True)
            if title not in self.SECTIONS:
                continue
            chunk = []
            sib = h2.find_next_sibling()
            guard = 0
            while sib is not None and sib.name not in ("h2",) and guard < 12:
                txt = sib.get_text("\n", strip=True)
                if txt:
                    chunk.append(txt)
                sib = sib.find_next_sibling()
                guard += 1
            body = normalize_text("\n".join(chunk))
            if body and body.lower() not in ("n/a", "na", "-", "not available"):
                parts.append(f"{title}\n\n{body}")
        return normalize_text("\n\n".join(parts))


# --------------------------------------------------------------------------- #
# Result + config
# --------------------------------------------------------------------------- #
@dataclass
class MatchResult:
    folder_id: str
    csv_name: str = ""
    status: str = STATUS_NOT_FOUND
    source: str = ""
    matched_name: str = ""
    score: float = 0.0
    url: str = ""
    description: str = ""
    note: str = ""
    selected: bool = True  # for UI preview / commit

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class ScrapeConfig:
    products_root: str
    csv_path: str
    header: str = "Details"               # xlsx column header to keep
    batch_size: int = 50
    source_order: tuple = (SOURCE_DWATSON, SOURCE_HEALTHWIRE)
    min_score: float = 80.0               # dwatson strict match threshold
    request_delay: float = 0.8            # polite delay between products (s)
    timeout: int = 40
    skip_filled: bool = False             # skip products that already have content
    overwrite: bool = True


# --------------------------------------------------------------------------- #
# Engine
# --------------------------------------------------------------------------- #
class ScraperEngine:
    def __init__(self, config: ScrapeConfig):
        self.cfg = config
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": USER_AGENT})
        self.dwatson = DwatsonClient(self.session, config.timeout)
        self.healthwire = HealthwireClient(self.session, config.timeout)
        self.catalog = load_catalog(config.csv_path)

    # -- per-source resolution ------------------------------------------------ #
    def _resolve_dwatson(self, item: CatalogItem) -> Optional[MatchResult]:
        results = self.dwatson.search(item.name)
        if not results:
            return None
        best = None
        best_score = -1.0
        for title, href in results:
            ok, score, reason = is_strong_match(item.name, title, self.cfg.min_score)
            if ok and score > best_score:
                best, best_score = (title, href), score
        if not best:
            return None
        desc = self.dwatson.fetch_description(best[1])
        if not desc:
            return None
        return MatchResult(
            folder_id=item.item_id, csv_name=item.name, status=STATUS_FOUND,
            source=SOURCE_DWATSON, matched_name=best[0], score=round(best_score, 1),
            url=best[1], description=desc,
        )

    def _resolve_healthwire(self, item: CatalogItem) -> Optional[MatchResult]:
        hits = self.healthwire.search(item.name)
        if not hits:
            return None
        # Authoritative: match by item_id == folder id.
        chosen = next((h for h in hits if str(h.get("item_id")) == item.item_id), None)
        score = 100.0
        if chosen is None:
            # Fall back to best fuzzy hit with strength validation.
            best = None
            best_score = -1.0
            for h in hits:
                ok, sc, _ = is_strong_match(item.name, h.get("name", ""), self.cfg.min_score)
                if ok and sc > best_score:
                    best, best_score = h, sc
            if best is None:
                return None
            chosen, score = best, round(best_score, 1)
        url = chosen.get("url") or ""
        if not url:
            return None
        desc = self.healthwire.fetch_description(url)
        if not desc:
            return None
        return MatchResult(
            folder_id=item.item_id, csv_name=item.name, status=STATUS_FOUND,
            source=SOURCE_HEALTHWIRE, matched_name=chosen.get("name", ""), score=score,
            url=self.healthwire.BASE + url if url.startswith("/") else url, description=desc,
        )

    def process_product(self, folder_id: str) -> MatchResult:
        item = self.catalog.get(folder_id)
        if not item or not item.name:
            return MatchResult(folder_id=folder_id, status=STATUS_NO_CSV,
                               note="No matching row in CSV")
        resolvers = {
            SOURCE_DWATSON: self._resolve_dwatson,
            SOURCE_HEALTHWIRE: self._resolve_healthwire,
        }
        last_error = ""
        for src in self.cfg.source_order:
            fn = resolvers.get(src)
            if not fn:
                continue
            try:
                res = fn(item)
                if res and res.description:
                    return res
            except Exception as exc:  # network / parse error -> try next source
                last_error = f"{src}: {exc}"
        return MatchResult(folder_id=folder_id, csv_name=item.name,
                           status=STATUS_ERROR if last_error else STATUS_NOT_FOUND,
                           note=last_error or "No match on either site")

    # -- batch ---------------------------------------------------------------- #
    def select_folders(self, offset: int = 0, limit: Optional[int] = None) -> list[str]:
        folders = list_product_folders(self.cfg.products_root)
        if self.cfg.skip_filled:
            folders = [f for f in folders if not _xlsx_has_content(self.cfg.products_root, f)]
        limit = limit if limit is not None else self.cfg.batch_size
        return folders[offset: offset + limit]

    def run_batch(
        self,
        folder_ids: list[str],
        on_result: Optional[Callable[[MatchResult], None]] = None,
        should_stop: Optional[Callable[[], bool]] = None,
    ) -> list[MatchResult]:
        out: list[MatchResult] = []
        for fid in folder_ids:
            if should_stop and should_stop():
                break
            res = self.process_product(fid)
            out.append(res)
            if on_result:
                on_result(res)
            if self.cfg.request_delay:
                time.sleep(self.cfg.request_delay)
        return out


# --------------------------------------------------------------------------- #
# xlsx writing
# --------------------------------------------------------------------------- #
def _xlsx_path(products_root: str, folder_id: str) -> str:
    return os.path.join(products_root, folder_id, "descriptions", "product_details.xlsx")


def _xlsx_has_content(products_root: str, folder_id: str) -> bool:
    path = _xlsx_path(products_root, folder_id)
    if not os.path.exists(path):
        return False
    try:
        wb = load_workbook(path, read_only=True)
        ws = wb.active
        val = ws.cell(2, 1).value
        wb.close()
        return bool(val and str(val).strip())
    except Exception:
        return False


def write_description(products_root: str, folder_id: str, description: str,
                      header: str = "Details") -> str:
    """Write the description into A2 (header in A1). Returns the file path."""
    path = _xlsx_path(products_root, folder_id)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path):
        try:
            wb = load_workbook(path)
        except Exception:
            wb = Workbook()
    else:
        wb = Workbook()
    ws = wb.active
    existing_header = ws.cell(1, 1).value
    ws.cell(row=1, column=1, value=existing_header if existing_header else header)
    cell = ws.cell(row=2, column=1, value=description)
    cell.alignment = cell.alignment.copy(wrap_text=True, vertical="top")
    wb.save(path)
    return path


def read_existing_description(products_root: str, folder_id: str) -> str:
    path = _xlsx_path(products_root, folder_id)
    if not os.path.exists(path):
        return ""
    try:
        wb = load_workbook(path, read_only=True)
        ws = wb.active
        val = ws.cell(2, 1).value
        wb.close()
        return str(val) if val else ""
    except Exception:
        return ""


# --------------------------------------------------------------------------- #
# CLI entry point
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    import argparse

    here = os.path.dirname(os.path.abspath(__file__))
    default_root = os.path.normpath(os.path.join(here, "..", "Products"))
    default_csv = os.path.normpath(os.path.join(here, "..", "Item List - 260327 (Items with ID).csv"))

    ap = argparse.ArgumentParser(description="Pulse Rx description scraper (CLI)")
    ap.add_argument("--root", default=default_root)
    ap.add_argument("--csv", default=default_csv)
    ap.add_argument("--offset", type=int, default=0)
    ap.add_argument("--limit", type=int, default=10)
    ap.add_argument("--write", action="store_true", help="write results to xlsx")
    ap.add_argument("--order", default="dwatson,healthwire")
    ap.add_argument("--min-score", type=float, default=80.0)
    ap.add_argument("--delay", type=float, default=0.8)
    args = ap.parse_args()

    cfg = ScrapeConfig(
        products_root=args.root, csv_path=args.csv,
        source_order=tuple(args.order.split(",")),
        min_score=args.min_score, request_delay=args.delay,
    )
    engine = ScraperEngine(cfg)
    folders = engine.select_folders(args.offset, args.limit)
    print(f"Processing {len(folders)} products (offset={args.offset})")

    def cb(res: MatchResult):
        print(f"[{res.status:11}] {res.folder_id} | {res.source or '-':10} "
              f"score={res.score:5} | {res.matched_name[:45]:45} | {res.csv_name[:40]}")
        if args.write and res.status == STATUS_FOUND and res.description:
            write_description(cfg.products_root, res.folder_id, res.description, cfg.header)

    results = engine.run_batch(folders, on_result=cb)
    found = sum(1 for r in results if r.status == STATUS_FOUND)
    print(f"\nDone: {found}/{len(results)} found.")
