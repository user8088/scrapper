"""
Pulse Rx Description Scraper - desktop UI.

Configure everything, run in batches, verify each matched description, then
commit the selected rows into each product's product_details.xlsx -- or let it
write directly while scraping.

Run:  python app.py
"""

from __future__ import annotations

import os
import sys
import json
import queue
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import scraper_core as sc
import updater
from version import __version__


def _app_dir() -> str:
    """Directory the app 'lives' in: the exe's folder when frozen, else this file's."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


HERE = _app_dir()
CONFIG_FILE = os.path.join(HERE, "config.json")
DEFAULT_ROOT = os.path.normpath(os.path.join(HERE, "..", "Products"))
DEFAULT_CSV = os.path.normpath(os.path.join(HERE, "..", "Item List - 260327 (Items with ID).csv"))

CHECK_ON = "\u2611"   # ballot box with check
CHECK_OFF = "\u2610"  # ballot box

STATUS_COLORS = {
    sc.STATUS_FOUND: "#0a7d28",
    sc.STATUS_NOT_FOUND: "#b06a00",
    sc.STATUS_NO_CSV: "#888888",
    sc.STATUS_ERROR: "#c0152f",
}


class ScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Pulse Rx - Description Scraper  (v{__version__})")
        self.geometry("1180x760")
        self.minsize(960, 640)

        self.results: dict[str, sc.MatchResult] = {}
        self.event_queue: "queue.Queue" = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop_flag = threading.Event()
        self.engine: sc.ScraperEngine | None = None
        self.total_in_run = 0
        self.done_in_run = 0

        self._build_menu()
        self._build_style()
        self._build_widgets()
        self._load_config()
        self.after(120, self._poll_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Silent check on launch; only the packaged exe can self-apply.
        self.after(1200, lambda: self._start_update_check(silent=True))

    # -------------------------------------------------------------------- menu
    def _build_menu(self):
        menubar = tk.Menu(self)
        helpm = tk.Menu(menubar, tearoff=0)
        helpm.add_command(label="Check for updates…",
                          command=lambda: self._start_update_check(silent=False))
        helpm.add_separator()
        helpm.add_command(label=f"About  (v{__version__})", command=self._about)
        menubar.add_cascade(label="Help", menu=helpm)
        self.config(menu=menubar)

    def _about(self):
        messagebox.showinfo(
            "About",
            f"Pulse Rx - Description Scraper\nVersion {__version__}\n\n"
            f"Updates: {updater.RELEASES_PAGE}",
        )

    # ------------------------------------------------------------------ style
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"))

    # ----------------------------------------------------------------- widgets
    def _build_widgets(self):
        self._build_config_panel()
        self._build_control_panel()
        self._build_results_table()
        self._build_status_bar()

    def _build_config_panel(self):
        f = ttk.LabelFrame(self, text="Configuration", padding=8)
        f.pack(fill="x", padx=10, pady=(10, 6))

        # Paths
        self.var_root = tk.StringVar(value=DEFAULT_ROOT)
        self.var_csv = tk.StringVar(value=DEFAULT_CSV)
        self._path_row(f, 0, "Products folder:", self.var_root, self._pick_root)
        self._path_row(f, 1, "Item List CSV:", self.var_csv, self._pick_csv)

        # numeric / option row
        opts = ttk.Frame(f)
        opts.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.var_batch = tk.IntVar(value=50)
        self.var_offset = tk.IntVar(value=0)
        self.var_minscore = tk.DoubleVar(value=80.0)
        self.var_delay = tk.DoubleVar(value=0.8)
        self.var_threads = tk.IntVar(value=4)
        self.var_header = tk.StringVar(value="Details")
        self.var_order = tk.StringVar(value="dwatson, healthwire")
        self.var_skip = tk.BooleanVar(value=False)

        self._spin(opts, "Batch size:", self.var_batch, 1, 5000, 0)
        self._spin(opts, "Start offset:", self.var_offset, 0, 100000, 1)
        self._spin(opts, "Min match score:", self.var_minscore, 0, 100, 2, inc=1.0, width=5)
        self._spin(opts, "Delay (s):", self.var_delay, 0, 30, 3, inc=0.1, width=5)
        self._spin(opts, "Threads:", self.var_threads, 1, 16, 4, width=5)

        opts2 = ttk.Frame(f)
        opts2.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        ttk.Label(opts2, text="Source priority:").grid(row=0, column=0, padx=(0, 4))
        cb = ttk.Combobox(opts2, textvariable=self.var_order, width=24, state="readonly",
                          values=["dwatson, healthwire", "healthwire, dwatson",
                                  "dwatson", "healthwire"])
        cb.grid(row=0, column=1, padx=(0, 16))

        ttk.Label(opts2, text="xlsx header:").grid(row=0, column=2, padx=(0, 4))
        ttk.Entry(opts2, textvariable=self.var_header, width=14).grid(row=0, column=3, padx=(0, 16))

        ttk.Checkbutton(opts2, text="Skip products already filled",
                        variable=self.var_skip).grid(row=0, column=4, padx=(0, 16))

        ttk.Button(opts2, text="Save settings", command=self._save_config).grid(row=0, column=5)

    def _path_row(self, parent, row, label, var, cmd):
        ttk.Label(parent, text=label, width=15).grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(parent, textvariable=var, width=110).grid(row=row, column=1, sticky="we", pady=2)
        ttk.Button(parent, text="Browse", command=cmd).grid(row=row, column=2, padx=6)
        parent.columnconfigure(1, weight=1)

    def _spin(self, parent, label, var, frm, to, col, inc=1, width=7):
        wrap = ttk.Frame(parent)
        wrap.grid(row=0, column=col, padx=(0, 16), sticky="w")
        ttk.Label(wrap, text=label).pack(side="left", padx=(0, 4))
        ttk.Spinbox(wrap, from_=frm, to=to, textvariable=var, width=width,
                    increment=inc).pack(side="left")

    def _build_control_panel(self):
        f = ttk.Frame(self)
        f.pack(fill="x", padx=10, pady=4)

        self.var_mode = tk.StringVar(value="preview")
        ttk.Label(f, text="Mode:").pack(side="left")
        ttk.Radiobutton(f, text="Preview then commit", value="preview",
                        variable=self.var_mode).pack(side="left", padx=4)
        ttk.Radiobutton(f, text="Write directly while scraping", value="direct",
                        variable=self.var_mode).pack(side="left", padx=4)

        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=10)

        self.btn_run = ttk.Button(f, text="\u25b6  Run batch", command=self._run_batch)
        self.btn_run.pack(side="left", padx=3)
        self.btn_stop = ttk.Button(f, text="\u25a0  Stop", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=3)

        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=10)

        self.btn_commit = ttk.Button(f, text="\U0001f4be  Write selected to xlsx",
                                     command=self._commit_selected, state="disabled")
        self.btn_commit.pack(side="left", padx=3)
        ttk.Button(f, text="Select all", command=lambda: self._set_all(True)).pack(side="left", padx=3)
        ttk.Button(f, text="Deselect all", command=lambda: self._set_all(False)).pack(side="left", padx=3)
        ttk.Button(f, text="Clear results", command=self._clear_results).pack(side="left", padx=3)

    def _build_results_table(self):
        f = ttk.LabelFrame(self, text="Results  (click \u2611 to toggle, double-click a row to view full description)",
                           padding=4)
        f.pack(fill="both", expand=True, padx=10, pady=6)

        cols = ("sel", "id", "csv_name", "source", "matched", "score", "status", "preview")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", selectmode="browse")
        headings = {
            "sel": ("\u2611", 36), "id": ("Item Id", 80), "csv_name": ("CSV Name", 230),
            "source": ("Source", 80), "matched": ("Matched Name", 230), "score": ("Score", 55),
            "status": ("Status", 90), "preview": ("Description preview", 360),
        }
        for c, (txt, w) in headings.items():
            self.tree.heading(c, text=txt)
            anchor = "center" if c in ("sel", "score", "source", "status") else "w"
            self.tree.column(c, width=w, anchor=anchor, stretch=(c in ("csv_name", "matched", "preview")))

        for st, color in STATUS_COLORS.items():
            self.tree.tag_configure(st, foreground=color)

        vsb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double)

    def _build_status_bar(self):
        f = ttk.Frame(self)
        f.pack(fill="x", padx=10, pady=(0, 4))
        self.progress = ttk.Progressbar(f, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.var_status = tk.StringVar(value="Ready.")
        ttk.Label(f, textvariable=self.var_status, width=46, anchor="e").pack(side="right")

        logf = ttk.LabelFrame(self, text="Log", padding=4)
        logf.pack(fill="x", padx=10, pady=(0, 10))
        self.log = tk.Text(logf, height=6, wrap="word", font=("Consolas", 9))
        self.log.pack(side="left", fill="both", expand=True)
        lsb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=lsb.set, state="disabled")
        lsb.pack(side="right", fill="y")

    # -------------------------------------------------------------- callbacks
    def _pick_root(self):
        d = filedialog.askdirectory(title="Select Products folder", initialdir=self.var_root.get() or HERE)
        if d:
            self.var_root.set(d)

    def _pick_csv(self):
        d = filedialog.askopenfilename(title="Select Item List CSV", filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                                       initialdir=os.path.dirname(self.var_csv.get() or HERE))
        if d:
            self.var_csv.set(d)

    def _logline(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------- config
    def _gather_config(self) -> sc.ScrapeConfig:
        order = tuple(s.strip() for s in self.var_order.get().split(",") if s.strip())
        return sc.ScrapeConfig(
            products_root=self.var_root.get().strip(),
            csv_path=self.var_csv.get().strip(),
            header=self.var_header.get().strip() or "Details",
            batch_size=int(self.var_batch.get()),
            source_order=order,
            min_score=float(self.var_minscore.get()),
            request_delay=float(self.var_delay.get()),
            skip_filled=bool(self.var_skip.get()),
            max_workers=int(self.var_threads.get()),
        )

    def _save_config(self):
        data = {
            "root": self.var_root.get(), "csv": self.var_csv.get(),
            "batch": int(self.var_batch.get()), "offset": int(self.var_offset.get()),
            "minscore": float(self.var_minscore.get()), "delay": float(self.var_delay.get()),
            "threads": int(self.var_threads.get()),
            "header": self.var_header.get(), "order": self.var_order.get(),
            "skip": bool(self.var_skip.get()), "mode": self.var_mode.get(),
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
            self._logline("Settings saved.")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    def _load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, encoding="utf-8") as fh:
                d = json.load(fh)
            self.var_root.set(d.get("root", DEFAULT_ROOT))
            self.var_csv.set(d.get("csv", DEFAULT_CSV))
            self.var_batch.set(d.get("batch", 50))
            self.var_offset.set(d.get("offset", 0))
            self.var_minscore.set(d.get("minscore", 80.0))
            self.var_delay.set(d.get("delay", 0.8))
            self.var_threads.set(d.get("threads", 4))
            self.var_header.set(d.get("header", "Details"))
            self.var_order.set(d.get("order", "dwatson, healthwire"))
            self.var_skip.set(d.get("skip", False))
            self.var_mode.set(d.get("mode", "preview"))
        except Exception:
            pass

    # -------------------------------------------------------------------- run
    def _run_batch(self):
        if self.worker and self.worker.is_alive():
            return
        cfg = self._gather_config()
        if not os.path.isdir(cfg.products_root):
            messagebox.showerror("Invalid path", "Products folder does not exist.")
            return
        if not os.path.isfile(cfg.csv_path):
            messagebox.showerror("Invalid path", "Item List CSV does not exist.")
            return

        try:
            self.engine = sc.ScraperEngine(cfg)
        except Exception as exc:
            messagebox.showerror("Engine error", str(exc))
            return

        offset = int(self.var_offset.get())
        folders = self.engine.select_folders(offset, cfg.batch_size)
        if not folders:
            messagebox.showinfo("Nothing to do", "No products in this range (or all filled).")
            return

        self.stop_flag.clear()
        self.total_in_run = len(folders)
        self.done_in_run = 0
        self.progress.configure(maximum=self.total_in_run, value=0)
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_commit.configure(state="disabled")
        direct = self.var_mode.get() == "direct"
        self._logline(f"Running {len(folders)} products (offset {offset}) | order={cfg.source_order} "
                      f"| threads={cfg.max_workers} | mode={'DIRECT WRITE' if direct else 'preview'}")

        self.worker = threading.Thread(target=self._worker, args=(folders, cfg, direct), daemon=True)
        self.worker.start()

    def _worker(self, folders, cfg, direct):
        def on_result(res: sc.MatchResult):
            if direct and res.status == sc.STATUS_FOUND and res.description:
                try:
                    sc.write_description(cfg.products_root, res.folder_id, res.description, cfg.header)
                    res.note = "written"
                except Exception as exc:
                    res.note = f"write failed: {exc}"
            self.event_queue.put(("result", res))

        try:
            self.engine.run_batch(folders, on_result=on_result, should_stop=self.stop_flag.is_set)
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))
        finally:
            self.event_queue.put(("done", None))

    def _stop(self):
        self.stop_flag.set()
        self._logline("Stopping after current item...")

    # ---------------------------------------------------------- queue polling
    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.event_queue.get_nowait()
                if kind == "result":
                    self._add_result(payload)
                elif kind == "error":
                    self._logline("ERROR: " + str(payload))
                elif kind == "done":
                    self._on_run_done()
                elif kind == "update":
                    self._on_update_check_result(*payload)
                elif kind == "upd_progress":
                    self._on_update_progress(*payload)
                elif kind == "upd_ready":
                    self._on_update_downloaded(payload)
                elif kind == "upd_failed":
                    self._on_update_failed(payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _add_result(self, res: sc.MatchResult):
        self.done_in_run += 1
        self.progress.configure(value=self.done_in_run)
        self.results[res.folder_id] = res
        res.selected = res.status == sc.STATUS_FOUND
        preview = (res.description[:140] + "...") if len(res.description) > 140 else res.description
        preview = preview.replace("\n", " ")
        self.tree.insert("", "end", iid=res.folder_id, tags=(res.status,), values=(
            CHECK_ON if res.selected else CHECK_OFF,
            res.folder_id, res.csv_name, res.source or "-",
            res.matched_name or "-", f"{res.score:g}" if res.score else "-",
            res.status, preview or (res.note or ""),
        ))
        self.var_status.set(f"{self.done_in_run}/{self.total_in_run} processed")
        flag = " [written]" if res.note == "written" else ""
        self._logline(f"[{res.status:11}] {res.folder_id} {res.source or '-':10} "
                      f"score={res.score:<5g} {res.matched_name[:40]}{flag}")

    def _on_run_done(self):
        self.btn_run.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        found = sum(1 for r in self.results.values() if r.status == sc.STATUS_FOUND)
        if self.var_mode.get() == "preview" and found:
            self.btn_commit.configure(state="normal")
        # auto-advance offset for the next batch
        self.var_offset.set(int(self.var_offset.get()) + self.total_in_run)
        self.var_status.set(f"Done. {self.done_in_run} processed, next offset {self.var_offset.get()}.")
        self._logline(f"Batch finished. {found} found. Next offset = {self.var_offset.get()}.")

    # --------------------------------------------------------- tree interaction
    def _on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            self._toggle(iid)

    def _toggle(self, iid):
        res = self.results.get(iid)
        if not res or res.status != sc.STATUS_FOUND:
            return
        res.selected = not res.selected
        self.tree.set(iid, "sel", CHECK_ON if res.selected else CHECK_OFF)

    def _set_all(self, value: bool):
        for iid, res in self.results.items():
            if res.status == sc.STATUS_FOUND:
                res.selected = value
                self.tree.set(iid, "sel", CHECK_ON if value else CHECK_OFF)

    def _on_tree_double(self, event):
        iid = self.tree.identify_row(event.y)
        res = self.results.get(iid)
        if res:
            self._show_detail(res)

    def _show_detail(self, res: sc.MatchResult):
        win = tk.Toplevel(self)
        win.title(f"{res.folder_id} - {res.csv_name}")
        win.geometry("720x620")
        head = ttk.Frame(win, padding=8)
        head.pack(fill="x")
        ttk.Label(head, text=res.csv_name, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        meta = f"Source: {res.source or '-'}    Score: {res.score:g}    Matched: {res.matched_name or '-'}"
        ttk.Label(head, text=meta).pack(anchor="w")
        if res.url:
            link = ttk.Label(head, text=res.url, foreground="#1155cc", cursor="hand2")
            link.pack(anchor="w")
            link.bind("<Button-1>", lambda e: webbrowser.open(res.url))

        body = tk.Text(win, wrap="word", font=("Segoe UI", 10), padx=10, pady=8)
        body.pack(fill="both", expand=True)
        body.insert("1.0", res.description or res.note or "(no description)")
        body.configure(state="disabled")

        btns = ttk.Frame(win, padding=8)
        btns.pack(fill="x")
        ttk.Button(btns, text="Write THIS to xlsx",
                   command=lambda: self._commit_one(res, win)).pack(side="right")
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="right", padx=6)

    # ------------------------------------------------------------------ commit
    def _commit_one(self, res: sc.MatchResult, win=None):
        if res.status != sc.STATUS_FOUND or not res.description:
            return
        cfg = self._gather_config()
        try:
            sc.write_description(cfg.products_root, res.folder_id, res.description, cfg.header)
            res.note = "written"
            self.tree.set(res.folder_id, "status", "written")
            self._logline(f"Wrote {res.folder_id}.")
            if win:
                win.destroy()
        except Exception as exc:
            messagebox.showerror("Write failed", str(exc))

    def _commit_selected(self):
        cfg = self._gather_config()
        chosen = [r for r in self.results.values()
                  if r.selected and r.status == sc.STATUS_FOUND and r.description]
        if not chosen:
            messagebox.showinfo("Nothing selected", "No found+selected rows to write.")
            return
        if not messagebox.askyesno("Confirm", f"Write {len(chosen)} descriptions into their xlsx files?"):
            return
        ok = 0
        for r in chosen:
            try:
                sc.write_description(cfg.products_root, r.folder_id, r.description, cfg.header)
                r.note = "written"
                self.tree.set(r.folder_id, "status", "written")
                ok += 1
            except Exception as exc:
                self._logline(f"Write failed {r.folder_id}: {exc}")
        self._logline(f"Committed {ok}/{len(chosen)} descriptions.")
        messagebox.showinfo("Done", f"Wrote {ok} of {len(chosen)} descriptions.")

    def _clear_results(self):
        self.results.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.progress.configure(value=0)
        self.btn_commit.configure(state="disabled")
        self.var_status.set("Cleared.")

    def _on_close(self):
        self.stop_flag.set()
        self.destroy()

    # ------------------------------------------------------------- auto-update
    def _start_update_check(self, silent: bool):
        """Check GitHub for a newer release in a background thread."""
        def work():
            info = updater.check_for_update()
            self.event_queue.put(("update", (info, silent)))
        threading.Thread(target=work, daemon=True).start()
        if not silent:
            self._logline("Checking for updates…")

    def _on_update_check_result(self, info, silent: bool):
        if info is None:
            if not silent:
                messagebox.showinfo("Up to date",
                                    f"You're running the latest version (v{__version__}).")
            return

        # A newer release exists. Running from source can't self-replace.
        if not updater.is_frozen():
            if messagebox.askyesno(
                "Update available",
                f"Version {info.version} is available (you have {__version__}).\n\n"
                "This is a source checkout, so it can't update itself.\n"
                "Open the releases page to download?",
            ):
                webbrowser.open(updater.RELEASES_PAGE)
            return

        if not messagebox.askyesno(
            "Update available",
            f"Version {info.version} is available (you have {__version__}).\n\n"
            f"{(info.notes[:500] + '…') if len(info.notes) > 500 else info.notes}\n\n"
            "Download and install it now? The app will restart.",
        ):
            return
        self._begin_download(info)

    def _begin_download(self, info):
        # Lightweight modal progress dialog.
        self._upd_win = tk.Toplevel(self)
        self._upd_win.title("Updating")
        self._upd_win.geometry("360x110")
        self._upd_win.resizable(False, False)
        self._upd_win.transient(self)
        self._upd_win.grab_set()
        ttk.Label(self._upd_win, text=f"Downloading version {info.version}…",
                  padding=10).pack(anchor="w")
        self._upd_bar = ttk.Progressbar(self._upd_win, mode="determinate", maximum=100)
        self._upd_bar.pack(fill="x", padx=10, pady=6)
        self._upd_pct = ttk.Label(self._upd_win, text="0%", padding=(10, 0))
        self._upd_pct.pack(anchor="w")

        def work():
            try:
                def prog(read, total):
                    self.event_queue.put(("upd_progress", (read, total)))
                path = updater.download_update(info, on_progress=prog)
                self.event_queue.put(("upd_ready", path))
            except Exception as exc:  # network / disk errors
                self.event_queue.put(("upd_failed", str(exc)))
        threading.Thread(target=work, daemon=True).start()

    def _on_update_progress(self, read, total):
        if not getattr(self, "_upd_bar", None):
            return
        if total > 0:
            pct = int(read * 100 / total)
            self._upd_bar.configure(value=pct)
            self._upd_pct.configure(text=f"{pct}%  ({read // 1024} / {total // 1024} KB)")
        else:
            self._upd_bar.configure(mode="indeterminate")
            self._upd_bar.start(20)

    def _on_update_downloaded(self, path):
        if getattr(self, "_upd_win", None):
            self._upd_win.destroy()
            self._upd_win = None
        self._logline("Update downloaded. Restarting to apply…")
        try:
            updater.apply_update_and_restart(path)
        except Exception as exc:
            messagebox.showerror("Update failed", str(exc))
            return
        self.stop_flag.set()
        self.destroy()

    def _on_update_failed(self, err):
        if getattr(self, "_upd_win", None):
            self._upd_win.destroy()
            self._upd_win = None
        messagebox.showerror("Update failed", f"Could not download the update:\n{err}")


if __name__ == "__main__":
    ScraperApp().mainloop()
