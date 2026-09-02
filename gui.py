#!/usr/bin/env python3
"""Tkinter front end for teamsexport -- same three steps, one window.

Everything here is glue: the work lives in teamsexport.py. This just runs it on a
worker thread and pipes its stdout/stderr into a log pane.
"""
from __future__ import annotations

import argparse
import io
import os
import pathlib
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk

import teamsexport as tx

HOME = pathlib.Path.home() / "TeamsExport"


class Pipe(io.TextIOBase):
    """stdout/stderr -> queue. Needed regardless: a --windowed exe has no console."""

    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, s: str) -> int:
        if s:
            self.q.put(s)
        return len(s)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.q: queue.Queue[str] = queue.Queue()
        sys.stdout = sys.stderr = Pipe(self.q)
        self.stop = threading.Event()
        self.busy = False
        self.lock = threading.Lock()  # a manual run and a watch tick share messages.jsonl

        root.title("teamsexport")
        root.geometry("820x520")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)

        top = ttk.Frame(root, padding=8)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text="Export folder").grid(row=0, column=0, padx=(0, 6))
        self.wd = tk.StringVar(value=str(HOME))
        ttk.Entry(top, textvariable=self.wd).grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Browse", command=self.browse).grid(row=0, column=2, padx=6)

        bar = ttk.Frame(root, padding=(8, 0))
        bar.grid(row=1, column=0, sticky="ew")
        self.btn_export = ttk.Button(bar, text="Export now", command=self.export)
        self.btn_zip = ttk.Button(bar, text="Parse a snapshot zip...", command=self.parse_zip)
        self.btn_open = ttk.Button(bar, text="Open export", command=self.open_html)
        self.btn_watch = ttk.Button(bar, text="Start watch", command=self.toggle_watch)
        for b in (self.btn_export, self.btn_zip, self.btn_open, self.btn_watch):
            b.pack(side="left", padx=(0, 6), pady=6)
        ttk.Label(bar, text="every").pack(side="left", padx=(12, 4))
        self.interval = tk.StringVar(value="15")
        ttk.Spinbox(bar, from_=1, to=1440, width=5, textvariable=self.interval).pack(side="left")
        ttk.Label(bar, text="min").pack(side="left", padx=4)

        self.log = tk.Text(root, wrap="word", height=20, bg="#14141a", fg="#e8e8ee",
                           insertbackground="#e8e8ee", font=("Consolas", 9))
        self.log.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        sb = ttk.Scrollbar(root, command=self.log.yview)
        sb.grid(row=2, column=1, sticky="ns", pady=(0, 8))
        self.log["yscrollcommand"] = sb.set

        print("Export now = snapshot the local Teams cache, merge it into your history, "
              "rebuild the HTML.\nRepeat runs only ever add messages -- nothing is lost when "
              "Teams evicts its cache.\n")
        self.drain()

    # ------------------------------------------------------------- plumbing
    def paths(self) -> tuple[str, str, str]:
        wd = pathlib.Path(self.wd.get())
        wd.mkdir(parents=True, exist_ok=True)
        return str(wd / "snapshots"), str(wd / "messages.jsonl"), str(wd / "teams-export.html")

    def drain(self) -> None:
        while True:
            try:
                self.log.insert("end", self.q.get_nowait())
            except queue.Empty:
                break
            else:
                self.log.see("end")
        self.root.after(100, self.drain)

    def run(self, fn) -> None:
        if self.busy:
            return
        self.busy = True
        for b in (self.btn_export, self.btn_zip):
            b.state(["disabled"])

        def wrapper():
            try:
                with self.lock:
                    fn()
            except SystemExit as e:
                print(f"\n{e}\n")
            except Exception as e:
                print(f"\nerror: {e!r}\n")
            finally:
                self.busy = False
                self.root.after(0, lambda: [b.state(["!disabled"])
                                            for b in (self.btn_export, self.btn_zip)])

        threading.Thread(target=wrapper, daemon=True).start()

    # ------------------------------------------------------------- actions
    def browse(self) -> None:
        d = filedialog.askdirectory(initialdir=self.wd.get())
        if d:
            self.wd.set(d)

    def export(self) -> None:
        self.run(lambda: tx.export_once(*self.paths()))

    def parse_zip(self) -> None:
        z = filedialog.askopenfilename(filetypes=[("Snapshot", "*.zip")])
        if not z:
            return
        _, store, out = self.paths()

        def job():
            tx.cmd_extract(argparse.Namespace(source=z, out=store))
            tx.cmd_render(argparse.Namespace(source=store, out=out))

        self.run(job)

    def open_html(self) -> None:
        html = self.paths()[2]
        if os.path.exists(html):
            os.startfile(html)  # noqa: S606 -- win32 only, same as double-clicking it
        else:
            print(f"{html} does not exist yet -- run Export now first.\n")

    def toggle_watch(self) -> None:
        if self.btn_watch["text"] == "Stop watch":
            self.stop.set()
            self.btn_watch["text"] = "Start watch"
            return
        self.stop.clear()
        self.btn_watch["text"] = "Stop watch"
        secs = max(60, int(self.interval.get() or 15) * 60)
        args = self.paths()

        def loop():
            while not self.stop.is_set():
                try:
                    with self.lock:
                        tx.export_once(*args)
                except SystemExit as e:
                    print(f"skip: {e}")
                except Exception as e:
                    print(f"error: {e!r}")
                print(f"-- next run in {secs // 60} min\n")
                self.stop.wait(secs)
            print("-- watch stopped\n")

        threading.Thread(target=loop, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
