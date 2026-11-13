"""
app.py - PPT to Word Converter (desktop UI, tkinter).

Run with:  python app.py
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from converter import (
    SEPARATOR_NONE,
    SEPARATOR_PAGE,
    SEPARATOR_RULE,
    ConvertOptions,
    convert_file,
    unique_path,
)

APP_NAME = "PPT to Word Converter"
APP_VERSION = "1.0.0"

# Palette
NAVY = "#1F3A5F"
BLUE = "#2F6FDE"
BLUE_DARK = "#2459B5"
BG = "#F4F6FA"
CARD = "#FFFFFF"
TEXT = "#1E2430"
MUTED = "#6B7385"
BORDER = "#DCE1EA"
GREEN = "#1E8E4E"
RED = "#C62828"


def resource_path(name: str) -> str:
    """Find bundled files both in development and inside a PyInstaller exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def open_path(path: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        messagebox.showerror(APP_NAME, f"Could not open:\n{path}\n\n{exc}")


def reveal_in_folder(path: Path):
    try:
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", str(path)])
        else:
            open_path(path.parent)
    except Exception:
        open_path(path.parent)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("820x700")
        self.minsize(640, 420)
        self.configure(bg=BG)
        try:
            self.iconbitmap(resource_path("app.ico"))
        except Exception:
            pass

        self.files: list[Path] = []
        self.outputs: list[Path] = []
        self.queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None

        # option variables
        self.var_numbers = tk.BooleanVar(value=True)
        self.var_images = tk.BooleanVar(value=True)
        self.var_tables = tk.BooleanVar(value=True)
        self.var_charts = tk.BooleanVar(value=True)
        self.var_notes = tk.BooleanVar(value=False)
        self.var_hidden = tk.BooleanVar(value=False)
        self.var_cover = tk.BooleanVar(value=True)
        self.var_pagenums = tk.BooleanVar(value=True)
        self.var_landscape = tk.BooleanVar(value=False)
        self.var_separator = tk.StringVar(value=SEPARATOR_RULE)
        self.var_lyrics = tk.BooleanVar(value=False)
        self.var_lyrics_size = tk.StringVar(value="48")
        self.var_lyrics_commas = tk.BooleanVar(value=True)
        self.var_same_folder = tk.BooleanVar(value=True)
        self.var_out_dir = tk.StringVar(value="")
        self.var_status = tk.StringVar(value="Add a PowerPoint file to get started.")

        self._build_styles()
        self._build_ui()
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------ UI
    def _build_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        base = ("Segoe UI", 10)
        style.configure(".", font=base, background=BG, foreground=TEXT)
        style.configure("Card.TFrame", background=CARD, relief="flat")
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Card.TLabel", background=CARD)
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Section.TLabel", background=CARD, foreground=NAVY, font=("Segoe UI Semibold", 11))
        style.configure("Status.TLabel", background=BG, foreground=MUTED)
        style.configure("TCheckbutton", background=CARD, foreground=TEXT)
        style.map("TCheckbutton", background=[("active", CARD)])
        style.configure("TRadiobutton", background=CARD, foreground=TEXT)
        style.map("TRadiobutton", background=[("active", CARD)])
        style.configure("TEntry", fieldbackground="#FFFFFF", bordercolor=BORDER, padding=4)

        style.configure("TButton", padding=(12, 6), background="#E8ECF4", foreground=TEXT,
                        bordercolor=BORDER, focusthickness=0, focuscolor="none")
        style.map("TButton", background=[("active", "#DCE3F0"), ("disabled", "#F0F2F6")],
                  foreground=[("disabled", "#A0A6B4")])
        style.configure("Accent.TButton", padding=(18, 10), background=BLUE, foreground="#FFFFFF",
                        font=("Segoe UI Semibold", 11), bordercolor=BLUE)
        style.map("Accent.TButton", background=[("active", BLUE_DARK), ("disabled", "#A9BCE6")],
                  foreground=[("disabled", "#FFFFFF")])
        style.configure("Success.TButton", padding=(12, 6), background="#E3F4EA", foreground=GREEN,
                        bordercolor="#BFE3CD")
        style.map("Success.TButton", background=[("active", "#D2ECDC"), ("disabled", "#F0F2F6")])

        style.configure("Horizontal.TProgressbar", troughcolor="#E3E8F1", background=BLUE,
                        bordercolor="#E3E8F1", lightcolor=BLUE, darkcolor=BLUE, thickness=10)

    def _card(self, parent, title: str, subtitle: str = "") -> ttk.Frame:
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill="x", padx=24, pady=(0, 12))
        card = ttk.Frame(outer, style="Card.TFrame", padding=(16, 12))
        card.pack(fill="both", expand=True, padx=1, pady=1)
        ttk.Label(card, text=title, style="Section.TLabel").pack(anchor="w")
        if subtitle:
            ttk.Label(card, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(0, 6))
        return card

    def _build_ui(self):
        # Header band
        header = tk.Frame(self, bg=NAVY, height=78)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="PPT  →  Word", bg=NAVY, fg="white",
                 font=("Segoe UI Semibold", 20)).pack(anchor="w", padx=24, pady=(12, 0))
        tk.Label(header, text="Turn PowerPoint presentations into clean, editable Word documents",
                 bg=NAVY, fg="#B8C7E0", font=("Segoe UI", 10)).pack(anchor="w", padx=24)

        # Everything below the header lives in a scrollable area, so shrinking
        # the window vertically shows a scrollbar instead of hiding buttons.
        scroll_outer = tk.Frame(self, bg=BG)
        scroll_outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(scroll_outer, bg=BG, highlightthickness=0, bd=0)
        vscroll = ttk.Scrollbar(scroll_outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vscroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")
        body = tk.Frame(self.canvas, bg=BG)
        self._body_window = self.canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        body.bind("<Configure>", _on_body_configure)

        def _on_canvas_configure(event):
            self.canvas.itemconfigure(self._body_window, width=event.width)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)        # Windows
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))  # Linux
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))

        tk.Frame(body, height=16, bg=BG).pack()

        # 1. Files
        card = self._card(body, "1.  Presentations", "Select one or more .pptx files")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x")
        list_wrap = tk.Frame(row, bg=BORDER)
        list_wrap.pack(side="left", fill="both", expand=True)
        self.listbox = tk.Listbox(list_wrap, height=4, selectmode="extended", activestyle="none",
                                  font=("Segoe UI", 10), bd=0, highlightthickness=0,
                                  selectbackground="#D6E4FF", selectforeground=TEXT)
        self.listbox.pack(fill="both", expand=True, padx=1, pady=1)
        btns = ttk.Frame(row, style="Card.TFrame")
        btns.pack(side="left", padx=(12, 0), anchor="n")
        self.btn_add = ttk.Button(btns, text="Add files…", command=self.add_files)
        self.btn_add.pack(fill="x")
        self.btn_remove = ttk.Button(btns, text="Remove", command=self.remove_selected)
        self.btn_remove.pack(fill="x", pady=4)
        self.btn_clear = ttk.Button(btns, text="Clear", command=self.clear_files)
        self.btn_clear.pack(fill="x")

        # 2. Output
        card = self._card(body, "2.  Save location")
        ttk.Checkbutton(card, text="Save next to each presentation", variable=self.var_same_folder,
                        command=self._toggle_out).pack(anchor="w")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", pady=(6, 0))
        self.entry_out = ttk.Entry(row, textvariable=self.var_out_dir, state="disabled")
        self.entry_out.pack(side="left", fill="x", expand=True)
        self.btn_out = ttk.Button(row, text="Browse…", command=self.pick_out_dir, state="disabled")
        self.btn_out.pack(side="left", padx=(8, 0))

        # 3. Song lyrics mode
        card = self._card(body, "3.  Song lyrics mode",
                          "For lyric slides that will be read from a screen: just the words, big and clear")
        lrow = ttk.Frame(card, style="Card.TFrame")
        lrow.pack(fill="x")
        ttk.Checkbutton(lrow, text="Lyrics mode (no 'Slide 1' headings, no bullets)",
                        variable=self.var_lyrics, command=self._toggle_lyrics).pack(side="left")
        lrow2 = ttk.Frame(card, style="Card.TFrame")
        lrow2.pack(fill="x", pady=(6, 0))
        ttk.Label(lrow2, text="Font size:", style="Card.TLabel").pack(side="left")
        self.spin_size = ttk.Spinbox(lrow2, from_=12, to=96, width=5, textvariable=self.var_lyrics_size,
                                     state="disabled")
        self.spin_size.pack(side="left", padx=(6, 18))
        self.chk_commas = ttk.Checkbutton(lrow2, text="New line after every comma",
                                          variable=self.var_lyrics_commas, state="disabled")
        self.chk_commas.pack(side="left")

        # 4. Options
        card = self._card(body, "4.  What to include")
        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill="x")
        opts = [
            ("Slide numbers in headings", self.var_numbers),
            ("Images", self.var_images),
            ("Tables", self.var_tables),
            ("Charts (as data tables)", self.var_charts),
            ("Speaker notes", self.var_notes),
            ("Hidden slides", self.var_hidden),
            ("Cover page from first slide", self.var_cover),
            ("Page numbers in footer", self.var_pagenums),
            ("Landscape pages", self.var_landscape),
        ]
        for i, (label, var) in enumerate(opts):
            ttk.Checkbutton(grid, text=label, variable=var).grid(
                row=i // 3, column=i % 3, sticky="w", padx=(0, 22), pady=2)

        sep = ttk.Frame(card, style="Card.TFrame")
        sep.pack(fill="x", pady=(10, 0))
        ttk.Label(sep, text="Between slides:", style="Card.TLabel").pack(side="left", padx=(0, 10))
        for text, val in (("Thin line (compact)", SEPARATOR_RULE),
                          ("New page for each slide", SEPARATOR_PAGE),
                          ("Just spacing", SEPARATOR_NONE)):
            ttk.Radiobutton(sep, text=text, value=val, variable=self.var_separator).pack(side="left", padx=(0, 14))

        # Convert + progress
        action = tk.Frame(body, bg=BG)
        action.pack(fill="x", padx=24, pady=(4, 0))
        self.btn_convert = ttk.Button(action, text="Convert to Word", style="Accent.TButton",
                                      command=self.start_conversion)
        self.btn_convert.pack(side="left")
        self.btn_open = ttk.Button(action, text="Open document", style="Success.TButton",
                                   command=self.open_result, state="disabled")
        self.btn_open.pack(side="left", padx=(12, 0))
        self.btn_folder = ttk.Button(action, text="Show in folder", command=self.show_result,
                                     state="disabled")
        self.btn_folder.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(body, mode="determinate", maximum=1000)
        self.progress.pack(fill="x", padx=24, pady=(14, 4))
        ttk.Label(body, textvariable=self.var_status, style="Status.TLabel").pack(anchor="w", padx=24)

        # Log
        log_wrap = tk.Frame(body, bg=BORDER)
        log_wrap.pack(fill="both", expand=True, padx=24, pady=(8, 16))
        self.log = tk.Text(log_wrap, height=6, wrap="word", state="disabled", bd=0,
                           font=("Consolas", 9), bg="#FBFCFE", fg=TEXT, padx=8, pady=6,
                           highlightthickness=0)
        scroll = ttk.Scrollbar(log_wrap, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y", padx=(0, 1), pady=1)
        self.log.pack(fill="both", expand=True, padx=1, pady=1)
        self.log.tag_configure("ok", foreground=GREEN)
        self.log.tag_configure("warn", foreground="#B26A00")
        self.log.tag_configure("err", foreground=RED)

    # ------------------------------------------------------------- actions
    def _toggle_out(self):
        state = "disabled" if self.var_same_folder.get() else "normal"
        self.entry_out.configure(state=state)
        self.btn_out.configure(state=state)

    def _toggle_lyrics(self):
        on = self.var_lyrics.get()
        self.spin_size.configure(state="normal" if on else "disabled")
        self.chk_commas.configure(state="normal" if on else "disabled")
        # Sensible defaults for projected lyrics: wide pages, one slide per page
        self.var_landscape.set(on)
        self.var_separator.set(SEPARATOR_PAGE if on else SEPARATOR_RULE)

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="Choose where to save the Word documents")
        if d:
            self.var_out_dir.set(d)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PowerPoint files",
            filetypes=[("PowerPoint presentations", "*.pptx"), ("All files", "*.*")],
        )
        for p in paths:
            path = Path(p)
            if path.suffix.lower() != ".pptx":
                messagebox.showwarning(
                    APP_NAME,
                    f"'{path.name}' is not a .pptx file.\n\nOlder .ppt files must first be saved "
                    "as .pptx from PowerPoint (File > Save As).")
                continue
            if path not in self.files:
                self.files.append(path)
                self.listbox.insert("end", f"  {path.name}    —    {path.parent}")
        if self.files:
            self.var_status.set(f"{len(self.files)} file(s) ready.")

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.files[idx]

    def clear_files(self):
        self.listbox.delete(0, "end")
        self.files.clear()

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        for b in (self.btn_add, self.btn_remove, self.btn_clear, self.btn_convert):
            b.configure(state=state)

    def _write_log(self, text: str, tag: str = ""):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def start_conversion(self):
        if not self.files:
            messagebox.showinfo(APP_NAME, "Please add at least one .pptx file first.")
            return
        if not self.var_same_folder.get() and not self.var_out_dir.get():
            messagebox.showinfo(APP_NAME, "Please choose a save folder, or tick 'Save next to each presentation'.")
            return

        try:
            lyrics_size = float(self.var_lyrics_size.get())
            if not 8 <= lyrics_size <= 200:
                raise ValueError
        except ValueError:
            messagebox.showinfo(APP_NAME, "Font size must be a number between 8 and 200.")
            return

        options = ConvertOptions(
            lyrics_mode=self.var_lyrics.get(),
            lyrics_font_size=lyrics_size,
            lyrics_break_commas=self.var_lyrics_commas.get(),
            include_slide_numbers=self.var_numbers.get(),
            include_images=self.var_images.get(),
            include_tables=self.var_tables.get(),
            include_charts=self.var_charts.get(),
            include_notes=self.var_notes.get(),
            include_hidden_slides=self.var_hidden.get(),
            cover_from_first_slide=self.var_cover.get(),
            page_numbers=self.var_pagenums.get(),
            landscape=self.var_landscape.get(),
            slide_separator=self.var_separator.get(),
        )
        files = list(self.files)
        out_dir = None if self.var_same_folder.get() else Path(self.var_out_dir.get())

        self.outputs.clear()
        self.btn_open.configure(state="disabled")
        self.btn_folder.configure(state="disabled")
        self.progress["value"] = 0
        self._write_log(f"--- Starting: {len(files)} file(s) ---")
        self._set_busy(True)
        self.worker = threading.Thread(target=self._run, args=(files, out_dir, options), daemon=True)
        self.worker.start()

    # The worker never touches widgets; it only posts messages to the queue.
    def _run(self, files: list[Path], out_dir: Path | None, options: ConvertOptions):
        n = len(files)
        succeeded = 0
        for i, src in enumerate(files):
            self.queue.put(("log", f"Converting {src.name} …", ""))
            try:
                target_dir = out_dir or src.parent
                target = unique_path(target_dir / (src.stem + ".docx"))

                def progress(cur, total, msg, _i=i, _name=src.name):
                    frac = (_i + (cur / total if total else 1)) / n
                    self.queue.put(("progress", frac, f"{_name}: {msg}"))

                res = convert_file(src, target, options, progress)
                succeeded += 1
                self.outputs.append(res.output_path)
                self.queue.put(("log",
                                f"  ✔ Saved {res.output_path.name}  "
                                f"({res.slides_converted}/{res.slides_total} slides, {res.images} images, "
                                f"{res.tables} tables, {res.charts} charts)", "ok"))
                for w in res.warnings:
                    self.queue.put(("log", f"  ! {w}", "warn"))
            except PermissionError:
                self.queue.put(("log", f"  ✘ {src.name}: cannot write the file. Is it open in Word, "
                                       "or is the folder read-only?", "err"))
            except Exception as exc:
                detail = traceback.format_exc(limit=2).strip().splitlines()[-1]
                self.queue.put(("log", f"  ✘ {src.name}: {exc or detail}", "err"))
        self.queue.put(("done", succeeded, n))

    def _poll_queue(self):
        try:
            while True:
                kind, *rest = self.queue.get_nowait()
                if kind == "log":
                    self._write_log(rest[0], rest[1])
                elif kind == "progress":
                    self.progress["value"] = rest[0] * 1000
                    self.var_status.set(rest[1])
                elif kind == "done":
                    ok, total = rest
                    self._set_busy(False)
                    self.progress["value"] = 1000 if ok else 0
                    if ok == total:
                        self.var_status.set(f"Finished: {ok} of {total} converted.")
                    else:
                        self.var_status.set(f"Finished with problems: {ok} of {total} converted.")
                    if self.outputs:
                        self.btn_open.configure(state="normal")
                        self.btn_folder.configure(state="normal")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def open_result(self):
        if self.outputs:
            open_path(self.outputs[-1])

    def show_result(self):
        if self.outputs:
            reveal_in_folder(self.outputs[-1])


def main():
    # Crisp text on high-DPI Windows screens
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()