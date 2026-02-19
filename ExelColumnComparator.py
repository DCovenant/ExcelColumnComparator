#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from pathlib import Path
import pandas as pd
import os
import subprocess
import sys

# ── Light Theme ──────────────────────────────────────────────────────────
C = {
    "bg":      "#f0f2f5", "surface":  "#ffffff", "surface2": "#f8f9fb",
    "bar":     "#2c3e50", "bar_text": "#ecf0f1", "accent":   "#3498db",
    "accent_dk": "#2980b9", "text": "#2c3e50", "dim": "#7f8c8d",
    "border":  "#dce1e6", "sel": "#3498db", "alt": "#f8f9fa",
    "hdr_bg":  "#34495e", "green": "#27ae60", "orange": "#e67e22",
    "red": "#e74c3c", "cyan": "#2980b9", "purple": "#8e44ad",
    "chk_sel": "#ffffff",
}
F      = ("Segoe UI", 10)
F_B    = ("Segoe UI", 10, "bold")
F_T    = ("Segoe UI", 16, "bold")
F_SUB  = ("Segoe UI", 12)
F_MONO = ("Consolas", 10)


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Column Comparator")
        self.root.geometry("1200x780")
        self.root.configure(bg=C["bg"])
        self.root.withdraw()
        self._styles()
        self.file_configs = []
        self.history = []
        self.temp_files = []
        self.template_mode = False
        self.template_files = []
        self.template_config = None
        self.pick_files()
        self.root.mainloop()

    def _styles(self):
        s = ttk.Style(); s.theme_use("clam")
        s.configure("T.Treeview", background=C["surface"], foreground=C["text"],
                     fieldbackground=C["surface"], font=F, rowheight=26)
        s.configure("T.Treeview.Heading", background=C["hdr_bg"],
                     foreground=C["cyan"], font=F_B, relief="flat")
        s.map("T.Treeview.Heading", background=[("active", C["surface2"])])
        s.map("T.Treeview", background=[("selected", C["sel"])],
              foreground=[("selected", "#fff")])
        s.configure("A.TButton", background=C["accent"], foreground="#fff",
                     font=F_B, padding=(20, 10))
        s.map("A.TButton", background=[("active", C["accent_dk"])])
        s.configure("TScrollbar", background=C["surface2"], troughcolor=C["bg"],
                     bordercolor=C["bg"], arrowcolor=C["dim"])

    def clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    # ── 1  File picker ───────────────────────────────────────────────────
    def pick_files(self):
        self.root.withdraw()
        paths = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not paths:
            if len(self.file_configs) >= 2:
                self.show_pair_selector(); return
            self.root.destroy(); return
        self.temp_files = list(paths)
        if len(paths) > 1 and not self.template_mode:
            if messagebox.askyesno("", "Do you want to create a template?"):
                self.template_mode = True
                self.template_files = list(paths)
        self.process_next_file()

    def process_next_file(self):
        if not self.temp_files:
            if len(self.file_configs) >= 2:
                self.show_pair_selector()
            else:
                self.root.destroy()
            return
        self.cur_path = self.temp_files.pop(0)
        self.wb = load_workbook(self.cur_path, data_only=True)
        self.show_sheet_and_header()

    # ── 2  Sheet + header row ────────────────────────────────────────────
    def show_sheet_and_header(self):
        self.root.deiconify()
        self.clear()
        self.history.append("sheet_and_header")
        bar = tk.Frame(self.root, bg=C["bar"], height=75)
        bar.pack(fill=tk.X); bar.pack_propagate(False)
        tk.Label(bar, text=Path(self.cur_path).name, font=F_T,
                 fg=C["accent"], bg=C["bar"]).pack(side=tk.LEFT, padx=20)
        tk.Label(bar, text=f"File {len(self.file_configs)+1}  |  Step 1 : Choose a sheet and click the header row",
                 font=F_SUB, fg=C["dim"], bg=C["bar"]).pack(side=tk.RIGHT, padx=20)

        mid = tk.Frame(self.root, bg=C["bg"])
        mid.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        # sheet list
        sb = tk.Frame(mid, bg=C["surface"], width=190,
                      highlightbackground=C["border"], highlightthickness=1)
        sb.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8)); sb.pack_propagate(False)
        tk.Label(sb, text="Sheets", font=F_B, bg=C["surface"],
                 fg=C["cyan"]).pack(padx=10, pady=(8, 4), anchor="w")
        self.sheet_lb = tk.Listbox(sb, font=F, relief="flat", bd=0,
                                   selectbackground=C["accent"], selectforeground="#fff",
                                   highlightthickness=0, bg=C["surface"], fg=C["text"])
        for name in self.wb.sheetnames:
            lock = "  [hidden]" if self.wb[name].sheet_state == "hidden" else ""
            self.sheet_lb.insert(tk.END, f"  {name}{lock}")
        self.sheet_lb.select_set(0)
        self.sheet_lb.bind("<<ListboxSelect>>", lambda _: self.load_sheet())
        self.sheet_lb.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # treeview
        ct = tk.Frame(mid, bg=C["surface"],
                      highlightbackground=C["border"], highlightthickness=1)
        ct.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        xscr = ttk.Scrollbar(ct, orient=tk.HORIZONTAL)
        yscr = ttk.Scrollbar(ct, orient=tk.VERTICAL)
        self.tree = ttk.Treeview(ct, style="T.Treeview", show="headings",
                                 xscrollcommand=xscr.set, yscrollcommand=yscr.set)
        xscr.config(command=self.tree.xview); yscr.config(command=self.tree.yview)
        yscr.pack(side=tk.RIGHT, fill=tk.Y)
        xscr.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_row)

        bot = tk.Frame(self.root, bg=C["bg"])
        bot.pack(fill=tk.X, padx=12, pady=(0, 8))
        back_btn = ttk.Button(bot, text="<- Back", style="A.TButton", command=self._back)
        back_btn.pack(side=tk.LEFT)
        self.status = tk.StringVar(value="Click a row to mark it as the table header")
        tk.Label(bot, textvariable=self.status, font=F, bg=C["bg"],
                 fg=C["dim"]).pack(side=tk.LEFT, expand=True)
        self.next_btn = ttk.Button(bot, text="Next  ->", style="A.TButton",
                                   command=self._header_next, state=tk.DISABLED)
        self.next_btn.pack(side=tk.RIGHT)
        self.sel_row = None
        self.load_sheet()

    def _sname(self):
        idx = self.sheet_lb.curselection()
        return self.wb.sheetnames[idx[0]] if idx else self.wb.sheetnames[0]

    def load_sheet(self):
        ws = self.wb[self._sname()]
        self.raw, self.rcols = [], []
        for row in ws.iter_rows(max_row=200, max_col=40):
            vals, clr = [], None
            for cell in row:
                vals.append(cell.value)
                if clr is None and cell.fill and cell.fill.start_color:
                    try:
                        rgb = cell.fill.start_color.rgb
                        if isinstance(rgb, str) and len(rgb) == 8 and rgb != "00000000":
                            clr = f"#{rgb[2:]}"
                    except Exception:
                        pass
            self.raw.append(tuple(vals)); self.rcols.append(clr)
        if not self.raw:
            return
        nc = max(len(r) for r in self.raw)
        cids = [f"c{i}" for i in range(nc)]
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cids
        for i, c in enumerate(cids):
            self.tree.heading(c, text=f"Col {i+1}")
            self.tree.column(c, width=120, minwidth=50)
        for i, row in enumerate(self.raw):
            v = [str(x) if x is not None else "" for x in row] + [""] * (nc - len(row))
            tag = f"r{i}"
            self.tree.insert("", "end", iid=str(i), values=v, tags=(tag,))
            bg = self.rcols[i] or (C["alt"] if i % 2 else C["surface"])
            self.tree.tag_configure(tag, background=bg)
        self.sel_row = None
        self.status.set("Click a row to mark it as the table header")
        self.next_btn.config(state=tk.DISABLED)

    def _on_row(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        self.sel_row = int(sel[0])
        for i, iid in enumerate(self.tree.get_children()):
            bg = self.rcols[i] or (C["alt"] if i % 2 else C["surface"])
            self.tree.item(iid, tags=(f"r{i}",))
            self.tree.tag_configure(f"r{i}", background=bg, foreground=C["text"])
        stag = "sel"
        self.tree.item(sel[0], tags=(stag,))
        self.tree.tag_configure(stag, background=C["sel"], foreground="#fff")
        pre = [str(v) for v in self.raw[self.sel_row] if v is not None][:6]
        self.status.set(f"Row {self.sel_row+1}:  {' | '.join(pre)}")
        self.next_btn.config(state=tk.NORMAL)

    def _header_next(self):
        self.cur_sheet = self._sname()
        self.cur_hdr_idx = self.sel_row
        self.cur_hdr_vals = self.raw[self.cur_hdr_idx]
        self.wb.close()
        self.show_col_sel()

    # ── 3  Column selection ──────────────────────────────────────────────
    def show_col_sel(self):
        self.root.deiconify()
        self.clear()
        self.history.append("col_sel")
        bar = tk.Frame(self.root, bg=C["bar"], height=75)
        bar.pack(fill=tk.X); bar.pack_propagate(False)
        tk.Label(bar, text=Path(self.cur_path).name, font=F_T,
                 fg=C["accent"], bg=C["bar"]).pack(side=tk.LEFT, padx=20)
        tk.Label(bar, text=f"File {len(self.file_configs)+1}  |  Step 2 : Select columns to compare",
                 font=F_SUB, fg=C["dim"], bg=C["bar"]).pack(side=tk.RIGHT, padx=20)

        card = tk.Frame(self.root, bg=C["surface"],
                        highlightbackground=C["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=12)
        tk.Label(card, text="Tick one or more columns:", font=F_B,
                 bg=C["surface"], fg=C["cyan"]).pack(padx=15, pady=(15, 8), anchor="w")

        wrap = tk.Frame(card, bg=C["surface"])
        wrap.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self.col_vars = []
        ci = 0
        per_row = 4
        for val in self.cur_hdr_vals:
            if val is None:
                continue
            name = str(val).strip()
            if not name or name.lower() == "nan":
                continue
            var = tk.BooleanVar()
            cb = tk.Checkbutton(wrap, text=name, variable=var, font=F,
                                bg=C["surface"], fg=C["text"],
                                selectcolor=C["chk_sel"],
                                activebackground=C["surface"],
                                activeforeground=C["accent"],
                                padx=8, pady=4, anchor="w")
            r, c = divmod(ci, per_row)
            cb.grid(row=r, column=c, sticky="w", padx=4, pady=2)
            self.col_vars.append((name, var))
            ci += 1
        for c in range(per_row):
            wrap.columnconfigure(c, weight=1)

        bot = tk.Frame(self.root, bg=C["bg"])
        bot.pack(fill=tk.X, padx=20, pady=(0, 12))
        ttk.Button(bot, text="<- Back", style="A.TButton", command=self._back).pack(side=tk.LEFT)
        ttk.Button(bot, text="Next  ->", style="A.TButton",
                   command=self._cols_next).pack(side=tk.RIGHT)

    def _cols_next(self):
        cols = [n for n, v in self.col_vars if v.get()]
        if not cols:
            messagebox.showwarning("", "Select at least one column"); return
        self.file_configs.append({"path": self.cur_path, "sheet": self.cur_sheet,
                                  "header_row": self.cur_hdr_idx, "columns": cols})
        if self.template_mode and not self.template_config:
            self.template_config = {"sheet": self.cur_sheet, "header_row": self.cur_hdr_idx, "columns": cols}
            for tfile in self.template_files[1:]:
                self.file_configs.append({"path": tfile, "sheet": self.cur_sheet,
                                         "header_row": self.cur_hdr_idx, "columns": cols})
            self.show_pair_selector()
        elif self.temp_files:
            self.process_next_file()
        elif messagebox.askyesno("", f"{len(self.file_configs)} file(s) added.\nAdd another file?"):
            self.pick_files()
        elif len(self.file_configs) >= 2:
            self.show_pair_selector()
        else:
            messagebox.showinfo("", "Need at least 2 files"); self.pick_files()

    # ── 4  What compares to what? ────────────────────────────────────────
    def show_pair_selector(self):
        self.root.deiconify()
        self.clear()
        self.history.append("pair_selector")
        bar = tk.Frame(self.root, bg=C["bar"], height=75)
        bar.pack(fill=tk.X); bar.pack_propagate(False)
        tk.Label(bar, text="What compares to what?", font=F_T,
                 fg=C["accent"], bg=C["bar"]).pack(side=tk.LEFT, padx=20)
        tk.Label(bar, text="Final step : Map columns from File #1 to each other file",
                 font=F_SUB, fg=C["dim"], bg=C["bar"]).pack(side=tk.RIGHT, padx=20)

        card = tk.Frame(self.root, bg=C["surface"],
                        highlightbackground=C["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=12)

        f1 = self.file_configs[0]
        others = self.file_configs[1:]
        SKIP = "-- skip --"

        # scrollable area
        canvas = tk.Canvas(card, bg=C["surface"], highlightthickness=0)
        vsb = ttk.Scrollbar(card, orient=tk.VERTICAL, command=canvas.yview)
        inner = tk.Frame(canvas, bg=C["surface"])
        inner.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", tags="inner")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure("inner", width=e.width))
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-e.delta // 120, "units"))
        canvas.bind_all("<Button-4>", lambda _: canvas.yview_scroll(-3, "units"))
        canvas.bind_all("<Button-5>", lambda _: canvas.yview_scroll(3, "units"))

        grid = tk.Frame(inner, bg=C["surface"])
        grid.pack(fill=tk.X, padx=8, pady=8)

        # header row
        f1_name = Path(f1["path"]).name
        tk.Label(grid, text=f"#{1}  {f1_name}", font=F_B,
                 bg=C["surface"], fg=C["accent"]).grid(row=0, column=0, sticky="w",
                                                       padx=(10, 20), pady=(0, 8))
        for k, oth in enumerate(others):
            name = Path(oth["path"]).name
            tk.Label(grid, text=f"#{k+2}  {name}", font=F_B,
                     bg=C["surface"], fg=C["orange"]).grid(row=0, column=k + 1,
                                                           sticky="w", padx=10, pady=(0, 8))

        # separator
        sep = tk.Frame(grid, bg=C["border"], height=1)
        sep.grid(row=1, column=0, columnspan=1 + len(others), sticky="ew",
                 padx=10, pady=(0, 6))

        # one row per File #1 column, with a dropdown per other file
        self.map_vars = []  # [(col1, [(file_idx, var), ...]), ...]
        for ri, col1 in enumerate(f1["columns"]):
            tk.Label(grid, text=col1, font=F, bg=C["surface"],
                     fg=C["text"]).grid(row=ri + 2, column=0, sticky="w",
                                        padx=(10, 20), pady=4)
            row_vars = []
            for k, oth in enumerate(others):
                options = [SKIP] + oth["columns"]
                var = tk.StringVar(value=SKIP)
                # auto-match by name (case-insensitive)
                for oc in oth["columns"]:
                    if oc.strip().lower() == col1.strip().lower():
                        var.set(oc); break
                cb = ttk.Combobox(grid, textvariable=var, values=options,
                                  state="readonly", width=28, font=F)
                cb.grid(row=ri + 2, column=k + 1, sticky="w", padx=10, pady=4)
                row_vars.append((k + 1, var))  # k+1 = file_configs index
            self.map_vars.append((col1, row_vars))

        for c in range(1 + len(others)):
            grid.columnconfigure(c, weight=1)

        bot = tk.Frame(self.root, bg=C["bg"])
        bot.pack(fill=tk.X, padx=20, pady=(0, 12))
        ttk.Button(bot, text="<- Back", style="A.TButton", command=self._back).pack(side=tk.LEFT)
        ttk.Button(bot, text="Compare  ->", style="A.TButton",
                   command=self._pairs_next).pack(side=tk.RIGHT)

    def _pairs_next(self):
        SKIP = "-- skip --"
        # build mapping: for each other file, which col1->colN pairs
        n_others = len(self.file_configs) - 1
        self.mappings = {}  # {file_idx: [(col1, colN), ...]}
        for col1, row_vars in self.map_vars:
            for fidx, var in row_vars:
                v = var.get()
                if v != SKIP:
                    self.mappings.setdefault(fidx, []).append((col1, v))
        if not self.mappings:
            messagebox.showwarning("", "Map at least one column"); return
        self.run_comparison()

    # ── 5  Comparison results ────────────────────────────────────────────
    def run_comparison(self):
        self.root.deiconify()
        self.clear()
        bar = tk.Frame(self.root, bg=C["bar"], height=70)
        bar.pack(fill=tk.X); bar.pack_propagate(False)
        tk.Label(bar, text="Comparison Results", font=F_T,
                 fg=C["accent"], bg=C["bar"]).pack(side=tk.LEFT, padx=20, pady=12)

        loading = tk.Toplevel(self.root)
        loading.title("Loading")
        loading.geometry("300x100")
        loading.resizable(False, False)
        tk.Label(loading, text="Loading Excel files...", font=F_B,
                 bg=C["bg"], fg=C["text"]).pack(expand=True)
        self.root.update()

        dfs = []
        for cfg in self.file_configs:
            df = pd.read_excel(cfg["path"], header=cfg["header_row"],
                               sheet_name=cfg["sheet"])
            dfs.append(df)

        loading.destroy()

        def col_vals(df, col, hdr_row):
            if col not in df.columns:
                return {}
            result = {}
            for idx, val in df[col].dropna().items():
                v = str(val).strip()
                if v and v.lower() != "nan":
                    excel_row = hdr_row + 2 + idx
                    if v not in result:
                        result[v] = excel_row
            return result

        canvas = tk.Canvas(self.root, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=canvas.yview)
        self._res_inner = tk.Frame(canvas, bg=C["bg"])
        self._res_inner.bind("<Configure>",
                             lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._res_inner, anchor="nw", tags="inner")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure("inner", width=e.width))
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        self._res_canvas = canvas
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-e.delta // 120, "units"))
        canvas.bind_all("<Button-4>", lambda _: canvas.yview_scroll(-3, "units"))
        canvas.bind_all("<Button-5>", lambda _: canvas.yview_scroll(3, "units"))

        PREVIEW = 20
        f1_cfg = self.file_configs[0]
        f1_name = Path(f1_cfg["path"]).name

        for fidx, col_pairs in sorted(self.mappings.items()):
            fN_cfg = self.file_configs[fidx]
            fN_name = Path(fN_cfg["path"]).name

            vals1 = {}
            valsN = {}
            for c1, cN in col_pairs:
                vals1.update(col_vals(dfs[0], c1, self.file_configs[0]["header_row"]))
                valsN.update(col_vals(dfs[fidx], cN, fN_cfg["header_row"]))

            only1_keys = sorted(set(vals1.keys()) - set(valsN.keys()))
            onlyN_keys = sorted(set(valsN.keys()) - set(vals1.keys()))
            common = set(vals1.keys()) & set(valsN.keys())
            only1 = [(k, vals1[k]) for k in only1_keys]
            onlyN = [(k, valsN[k]) for k in onlyN_keys]

            card = tk.Frame(self._res_inner, bg=C["surface"],
                            highlightbackground=C["border"], highlightthickness=1)
            card.pack(fill=tk.X, pady=6, padx=4)

            tk.Label(card, text=f"{f1_name} [{f1_cfg['sheet']}]   vs   {fN_name} [{fN_cfg['sheet']}]",
                     font=F_B, bg=C["surface"], fg=C["text"]).pack(
                padx=15, pady=(12, 2), anchor="w")

            # show which columns were mapped
            mapped_str = "    ".join(f"{c1}  ->  {cN}" for c1, cN in col_pairs)
            tk.Label(card, text=mapped_str, font=F,
                     bg=C["surface"], fg=C["dim"]).pack(padx=15, pady=(0, 4), anchor="w")

            st = tk.Frame(card, bg=C["surface"])
            st.pack(fill=tk.X, padx=15, pady=4)
            tk.Label(st, text=f"Common: {len(common)}", font=F_B,
                     bg=C["surface"], fg=C["green"]).pack(side=tk.LEFT, padx=(0, 18))
            tk.Label(st, text=f"Only in {f1_name}: {len(only1)}", font=F_B,
                     bg=C["surface"], fg=C["accent"]).pack(side=tk.LEFT, padx=(0, 18))
            tk.Label(st, text=f"Only in {fN_name}: {len(onlyN)}", font=F_B,
                     bg=C["surface"], fg=C["orange"]).pack(side=tk.LEFT)

            if only1:
                self._exp(card, f"Only in {f1_name}", only1, PREVIEW, C["accent"],
                         f1_cfg["path"], col_pairs[0][0], f1_cfg["sheet"], f1_cfg["header_row"], dfs[0], [k for k,r in only1])
            if onlyN:
                self._exp(card, f"Only in {fN_name}", onlyN, PREVIEW, C["orange"],
                         fN_cfg["path"], col_pairs[0][1], fN_cfg["sheet"], fN_cfg["header_row"], dfs[fidx], [k for k,r in onlyN])
            tk.Frame(card, bg=C["surface"], height=8).pack()

        # summary
        sm = tk.Frame(self._res_inner, bg=C["hdr_bg"],
                      highlightbackground=C["border"], highlightthickness=1)
        sm.pack(fill=tk.X, pady=6, padx=4)
        tk.Label(sm, text="Summary", font=F_T, fg=C["purple"],
                 bg=C["hdr_bg"]).pack(padx=15, pady=(12, 4), anchor="w")
        for fidx, col_pairs in sorted(self.mappings.items()):
            fN_name = Path(self.file_configs[fidx]["path"]).name
            vals1 = {}
            valsN = {}
            for c1, cN in col_pairs:
                vals1.update(col_vals(dfs[0], c1, self.file_configs[0]["header_row"]))
                valsN.update(col_vals(dfs[fidx], cN, self.file_configs[fidx]["header_row"]))
            common = set(vals1.keys()) & set(valsN.keys())
            u1 = set(vals1.keys()) - set(valsN.keys())
            uN = set(valsN.keys()) - set(vals1.keys())
            tk.Label(sm, text=f"{f1_name} vs {fN_name}:  "
                     f"{len(common)} common,  {len(u1)} unique to #1,  {len(uN)} unique to #{fidx+1}",
                     font=F, fg=C["dim"], bg=C["hdr_bg"]).pack(padx=15, anchor="w")
        tk.Frame(sm, bg=C["hdr_bg"], height=12).pack()

        bot = tk.Frame(self.root, bg=C["bg"])
        bot.pack(fill=tk.X, padx=12, pady=8)
        ttk.Button(bot, text="New Comparison", style="A.TButton",
                   command=self._new_comparison).pack(side=tk.RIGHT)

    def _copy_to_clipboard(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("", "Copied to clipboard")

    def _open_excel_file(self, file_path):
        if sys.platform == 'win32':
            os.startfile(file_path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', file_path])
        else:
            subprocess.Popen(['xdg-open', file_path])

    def _show_excel_window(self, file_path, column, sheet, hdr_row, df, unique_vals, scroll_to=None):
        win = tk.Toplevel(self.root)
        win.title(f"Excel: {Path(file_path).name} - {column}")
        win.geometry("900x600")

        tree = ttk.Treeview(win, style="T.Treeview", show="headings")
        xscr = ttk.Scrollbar(win, orient=tk.HORIZONTAL, command=tree.xview)
        yscr = ttk.Scrollbar(win, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(xscrollcommand=xscr.set, yscrollcommand=yscr.set)

        cols = [column]
        tree["columns"] = cols
        tree.heading("#0", text="Row")
        tree.column("#0", width=50, anchor="center")
        tree.heading(column, text=column)
        tree.column(column, width=200)

        unique_set = set(unique_vals)
        for idx, val in df[column].dropna().items():
            v = str(val).strip()
            if v and v.lower() != "nan":
                excel_row = hdr_row + 1 + idx
                tag = "unique" if v in unique_set else "normal"
                tree.insert("", "end", iid=str(excel_row), text=str(excel_row), values=(v,), tags=(tag,))

        tree.tag_configure("unique", background=C["red"], foreground="#fff")
        tree.tag_configure("normal", background=C["surface"])

        yscr.pack(side=tk.RIGHT, fill=tk.Y)
        xscr.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        if scroll_to:
            tree.see(str(scroll_to))
            tree.selection_set(str(scroll_to))

    def _new_comparison(self):
        self.file_configs = []
        self.history = []
        self.temp_files = []
        self.template_mode = False
        self.template_files = []
        self.template_config = None
        self.pick_files()

    def _exp(self, parent, title, items, n, color, file_path=None, column=None, sheet=None, hdr_row=None, df=None, unique_vals=None):
        fr = tk.Frame(parent, bg=C["surface"])
        fr.pack(fill=tk.X, padx=15, pady=2)
        hdr_fr = tk.Frame(fr, bg=C["surface"])
        hdr_fr.pack(fill=tk.X)
        tk.Label(hdr_fr, text=title + ":", font=F_B, bg=C["surface"],
                 fg=color).pack(side=tk.LEFT, anchor="w")
        if file_path:
            tk.Button(hdr_fr, text="Open Excel", font=F, bg=C["surface"],
                     fg=color, relief="flat", cursor="hand2",
                     command=lambda fp=file_path: self._open_excel_file(fp),
                     activeforeground=color, activebackground=C["alt"]).pack(side=tk.LEFT, padx=(10, 0))
        inner = tk.Frame(fr, bg=C["surface"])
        inner.pack(fill=tk.X, padx=16)
        for v, row in items[:n]:
            lbl = tk.Label(inner, text=f"  Row {row}: {v}", font=F_MONO, bg=C["surface"],
                     fg=C["text"], cursor="hand2")
            lbl.pack(anchor="w")
            lbl.bind("<Button-1>", lambda e, val=v: self._copy_to_clipboard(val))
        if len(items) > n:
            rest = items[n:]
            def show(rest=rest):
                btn.destroy()
                for v, row in rest:
                    lbl = tk.Label(inner, text=f"  Row {row}: {v}", font=F_MONO,
                             bg=C["surface"], fg=C["text"], cursor="hand2")
                    lbl.pack(anchor="w")
                    lbl.bind("<Button-1>", lambda e, val=v: self._copy_to_clipboard(val))
                self._res_inner.update_idletasks()
                self._res_canvas.configure(
                    scrollregion=self._res_canvas.bbox("all"))
            btn = tk.Button(inner, text=f"Show all {len(items)} items  v",
                            command=show, font=F, fg=color, bg=C["surface"],
                            relief="flat", cursor="hand2",
                            activeforeground=color, activebackground=C["alt"])
            btn.pack(anchor="w", pady=3)

    def _back(self):
        if len(self.history) < 2:
            return
        self.history.pop()
        if self.history[-1] == "sheet_and_header":
            self.show_sheet_and_header()
        elif self.history[-1] == "col_sel":
            self.show_col_sel()
        elif self.history[-1] == "pair_selector":
            self.show_pair_selector()


if __name__ == "__main__":
    App()
# Requires: pip install openpyxl pandas
# On Linux you may also need: sudo apt install python3-tk
