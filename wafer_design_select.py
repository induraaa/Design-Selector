"""
Wafer Map Export Tool
- Fast NumPy/Pillow rendering (no canvas freeze)
- Export to Excel (.xlsx) as default  — matches original file format
- Export to Unicode Text (.txt)       — optional, for assembly site upload
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, re, time
from collections import Counter

# ── dependency bootstrap ──────────────────────────────────────────────────────
def _ensure(pkg, import_name=None):
    import importlib, subprocess, sys
    try:
        importlib.import_module(import_name or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg,
                               "--break-system-packages", "-q"])

_ensure("openpyxl");  _ensure("pillow", "PIL");  _ensure("numpy")

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.cell import WriteOnlyCell
import numpy as np
from PIL import Image, ImageTk

# ── palettes ─────────────────────────────────────────────────────────────────
PALETTE_RGB = [
    (  0,128,  0),(220, 50, 47),( 65,105,225),(255,140,  0),(128,  0,128),
    (  0,128,128),(220, 20, 60),( 70,130,180),(218,165, 32),( 46,139, 87),
    (255, 99, 71),( 70,130,180),(210,105, 30),(106, 90,205),( 32,178,170),
    (184,134, 11),( 95,158,160),(188,143,143),( 85,107, 47),(139,  0,139),
]
PALETTE_HEX = ["#%02x%02x%02x" % c for c in PALETTE_RGB]
PALETTE_XLSX = ["%02X%02X%02X" % c for c in PALETTE_RGB]   # openpyxl needs no #

SEL_RGB  = (  0,170,  0);  SEL_XLSX  = "00AA00"
X_RGB    = (100,100,100);  X_XLSX    = "646464"
x_RGB    = (210,210,210);  x_XLSX    = "D2D2D2"


# ── Export format dialog ──────────────────────────────────────────────────────
class ExportDialog(tk.Toplevel):
    """Small dialog to choose export format and options."""
    def __init__(self, parent, designs_selected, die_count):
        super().__init__(parent)
        self.title("Export Options")
        self.resizable(False, False)
        self.grab_set()
        self.result = None   # 'xlsx' | 'txt' | None

        # Centre over parent
        self.geometry("+%d+%d" % (parent.winfo_x()+200, parent.winfo_y()+180))

        tk.Label(self, text="Export Format", font=("Segoe UI", 10, "bold"),
                 pady=8).pack()

        info = f"Selected: {', '.join(sorted(designs_selected))}  |  {die_count:,} die → Bin 1"
        tk.Label(self, text=info, font=("Segoe UI", 8),
                 fg="#444444", padx=12).pack()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # Format choice
        self.fmt = tk.StringVar(value="xlsx")
        frm = tk.Frame(self); frm.pack(padx=16, pady=4, anchor="w")

        tk.Radiobutton(frm, text="Excel Workbook  (.xlsx)",
                       variable=self.fmt, value="xlsx",
                       font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=3)
        tk.Label(frm, text="Matches original file format, with colour-coded map sheet",
                 font=("Segoe UI", 7), fg="#666").grid(row=1, column=0, sticky="w", padx=20)

        tk.Radiobutton(frm, text="Unicode Text  (.txt)",
                       variable=self.fmt, value="txt",
                       font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=(8,3))
        tk.Label(frm, text="Plain character map for assembly site upload",
                 font=("Segoe UI", 7), fg="#666").grid(row=3, column=0, sticky="w", padx=20)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # xlsx-only option: include colour fills
        self.add_colour = tk.BooleanVar(value=True)
        self.colour_cb = tk.Checkbutton(self,
            text="Include colour fills in Excel  (slower export, ~8s)",
            variable=self.add_colour,
            font=("Segoe UI", 8))
        self.colour_cb.pack(anchor="w", padx=16, pady=2)

        # txt-only option: line ending
        self.line_ending = tk.StringVar(value="crlf")
        self.le_frame = tk.Frame(self)
        self.le_frame.pack(anchor="w", padx=16, pady=2)
        tk.Label(self.le_frame, text="Line ending:", font=("Segoe UI",8)).pack(side="left")
        tk.Radiobutton(self.le_frame, text="CRLF (Windows)", variable=self.line_ending,
                       value="crlf", font=("Segoe UI",8)).pack(side="left", padx=4)
        tk.Radiobutton(self.le_frame, text="LF (Unix)", variable=self.line_ending,
                       value="lf", font=("Segoe UI",8)).pack(side="left")

        self.fmt.trace_add("write", self._on_fmt_change)
        self._on_fmt_change()

        # Buttons
        btn_row = tk.Frame(self); btn_row.pack(pady=10)
        tk.Button(btn_row, text="Export...", command=self._ok,
                  font=("Segoe UI",9), relief="raised", bd=2,
                  width=10, default="active").pack(side="left", padx=6)
        tk.Button(btn_row, text="Cancel", command=self.destroy,
                  font=("Segoe UI",9), relief="raised", bd=2,
                  width=8).pack(side="left", padx=6)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def _on_fmt_change(self, *_):
        is_xlsx = self.fmt.get() == "xlsx"
        self.colour_cb.config(state="normal" if is_xlsx else "disabled")
        for w in self.le_frame.winfo_children():
            w.config(state="disabled" if is_xlsx else "normal")

    def _ok(self):
        self.result = {
            "fmt":        self.fmt.get(),
            "add_colour": self.add_colour.get(),
            "line_ending": self.line_ending.get(),
        }
        self.destroy()


# ── Main app ──────────────────────────────────────────────────────────────────
class WaferMapTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Wafer Map Export Tool")
        self.geometry("1200x780")
        self.minsize(900, 600)

        style = ttk.Style(self)
        for theme in ("winnative", "vista", "clam", "default"):
            try: style.theme_use(theme); break
            except: pass

        self.filepath         = None
        self.source_sheet     = None
        self.grid_data        = []
        self.enc              = None
        self.designs          = []
        self.design_counts    = {}
        self.selected_designs = set()
        self.colour_map_hex   = {}
        self.colour_map_rgb   = {}
        self.colour_map_xlsx  = {}
        self.design_vars      = {}
        self.zoom             = 1.0
        self._photo           = None
        self._redraw_pending  = False

        self._build_menu()
        self._build_toolbar()
        self._build_body()
        self._build_statusbar()

    # ── menu ─────────────────────────────────────────────────────────────────
    def _build_menu(self):
        mb = tk.Menu(self); self.config(menu=mb)
        fm = tk.Menu(mb, tearoff=0)
        fm.add_command(label="Open...",   command=self._open_file)
        fm.add_separator()
        fm.add_command(label="Export...", command=self._export)
        fm.add_separator()
        fm.add_command(label="Exit",      command=self.destroy)
        mb.add_cascade(label="File", menu=fm)
        mb.add_cascade(label="View", menu=tk.Menu(mb, tearoff=0))
        hm = tk.Menu(mb, tearoff=0)
        hm.add_command(label="About", command=lambda: messagebox.showinfo(
            "About", "Wafer Map Export Tool\nPython · Tkinter · NumPy · Pillow · openpyxl"))
        mb.add_cascade(label="Help", menu=hm)

    # ── toolbar ───────────────────────────────────────────────────────────────
    def _build_toolbar(self):
        tb = tk.Frame(self, relief="raised", bd=1, bg="#f0f0f0")
        tb.pack(fill="x", side="top")

        def tbtn(text, cmd, w=8):
            b = tk.Button(tb, text=text, command=cmd,
                          relief="raised", bd=2, bg="#f0f0f0",
                          font=("Segoe UI", 9), activebackground="#d9d9d9",
                          padx=6, pady=3, width=w)
            b.pack(side="left", padx=2, pady=3)
            return b

        tbtn("Open",    self._open_file,   6)
        tbtn("Convert", self._do_convert,  7)
        self.tb_export = tbtn("Export", self._export, 6)
        self.tb_export.config(state="disabled")

        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=4, pady=4)
        tbtn("Zoom In",  self._zoom_in,  7)
        tbtn("Zoom Out", self._zoom_out, 8)
        tbtn("Fit",      self._zoom_fit, 4)
        tbtn("Reset",    self._zoom_reset, 6)

        self.zoom_lbl = tk.Label(tb, text="Zoom: 100%",
                                  bg="#f0f0f0", font=("Segoe UI", 9))
        self.zoom_lbl.pack(side="right", padx=10)
        self.show_grid = tk.BooleanVar(value=True)
        tk.Checkbutton(tb, text="Grid Lines", variable=self.show_grid,
                       command=self._schedule_redraw,
                       bg="#f0f0f0", font=("Segoe UI", 9)).pack(side="right", padx=4)

    # ── body ──────────────────────────────────────────────────────────────────
    def _build_body(self):
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left_outer = tk.Frame(paned, width=215, bg="#f0f0f0",
                               relief="sunken", bd=1)
        left_outer.pack_propagate(False)
        paned.add(left_outer, weight=0)

        lc = tk.Canvas(left_outer, bg="#f0f0f0", highlightthickness=0)
        ls = ttk.Scrollbar(left_outer, orient="vertical", command=lc.yview)
        lc.configure(yscrollcommand=ls.set)
        ls.pack(side="right", fill="y"); lc.pack(side="left", fill="both", expand=True)
        self.left_inner = tk.Frame(lc, bg="#f0f0f0")
        lc.create_window((0, 0), window=self.left_inner, anchor="nw")
        self.left_inner.bind("<Configure>",
            lambda e: lc.configure(scrollregion=lc.bbox("all")))
        self._build_left(self.left_inner)

        rp = ttk.PanedWindow(paned, orient="horizontal")
        paned.add(rp, weight=1)

        map_frm = tk.LabelFrame(rp, text=" Map View ",
                                  font=("Segoe UI", 9, "bold"),
                                  relief="sunken", bd=2)
        rp.add(map_frm, weight=3)

        self.canvas = tk.Canvas(map_frm, bg="white",
                                 cursor="crosshair", highlightthickness=0)
        hsc = ttk.Scrollbar(map_frm, orient="horizontal", command=self.canvas.xview)
        vsc = ttk.Scrollbar(map_frm, orient="vertical",   command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=hsc.set, yscrollcommand=vsc.set)
        hsc.pack(side="bottom", fill="x"); vsc.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)
        self._draw_placeholder()

        raw_frm = tk.LabelFrame(rp,
                                  text=" Raw Output  (editable — updates map live) ",
                                  font=("Segoe UI", 9, "bold"),
                                  relief="sunken", bd=2)
        rp.add(raw_frm, weight=2)

        self.raw_text = tk.Text(raw_frm, wrap="none",
                                 font=("Courier New", 7),
                                 bg="white", state="disabled")
        rh = ttk.Scrollbar(raw_frm, orient="horizontal", command=self.raw_text.xview)
        rv = ttk.Scrollbar(raw_frm, orient="vertical",   command=self.raw_text.yview)
        self.raw_text.configure(xscrollcommand=rh.set, yscrollcommand=rv.set)
        rh.pack(side="bottom", fill="x"); rv.pack(side="right", fill="y")
        self.raw_text.pack(fill="both", expand=True)

    def _build_left(self, parent):
        # Input File
        g1 = tk.LabelFrame(parent, text=" Input File ",
                            font=("Segoe UI",9,"bold"),
                            bg="#f0f0f0", relief="groove", bd=2)
        g1.pack(fill="x", padx=6, pady=(8,4))
        self.file_lbl = tk.Label(g1, text="No file selected",
                                  font=("Segoe UI",8), bg="#f0f0f0",
                                  fg="#555", wraplength=185, justify="left")
        self.file_lbl.pack(fill="x", padx=8, pady=(4,2))
        tk.Button(g1, text="Browse...", command=self._open_file,
                  font=("Segoe UI",9), relief="raised", bd=2,
                  padx=6, pady=2).pack(anchor="w", padx=8, pady=(2,6))

        # Designs Detected
        g2 = tk.LabelFrame(parent, text=" Designs Detected ",
                            font=("Segoe UI",9,"bold"),
                            bg="#f0f0f0", relief="groove", bd=2)
        g2.pack(fill="x", padx=6, pady=4)

        hdr = tk.Frame(g2, bg="#f0f0f0"); hdr.pack(fill="x", padx=6, pady=(2,0))
        tk.Label(hdr, text="Design", font=("Courier New",7,"bold"),
                 bg="#f0f0f0", width=8, anchor="w").pack(side="left")
        tk.Label(hdr, text="Die Count", font=("Courier New",7,"bold"),
                 bg="#f0f0f0", anchor="w").pack(side="left")
        ttk.Separator(g2, orient="horizontal").pack(fill="x", padx=6, pady=1)

        self.design_list_frame = tk.Frame(g2, bg="#f0f0f0")
        self.design_list_frame.pack(fill="x", padx=4, pady=2)
        tk.Label(self.design_list_frame, text="(open a file to scan)",
                 font=("Segoe UI",8,"italic"),
                 bg="#f0f0f0", fg="#888").pack(anchor="w", padx=4)

        ttk.Separator(g2, orient="horizontal").pack(fill="x", padx=6, pady=2)
        br = tk.Frame(g2, bg="#f0f0f0"); br.pack(fill="x", padx=6, pady=(0,6))
        tk.Button(br, text="Select All",  command=self._select_all,
                  font=("Segoe UI",8), relief="raised", bd=2,
                  padx=4, pady=1).pack(side="left", padx=2)
        tk.Button(br, text="Select None", command=self._select_none,
                  font=("Segoe UI",8), relief="raised", bd=2,
                  padx=4, pady=1).pack(side="left", padx=2)

        # Actions
        g3 = tk.LabelFrame(parent, text=" Actions ",
                            font=("Segoe UI",9,"bold"),
                            bg="#f0f0f0", relief="groove", bd=2)
        g3.pack(fill="x", padx=6, pady=4)
        for text, cmd, attr in [
            ("Convert",        self._do_convert, "convert_btn"),
            ("Export...",      self._export,     "export_btn"),
            ("Reset",          self._reset,       None),
        ]:
            b = tk.Button(g3, text=text, command=cmd,
                          font=("Segoe UI",9), relief="raised", bd=2,
                          width=22, pady=3,
                          state="disabled" if attr == "export_btn" else "normal")
            b.pack(fill="x", padx=6, pady=2)
            if attr: setattr(self, attr, b)

        # Statistics
        g4 = tk.LabelFrame(parent, text=" Statistics ",
                            font=("Segoe UI",9,"bold"),
                            bg="#f0f0f0", relief="groove", bd=2)
        g4.pack(fill="x", padx=6, pady=4)
        self.stat_vars = {}
        for lbl in ("Dies Found:","Map Rows:","Map Cols:",
                     "File Size:","Load Time:","Selected Die:"):
            fr = tk.Frame(g4, bg="#f0f0f0"); fr.pack(fill="x", padx=6, pady=1)
            tk.Label(fr, text=lbl, font=("Segoe UI",8),
                     bg="#f0f0f0", width=13, anchor="w").pack(side="left")
            v = tk.StringVar(value="—"); self.stat_vars[lbl] = v
            tk.Label(fr, textvariable=v, font=("Segoe UI",8,"bold"),
                     bg="#f0f0f0", fg="#003399").pack(side="left")

        # Legend
        g5 = tk.LabelFrame(parent, text=" Legend ",
                            font=("Segoe UI",9,"bold"),
                            bg="#f0f0f0", relief="groove", bd=2)
        g5.pack(fill="x", padx=6, pady=(4,10))
        for colour, label in [
            ("#00AA00", '"1"  Selected design'),
            ("#646464", '"X"  Edge / drop-in'),
            ("#FFFFFF", '"."  Outer null'),
            ("#D2D2D2", '"x"  Opted-out'),
        ]:
            fr = tk.Frame(g5, bg="#f0f0f0"); fr.pack(fill="x", padx=6, pady=2)
            tk.Canvas(fr, width=16, height=13, bg=colour,
                      relief="solid", bd=1, highlightthickness=0).pack(side="left", padx=(0,6))
            tk.Label(fr, text=label, font=("Segoe UI",8),
                     bg="#f0f0f0").pack(side="left")

    # ── status bar ────────────────────────────────────────────────────────────
    def _build_statusbar(self):
        sb = tk.Frame(self, relief="sunken", bd=1, bg="#f0f0f0")
        sb.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(sb, textvariable=self.status_var,
                 font=("Segoe UI",8), bg="#f0f0f0",
                 anchor="w").pack(side="left", padx=6, pady=2)
        self.progress = ttk.Progressbar(sb, mode="indeterminate", length=150)

    def _set_status(self, msg): self.status_var.set(msg)

    # ── load ─────────────────────────────────────────────────────────────────
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Open Wafer Map",
            filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")])
        if not path: return
        self.filepath = path
        self.file_lbl.config(text=os.path.basename(path))
        self._set_status(f"Loading {os.path.basename(path)} …")
        self.progress.pack(side="right", padx=6, pady=2); self.progress.start(10)
        threading.Thread(target=self._load_thread, daemon=True).start()

    def _load_thread(self):
        t0 = time.time()
        try:
            wb   = openpyxl.load_workbook(self.filepath, data_only=True)
            best = max(wb.sheetnames,
                       key=lambda s: wb[s].max_row * wb[s].max_column)
            ws   = wb[best]
            grid = [list(row) for row in ws.iter_rows(values_only=True)]

            counts = Counter()
            for row in grid:
                for cell in row:
                    if cell not in (None, '.', 'X', 'x'):
                        counts[str(cell)] += 1
            designs = sorted(counts.keys(),
                key=lambda d: (re.sub(r'[^0-9]','',d) or '0',
                               re.sub(r'[0-9]','',d)))

            cm_hex  = {d: PALETTE_HEX[i%len(PALETTE_HEX)]  for i,d in enumerate(designs)}
            cm_rgb  = {d: PALETTE_RGB[i%len(PALETTE_RGB)]   for i,d in enumerate(designs)}
            cm_xlsx = {d: PALETTE_XLSX[i%len(PALETTE_XLSX)] for i,d in enumerate(designs)}

            rows = len(grid)
            cols = max(len(r) for r in grid)
            design_idx = {d: i+1 for i,d in enumerate(designs)}
            enc = np.zeros((rows, cols), dtype=np.int16)
            for r, row in enumerate(grid):
                for c, val in enumerate(row):
                    if val is None or str(val) == '.': pass
                    elif str(val) == 'X': enc[r,c] = -1
                    elif str(val) == 'x': enc[r,c] = -2
                    else: enc[r,c] = design_idx.get(str(val), 0)

            load_time = time.time()-t0
            fsize = os.path.getsize(self.filepath)
            fsize_s = (f"{fsize/1024:.1f} KB" if fsize<1048576 else f"{fsize/1048576:.1f} MB")

            self.after(0, lambda: self._on_load_done(
                grid, enc, best, designs, dict(counts),
                cm_hex, cm_rgb, cm_xlsx, rows, cols, fsize_s, load_time))
        except Exception as e:
            self.after(0, lambda: self._on_load_error(str(e)))

    def _on_load_done(self, grid, enc, sheet_name, designs, counts,
                      cm_hex, cm_rgb, cm_xlsx, rows, cols, fsize_s, load_time):
        self.progress.stop(); self.progress.pack_forget()
        self.grid_data = grid; self.enc = enc; self.source_sheet = sheet_name
        self.designs = designs; self.design_counts = counts
        self.colour_map_hex = cm_hex; self.colour_map_rgb = cm_rgb
        self.colour_map_xlsx = cm_xlsx
        self.selected_designs.clear(); self.zoom = 1.0

        self.stat_vars["Dies Found:"].set(f"{sum(counts.values()):,}")
        self.stat_vars["Map Rows:"].set(str(rows))
        self.stat_vars["Map Cols:"].set(str(cols))
        self.stat_vars["File Size:"].set(fsize_s)
        self.stat_vars["Load Time:"].set(f"{load_time:.2f}s")
        self.stat_vars["Selected Die:"].set("0")
        self._set_status(
            f"Loaded — {len(designs)} designs  ({rows}×{cols})  "
            f"— Select design(s) then Convert")
        self._populate_designs()
        self._render_map()

    def _on_load_error(self, msg):
        self.progress.stop(); self.progress.pack_forget()
        self._set_status(f"Error: {msg}")
        messagebox.showerror("Load Error", msg)

    # ── design list ───────────────────────────────────────────────────────────
    def _populate_designs(self):
        for w in self.design_list_frame.winfo_children(): w.destroy()
        self.design_vars = {}
        for i, design in enumerate(self.designs):
            colour = PALETTE_HEX[i % len(PALETTE_HEX)]
            var = tk.BooleanVar(value=False)
            self.design_vars[design] = var
            fr = tk.Frame(self.design_list_frame, bg="#f0f0f0")
            fr.pack(fill="x", pady=1)
            tk.Canvas(fr, width=14, height=13, bg=colour,
                      relief="solid", bd=1, highlightthickness=0).pack(side="left", padx=(2,4))
            count = self.design_counts.get(design, 0)
            tk.Checkbutton(fr,
                           text=f"{str(design):<8}  {count:>8,}",
                           font=("Courier New",8), variable=var,
                           bg="#f0f0f0", activebackground="#f0f0f0",
                           command=self._on_sel_change).pack(side="left")

    def _on_sel_change(self):
        self.selected_designs = {d for d,v in self.design_vars.items() if v.get()}
        total = sum(self.design_counts.get(d,0) for d in self.selected_designs)
        self.stat_vars["Selected Die:"].set(f"{total:,}")
        en = "normal" if self.selected_designs else "disabled"
        self.export_btn.config(state=en); self.tb_export.config(state=en)
        self._schedule_redraw()

    def _select_all(self):
        for v in self.design_vars.values(): v.set(True); self._on_sel_change()
    def _select_none(self):
        for v in self.design_vars.values(): v.set(False); self._on_sel_change()

    # ── fast image rendering ──────────────────────────────────────────────────
    def _build_pil_image(self):
        enc = self.enc; h,w = enc.shape
        img = np.full((h,w,3), 255, dtype=np.uint8)
        img[enc == -1] = X_RGB
        img[enc == -2] = x_RGB
        for i,d in enumerate(self.designs):
            mask = enc == (i+1)
            img[mask] = SEL_RGB if d in self.selected_designs else self.colour_map_rgb.get(d,(150,150,150))
        return Image.fromarray(img, 'RGB')

    def _render_map(self):
        if self.enc is None: return
        cw = max(self.canvas.winfo_width(),  200)
        ch = max(self.canvas.winfo_height(), 200)
        h,w = self.enc.shape
        scale = min(cw/w, ch/h)
        disp_w = max(int(w * scale * self.zoom), 1)
        disp_h = max(int(h * scale * self.zoom), 1)
        disp_w = min(disp_w, w*8); disp_h = min(disp_h, h*8)

        pil_img  = self._build_pil_image()
        pil_disp = pil_img.resize((disp_w, disp_h), Image.NEAREST)

        if self.show_grid.get() and (disp_w/w) >= 4:
            from PIL import ImageDraw
            draw = ImageDraw.Draw(pil_disp)
            step = int(disp_w/w)
            for x in range(0, disp_w, step): draw.line([(x,0),(x,disp_h)], fill=(200,200,200), width=1)
            for y in range(0, disp_h, step): draw.line([(0,y),(disp_w,y)], fill=(200,200,200), width=1)

        self._photo = ImageTk.PhotoImage(pil_disp)
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0,0,disp_w,disp_h))
        self.canvas.create_image(0,0,anchor="nw",image=self._photo)
        self.zoom_lbl.config(text=f"Zoom: {int(self.zoom*100)}%")

    def _schedule_redraw(self):
        if self._redraw_pending: return
        self._redraw_pending = True; self.after(40, self._do_redraw)
    def _do_redraw(self):
        self._redraw_pending = False
        if self.enc is not None: self._render_map()

    # ── convert + raw output ─────────────────────────────────────────────────
    def _do_convert(self):
        if not self.grid_data:
            messagebox.showinfo("No File","Please open a file first."); return
        self._render_map(); self._update_raw()
        self._set_status("Converted — select design(s) and click Export...")

    def _update_raw(self):
        sel = self.selected_designs; lines = []
        for row in self.grid_data:
            parts = []
            for val in row:
                s = str(val) if val is not None else '.'
                if s=='.': parts.append('.')
                elif s=='X': parts.append('X')
                elif s in sel: parts.append('1')
                else: parts.append('x')
            lines.append(''.join(parts))
        self.raw_text.config(state="normal"); self.raw_text.delete("1.0","end")
        preview = lines[:200]
        if len(lines)>200: preview.append(f"... ({len(lines)-200} more rows) ...")
        self.raw_text.insert("1.0",'\n'.join(preview))
        self.raw_text.config(state="disabled")

    # ── zoom ─────────────────────────────────────────────────────────────────
    def _zoom_in(self):    self.zoom=min(self.zoom*1.5,16.0); self._schedule_redraw()
    def _zoom_out(self):   self.zoom=max(self.zoom/1.5,0.05); self._schedule_redraw()
    def _zoom_fit(self):   self.zoom=1.0; self._schedule_redraw()
    def _zoom_reset(self): self.zoom=1.0; self._schedule_redraw()

    def _draw_placeholder(self):
        self.canvas.delete("all")
        self.canvas.create_text(350,280,
            text="Open a wafer map file and click Convert",
            fill="#888888", font=("Segoe UI",11))

    # ── export ────────────────────────────────────────────────────────────────
    def _export(self):
        if not self.selected_designs:
            messagebox.showwarning("No Selection",
                "Please select at least one design first."); return

        total_die = sum(self.design_counts.get(d,0) for d in self.selected_designs)
        dlg = ExportDialog(self, self.selected_designs, total_die)
        self.wait_window(dlg)
        if dlg.result is None: return

        opts  = dlg.result
        fmt   = opts["fmt"]
        label = "_".join(sorted(self.selected_designs))

        if fmt == "xlsx":
            default = f"Design_{label}_R1.00.xlsx"
            path = filedialog.asksaveasfilename(
                title="Save Excel Wafer Map",
                defaultextension=".xlsx",
                initialfile=default,
                filetypes=[("Excel Workbook","*.xlsx"),("All files","*.*")])
        else:
            default = f"Design_{label}_R1.00.txt"
            path = filedialog.asksaveasfilename(
                title="Save Text Wafer Map",
                defaultextension=".txt",
                initialfile=default,
                filetypes=[("Text files","*.txt"),("All files","*.*")])

        if not path: return

        self._set_status("Exporting …")
        self.progress.pack(side="right", padx=6, pady=2); self.progress.start(10)

        if fmt == "xlsx":
            threading.Thread(target=self._export_xlsx,
                             args=(path, opts["add_colour"]), daemon=True).start()
        else:
            threading.Thread(target=self._export_txt,
                             args=(path, opts["line_ending"]), daemon=True).start()

    # ── xlsx export ───────────────────────────────────────────────────────────
    def _export_xlsx(self, path, add_colour):
        try:
            sel  = self.selected_designs
            grid = self.grid_data

            # Pre-build fill objects once (not per cell)
            FILL_SEL  = PatternFill("solid", fgColor=SEL_XLSX)
            FILL_X    = PatternFill("solid", fgColor=X_XLSX)
            FILL_x    = PatternFill("solid", fgColor=x_XLSX)
            design_fills = {d: PatternFill("solid", fgColor=self.colour_map_xlsx.get(d,"AAAAAA"))
                            for d in self.designs}

            wb  = openpyxl.Workbook(write_only=True)
            ws  = wb.create_sheet(self.source_sheet or "MAP")

            # Narrow column widths to match typical wafer map view
            # write_only doesn't support column_dimensions directly — use sheet properties
            die_count = 0

            for row in grid:
                out_row = []
                for val in row:
                    s = str(val) if val is not None else '.'
                    cell = WriteOnlyCell(ws)

                    if s == '.':
                        cell.value = '.'
                    elif s == 'X':
                        cell.value = 'X'
                        if add_colour: cell.fill = FILL_X
                    elif s in sel:
                        cell.value = 1
                        die_count += 1
                        if add_colour: cell.fill = FILL_SEL
                    else:
                        cell.value = 'x'
                        if add_colour: cell.fill = FILL_x

                    out_row.append(cell)
                ws.append(out_row)

            # Add a Die_Counts summary sheet
            ws_counts = wb.create_sheet("Die_Counts")
            ws_counts.append(["Design", "Die Count", "Selected"])
            for d in self.designs:
                cnt = self.design_counts.get(d, 0)
                is_sel = "YES" if d in sel else ""
                ws_counts.append([d, cnt, is_sel])
            ws_counts.append([])
            ws_counts.append(["TOTAL SELECTED", die_count, ""])

            wb.save(path)
            self.after(0, lambda: self._on_export_done(path, die_count, "xlsx"))
        except Exception as e:
            self.after(0, lambda: self._on_export_error(str(e)))

    # ── txt export ────────────────────────────────────────────────────────────
    def _export_txt(self, path, line_ending):
        try:
            sel = self.selected_designs
            nl  = "\r\n" if line_ending == "crlf" else "\n"
            lines = []; die_count = 0
            for row in self.grid_data:
                parts = []
                for val in row:
                    s = str(val) if val is not None else '.'
                    if s=='.':       parts.append('.')
                    elif s=='X':     parts.append('X')
                    elif s in sel:   parts.append('1'); die_count+=1
                    else:            parts.append('x')
                lines.append(''.join(parts))
            with open(path, 'w', encoding='utf-8', newline='') as f:
                f.write(nl.join(lines))
            self.after(0, lambda: self._on_export_done(path, die_count, "txt"))
        except Exception as e:
            self.after(0, lambda: self._on_export_error(str(e)))

    def _on_export_done(self, path, die_count, fmt):
        self.progress.stop(); self.progress.pack_forget()
        fname = os.path.basename(path)
        self._set_status(f"Exported: {fname}  —  {die_count:,} die as Bin 1")
        if fmt == "xlsx":
            messagebox.showinfo("Export Complete",
                f"Saved:  {fname}\n\n"
                f"Die exported as Bin 1:  {die_count:,}\n\n"
                f"Contains:\n"
                f"  • MAP sheet — full wafer grid\n"
                f"  • Die_Counts sheet — design summary\n\n"
                f"For assembly upload: use File › Export › Text (.txt)\n"
                f"and add site header/footer before sending.")
        else:
            messagebox.showinfo("Export Complete",
                f"Saved:  {fname}\n\n"
                f"Die exported as Bin 1:  {die_count:,}\n\n"
                f"Next step: add your site header/footer,\n"
                f"then save as Unicode UTF-8.")

    def _on_export_error(self, msg):
        self.progress.stop(); self.progress.pack_forget()
        self._set_status(f"Export error: {msg}")
        messagebox.showerror("Export Error", msg)

    # ── reset ─────────────────────────────────────────────────────────────────
    def _reset(self):
        self.filepath=None; self.grid_data=[]; self.enc=None
        self.designs=[]; self.design_counts={}; self.selected_designs.clear()
        self.design_vars={}; self.zoom=1.0
        self.file_lbl.config(text="No file selected")
        for w in self.design_list_frame.winfo_children(): w.destroy()
        tk.Label(self.design_list_frame, text="(open a file to scan)",
                 font=("Segoe UI",8,"italic"),
                 bg="#f0f0f0", fg="#888").pack(anchor="w", padx=4)
        for k in self.stat_vars: self.stat_vars[k].set("—")
        self.export_btn.config(state="disabled"); self.tb_export.config(state="disabled")
        self.raw_text.config(state="normal"); self.raw_text.delete("1.0","end")
        self.raw_text.config(state="disabled")
        self._draw_placeholder(); self._set_status("Ready")

    def _on_resize(self, event=None):
        if self.enc is not None: self.after(120, self._do_redraw)


if __name__ == "__main__":
    app = WaferMapTool()
    app.bind("<Configure>", app._on_resize)
    app.mainloop()  