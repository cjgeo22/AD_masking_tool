import os
import json
import argparse
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
from PIL.ImageDraw import floodfill
from openpyxl import Workbook
import csv

from style_config import (
    COLUMN_WEIGHTS,
    ROW_WEIGHTS,
    PAD_SMALL, PAD_MEDIUM, PAD_LARGE,
    BUTTON_FONT, LABEL_FONT, ENTRY_FONT, LARGE_TEXT,
    BUTTON_PAD,
    CANVAS_BG, CANVAS_CURSOR,
    TS_DISPLAY_FMT, TS_FILENAME_FMT,
    PEN1_COLOR, PEN2_COLOR, PEN3_COLOR, PEN4_COLOR, PEN5_COLOR
)

class DrawApp:
    def __init__(self, input_dir, output_dir, config_path, dataset):
        self.input_dir   = input_dir
        self.output_dir  = output_dir

        # Load defect list from JSON config
        cfg = json.load(open(config_path))
        self.defects = cfg.get(dataset, cfg.get("default", []))

        # Gather image paths
        exts = {".jpg", ".jpeg", ".png", ".bmp", ".tiff"}
        self.image_paths = sorted(
            os.path.join(input_dir, f)
            for f in os.listdir(input_dir)
            if os.path.splitext(f.lower())[1] in exts
        )
        if not self.image_paths:
            raise RuntimeError(f"No images in {input_dir}")

        # Annotation state
        self.idx         = 0
        self.mode        = "draw"
        self.history     = []
        self.redo_stack  = []
        self.pan_x = self.pan_y = 0
        self.zoom        = 1.0
        self.drawing     = False   # track whether a draw‐stroke is in progress

        # Pen state
        self.current_pen = "pen1"
        self.pen_color   = PEN1_COLOR   # (R,G,B,A) tuple
        self.pen_width   = 3

        # Build UI and load first image
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Load first image and metadata
        self._load_image()
        # After loading, adjust initial zoom so entire image fits in canvas
        self.root.update()
        self._fit_image_to_canvas()
        self._show_image()

        self.root.mainloop()

    def _wrap_cmd(self, fn):
        #Wrapper that first disables 'fill mode' before running any action
        def wrapped(*args, **kwargs):
            self.fill_var.set(False)
            return fn(*args, **kwargs)
        return wrapped

    def _build_ui(self):
        #Construct all Tkinter widgets
        self.root = tk.Tk()
        self.root.title("Anomaly Detection Masking Tool")

        # Configure styles
        style = ttk.Style(self.root)
        style.configure("Fill.Toolbutton", padding=(5, 5))
        style.configure(
            "Fill.Toolbutton",
            background="#d9d9d9",    
            foreground="#000000",    
            borderwidth=2,           
            relief="raised",         
            font=("Arial",14)        
        )
        style.map(
            "Fill.Toolbutton",
            background=[
                ("active", "#bfbfbf"),    
                ("selected", "#2c69c4"),  
            ],
            relief=[
                ("selected", "sunken"),   
                ("!selected", "raised"),  
            ],
            foreground=[
                ("selected", "#000000"),  
                ("!selected", "#000000")
            ],
            bordercolor=[
                ("active", "#8fa9d8"),    
                ("!active", "#a4a4a4")
            ]
        )

        style.configure("App.TButton", font=BUTTON_FONT)
        style.configure("App.TLabel",  font=LABEL_FONT)
        style.configure("App.TEntry",  font=ENTRY_FONT)

        # Layout weights
        for col, w in COLUMN_WEIGHTS.items():
            self.root.columnconfigure(col, weight=w)
        for row, w in ROW_WEIGHTS.items():
            self.root.rowconfigure(row, weight=w)

        # ——— Canvas for drawing ———
        self.canvas = tk.Canvas(self.root, bg=CANVAS_BG, cursor=CANVAS_CURSOR)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        # Bind left‐click press, drag, and release:
        self.canvas.bind("<ButtonPress-1>",    self._on_left_press)
        self.canvas.bind("<B1-Motion>",        self._on_left_move)
        self.canvas.bind("<ButtonRelease-1>",  self._on_left_release)
        # Bind right‐click (for panning)
        self.canvas.bind("<ButtonPress-3>",    self._on_right_press)
        self.canvas.bind("<B3-Motion>",        self._on_right_move)
        self.canvas.bind("<ButtonRelease-3>",  self._on_right_release)

        # ——— Control panel container with scrollbar ———
        ctrl_container = ttk.Frame(self.root)
        ctrl_container.grid(row=0, column=1, sticky="nsew", padx=PAD_MEDIUM, pady=PAD_MEDIUM)

        # Canvas for scrollbar
        self.ctrl_canvas = tk.Canvas(ctrl_container, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(ctrl_container, orient="vertical", command=self.ctrl_canvas.yview)
        self.ctrl_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.ctrl_frame = ttk.Frame(self.ctrl_canvas)
        self.ctrl_canvas.create_window((0,0), window=self.ctrl_frame, anchor="nw")
        self.ctrl_canvas.pack(side="left", fill="both", expand=True)
        self.ctrl_frame.bind("<Configure>", lambda e: self.ctrl_canvas.configure(scrollregion=self.ctrl_canvas.bbox("all")))

        # ——— Variables ———
        self.note_var      = tk.StringVar(master=self.root)
        self.inspector_var = tk.StringVar(master=self.root)
        self.tray_var      = tk.StringVar(master=self.root)
        self.fill_var      = tk.BooleanVar(master=self.root)
        self.save_ts_var   = tk.StringVar(master=self.root, value="never")

        # Pen label vars (for the dropdowns PEN1–PEN5)
        self.pen_vars = {
            "pen1": tk.StringVar(master=self.root, value=""),
            "pen2": tk.StringVar(master=self.root, value=""),
            "pen3": tk.StringVar(master=self.root, value=""),
            "pen4": tk.StringVar(master=self.root, value=""),
            "pen5": tk.StringVar(master=self.root, value=""),
        }

        # ——— Navigation Group ———
        nav_frame = ttk.LabelFrame(self.ctrl_frame, text="Main Controls", padding=(50,5))
        nav_frame.grid(row=0, column=0, sticky="nsew")
        actions = [
            ("◀ Prev",   self.prev_image),
            ("Next ▶",   self.next_image),
            ("Clear",    self.clear_all),
            ("Undo",     self.undo),
            ("Redo",     self.redo),
            ("Draw",     lambda: self.set_mode("draw")),
            ("Erase",    lambda: self.set_mode("erase")),
            ("Save",     self.save),
        ]
        for i, (txt, fn) in enumerate(actions):
            r, c = divmod(i, 2)
            cmd = fn if txt in ("Draw", "Erase") else self._wrap_cmd(fn)
            ttk.Button(nav_frame, text=txt, command=cmd, style="App.TButton", width=25) \
               .grid(row=r, column=c, sticky="w", padx=BUTTON_PAD[0], pady=BUTTON_PAD[1])
        # Configure equal expansion
        for c in range(2):
            nav_frame.columnconfigure(c, weight=1)

        # ——— Zoom Group ———
        zoom_frame = ttk.LabelFrame(self.ctrl_frame, text="Zoom", padding=(10,5))
        zoom_frame.grid(row=1, column=0, sticky="ew")
        for symbol, factor in [("−", 0.9), ("+", 1.1)]:
            ttk.Button(
                zoom_frame,
                text=symbol,
                command=self._wrap_cmd(lambda f=factor: self._set_zoom(f)),
                style="App.TButton",
                width=15
            ).pack(side="top", expand=False, fill="none")

        # ——— Metadata Group ———
        meta_frame = ttk.LabelFrame(self.ctrl_frame, text="Information", padding=PAD_SMALL)
        meta_frame.grid(row=2, column=0, sticky="ew", pady=PAD_SMALL)
        # Inspector
        ttk.Label(meta_frame, text="Inspector:", style="App.TLabel") \
           .grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(meta_frame, textvariable=self.inspector_var, style="App.TEntry") \
           .grid(row=1, column=1, sticky="w", padx=5, pady=5)
        # Tray/Directory
        ttk.Label(meta_frame, text="Tray/Directory:", style="App.TLabel") \
           .grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(meta_frame, textvariable=self.tray_var, style="App.TEntry") \
           .grid(row=2, column=1, sticky="w", padx=5, pady=5)
        # Note field
        ttk.Label(meta_frame, text="Note:", style="App.TLabel") \
           .grid(row=0, column=0, sticky="w", padx=5)
        self.note_text = tk.Text(meta_frame, wrap="word", font=ENTRY_FONT, height=3, width=65)
        self.note_text.grid(row=0, column=1, sticky="w", padx=5)
        self.note_text.bind("<KeyRelease>", self._adjust_note_height)
        meta_frame.rowconfigure(0, weight=1)

        # Last-saved timestamp
        ttk.Label(meta_frame, text="Last saved:", style="App.TLabel") \
           .grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(meta_frame, textvariable=self.save_ts_var, style="App.TLabel") \
           .grid(row=3, column=1, sticky="w", padx=5, pady=5)
        # Configure column expansion
        meta_frame.columnconfigure(1, weight=1)

        # ——— Defects Group ———
        def_frame = ttk.LabelFrame(self.ctrl_frame, text="Categories", padding=PAD_SMALL)
        def_frame.grid(row=3, column=0, sticky="ew", pady=PAD_SMALL)
        self.defect_vars = {}
        for i, d in enumerate(self.defects):
            v = tk.BooleanVar(master=self.root)
            ttk.Checkbutton(def_frame, text=d, variable=v) \
               .grid(row=i, column=0, sticky="w", padx=PAD_SMALL)
            self.defect_vars[d] = v

        # ——— Fill Group ———
        fill_frame = ttk.LabelFrame(self.ctrl_frame, text="Fill", padding=(225,5))
        fill_frame.grid(row=4, column=0, sticky="ew")
        # Fill-mode toggle
        ttk.Checkbutton(fill_frame, text="PEN FILL MODE", variable=self.fill_var, style="Fill.Toolbutton") \
           .grid(row=len(self.defects), column=0, sticky="w", padx=PAD_SMALL, pady=(PAD_SMALL,0))

        # ——— Pen Labels Group ———
        pen_frame = ttk.LabelFrame(self.ctrl_frame, text="Pen Labels", padding=PAD_SMALL)
        pen_frame.grid(row=5, column=0, sticky="ew", pady=PAD_SMALL)
        for i, (key, color) in enumerate([("pen1", "Red"), ("pen2", "Blue"), ("pen3", "Green"), ("pen4", "Yellow"), ("pen5", "Magenta")]):
            ttk.Label(pen_frame, text=f"{key.upper()} label:", style="App.TLabel") \
               .grid(row=2*i, column=0, sticky="w", padx=PAD_SMALL)
            cb = ttk.Combobox(
                pen_frame,
                textvariable=self.pen_vars[key],
                values=self.defects,
                state="readonly",
                style="App.TEntry",
                width=35
            )
            cb.grid(row=2*i, column=1, sticky="w", padx=20)
            ttk.Button(
                pen_frame,
                text=f"Use {key.upper()} ({color})",
                command=self._wrap_cmd(lambda k=key: self.select_pen(k)),
                style="Button.Toolbutton"
            ).grid(row=2*i+1, column=0, columnspan=2, sticky="w", padx=50, pady=5)
        pen_frame.columnconfigure(1, weight=1)

        # ——— Export Group ———
        export_frame = ttk.LabelFrame(self.ctrl_frame, text="Export Information", padding=(210,5))
        export_frame.grid(row=6, column=0, sticky="ew", pady=PAD_SMALL)

        # Excel export button
        ttk.Button(
            export_frame,
            text="Export as XLSX",
            command=self._wrap_cmd(self.export_excel),
            style="App.TButton"
        ).pack(side="left", padx=BUTTON_PAD[0], pady=BUTTON_PAD[1])

        # CSV export button (NEW)
        ttk.Button(
            export_frame,
            text="Export as CSV",
            command=self._wrap_cmd(self.export_csv),
            style="App.TButton"
        ).pack(side="left", padx=BUTTON_PAD[0], pady=BUTTON_PAD[1])


    # ——— Helper to auto‐grow the Note Text widget’s height as lines wrap ———
    def _adjust_note_height(self, event=None):
        text = self.note_text
        line_count = int(text.index('end-1c').split('.')[0])
        new_height = max(3, line_count)
        text.configure(height=new_height)

    def select_pen(self, which):
        #Switch to a different pen (pen1…pen5).
        self.current_pen = which
        self.pen_color   = {
            "pen1": PEN1_COLOR,
            "pen2": PEN2_COLOR,
            "pen3": PEN3_COLOR,
            "pen4": PEN4_COLOR,
            "pen5": PEN5_COLOR
        }[which]
        self.pen_width = 3
        self.mode      = "draw"
        self.canvas.config(cursor="pencil" if self.mode=="draw" else "circle")

    def _load_image(self):
        #Load the current image at its original resolution (no thumbnailing).
        #If a mask already exists, force‐resize it to match exactly that resolution.
        #Load the full, original image (RGB mode)
        img = Image.open(self.image_paths[self.idx]).convert("RGB")
        self.base = img

        #Build the path to the saved mask
        name = os.path.splitext(os.path.basename(self.image_paths[self.idx]))[0]
        masks_folder = os.path.join(self.output_dir, "masks")
        os.makedirs(masks_folder, exist_ok=True)
        mask_path = os.path.join(masks_folder, f"{name}_mask.png")

        #If a mask exists, load and resize it to match `self.base.size`. If not, create a blank RGBA layer
        if os.path.isfile(mask_path):
            loaded_mask = Image.open(mask_path).convert("RGBA")
            if loaded_mask.size != self.base.size:
                # Resize (up or down) so it exactly matches the original image’s dimensions
                loaded_mask = loaded_mask.resize(self.base.size, Image.LANCZOS)
            self.layer = loaded_mask
        else:
            # No existing mask -> create a brand‐new, fully transparent RGBA layer at full resolution
            self.layer = Image.new("RGBA", self.base.size, (0, 0, 0, 0))

        #Prepare the drawing context on that layer
        self.drawer = ImageDraw.Draw(self.layer)

        #Reset undo/redo history and pan/zoom state
        self.history    = [self.layer.copy()]
        self.redo_stack = []
        self.pan_x = self.pan_y = 0
        self.zoom  = 1.0
        self.drawing = False

        #Load any saved metadata (notes, inspector, defects, pen_labels, etc.)
        meta = self._load_meta().get(os.path.basename(self.image_paths[self.idx]), {})

        # Populate text fields / checkboxes from the metadata JSON
        self.note_text.delete("1.0", "end")
        self.note_text.insert("1.0", meta.get("note", ""))
        self.inspector_var.set(meta.get("inspector", ""))
        self.tray_var.set(meta.get("tray", ""))

        for d, var in self.defect_vars.items():
            var.set(d in meta.get("defects", []))

        for k in self.pen_vars:
            self.pen_vars[k].set("")
        pens = meta.get("pen_labels", {})
        for k, lab in pens.items():
            if k in self.pen_vars:
                self.pen_vars[k].set(lab)

        self.save_ts_var.set(meta.get("last_saved", "never"))
        self._unsaved_clear = False


    def _show_image(self):
        comp = Image.alpha_composite(self.base.convert("RGBA"), self.layer)
        w, h = comp.size
        disp = comp.resize((int(w * self.zoom), int(h * self.zoom)), Image.LANCZOS)
        self.tkimg = ImageTk.PhotoImage(disp)

        self.canvas.delete("IMG")
        self.canvas.create_image(self.pan_x, self.pan_y, anchor="nw", image=self.tkimg, tags="IMG")
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _record_history(self):
        if self.history:
            self.history.append(self.layer.copy())
            if len(self.history) > 50:
                self.history.pop(0)
        self.redo_stack.clear()

    def _on_left_press(self, ev):
        # Convert canvas coords → image coords
        cx = self.canvas.canvasx(ev.x) - self.pan_x
        cy = self.canvas.canvasy(ev.y) - self.pan_y
        ix, iy = int(cx / self.zoom), int(cy / self.zoom)

        if self.fill_var.get():
            floodfill(self.layer, (ix, iy), self.pen_color)
            self._record_history()
            self._show_image()
            return

        if not self.drawing:
            self.drawing = True
            self.last_x, self.last_y = ix, iy
            self._show_image()

    def _on_left_move(self, ev):
        if self.fill_var.get() or not self.drawing:
            return
        cx = self.canvas.canvasx(ev.x) - self.pan_x
        cy = self.canvas.canvasy(ev.y) - self.pan_y
        ix, iy = int(cx / self.zoom), int(cy / self.zoom)
        if self.mode == "draw":
            color, width = (self.pen_color, self.pen_width)
        else:
            # erase mode
            color, width = ((0, 0, 0, 0), 20)
        self.drawer.line([(self.last_x, self.last_y), (ix, iy)], fill=color, width=width)
        self.last_x, self.last_y = ix, iy
        self._show_image()

    def _on_left_release(self, ev):
        if self.drawing:
            self._record_history()
        self.drawing = False

    def _on_right_press(self, ev):
        self._prx, self._pry = ev.x, ev.y

    def _on_right_move(self, ev):
        dx, dy = ev.x - self._prx, ev.y - self._pry
        self._prx, self._pry = ev.x, ev.y
        self.pan_x += dx
        self.pan_y += dy
        self._show_image()

    def _on_right_release(self, ev):
        self._prx = self._pry = None

    def set_mode(self, m):
        self.mode = m
        if m == "erase":
            self.fill_var.set(False)
        self.canvas.config(cursor="pencil" if m=="draw" else "circle")

    def _set_zoom(self, factor):
        self.zoom *= factor
        self.zoom = max(0.1, min(self.zoom, 10.0))
        self._show_image()

    def clear_all(self):
        self._unsaved_clear = True
        self.layer = Image.new("RGBA", self.base.size, (0,0,0,0))
        self.drawer = ImageDraw.Draw(self.layer)
        self.history = [self.layer.copy()]
        self.redo_stack.clear()
        self.note_text.delete("1.0", "end")
        self.inspector_var.set("")
        self.tray_var.set("")
        for v in self.defect_vars.values():
            v.set(False)
        for v in self.pen_vars.values():
            v.set("")
        self._show_image()

    def undo(self):
        if len(self.history) > 1:
            self.redo_stack.append(self.history.pop())
            self.layer = self.history[-1].copy()
            self.drawer = ImageDraw.Draw(self.layer)
            self._show_image()

    def redo(self):
        if self.redo_stack:
            img = self.redo_stack.pop()
            self.history.append(img.copy())
            self.layer = img.copy()
            self.drawer = ImageDraw.Draw(self.layer)
            self._show_image()

    def prev_image(self):
        if not self._unsaved_clear:
            self._save_current()
        else:
            self._unsaved_clear = False
        self.idx = (self.idx - 1) % len(self.image_paths)
        self._load_image()
        self._fit_image_to_canvas()
        self._show_image()

    def next_image(self):
        if not self._unsaved_clear:
            self._save_current()
        else:
            self._unsaved_clear = False
        self.idx = (self.idx + 1) % len(self.image_paths)
        self._load_image()
        self._fit_image_to_canvas()
        self._show_image()

    def _on_close(self):
        if not self._unsaved_clear:
            self._save_current()
        self.root.destroy()

    def save(self):
        self._save_current()
        now       = datetime.datetime.now()
        file_ts   = now.strftime(TS_FILENAME_FMT)
        disp_ts   = now.strftime(TS_DISPLAY_FMT)
        inspector = self.inspector_var.get().strip() or "unknown"
        tray      = self.tray_var.get().strip() or ""
        base_name = os.path.splitext(os.path.basename(self.image_paths[self.idx]))[0]
        sels = [d for d, v in self.defect_vars.items() if v.get()]
        if not sels:
            sels = ["good"]
        pen_labels = {k: self.pen_vars[k].get() for k in self.pen_vars}
        pen_to_color = {
            "pen1": PEN1_COLOR,
            "pen2": PEN2_COLOR,
            "pen3": PEN3_COLOR,
            "pen4": PEN4_COLOR,
            "pen5": PEN5_COLOR
        }
        COLOR_TOL = 10
        for defect in sels:
            outd = os.path.join(self.output_dir, defect)
            os.makedirs(outd, exist_ok=True)
            matching_colors = []
            for pen_key, label_text in pen_labels.items():
                if label_text.strip().lower() == defect.strip().lower():
                    color_tuple = pen_to_color[pen_key]
                    if len(color_tuple) > 3:
                        color_tuple = tuple(color_tuple[:3])
                    matching_colors.append(color_tuple)
            mask_rgba = Image.new("RGBA", self.base.size, (0,0,0,0))
            src = self.layer.load()
            dst = mask_rgba.load()
            w, h = self.base.size
            for x in range(w):
                for y in range(h):
                    r, g, b, a = src[x, y]
                    if a == 0:
                        continue
                    for (pr, pg, pb) in matching_colors:
                        if (abs(r - pr) <= COLOR_TOL
                          and abs(g - pg) <= COLOR_TOL
                          and abs(b - pb) <= COLOR_TOL):
                            dst[x, y] = (pr, pg, pb, a)
                            break
            rgb_mask = Image.new("RGB", self.base.size, (0,0,0))
            rgb_mask.paste(mask_rgba, mask=mask_rgba.split()[3])
            draw_legend = ImageDraw.Draw(rgb_mask)
            font = ImageFont.load_default()
            bbox = draw_legend.textbbox((0, 0), defect, font=font)
            text_h = bbox[3] - bbox[1]
            draw_legend.text(
                (5, self.base.size[1] - PAD_SMALL - text_h),
                defect,
                fill=(255, 255, 255),
                font=font
            )
            out_path = os.path.join(outd, f"{file_ts}_{inspector}_{base_name}.png")
            rgb_mask.save(out_path)
        meta_path = self._meta_path()
        meta = json.load(open(meta_path)) if os.path.isfile(meta_path) else {}
        key = os.path.basename(self.image_paths[self.idx])
        entry = meta.get(key, {})
        entry.update({
            "note": self.note_text.get("1.0", "end-1c"),
            "inspector": self.inspector_var.get(),
            "tray": self.tray_var.get(),
            "defects": [d for d, v in self.defect_vars.items() if v.get()],
            "pen_labels": {k: self.pen_vars[k].get() for k in self.pen_vars},
            "last_saved": disp_ts
        })
        meta[key] = entry
        json.dump(meta, open(meta_path, "w"), indent=2)
        self.save_ts_var.set(disp_ts)
        messagebox.showinfo("Saved", f"Masks exported to: {', '.join(sels)}")

    def _meta_path(self):
        return os.path.join(self.output_dir, "metadata.json")

    def _load_meta(self):
        if os.path.isfile(self._meta_path()):
            return json.load(open(self._meta_path()))
        return {}

    def _save_current(self):
        name = os.path.splitext(os.path.basename(self.image_paths[self.idx]))[0]
        mp = os.path.join(self.output_dir, "masks", f"{name}_mask.png")
        os.makedirs(os.path.dirname(mp), exist_ok=True)
        self.layer.save(mp)
        meta_path = self._meta_path()
        meta = json.load(open(meta_path)) if os.path.isfile(meta_path) else {}
        key = os.path.basename(self.image_paths[self.idx])
        entry = meta.get(key, {})
        entry.update({
            "note": self.note_text.get("1.0", "end-1c"),
            "inspector": self.inspector_var.get(),
            "tray": self.tray_var.get(),
            "defects": [d for d, v in self.defect_vars.items() if v.get()],
            "pen_labels": {k: self.pen_vars[k].get() for k in self.pen_vars},
            "last_saved": self.save_ts_var.get()
        })
        meta[key] = entry
        json.dump(meta, open(meta_path, "w"), indent=2)

    def export_excel(self):
        base_name = os.path.basename(self.input_dir.rstrip(os.sep))
        excel_path = os.path.join(self.output_dir, f"{base_name}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Tray Information"

        headers = ["Filename", "Tray/Directory", "Inspector", "Accept/Reject", "Defect(s)", "Notes"]
        ws.append(headers)

        meta = self._load_meta()
        for fullpath in self.image_paths:
            basefile = os.path.basename(fullpath)
            name_only, _ = os.path.splitext(basefile)
            entry = meta.get(basefile, {})

            defects_list = entry.get("defects", [])
            inspector_txt = entry.get("inspector", "").strip()
            tray_txt      = entry.get("tray", "").strip()
            note_txt      = entry.get("note", "").strip()

            if not defects_list and not inspector_txt and not tray_txt and not note_txt:
                accept_reject = "Unlabeled"
            elif not defects_list:
                accept_reject = "Accept"
            else:
                accept_reject = "Reject"

            row = [
                name_only,
                entry.get("tray", ""),
                entry.get("inspector", ""),
                accept_reject,            
                ", ".join(defects_list),
                entry.get("note", "")
            ]
            ws.append(row)

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    val = str(cell.value)
                    if len(val) > max_length:
                        max_length = len(val)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        wb.save(excel_path)
        messagebox.showinfo("Export Excel", f"Excel file saved to:\n{excel_path}")

    def export_csv(self):
        base_name = os.path.basename(self.input_dir.rstrip(os.sep))
        csv_path = os.path.join(self.output_dir, f"{base_name}.csv")
        meta = self._load_meta()

        with open(csv_path, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

            # WRITE HEADER (same as Excel)
            headers = ["Filename", "Tray/Directory", "Inspector", "Accept/Reject", "Defect(s)", "Notes"]
            writer.writerow(headers)

            for fullpath in self.image_paths:
                basefile = os.path.basename(fullpath)
                name_only, _ = os.path.splitext(basefile)
                entry = meta.get(basefile, {})

                defects_list = entry.get("defects", [])
                inspector_txt = entry.get("inspector", "").strip()
                tray_txt      = entry.get("tray", "").strip()
                note_txt      = entry.get("note", "").strip()

                # Determine Unlabeled / Accept / Reject
                if not defects_list and not inspector_txt and not tray_txt and not note_txt:
                    accept_reject = "Unlabeled"
                elif not defects_list:
                    accept_reject = "Accept"
                else:
                    accept_reject = "Reject"

                row = [
                    name_only,
                    entry.get("tray", ""),
                    entry.get("inspector", ""),
                    accept_reject,
                    ", ".join(defects_list),
                    entry.get("note", "")
                ]
                writer.writerow(row)

        messagebox.showinfo("Export CSV", f"CSV file saved to:\n{csv_path}")

    def _fit_image_to_canvas(self):
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        bw, bh = self.base.size
        fx = cw / bw
        fy = ch / bh
        self.zoom = min(fx, fy, 1.0)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_dir",  required=True)
    parser.add_argument("--output_dir", required=True)
    parser.add_argument("--config",     default="defects_config.json")
    parser.add_argument("--dataset",    default="default")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    DrawApp(
        args.input_dir,
        args.output_dir,
        args.config,
        args.dataset
    )
