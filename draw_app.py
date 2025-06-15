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
import glob
import configparser
import tkinter.filedialog as filedialog

from style_config import (
    COLUMN_WEIGHTS,
    ROW_WEIGHTS,
    PAD_SMALL,
    PAD_MEDIUM,
    PAD_LARGE,
    BUTTON_FONT,
    LABEL_FONT,
    ENTRY_FONT,
    BUTTON_PAD,
    CANVAS_BG,
    CANVAS_CURSOR,
    TS_DISPLAY_FMT,
    TS_FILENAME_FMT,
    PEN1_COLOR,
    PEN2_COLOR,
    PEN3_COLOR,
    PEN4_COLOR,
    PEN5_COLOR,
    PEN6_COLOR,
)

# ─── Configuration for user paths ─────────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.ini")
config = configparser.ConfigParser()
if not os.path.exists(CONFIG_PATH):
    config["Paths"] = {
        "input_folder": "",
        "defects_config": "defects_config.json",
        "output_subdir": "output",
    }
    with open(CONFIG_PATH, "w") as cfg_file:
        config.write(cfg_file)
else:
    config.read(CONFIG_PATH)


def save_config():
    """Save configuration to config.ini"""
    with open(CONFIG_PATH, "w") as cfg_file:
        config.write(cfg_file)

class DrawApp:
    def __init__(self, input_dir, output_dir, config_path, dataset):
        self.input_dir = input_dir
        self.output_dir = output_dir

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
        self.idx = 0
        self.mode = "draw"
        self.history = []
        self.redo_stack = []
        self.pan_x = self.pan_y = 0
        self.zoom = 1.0
        self.drawing = False

        # Pen state
        self.current_pen = "pen1"
        self.pen_color = PEN1_COLOR
        self.pen_width = 8

        # Build UI
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Load first image
        self._load_image()
        self.root.update()
        self._fit_image_to_canvas()
        self._show_image()

        self.root.mainloop()

    # ─── New methods for Settings menu ──────────────────────────────────────────
    def choose_input_folder(self):
        """Open dialog to select input folder and save to config"""
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            config["Paths"]["input_folder"] = folder
            save_config()
            messagebox.showinfo("Config Saved", f"Input folder set to:\n{folder}")

    def choose_defects_config(self):
        """Open dialog to select defects JSON and save to config"""
        fpath = filedialog.askopenfilename(
            title="Select Defects JSON", filetypes=[("JSON", "*.json")]
        )
        if fpath:
            config["Paths"]["defects_config"] = fpath
            save_config()
            messagebox.showinfo("Config Saved", f"Defects config set to:\n{fpath}")


    def _wrap_cmd(self, fn):
        # Wrapper to disable fill mode before executing the given function
        def wrapped(*args, **kwargs):
            self.fill_var.set(False)
            return fn(*args, **kwargs)

        return wrapped

    def _build_ui(self):
        # Root window
        self.root = tk.Tk()
        self.root.title("AnnoMate")
        self.root.iconbitmap(os.path.join(os.path.dirname(__file__), "imgs", "AnnoMate.ico"))

        # ─── Settings menu for configuration ─────────────────────────────────
        menubar = tk.Menu(self.root)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(
            label="Input Folder…", command=self.choose_input_folder
        )
        settings_menu.add_command(
            label="Defects JSON…", command=self.choose_defects_config
        )
        menubar.add_cascade(label="Settings", menu=settings_menu)
        self.root.config(menu=menubar)

        # Styles
        style = ttk.Style(self.root)
        style.configure("App.TButton", font=BUTTON_FONT)
        style.configure("App.TLabel", font=LABEL_FONT)
        style.configure("App.TEntry", font=ENTRY_FONT)
        style.configure(
            "Fill.Toolbutton",
            padding=(5, 5),
            background="#d9d9d9",
            foreground="#000000",
            borderwidth=2,
            relief="raised",
            font=("Arial", 14),
        )
        style.map(
            "Fill.Toolbutton",
            background=[("active", "#bfbfbf"), ("selected", "#2c69c4")],
            relief=[("selected", "sunken"), ("!selected", "raised")],
            foreground=[("selected", "#000000"), ("!selected", "#000000")],
        )

        # Layout weights
        for col, w in COLUMN_WEIGHTS.items():
            self.root.columnconfigure(col, weight=w)
        for row, w in ROW_WEIGHTS.items():
            self.root.rowconfigure(row, weight=w)

        # Canvas for image and drawing
        # Paned window to allow resizing of the control panel
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.grid(row=0, column=0, columnspan=2, sticky="nsew")
        self.canvas = tk.Canvas(self.root, bg=CANVAS_BG, cursor=CANVAS_CURSOR)
        self.paned.add(self.canvas)
        self.canvas.bind("<ButtonPress-1>", self._on_left_press)
        self.canvas.bind("<B1-Motion>", self._on_left_move)
        self.canvas.bind("<ButtonRelease-1>", self._on_left_release)
        self.canvas.bind("<ButtonPress-3>", self._on_right_press)
        self.canvas.bind("<B3-Motion>", self._on_right_move)
        self.canvas.bind("<ButtonRelease-3>", self._on_right_release)
        self.canvas.bind(
            "<MouseWheel>", lambda ev: self._wrap_cmd(self._on_mousewheel)(ev)
        )
        self.canvas.bind(
            "<Button-4>", lambda ev: self._wrap_cmd(lambda e: self._set_zoom(1.1))(ev)
        )
        self.canvas.bind(
            "<Button-5>", lambda ev: self._wrap_cmd(lambda e: self._set_zoom(0.9))(ev)
        )

        # Control panel with scrollbar
        # Control panel container now lives in the PanedWindow
        ctrl_container = ttk.Frame(self.paned, padding=(0, 0))
        # Logo
        logo_path = os.path.join(os.path.dirname(__file__), "imgs", "AnnoMate.png")
        try:
            logo_img = Image.open(logo_path)
            logo_img = logo_img.resize((100, 100), Image.LANCZOS)
            self.logo_tk = ImageTk.PhotoImage(logo_img)
            logo_label = ttk.Label(ctrl_container, image=self.logo_tk)
            logo_label.pack(side="top", anchor="center", pady=(0, PAD_SMALL))
        except Exception as e:
            print(f"Failed to load logo: {e}")

        self.paned.add(ctrl_container)
        self.ctrl_canvas = tk.Canvas(
            ctrl_container, borderwidth=0, highlightthickness=0
        )
        vsb = ttk.Scrollbar(
            ctrl_container, orient="vertical", command=self.ctrl_canvas.yview
        )
        hsb = ttk.Scrollbar(
            ctrl_container, orient=tk.HORIZONTAL, command=self.ctrl_canvas.xview
        )
        self.ctrl_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.ctrl_frame = ttk.Frame(self.ctrl_canvas)
        self.ctrl_canvas.create_window((0, 0), window=self.ctrl_frame, anchor="nw")
        self.ctrl_canvas.pack(side="top", fill="both", expand=True, pady=(0,0))
        self.ctrl_frame.bind(
            "<Configure>",
            lambda e: self.ctrl_canvas.configure(
                scrollregion=self.ctrl_canvas.bbox("all")
            ),
        )
        # ─── enable scroll-wheel on control panel ─────────────────────────────────

        # when mouse enters the control frame, bind all wheel events to our handler
        self.ctrl_frame.bind(
            "<Enter>",
            lambda e: self.ctrl_canvas.bind_all(
                "<MouseWheel>", self._on_control_panel_wheel
            ),
        )
        # when it leaves, unbind so wheel goes back to default
        self.ctrl_frame.bind(
            "<Leave>", lambda e: self.ctrl_canvas.unbind_all("<MouseWheel>")
        )

        # (Optional for Linux scrollpads)
        self.ctrl_frame.bind(
            "<Enter>",
            lambda e: (
                self.ctrl_canvas.bind_all("<Button-4>", self._on_control_panel_wheel),
                self.ctrl_canvas.bind_all("<Button-5>", self._on_control_panel_wheel),
            ),
        )
        self.ctrl_frame.bind(
            "<Leave>",
            lambda e: (
                self.ctrl_canvas.unbind_all("<Button-4>"),
                self.ctrl_canvas.unbind_all("<Button-5>"),
            ),
        )

        # ─── global wheel binding ──────────────────────────────────────────────────
        # anywhere in the window, catch all wheel events...
        self.root.bind_all("<MouseWheel>", self._on_global_wheel)
        self.root.bind_all("<Button-4>", self._on_global_wheel)  # Linux scroll up
        self.root.bind_all("<Button-5>", self._on_global_wheel)  # Linux scroll down

        # Variables for metadata
        self.note_var = tk.StringVar(master=self.root)
        # create a single StringVar for inspector
        self.inspector_var = tk.StringVar(master=self.root)
        # whenever it changes, call the handler
        # (trace_add gives you name, index, mode args; your handler can accept *args)

        self.tray_var = tk.StringVar(master=self.root)
        self.filename_var = tk.StringVar(master=self.root)
        self.save_ts_var = tk.StringVar(master=self.root, value="never")
        self.fill_var = tk.BooleanVar(master=self.root)

        # Navigation buttons
        nav_frame = ttk.LabelFrame(
            self.ctrl_frame, text="Main Controls", padding=(5, 5)
        )
        nav_frame.grid(row=0, column=0, sticky="ew")
        actions = [
            ("◀ Prev", self.prev_image),
            ("Next ▶", self.next_image),
            ("Clear", self.clear_all),
            ("Undo", self.undo),
            ("Redo", self.redo),
            ("Draw", lambda: self.set_mode("draw")),
            ("Erase", lambda: self.set_mode("erase")),
            ("Save", self.save),
        ]
        for i, (txt, fn) in enumerate(actions):
            btn = ttk.Button(
                nav_frame, text=txt, command=self._wrap_cmd(fn), style="App.TButton"
            )
            btn.grid(
                row=i // 2,
                column=i % 2,
                sticky="ew",
                padx=BUTTON_PAD[0],
                pady=BUTTON_PAD[1],
            )
        for c in range(2):
            nav_frame.columnconfigure(c, weight=1)

        # Zoom controls
        zoom_frame = ttk.LabelFrame(self.ctrl_frame, text="Zoom", padding=(5, 5))
        zoom_frame.grid(row=1, column=0, sticky="ew")
        for symbol, factor in [("−", 0.9), ("+", 1.1)]:
            ttk.Button(
                zoom_frame,
                text=symbol,
                command=self._wrap_cmd(lambda f=factor: self._set_zoom(f)),
                style="App.TButton",
            ).pack(side="left", padx=BUTTON_PAD[0], pady=BUTTON_PAD[1])

        # Metadata display
        meta_frame = ttk.LabelFrame(self.ctrl_frame, text="Information", padding=(5, 5))
        meta_frame.grid(row=2, column=0, sticky="ew")
        # Note field
        ttk.Label(meta_frame, text="Note:").grid(row=0, column=0, sticky="nw", padx=5)
        self.note_text = tk.Text(meta_frame, wrap="word", font=ENTRY_FONT, height=3)
        self.note_text.grid(row=0, column=1, sticky="ew", padx=5)
        self.note_text.bind("<KeyRelease>", lambda e: self._adjust_note_height())
        # Inspector combobox

        ttk.Label(meta_frame, text="Inspector:").grid(
            row=1, column=0, sticky="w", padx=5
        )
        self.inspector_entry = ttk.Entry(
            meta_frame, textvariable=self.inspector_var, style="App.TEntry"
        )
        self.inspector_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # Tray
        ttk.Label(meta_frame, text="Tray:").grid(row=2, column=0, sticky="w", padx=5)
        ttk.Entry(
            meta_frame, textvariable=self.tray_var, style="App.TEntry", state="readonly"
        ).grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        # Filename
        ttk.Label(meta_frame, text="Filename:").grid(
            row=3, column=0, sticky="w", padx=5
        )
        ttk.Entry(
            meta_frame,
            textvariable=self.filename_var,
            style="App.TEntry",
            state="readonly",
        ).grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        # Last saved timestamp
        ttk.Label(meta_frame, text="Last saved:").grid(
            row=4, column=0, sticky="w", padx=5
        )
        ttk.Label(meta_frame, textvariable=self.save_ts_var).grid(
            row=4, column=1, sticky="w", padx=5, pady=2
        )
        meta_frame.columnconfigure(1, weight=1)

        # Categories (defects)
        def_frame = ttk.LabelFrame(self.ctrl_frame, text="Categories", padding=(5, 5))
        def_frame.grid(row=3, column=0, sticky="ew")
        self.defect_vars = {}
        for i, d in enumerate(self.defects):
            var = tk.BooleanVar(master=self.root)
            cb = ttk.Checkbutton(def_frame, text=d, variable=var)
            cb.grid(row=i, column=0, sticky="w", padx=PAD_SMALL)
            var.trace_add("write", self._update_pen_labels)
            self.defect_vars[d] = var

        # Fill mode toggle
        fill_frame = ttk.LabelFrame(self.ctrl_frame, text="Fill", padding=(5, 5))
        fill_frame.grid(row=4, column=0, sticky="ew")
        ttk.Checkbutton(
            fill_frame,
            text="PEN FILL MODE",
            variable=self.fill_var,
            style="Fill.Toolbutton",
        ).pack(padx=PAD_SMALL, pady=PAD_SMALL)

        # Pen Labels (6 pens)
        self.pen_frame = ttk.LabelFrame(
            self.ctrl_frame, text="Pen Labels", padding=(5, 5)
        )
        self.pen_frame.grid(row=5, column=0, sticky="ew")
        self.pen_vars = {
            "pen1": tk.StringVar(master=self.root),
            "pen2": tk.StringVar(master=self.root),
            "pen3": tk.StringVar(master=self.root),
            "pen4": tk.StringVar(master=self.root),
            "pen5": tk.StringVar(master=self.root),
            "pen6": tk.StringVar(master=self.root),
        }
        pen_order = [
            ("pen1", "Red"),
            ("pen2", "Blue"),
            ("pen3", "Green"),
            ("pen4", "Yellow"),
            ("pen5", "Magenta"),
            ("pen6", "Cyan"),
        ]
        for i, (key, color) in enumerate(pen_order):
            ttk.Label(self.pen_frame, text=f"{key.upper()} label:").grid(
                row=2 * i, column=0, sticky="w", padx=PAD_SMALL
            )
            combo = ttk.Combobox(
                self.pen_frame,
                textvariable=self.pen_vars[key],
                state="readonly",
                style="App.TEntry",
                width=30,
            )
            combo.grid(row=2 * i, column=1, sticky="ew", padx=PAD_SMALL)
            btn = ttk.Button(
                self.pen_frame,
                text=f"Use {key.upper()} ({color})",
                command=self._wrap_cmd(lambda k=key: self.select_pen(k)),
                style="App.TButton",
            )
            btn.grid(
                row=2 * i + 1,
                column=0,
                columnspan=2,
                sticky="w",
                padx=PAD_SMALL,
                pady=(0, 5),
            )
        self.pen_frame.columnconfigure(1, weight=1)

        # Export buttons
        export_frame = ttk.LabelFrame(
            self.ctrl_frame, text="Export Information", padding=(5, 5)
        )
        export_frame.grid(row=6, column=0, sticky="ew")
        ttk.Button(
            export_frame,
            text="Export as XLSX",
            command=self._wrap_cmd(self.export_excel),
            style="App.TButton",
        ).pack(side="left", padx=PAD_SMALL)
        ttk.Button(
            export_frame,
            text="Export as CSV",
            command=self._wrap_cmd(self.export_csv),
            style="App.TButton",
        ).pack(side="left", padx=PAD_SMALL)

    def _on_global_wheel(self, event):
        """
        If the cursor is over the image canvas → zoom;
        otherwise (i.e. it’s anywhere in the control panel side) → scroll the panel.
        """
        # get pointer in screen coords
        x, y = event.x_root, event.y_root

        # image‐canvas screen bbox
        ix1 = self.canvas.winfo_rootx()
        iy1 = self.canvas.winfo_rooty()
        ix2 = ix1 + self.canvas.winfo_width()
        iy2 = iy1 + self.canvas.winfo_height()

        if ix1 <= x <= ix2 and iy1 <= y <= iy2:
            # inside image → zoom
            if hasattr(event, "delta"):
                factor = 1.1 if event.delta > 0 else 0.9
            else:
                # Linux: Button-4 scroll up, 5 scroll down
                factor = 1.1 if event.num == 4 else 0.9
            self._set_zoom(factor)
        else:
            # anywhere else (your control panel) → scroll it
            # reuse your existing handler (it handles both delta & Button-4/5)
            self._on_control_panel_wheel(event)

    def _on_control_panel_wheel(self, event):
        """
        Scroll the control‐panel canvas up/down when the mousewheel
        is used while the cursor is over the control frame.
        """
        # Windows & macOS: event.delta is ±120 per notch
        if hasattr(event, "delta") and event.delta:
            # negative delta → scroll down, positive → scroll up
            units = int(-1 * (event.delta / 120))
        else:
            # Linux: event.num 4=up, 5=down
            units = -1 if event.num == 4 else 1

        # Scroll the canvas
        self.ctrl_canvas.yview_scroll(units, "units")

    def _update_pen_labels(self, *args):
        # Fixed defect→pen assignment regardless of order
        mapping = {
            "chip": "pen1",
            "scratch": "pen2",
            "gouge": "pen3",
            "inclusion": "pen4",
            "void": "pen5",
            "other": "pen6",
        }
        # Clear all pens
        for k in self.pen_vars.keys():
            self.pen_vars[k].set("")
        # Assign each selected defect to its dedicated pen
        for defect, var in self.defect_vars.items():
            if var.get() and defect in mapping:
                self.pen_vars[mapping[defect]].set(defect)

    def _adjust_note_height(self, event=None):
        text = self.note_text
        line_count = int(text.index("end-1c").split(".")[0])
        text.configure(height=max(3, line_count))

    def select_pen(self, which):
        self.current_pen = which
        self.pen_color = {
            "pen1": PEN1_COLOR,
            "pen2": PEN2_COLOR,
            "pen3": PEN3_COLOR,
            "pen4": PEN4_COLOR,
            "pen5": PEN5_COLOR,
            "pen6": PEN6_COLOR,
        }[which]
        self.pen_width = 8
        self.mode = "draw"
        self.canvas.config(cursor="pencil")

    def _load_image(self):
        full_path = self.image_paths[self.idx]
        self.base = Image.open(full_path).convert("RGB")

        folder = os.path.basename(os.path.dirname(full_path))
        file_name = os.path.basename(full_path)
        base_name, _ = os.path.splitext(file_name)

        # Tray and filename display
        self.tray_var.set(folder.split("_")[0])
        self.filename_var.set(f"{folder}/{file_name}")

        # ─── Load existing combined mask (if any) ────────────────────────────────
        masks_folder = os.path.join(self.output_dir, "masks")
        os.makedirs(masks_folder, exist_ok=True)

        # look for any "*-COMBINED-{base_name}.png" in masks_folder
        pattern = os.path.join(masks_folder, f"*-COMBINED-{base_name}.png")
        mask_files = glob.glob(pattern)

        if mask_files:
            # pick the most recently modified mask
            latest_mask = max(mask_files, key=os.path.getmtime)
            lm = Image.open(latest_mask).convert("RGBA")
            # if the saved mask size doesn’t match the base, resize it
            if lm.size != self.base.size:
                lm = lm.resize(self.base.size, Image.LANCZOS)
            self.layer = lm
        else:
            # no existing mask: start with a blank transparent layer
            self.layer = Image.new("RGBA", self.base.size, (0, 0, 0, 0))

        self.drawer = ImageDraw.Draw(self.layer)

        # Reset history/state
        self.history = [self.layer.copy()]
        self.redo_stack.clear()
        self.pan_x = self.pan_y = 0
        self.zoom = 1.0
        self.drawing = False

        # Load metadata
        key = os.path.basename(full_path)
        meta = self._load_meta().get(key, {})
        self.note_text.delete("1.0", "end")
        self.note_text.insert("1.0", meta.get("note", ""))

        if meta.get("tray"):
            self.tray_var.set(meta["tray"])

        for d, var in self.defect_vars.items():
            var.set(d in meta.get("defects", []))
        self._update_pen_labels()
        for pen_key, lab in meta.get("pen_labels", {}).items():
            if pen_key in self.pen_vars:
                self.pen_vars[pen_key].set(lab)
        self.save_ts_var.set(meta.get("last_saved", "never"))
        self._unsaved_clear = False

    def _show_image(self):
        comp = Image.alpha_composite(self.base.convert("RGBA"), self.layer)
        w, h = comp.size
        disp = comp.resize((int(w * self.zoom), int(h * self.zoom)), Image.LANCZOS)
        self.tkimg = ImageTk.PhotoImage(disp)
        self.canvas.delete("IMG")
        self.canvas.create_image(
            self.pan_x, self.pan_y, anchor="nw", image=self.tkimg, tags="IMG"
        )
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _record_history(self):
        self.history.append(self.layer.copy())
        if len(self.history) > 50:
            self.history.pop(0)
        self.redo_stack.clear()

    def _on_left_press(self, ev):
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
            self.stroke_start_x, self.stroke_start_y = ix, iy
            self.last_x, self.last_y = ix, iy
            self._show_image()

    def _on_left_move(self, ev):
        if self.fill_var.get() or not self.drawing:
            return
        cx = self.canvas.canvasx(ev.x) - self.pan_x
        cy = self.canvas.canvasy(ev.y) - self.pan_y
        ix, iy = int(cx / self.zoom), int(cy / self.zoom)
        color, width = (
            (self.pen_color, self.pen_width)
            if self.mode == "draw"
            else ((0, 0, 0, 0), 20)
        )
        self.drawer.line(
            [(self.last_x, self.last_y), (ix, iy)], fill=color, width=width
        )
        self.last_x, self.last_y = ix, iy
        self._show_image()

    def _on_left_release(self, ev):
        if self.drawing and self.mode == "draw" and not self.fill_var.get():
            # Snap-close any freehand shape by drawing a line from the last point back to the start
            self.drawer.line(
                [
                    (self.last_x, self.last_y),
                    (self.stroke_start_x, self.stroke_start_y),
                ],
                fill=self.pen_color,
                width=self.pen_width,
            )
            self._show_image()
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

    def _on_mousewheel(self, ev):
        factor = 1.1 if ev.delta > 0 else 0.9
        self._set_zoom(factor)

    def set_mode(self, m):
        self.mode = m
        self.canvas.config(cursor="pencil" if m == "draw" else "circle")

    def _set_zoom(self, factor):
        self.zoom = max(0.1, min(self.zoom * factor, 10.0))
        self._show_image()

    def clear_all(self):
        self.layer = Image.new("RGBA", self.base.size, (0, 0, 0, 0))
        self.drawer = ImageDraw.Draw(self.layer)
        self.history = [self.layer.copy()]
        self.redo_stack.clear()
        self.note_text.delete("1.0", "end")
        self.inspector_var.set("")
        for v in self.defect_vars.values():
            v.set(False)
        for v in self.pen_vars.values():
            v.set("")
        self._show_image()
        self._unsaved_clear = True

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
        now = datetime.datetime.now()
        file_ts = now.strftime(TS_FILENAME_FMT)
        disp_ts = now.strftime(TS_DISPLAY_FMT)
        inspector = self.inspector_var.get().strip() or "unknown"
        tray = self.tray_var.get().strip() or ""
        input_folder = os.path.basename(self.input_dir.rstrip(os.sep))
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
            "pen5": PEN5_COLOR,
            "pen6": PEN6_COLOR,
        }
        COLOR_TOL = 10

        # Save individual masks, overwriting any existing ones
        for defect in sels:
            outd = os.path.join(self.output_dir, defect)
            os.makedirs(outd, exist_ok=True)
            # Remove any previous mask files for this image & defect
            pattern = os.path.join(outd, f"{input_folder}-*-{defect}-{base_name}.png")
            for fp in glob.glob(pattern):
                try:
                    os.remove(fp)
                except OSError:
                    pass

            matching_colors = []
            for pen_key, label_text in pen_labels.items():
                if label_text.strip().lower() == defect.strip().lower():
                    color_tuple = pen_to_color[pen_key]
                    if len(color_tuple) > 3:
                        color_tuple = tuple(color_tuple[:3])
                    matching_colors.append(color_tuple)

            mask_rgba = Image.new("RGBA", self.base.size, (0, 0, 0, 0))
            src = self.layer.load()
            dst = mask_rgba.load()
            w, h = self.base.size
            for x in range(w):
                for y in range(h):
                    r, g, b, a = src[x, y]
                    if a == 0:
                        continue
                    for pr, pg, pb in matching_colors:
                        if (
                            abs(r - pr) <= COLOR_TOL
                            and abs(g - pg) <= COLOR_TOL
                            and abs(b - pb) <= COLOR_TOL
                        ):
                            dst[x, y] = (pr, pg, pb, a)
                            break

            rgb_mask = Image.new("RGB", self.base.size, (0, 0, 0))
            rgb_mask.paste(mask_rgba, mask=mask_rgba.split()[3])
            draw_legend = ImageDraw.Draw(rgb_mask)
            font = ImageFont.load_default()
            bbox = draw_legend.textbbox((0, 0), defect, font=font)
            text_h = bbox[3] - bbox[1]
            draw_legend.text(
                (5, self.base.size[1] - PAD_SMALL - text_h),
                defect,
                fill=(255, 255, 255),
                font=font,
            )

            out_filename = (
                f"{input_folder}-{inspector}-{file_ts}-{defect}-{base_name}.png"
            )
            out_path = os.path.join(outd, out_filename)
            rgb_mask.save(out_path)

        # Save combined mask, overwriting any existing ones
        combined_dir = os.path.join(self.output_dir, "masks")
        os.makedirs(combined_dir, exist_ok=True)
        # Remove any previous combined mask files for this image
        pattern = os.path.join(
            combined_dir, f"{input_folder}*-COMBINED-{base_name}.png"
        )
        for fp in glob.glob(pattern):
            try:
                os.remove(fp)
            except OSError:
                pass

        all_defects = "_".join(sels)
        combined_filename = f"{input_folder}-{inspector}-{file_ts}-{all_defects}-COMBINED-{base_name}.png"
        combined_path = os.path.join(combined_dir, combined_filename)
        self.layer.save(combined_path)

        # metadata update (unchanged)
        meta_path = self._meta_path()
        meta = json.load(open(meta_path)) if os.path.isfile(meta_path) else {}
        key = os.path.basename(self.image_paths[self.idx])
        entry = meta.get(key, {})
        entry.update(
            {
                "note": self.note_text.get("1.0", "end-1c"),
                "inspector": self.inspector_var.get(),
                "tray": self.tray_var.get(),
                "defects": [d for d, v in self.defect_vars.items() if v.get()],
                "pen_labels": {k: self.pen_vars[k].get() for k in self.pen_vars},
                "last_saved": disp_ts,
            }
        )
        meta[key] = entry
        json.dump(meta, open(meta_path, "w"), indent=2)
        self.save_ts_var.set(disp_ts)
        messagebox.showinfo("Saved", f"Masks exported to: {', '.join(sels)}")

    def _save_current(self):
        # Determine base filename without extension
        name = os.path.splitext(os.path.basename(self.image_paths[self.idx]))[0]
        # Get input folder name
        input_folder = os.path.basename(self.input_dir.rstrip(os.sep))
        # Prepare masks directory
        mask_dir = os.path.join(self.output_dir, "masks")
        os.makedirs(mask_dir, exist_ok=True)
        # Build filename: [input_folder]-COMBINED-[name].png
        mask_name = f"{input_folder}-COMBINED-{name}.png"
        mp = os.path.join(mask_dir, mask_name)
        # Save the combined mask layer
        self.layer.save(mp)

        # Update metadata without altering the last_saved timestamp
        meta_path = self._meta_path()
        meta = json.load(open(meta_path)) if os.path.isfile(meta_path) else {}
        key = os.path.basename(self.image_paths[self.idx])
        entry = meta.get(key, {})
        entry.update(
            {
                "note": self.note_text.get("1.0", "end-1c"),
                "inspector": self.inspector_var.get(),
                "tray": self.tray_var.get(),
                "defects": [d for d, v in self.defect_vars.items() if v.get()],
                "pen_labels": {k: self.pen_vars[k].get() for k in self.pen_vars},
                "last_saved": entry.get("last_saved", "never"),
            }
        )
        meta[key] = entry
        json.dump(meta, open(meta_path, "w"), indent=2)

    def export_excel(self):
        base_name = os.path.basename(self.input_dir.rstrip(os.sep))
        excel_path = os.path.join(self.output_dir, f"{base_name}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Tray Information"
        headers = [
            "Filename",
            "Tray/Directory",
            "Inspector",
            "Accept/Reject",
            "Defect(s)",
            "Notes",
        ]
        ws.append(headers)
        meta = self._load_meta()
        for fp in self.image_paths:
            bf = os.path.basename(fp)
            name, _ = os.path.splitext(bf)
            e = meta.get(bf, {})
            defects = e.get("defects", [])
            insp = e.get("inspector", "").strip()
            tray = e.get("tray", "").strip()
            note = e.get("note", "").strip()
            if not defects and not insp and not tray and not note:
                status = "Unlabeled"
            elif not defects:
                status = "Accept"
            else:
                status = "Reject"
            ws.append([name, tray, insp, status, ", ".join(defects), note])
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2
        wb.save(excel_path)
        messagebox.showinfo("Export Excel", f"Excel saved to:\n{excel_path}")

    def export_csv(self):
        base_name = os.path.basename(self.input_dir.rstrip(os.sep))
        csv_path = os.path.join(self.output_dir, f"{base_name}.csv")
        meta = self._load_meta()
        with open(csv_path, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            headers = [
                "Filename",
                "Tray/Directory",
                "Inspector",
                "Accept/Reject",
                "Defect(s)",
                "Notes",
            ]
            writer.writerow(headers)
            for fp in self.image_paths:
                bf = os.path.basename(fp)
                name, _ = os.path.splitext(bf)
                e = meta.get(bf, {})
                defects = e.get("defects", [])
                insp = e.get("inspector", "").strip()
                tray = e.get("tray", "").strip()
                note = e.get("note", "").strip()
                if not defects and not insp and not tray and not note:
                    status = "Unlabeled"
                elif not defects:
                    status = "Accept"
                else:
                    status = "Reject"
                writer.writerow([name, tray, insp, status, ", ".join(defects), note])
        messagebox.showinfo("Export CSV", f"CSV saved to:\n{csv_path}")

    def _fit_image_to_canvas(self):
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw > 1 and ch > 1:
            bw, bh = self.base.size
            self.zoom = min(cw / bw, ch / bh, 1.0)

    def _meta_path(self):
        return os.path.join(self.output_dir, "metadata.json")

    def _load_meta(self):
        if os.path.isfile(self._meta_path()):
            return json.load(open(self._meta_path()))
        return {}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_dir", required=True)
    parser.add_argument("--output_dir", required=True)
    parser.add_argument("--config", default="defects_config.json")
    parser.add_argument("--dataset", default="default")
    args = parser.parse_args()
     # ─── Override args with config values ─────────────────────────────────────
    input_folder = config["Paths"].get("input_folder") or args.input_dir
    defects_cfg = config["Paths"].get("defects_config") or args.config
    output_subdir = config["Paths"].get("output_subdir", "output")
    output_folder = os.path.join(input_folder, output_subdir)
    os.makedirs(output_folder, exist_ok=True)
    DrawApp(args.input_dir, args.output_dir, args.config, args.dataset)
