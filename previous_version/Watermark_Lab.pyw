"""Watermark Lab - Tkinter GUI for PowerPoint, Word and video watermarking."""
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, simpledialog

# This legacy UI was archived under previous_version/ during the 2.0.0 switch,
# but it still uses the shared backend modules that live in the repository root.
# Put the root on sys.path so those imports resolve no matter how this file is
# launched (directly, or relaunched by the modern UI's switch).
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from _xpowerpoint import add_watermark  # OneDrive-safe save (shared with the modern UI)
from _word import add_word_watermark, WORD_EXTS
from _video import add_video_watermark, is_video_file, VIDEO_EXTS
from _prefs import load_presets, save_preset, delete_preset, load_recent, add_recent, load_last_dir, save_last_dir
from _version import APP_VERSION
from _updater import check_for_update_async, apply_update, cleanup_old_exe
from _ffmpeg import is_ffmpeg_cached, download_ffmpeg, FfmpegNotReadyError

PPT_EXTS = {".pptx", ".ppt"}
ALL_SUPPORTED = PPT_EXTS | WORD_EXTS | VIDEO_EXTS
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Brand assets live in the repo root, not in previous_version/. Resolve them via
# the shared _uiswitch constants so the window icon and splash always load; fall
# back to the repo root (parent of this folder) if _uiswitch is unavailable.
try:
    from _uiswitch import ICON_ICO, ICON_PNG, SPLASH_PNG as SPLASH_IMAGE
except Exception:
    _ASSET_ROOT = os.path.dirname(SCRIPT_DIR)
    SPLASH_IMAGE = os.path.join(_ASSET_ROOT, "SplashLab.png")
    ICON_ICO = os.path.join(_ASSET_ROOT, "Watermark.ico")
    ICON_PNG = os.path.join(_ASSET_ROOT, "Watermark.png")
SPLASH_DURATION_MS = 750

# Dark theme palette
DARK_BG = "#1e1e1e"
DARK_PANEL = "#252526"
DARK_FIELD = "#2d2d30"
DARK_BORDER = "#3f3f46"
DARK_FG = "#e6e6e6"
DARK_MUTED = "#a0a0a0"
ACCENT = "#FDC301"          # gold
ACCENT_HOVER = "#ffce33"     # lighter gold
ACCENT_ACTIVE = "#d9a600"    # darker gold
ACCENT_FG = "#1e1e1e"        # text/icon color on top of the accent (dark for gold)


def _apply_dark_theme(root: tk.Tk) -> None:
    """Configure ttk + tk widgets for a dark appearance."""
    root.configure(bg=DARK_BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")  # most themable built-in
    except tk.TclError:
        pass

    style.configure(".",
        background=DARK_BG, foreground=DARK_FG,
        fieldbackground=DARK_FIELD, bordercolor=DARK_BORDER,
        lightcolor=DARK_BORDER, darkcolor=DARK_BORDER,
        troughcolor=DARK_FIELD, focuscolor=ACCENT,
        insertcolor=DARK_FG,
    )
    style.configure("TFrame", background=DARK_BG, borderwidth=0, relief="flat")
    style.configure("TLabel", background=DARK_BG, foreground=DARK_FG)
    style.configure("Muted.TLabel", background=DARK_BG, foreground=DARK_MUTED)
    style.configure("TEntry",
        fieldbackground=DARK_FIELD, foreground=DARK_FG,
        bordercolor=DARK_BORDER, insertcolor=DARK_FG, padding=4,
        lightcolor=DARK_BORDER, darkcolor=DARK_BORDER,
        relief="flat",
    )
    style.map("TEntry",
        fieldbackground=[("focus", DARK_FIELD)],
        bordercolor=[("focus", ACCENT)],
    )
    style.configure("TCombobox",
        fieldbackground=DARK_FIELD, foreground=DARK_FG,
        background=DARK_PANEL, selectbackground=DARK_FIELD,
        selectforeground=DARK_FG, bordercolor=DARK_BORDER,
        lightcolor=DARK_BORDER, darkcolor=DARK_BORDER,
        arrowcolor=DARK_FG, padding=4,
    )
    style.map("TCombobox",
        fieldbackground=[("readonly", DARK_FIELD), ("focus", DARK_FIELD)],
        foreground=[("readonly", DARK_FG)],
        bordercolor=[("focus", ACCENT)],
        selectbackground=[("readonly", DARK_FIELD)],
    )
    style.configure("TButton",
        background=DARK_PANEL, foreground=DARK_FG,
        bordercolor=DARK_BORDER, padding=(10, 5), relief="flat",
    )
    style.map("TButton",
        background=[("active", DARK_BORDER), ("pressed", DARK_FIELD)],
        foreground=[("disabled", DARK_MUTED)],
    )
    style.configure("Accent.TButton",
        background=ACCENT, foreground=ACCENT_FG,
        bordercolor=ACCENT, padding=(12, 6), relief="flat",
    )
    style.map("Accent.TButton",
        background=[("active", ACCENT_HOVER), ("pressed", ACCENT_ACTIVE),
                    ("disabled", DARK_BORDER)],
        foreground=[("disabled", DARK_MUTED)],
    )
    style.configure("Horizontal.TScale",
        background=ACCENT, troughcolor=DARK_FIELD,
        bordercolor=DARK_BORDER, lightcolor=ACCENT, darkcolor=ACCENT,
    )
    style.map("Horizontal.TScale",
        background=[("active", ACCENT_HOVER), ("pressed", ACCENT_ACTIVE)],
    )
    # Style the combobox dropdown listbox popup
    root.option_add("*TCombobox*Listbox.background",   DARK_FIELD)
    root.option_add("*TCombobox*Listbox.foreground",   DARK_FG)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", ACCENT_FG)
    style.configure("Dark.TCheckbutton",
        background=DARK_BG, foreground=DARK_FG,
        focuscolor=ACCENT, indicatorcolor=DARK_FIELD,
        indicatorbackground=DARK_FIELD, indicatorforeground=ACCENT,
        indicatorsize=16,
    )
    style.map("Dark.TCheckbutton",
        background=[("active", DARK_BG)],
        foreground=[("disabled", DARK_MUTED)],
        indicatorcolor=[("selected", ACCENT), ("pressed", ACCENT_ACTIVE)],
    )
    # Best-effort dark title bar on Windows 10/11.
    try:
        root.update_idletasks()
        from ctypes import windll, byref, c_int, sizeof
        hwnd = windll.user32.GetParent(root.winfo_id())
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE
            try:
                windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, byref(c_int(1)), sizeof(c_int)
                )
            except Exception:
                pass
    except Exception:
        pass


class _Tooltip:
    """Dark-themed tooltip that appears after a short hover delay.

    Pass either text_fn (plain string) or render_fn(parent_frame) for a
    custom, richly-formatted body.
    """
    _DELAY_MS = 500

    def __init__(self, widget: tk.Widget, text_fn=None, render_fn=None):
        self._widget    = widget
        self._text_fn   = text_fn
        self._render_fn = render_fn
        self._win       = None
        self._after     = None
        widget.bind("<Enter>",       self._on_enter,  add="+")
        widget.bind("<Leave>",       self._on_leave,  add="+")
        widget.bind("<ButtonPress>", self._on_leave,  add="+")

    def _on_enter(self, _evt=None):
        self._after = self._widget.after(self._DELAY_MS, self._show)

    def _on_leave(self, _evt=None):
        if self._after:
            self._widget.after_cancel(self._after)
            self._after = None
        self._hide()

    def _show(self):
        if self._render_fn is None and not (self._text_fn and self._text_fn()):
            return
        self._hide()
        x = self._widget.winfo_rootx()
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 6
        self._win = tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        try:
            tw.wm_attributes("-alpha", 0.97)
        except tk.TclError:
            pass

        # Outer border frame (acts as a thin 1px outline)
        outer = tk.Frame(tw, bg=DARK_BORDER, bd=0)
        outer.pack()
        # Accent strip down the left edge
        strip = tk.Frame(outer, bg=ACCENT, width=4)
        strip.pack(side="left", fill="y")
        # Inner content panel
        body = tk.Frame(outer, bg=DARK_PANEL, bd=0)
        body.pack(side="left", fill="both")

        if self._render_fn is not None:
            self._render_fn(body)
        else:
            tk.Label(
                body, text=self._text_fn(), bg=DARK_PANEL, fg=DARK_FG,
                font=("Segoe UI", 9), padx=12, pady=8,
                wraplength=620, justify="left",
            ).pack()
        tw.lift()

    def _hide(self):
        if self._win:
            try:
                self._win.destroy()
            except Exception:
                pass
            self._win = None


# Structured supported-types data for the help tooltip
_SUPPORTED_TYPES = [
    ("PowerPoint", ".pptx  .ppt", "tiled diagonal watermark on every slide"),
    ("Word",       ".docx  .doc", "native diagonal watermark on every page"),
    ("Video",      ".mp4  .mov  .m4v  .mkv  .avi  .webm", "overlay via ffmpeg"),
]


def _render_supported_types(parent: tk.Frame) -> None:
    """Build the richly-formatted supported file types tooltip body."""
    pad_x = 14
    # Title
    tk.Label(
        parent, text="Supported file types", bg=DARK_PANEL, fg=DARK_FG,
        font=("Segoe UI Semibold", 10, "bold"), justify="left",
    ).grid(row=0, column=0, sticky="w", padx=pad_x, pady=(10, 2))
    # Divider
    tk.Frame(parent, bg=DARK_BORDER, height=1).grid(
        row=1, column=0, sticky="we", padx=pad_x, pady=(2, 6))

    r = 2
    for name, exts, desc in _SUPPORTED_TYPES:
        row = tk.Frame(parent, bg=DARK_PANEL)
        row.grid(row=r, column=0, sticky="w", padx=pad_x, pady=(0, 6))
        tk.Label(row, text=name, bg=DARK_PANEL, fg=ACCENT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        line = tk.Frame(row, bg=DARK_PANEL)
        line.pack(anchor="w")
        tk.Label(line, text=exts, bg=DARK_PANEL, fg=DARK_FG,
                 font=("Consolas", 9)).pack(side="left")
        tk.Label(line, text=f"   {desc}", bg=DARK_PANEL, fg=DARK_MUTED,
                 font=("Segoe UI", 9, "italic")).pack(side="left")
        r += 1

    # Footer note
    tk.Frame(parent, bg=DARK_BORDER, height=1).grid(
        row=r, column=0, sticky="we", padx=pad_x, pady=(2, 6)); r += 1
    tk.Label(
        parent, text="PDF export available for PowerPoint and Word",
        bg=DARK_PANEL, fg=DARK_MUTED, font=("Segoe UI", 8, "italic"),
        justify="left",
    ).grid(row=r, column=0, sticky="w", padx=pad_x, pady=(0, 10))


def _shorten_path(path: str, max_len: int = 70) -> str:
    """Middle-truncate a path so it fits the app width while keeping the
    drive root and filename visible. e.g.
    C:\\Users\\bruce...\\Update Monitor\\file_watermarked.mp4
    """
    if len(path) <= max_len:
        return path
    head = path[:12]                       # drive + start, e.g. "C:\Users\bru"
    tail = path[-(max_len - len(head) - 3):]  # keep the end (folder + filename)
    return f"{head}...{tail}"


def _set_window_icon(win: tk.Misc) -> None:
    """Apply Watermark.ico (preferred) or Watermark.png as the window icon."""
    try:
        if os.path.isfile(ICON_ICO):
            win.iconbitmap(default=ICON_ICO)
            return
    except tk.TclError:
        pass
    if os.path.isfile(ICON_PNG):
        try:
            img = tk.PhotoImage(file=ICON_PNG)
            win.iconphoto(True, img)
            win._icon_ref = img  # type: ignore[attr-defined]
        except tk.TclError:
            pass


class WatermarkApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Watermark Lab  v{APP_VERSION}")
        self.resizable(False, False)
        _apply_dark_theme(self)
        _set_window_icon(self)

        self.file_var         = tk.StringVar()
        self.text_var         = tk.StringVar(value="CONFIDENTIAL")
        self.color_hex        = "#A6A6A6"
        self.transparency_var = tk.DoubleVar(value=70.0)
        self.export_pdf_var   = tk.BooleanVar(value=False)
        self.open_file_var    = tk.BooleanVar(value=True)
        self._last_output_path = None
        self._presets         = {}
        self._preset_var      = tk.StringVar()

        self._build_ui()
        self._refresh_presets()
        self._refresh_recent()
        self._center_on_screen()
        self._bring_to_front()

    def _center_on_screen(self):
        """Position the window centred on screen, matching the splash location."""
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"+{x}+{y}")

    def _bring_to_front(self):
        """Land in the foreground on launch and after a UI-switch relaunch
        (a child process' window can otherwise open behind its parent)."""
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            # Drop the always-on-top flag shortly after so the window behaves
            # normally once it has surfaced.
            self.after(400, lambda: self.attributes("-topmost", False))
        except tk.TclError:
            pass

    def _build_ui(self):
        pad = {"padx": 8, "pady": 5}
        frm = ttk.Frame(self, borderwidth=0, relief="flat")
        frm.grid(row=0, column=0, padx=14, pady=(16, 12))

        # ── Row 0: File picker ───────────────────────────────────────────
        file_lbl_frm = ttk.Frame(frm, borderwidth=0, relief="flat")
        file_lbl_frm.grid(row=0, column=0, sticky="w", **pad)
        ttk.Label(file_lbl_frm, text="File:").pack(side="left")
        help_badge = tk.Label(
            file_lbl_frm, text="\u24D8",  # circled lower-case i
            font=("Segoe UI", 9),
            bg=DARK_BG, fg=DARK_MUTED,
            relief="flat", bd=0, padx=0, pady=0, cursor="arrow",
        )
        help_badge.pack(side="left", padx=(4, 0), pady=(3, 0))
        _Tooltip(help_badge, render_fn=_render_supported_types)

        self.file_combo = ttk.Combobox(frm, textvariable=self.file_var,
                                       width=48, state="normal",
                                       style="File.TCombobox",
                                       postcommand=lambda: self._widen_dropdown(self.file_combo, "File.TCombobox"))
        self.file_combo.grid(row=0, column=1, **pad)
        self.file_combo.bind("<<ComboboxSelected>>", self._on_recent_selected)
        ttk.Button(frm, text="Browse…", command=self._pick_file).grid(row=0, column=2, sticky="we", **pad)
        ttk.Button(frm, text="Folder…", command=self._pick_batch).grid(row=0, column=3, sticky="we", **pad)

        # ── Row 1: Watermark text ────────────────────────────────────────
        ttk.Label(frm, text="Watermark text:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.text_var, width=48).grid(
            row=1, column=1, sticky="we", **pad)

        # ── Row 2: Presets ───────────────────────────────────────────────
        ttk.Label(frm, text="Saved presets:").grid(row=2, column=0, sticky="w", **pad)
        self.preset_combo = ttk.Combobox(frm, textvariable=self._preset_var,
                                         width=48, state="readonly",
                                         style="Preset.TCombobox",
                                         postcommand=lambda: self._widen_dropdown(self.preset_combo, "Preset.TCombobox"))
        self.preset_combo.grid(row=2, column=1, **pad)
        self.preset_combo.bind("<<ComboboxSelected>>", self._load_preset)
        ttk.Button(frm, text="Save",   command=self._save_preset).grid(row=2, column=2, sticky="we", **pad)
        ttk.Button(frm, text="Delete", command=self._delete_preset).grid(row=2, column=3, sticky="we", **pad)

        # Tooltips for truncated fields — hover shows the full text
        _Tooltip(self.file_combo,   lambda: self.file_var.get())
        _Tooltip(self.preset_combo, lambda: self._preset_var.get())

        # ── Row 3: Color picker ──────────────────────────────────────────
        ttk.Label(frm, text="Text color:").grid(row=3, column=0, sticky="w", **pad)
        self.color_swatch = tk.Label(
            frm, text=self.color_hex, bg=self.color_hex, fg="#000000",
            width=12, relief="flat", bd=1, highlightthickness=1,
            highlightbackground=DARK_BORDER,
        )
        self.color_swatch.grid(row=3, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Color…", command=self._pick_color).grid(row=3, column=2, sticky="we", **pad)

        # ── Row 4: Transparency ──────────────────────────────────────────
        ttk.Label(frm, text="Transparency (%):").grid(row=4, column=0, sticky="w", **pad)
        slider = ttk.Scale(frm, from_=0, to=100, orient="horizontal",
                           variable=self.transparency_var,
                           command=self._update_trans_label)
        slider.grid(row=4, column=1, columnspan=2, sticky="we", **pad)
        self.trans_label = ttk.Label(frm, text="70%")
        self.trans_label.grid(row=4, column=3, sticky="w", **pad)

        # ── Row 5: PDF export ────────────────────────────────────────────
        self.pdf_check = tk.Checkbutton(
            frm, text="Also export PDF (PowerPoint / Word only)",
            variable=self.export_pdf_var,
            bg=DARK_BG, fg=DARK_FG,
            activebackground=DARK_BG, activeforeground=DARK_FG,
            selectcolor=DARK_FIELD, relief="flat", bd=0, highlightthickness=0,
        )
        self.pdf_check.grid(row=5, column=1, columnspan=3, sticky="w", **pad)

        # ── Row 6: Open after ────────────────────────────────────────────
        self.open_file_check = tk.Checkbutton(
            frm, text="Open file(s) after watermarking",
            variable=self.open_file_var,
            bg=DARK_BG, fg=DARK_FG,
            activebackground=DARK_BG, activeforeground=DARK_FG,
            selectcolor=DARK_FIELD, relief="flat", bd=0, highlightthickness=0,
            disabledforeground=DARK_MUTED,
        )
        self.open_file_check.grid(row=6, column=1, columnspan=3, sticky="w", **pad)

        # ── Row 7: Action buttons ────────────────────────────────────────
        btns = ttk.Frame(frm, borderwidth=0, relief="flat")
        btns.grid(row=7, column=0, columnspan=4, pady=(12, 4))
        self.run_btn = ttk.Button(btns, text="Apply Watermark",
                                  command=self._run, style="Accent.TButton")
        self.run_btn.pack(side="left", padx=6)
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="left", padx=6)

        # ── Row 8: Status row ────────────────────────────────────────────
        status_frm = ttk.Frame(frm, borderwidth=0, relief="flat")
        status_frm.grid(row=8, column=0, columnspan=4, sticky="ew", **pad)
        status_frm.columnconfigure(0, weight=1)

        # Inner frame holds icon + text together and is centred in the row
        inner = ttk.Frame(status_frm, borderwidth=0, relief="flat")
        inner.grid(row=0, column=0)

        self.folder_btn = tk.Button(
            inner, text="\U0001F4C2",
            font=("Segoe UI Emoji", 11),
            bg=DARK_BG, fg=DARK_BG,
            activebackground=DARK_BG, activeforeground=DARK_BG,
            relief="flat", bd=0, padx=0, pady=0, cursor="",
            command=self._open_output_folder,
        )
        self.folder_btn.grid(row=0, column=0, padx=(0, 4), pady=(5, 0), sticky="s")

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(inner, textvariable=self.status_var,
                  style="Muted.TLabel").grid(row=0, column=1, pady=(0, 2), sticky="ws")

        # Switch to the modern UI (bottom-right); remembers the choice.
        self.modern_btn = ttk.Button(
            status_frm, text="Modern UI ›", command=self._switch_to_modern)
        self.modern_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))

    def _open_output_folder(self):
        target = self._last_output_path
        if not target or not os.path.exists(target):
            return
        target = os.path.normpath(target)  # explorer needs backslashes
        if os.path.isdir(target):
            # Batch result — open the folder itself
            subprocess.Popen(["explorer", target])
        else:
            # Single file — open its folder with the file selected
            subprocess.Popen(["explorer", "/select,", target])

    def _switch_to_modern(self):
        """Remember the choice, launch the modern Qt UI, and close this one."""
        try:
            import _uiswitch
        except Exception:
            self.status_var.set("Modern UI is unavailable.")
            return
        if _uiswitch.relaunch(_uiswitch.MODE_MODERN):
            self.destroy()
        else:
            self.status_var.set("Could not start the modern UI.")

    def _show_folder_btn(self):
        """Make the output-folder icon visible and clickable."""
        self.folder_btn.configure(
            fg=DARK_FG, activebackground=DARK_BORDER,
            activeforeground=DARK_FG, cursor="hand2",
        )

    def _hide_folder_btn(self):
        """Hide the output-folder icon (blends into the background)."""
        self._last_output_path = None
        self.folder_btn.configure(
            fg=DARK_BG, activebackground=DARK_BG,
            activeforeground=DARK_BG, cursor="",
        )

    def _widen_dropdown(self, combo: ttk.Combobox, style_name: str):
        """Extend the dropdown width to fit the longest item using ttk's
        documented 'postoffset' style option (x, y, width_delta, height_delta)."""
        values = combo["values"]
        if not values:
            return
        try:
            import tkinter.font as tkfont
            font = tkfont.nametofont("TkTextFont")
            longest_px = max((font.measure(str(v)) for v in values), default=0)
            entry_px   = combo.winfo_width()
            # Extra width needed beyond the entry box (plus scrollbar + padding)
            extra = max(0, longest_px - entry_px + 40)
            extra = min(extra, 500)  # never run absurdly wide
            style = ttk.Style(self)
            style.configure(style_name, postoffset=(0, 0, extra, 0))
        except Exception:
            pass

    # ── Recent files ─────────────────────────────────────────────────────

    def _refresh_recent(self):
        recent = load_recent()
        self.file_combo["values"] = recent

    def _on_recent_selected(self, _evt=None):
        pass  # value already set by combobox selection

    # ── Presets ─────────────────────────────────────────

    def _refresh_presets(self):
        self._presets = load_presets()
        names = list(self._presets.keys())
        self.preset_combo["values"] = names

    def _load_preset(self, _evt=None):
        name = self._preset_var.get()
        p = self._presets.get(name)
        if not p:
            return
        self.text_var.set(p.get("text", "CONFIDENTIAL"))
        hex_ = p.get("color", "#A6A6A6")
        self.color_hex = hex_
        r, g, b = int(hex_[1:3], 16), int(hex_[3:5], 16), int(hex_[5:7], 16)
        fg = "#000000" if (0.299*r + 0.587*g + 0.114*b) > 140 else "#ffffff"
        self.color_swatch.configure(bg=hex_, fg=fg, text=hex_)
        self.transparency_var.set(p.get("transparency", 70.0))
        self._update_trans_label()

    def _save_preset(self):
        name = simpledialog.askstring(
            "Save Preset", "Preset name:", parent=self,
            initialvalue=self._preset_var.get() or self.text_var.get(),
        )
        if not name:
            return
        save_preset(name, self.text_var.get(), self.color_hex,
                    self.transparency_var.get())
        self._refresh_presets()
        self._preset_var.set(name)

    def _delete_preset(self):
        name = self._preset_var.get()
        if not name:
            return
        if messagebox.askyesno("Delete Preset", f'Delete preset "{name}"?', parent=self):
            delete_preset(name)
            self._refresh_presets()
            self._preset_var.set("")

    # ── Batch processing ──────────────────────────────────────────────────

    def _pick_batch(self):
        folder = filedialog.askdirectory(
            title="Select folder to batch watermark", initialdir=load_last_dir())
        if not folder:
            return
        folder = os.path.normpath(folder)  # askdirectory returns forward slashes
        save_last_dir(folder)
        files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if os.path.splitext(f)[1].lower() in ALL_SUPPORTED
        ]
        if not files:
            messagebox.showinfo("Batch", "No supported files found in that folder.")
            return
        msg = f"Found {len(files)} supported file(s) in:\n{folder}\n\nWatermark all of them?"
        if not messagebox.askyesno("Batch Watermark", msg, parent=self):
            return
        self._run_batch(files, folder)

    def _run_batch(self, files, folder):
        text         = self.text_var.get().strip()
        color_rgb    = int(self.color_hex.lstrip("#"), 16)
        transparency = max(0.0, min(1.0, self.transparency_var.get() / 100.0))
        export_pdf   = bool(self.export_pdf_var.get())

        if not text:
            messagebox.showerror("Missing text", "Please enter watermark text.")
            return

        self.run_btn.configure(state="disabled")
        self.open_file_check.configure(state="disabled")
        self._hide_folder_btn()
        self.status_var.set(f"Batch: processing 0 / {len(files)}…")
        self.configure(cursor="watch")
        self.update_idletasks()

        def worker():
            done, errors = 0, []
            for i, path in enumerate(files, 1):
                self.after(0, lambda i=i: self.status_var.set(
                    f"Batch: processing {i} / {len(files)}…"))
                ext = os.path.splitext(path)[1].lower()
                try:
                    if ext in PPT_EXTS:
                        add_watermark(path, text, color_rgb=color_rgb,
                                      transparency=transparency, export_pdf=export_pdf)
                    elif ext in WORD_EXTS:
                        add_word_watermark(path, text, color_rgb=color_rgb,
                                           transparency=transparency, export_pdf=export_pdf)
                    elif ext in VIDEO_EXTS or is_video_file(path):
                        add_video_watermark(path, text, color_rgb=color_rgb,
                                            transparency=transparency)
                    done += 1
                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")
            self.after(0, self._on_batch_done, done, errors, folder)

        threading.Thread(target=worker, daemon=True).start()

    def _on_batch_done(self, done, errors, folder):
        self.run_btn.configure(state="normal")
        self.open_file_check.configure(state="normal")
        self.configure(cursor="")
        self.bell()
        if done:
            self._last_output_path = folder
            self._show_folder_btn()
        msg = f"Batch complete: {done} file(s) watermarked."
        if errors:
            msg += f"\n\n{len(errors)} error(s):\n" + "\n".join(errors[:10])
        messagebox.showinfo("Batch Complete", msg)
        self.status_var.set(f"Batch done. {done} file(s) processed.")

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Select file to watermark",
            initialdir=load_last_dir(),
            filetypes=[
                ("Supported files",
                 "*.pptx *.ppt *.docx *.doc "
                 "*.mp4 *.mov *.m4v *.mkv *.avi *.webm"),
                ("PowerPoint files", "*.pptx *.ppt"),
                ("Word documents",   "*.docx *.doc"),
                ("Video files",      "*.mp4 *.mov *.m4v *.mkv *.avi *.webm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            save_last_dir(path)
            self.file_var.set(path)
            add_recent(path)
            self._refresh_recent()

    def _pick_color(self):
        rgb, hex_ = colorchooser.askcolor(color=self.color_hex, title="Pick watermark color")
        if hex_:
            self.color_hex = hex_
            # Pick readable foreground based on luminance.
            r, g, b = int(hex_[1:3], 16), int(hex_[3:5], 16), int(hex_[5:7], 16)
            lum = (0.299 * r + 0.587 * g + 0.114 * b)
            fg = "#000000" if lum > 140 else "#ffffff"
            self.color_swatch.configure(bg=hex_, fg=fg, text=hex_)

    def _update_trans_label(self, _evt=None):
        self.trans_label.configure(text=f"{int(self.transparency_var.get())}%")

    def _run(self):
        path = self.file_var.get().strip()
        text = self.text_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Missing file", "Please select a valid file.")
            return
        if not text:
            messagebox.showerror("Missing text", "Please enter watermark text.")
            return

        ext = os.path.splitext(path)[1].lower()
        if ext in PPT_EXTS:
            mode = "ppt"
        elif ext in WORD_EXTS:
            mode = "word"
        elif ext in VIDEO_EXTS or is_video_file(path):
            mode = "video"
        else:
            messagebox.showerror(
                "Unsupported file",
                f"Unsupported file type: {ext}\n\n"
                "Supported: .pptx .ppt .docx .doc "
                ".mp4 .mov .m4v .mkv .avi .webm",
            )
            return

        color_rgb = int(self.color_hex.lstrip("#"), 16)
        transparency = max(0.0, min(1.0, self.transparency_var.get() / 100.0))

        self.run_btn.configure(state="disabled")
        self._hide_folder_btn()
        if mode == "ppt":
            self.status_var.set("Working… PowerPoint is processing the file.")
        elif mode == "word":
            self.status_var.set("Working… Word is processing the document.")
        else:
            self.status_var.set("Working… ffmpeg is encoding the video.")
        self.configure(cursor="watch")
        self.update_idletasks()
        add_recent(path)
        self._refresh_recent()

        def worker():
            try:
                if mode == "ppt":
                    output_path = add_watermark(
                        path, text, color_rgb=color_rgb, transparency=transparency,
                        export_pdf=bool(self.export_pdf_var.get()),
                    )
                elif mode == "word":
                    output_path = add_word_watermark(
                        path, text, color_rgb=color_rgb, transparency=transparency,
                        export_pdf=bool(self.export_pdf_var.get()),
                    )
                else:
                    output_path = add_video_watermark(
                        path, text,
                        color_rgb=color_rgb,
                        transparency=transparency,
                        progress_cb=self._on_progress,
                    )
                self.after(0, self._on_done, None, output_path)
            except FfmpegNotReadyError:
                self.after(0, self._prompt_ffmpeg_download, path, text, color_rgb, transparency)
            except Exception as e:  # noqa: BLE001
                self.after(0, self._on_done, e, None)

        threading.Thread(target=worker, daemon=True).start()

    def _prompt_ffmpeg_download(self, path, text, color_rgb, transparency):
        """Ask the user to download ffmpeg, then retry the video job."""
        self.run_btn.configure(state="normal")
        self.configure(cursor="")
        if not messagebox.askyesno(
            "ffmpeg required",
            "Video watermarking needs ffmpeg, which hasn't been downloaded yet.\n\n"
            "It will be downloaded once (~30 MB) from GitHub and saved alongside\n"
            "this application so it travels with it.\n\n"
            "Download now?",
            parent=self,
        ):
            self.status_var.set("Video job cancelled — ffmpeg not downloaded.")
            return
        self._run_ffmpeg_download(path, text, color_rgb, transparency)

    def _run_ffmpeg_download(self, path, text, color_rgb, transparency):
        """Show a progress dialog, download ffmpeg, then kick off the video job."""
        dlg = tk.Toplevel(self)
        dlg.title("Downloading ffmpeg…")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(bg=DARK_BG)
        try:
            dlg.iconbitmap(ICON_ICO)
        except Exception:
            pass

        ttk.Label(dlg, text="Downloading ffmpeg — please wait…").pack(padx=20, pady=(16, 4))
        bar = ttk.Progressbar(dlg, length=340, mode="determinate")
        bar.pack(padx=20, pady=4)
        lbl = ttk.Label(dlg, text="Starting…", style="Muted.TLabel")
        lbl.pack(padx=20, pady=(4, 16))
        dlg.update_idletasks()
        # Centre over the main window
        dlg.geometry(
            f"+{self.winfo_x() + (self.winfo_width() - dlg.winfo_reqwidth()) // 2}"
            f"+{self.winfo_y() + (self.winfo_height() - dlg.winfo_reqheight()) // 2}"
        )

        def _progress(done, total):
            if total:
                pct = done * 100 // total
                mb_done = done / 1_048_576
                mb_total = total / 1_048_576
                self.after(0, lambda p=pct, d=mb_done, t=mb_total: (
                    bar.configure(value=p),
                    lbl.configure(text=f"{d:.1f} / {t:.1f} MB"),
                ))
            else:
                mb_done = done / 1_048_576
                self.after(0, lambda d=mb_done: lbl.configure(text=f"{d:.1f} MB downloaded…"))

        def _worker():
            try:
                download_ffmpeg(progress_cb=_progress)
                self.after(0, _on_done, None)
            except Exception as exc:
                self.after(0, _on_done, exc)

        def _on_done(err):
            dlg.destroy()
            if err:
                messagebox.showerror(
                    "Download failed",
                    f"Could not download ffmpeg:\n{err}\n\n"
                    "Download it manually from https://www.gyan.dev/ffmpeg/builds/ "
                    "and place ffmpeg.exe in the same folder as this application.",
                    parent=self,
                )
                self.status_var.set("ffmpeg download failed.")
                return
            self.status_var.set("ffmpeg downloaded. Starting video job…")
            self.run_btn.configure(state="disabled")
            self.configure(cursor="watch")

            def _video_worker():
                try:
                    output_path = add_video_watermark(
                        path, text,
                        color_rgb=color_rgb,
                        transparency=transparency,
                        progress_cb=self._on_progress,
                    )
                    self.after(0, self._on_done, None, output_path)
                except Exception as e:
                    self.after(0, self._on_done, e, None)

            threading.Thread(target=_video_worker, daemon=True).start()

        threading.Thread(target=_worker, daemon=True).start()

    def _on_progress(self, seconds_done, _total):
        msg = f"Encoding… {seconds_done:0.1f}s processed"
        self.after(0, lambda: self.status_var.set(msg))

    def _on_done(self, err, output_path):
        self.run_btn.configure(state="normal")
        self.configure(cursor="")
        if err is None:
            self._last_output_path = output_path
            self._show_folder_btn()
            self.status_var.set(f"Done. Saved: {_shorten_path(output_path)}")
            self.bell()
            pdf_path = os.path.splitext(output_path)[0] + ".pdf"
            pdf_exported = self.export_pdf_var.get() and os.path.isfile(pdf_path)
            msg = f"Watermark applied successfully.\n\nSaved to:\n{output_path}"
            if pdf_exported:
                msg += f"\n\nPDF also saved to:\n{pdf_path}"
            messagebox.showinfo("Finished", msg)
            if self.open_file_var.get():
                try:
                    os.startfile(output_path)  # type: ignore[attr-defined]
                except Exception as open_err:  # noqa: BLE001
                    messagebox.showwarning(
                        "Could not open file",
                        f"Watermarking succeeded but the file could not be opened automatically:\n{open_err}",
                    )
                # Also open the PDF if export was requested and the file exists
                if self.export_pdf_var.get():
                    pdf_path = os.path.splitext(output_path)[0] + ".pdf"
                    if os.path.isfile(pdf_path):
                        try:
                            os.startfile(pdf_path)  # type: ignore[attr-defined]
                        except Exception:
                            pass
        else:
            self._hide_folder_btn()
            self.status_var.set(f"Error: {err}")
            messagebox.showerror("Error", str(err))


def _show_splash(image_path: str, duration_ms: int) -> None:
    """Display SplashLab.png in a borderless window for duration_ms, then close."""
    if not os.path.isfile(image_path):
        return
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.configure(bg=DARK_BG)
    _set_window_icon(splash)
    try:
        splash.attributes("-topmost", True)
    except tk.TclError:
        pass
    try:
        img = tk.PhotoImage(file=image_path)
    except tk.TclError:
        splash.destroy()
        return
    label = tk.Label(
        splash, image=img, borderwidth=0, highlightthickness=0, bg=DARK_BG
    )
    label.image = img  # keep reference
    label.pack()
    splash.update_idletasks()
    w, h = img.width(), img.height()
    sw, sh = splash.winfo_screenwidth(), splash.winfo_screenheight()
    splash.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
    splash.after(duration_ms, splash.destroy)
    splash.mainloop()


def _on_update_available(root: tk.Tk, remote_version: str, notes: str) -> None:
    """Prompt the user and apply the update if they agree."""
    msg = f"Watermark Lab {remote_version} is available."
    if notes:
        msg += f"\n\nWhat's new:\n{notes}"
    msg += "\n\nInstall now and restart?"
    if messagebox.askyesno("Update available", msg, parent=root):
        root.configure(cursor="watch")
        root.update_idletasks()
        try:
            apply_update()
        except Exception as exc:
            root.configure(cursor="")
            messagebox.showerror(
                "Update failed",
                f"Could not apply the update:\n{exc}\n\n"
                "You can download the latest version manually from the releases page.",
                parent=root,
            )


if __name__ == "__main__":
    cleanup_old_exe()
    _show_splash(SPLASH_IMAGE, SPLASH_DURATION_MS)
    app = WatermarkApp()
    app.after(
        3000,
        lambda: check_for_update_async(
            app,
            lambda ver, notes: _on_update_available(app, ver, notes),
        ),
    )
    app.mainloop()
