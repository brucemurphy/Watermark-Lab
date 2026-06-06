"""Watermark Lab — Tkinter GUI for PowerPoint and video watermarking."""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

from _powerpoint import add_watermark
from _video import add_video_watermark, is_video_file, VIDEO_EXTS

PPT_EXTS = {".pptx", ".ppt"}
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SPLASH_IMAGE = os.path.join(SCRIPT_DIR, "SplashLab.png")
ICON_ICO = os.path.join(SCRIPT_DIR, "Watermark.ico")
ICON_PNG = os.path.join(SCRIPT_DIR, "Watermark.png")
SPLASH_DURATION_MS = 750

# Dark theme palette
DARK_BG = "#1e1e1e"
DARK_PANEL = "#252526"
DARK_FIELD = "#2d2d30"
DARK_BORDER = "#3f3f46"
DARK_FG = "#e6e6e6"
DARK_MUTED = "#a0a0a0"
ACCENT = "#0a84ff"
ACCENT_HOVER = "#1f8fff"
ACCENT_ACTIVE = "#0066cc"


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
    style.configure("TFrame", background=DARK_BG)
    style.configure("TLabel", background=DARK_BG, foreground=DARK_FG)
    style.configure("Muted.TLabel", background=DARK_BG, foreground=DARK_MUTED)
    style.configure("TEntry",
        fieldbackground=DARK_FIELD, foreground=DARK_FG,
        bordercolor=DARK_BORDER, insertcolor=DARK_FG, padding=4,
    )
    style.map("TEntry",
        fieldbackground=[("focus", DARK_FIELD)],
        bordercolor=[("focus", ACCENT)],
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
        background=ACCENT, foreground="#ffffff",
        bordercolor=ACCENT, padding=(12, 6), relief="flat",
    )
    style.map("Accent.TButton",
        background=[("active", ACCENT_HOVER), ("pressed", ACCENT_ACTIVE),
                    ("disabled", DARK_BORDER)],
        foreground=[("disabled", DARK_MUTED)],
    )
    style.configure("Horizontal.TScale",
        background=DARK_BG, troughcolor=DARK_FIELD,
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
        self.title("Watermark Lab")
        self.resizable(False, False)
        _apply_dark_theme(self)
        _set_window_icon(self)

        self.file_var = tk.StringVar()
        self.text_var = tk.StringVar(value="CONFIDENTIAL")
        self.color_hex = "#A6A6A6"  # default gray
        self.transparency_var = tk.DoubleVar(value=70.0)  # percent

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}
        frm = ttk.Frame(self)
        frm.grid(row=0, column=0, padx=10, pady=10)

        # File picker
        ttk.Label(frm, text="File (.pptx / .mp4):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.file_var, width=50).grid(row=0, column=1, **pad)
        ttk.Button(frm, text="Browse…", command=self._pick_file).grid(row=0, column=2, **pad)

        # Watermark text
        ttk.Label(frm, text="Watermark text:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.text_var, width=50).grid(
            row=1, column=1, columnspan=2, sticky="we", **pad
        )

        # Color picker
        ttk.Label(frm, text="Text color:").grid(row=2, column=0, sticky="w", **pad)
        self.color_swatch = tk.Label(
            frm, text=self.color_hex, bg=self.color_hex, fg="#000000",
            width=12, relief="flat", bd=1, highlightthickness=1,
            highlightbackground=DARK_BORDER,
        )
        self.color_swatch.grid(row=2, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Choose color…", command=self._pick_color).grid(
            row=2, column=2, **pad
        )

        # Transparency slider
        ttk.Label(frm, text="Transparency (%):").grid(row=3, column=0, sticky="w", **pad)
        slider = ttk.Scale(
            frm, from_=0, to=100, orient="horizontal",
            variable=self.transparency_var, command=self._update_trans_label,
        )
        slider.grid(row=3, column=1, sticky="we", **pad)
        self.trans_label = ttk.Label(frm, text="70%")
        self.trans_label.grid(row=3, column=2, sticky="w", **pad)

        # Action buttons
        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=3, pady=(12, 4))
        self.run_btn = ttk.Button(
            btns, text="Apply Watermark", command=self._run, style="Accent.TButton"
        )
        self.run_btn.pack(side="left", padx=6)
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="left", padx=6)

        # Status
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(frm, textvariable=self.status_var, style="Muted.TLabel").grid(
            row=5, column=0, columnspan=3, sticky="w", **pad
        )

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Select PowerPoint or video file",
            filetypes=[
                ("Supported files", "*.pptx *.ppt *.mp4 *.mov *.m4v *.mkv *.avi *.webm"),
                ("PowerPoint files", "*.pptx *.ppt"),
                ("Video files", "*.mp4 *.mov *.m4v *.mkv *.avi *.webm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.file_var.set(path)

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
        elif ext in VIDEO_EXTS or is_video_file(path):
            mode = "video"
        else:
            messagebox.showerror(
                "Unsupported file",
                f"Unsupported file type: {ext}\n\n"
                "Supported: .pptx, .ppt, .mp4, .mov, .m4v, .mkv, .avi, .webm",
            )
            return

        color_rgb = int(self.color_hex.lstrip("#"), 16)
        transparency = max(0.0, min(1.0, self.transparency_var.get() / 100.0))

        self.run_btn.configure(state="disabled")
        if mode == "ppt":
            self.status_var.set("Working… PowerPoint is processing the file.")
        else:
            self.status_var.set("Working… ffmpeg is encoding the video.")
        self.configure(cursor="watch")
        self.update_idletasks()

        def worker():
            try:
                if mode == "ppt":
                    output_path = add_watermark(
                        path, text, color_rgb=color_rgb, transparency=transparency
                    )
                else:
                    output_path = add_video_watermark(
                        path, text,
                        color_rgb=color_rgb,
                        transparency=transparency,
                        progress_cb=self._on_progress,
                    )
                self.after(0, self._on_done, None, output_path)
            except Exception as e:  # noqa: BLE001
                self.after(0, self._on_done, e, None)

        threading.Thread(target=worker, daemon=True).start()

    def _on_progress(self, seconds_done, _total):
        msg = f"Encoding… {seconds_done:0.1f}s processed"
        self.after(0, lambda: self.status_var.set(msg))

    def _on_done(self, err, output_path):
        self.run_btn.configure(state="normal")
        self.configure(cursor="")
        if err is None:
            self.status_var.set(f"Done. Saved: {output_path}")
            self.bell()
            messagebox.showinfo(
                "Finished",
                f"Watermark applied successfully.\n\nSaved to:\n{output_path}\n\n"
                "Opening the new file now…",
            )
            try:
                os.startfile(output_path)  # type: ignore[attr-defined]
            except Exception as open_err:  # noqa: BLE001
                messagebox.showwarning(
                    "Could not open file",
                    f"Watermarking succeeded but the file could not be opened automatically:\n{open_err}",
                )
        else:
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


if __name__ == "__main__":
    _show_splash(SPLASH_IMAGE, SPLASH_DURATION_MS)
    WatermarkApp().mainloop()
