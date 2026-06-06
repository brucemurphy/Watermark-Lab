"""Video watermarking via ffmpeg.

Generates a tiled, diagonal, semi-transparent text watermark as a PNG
using Pillow, then overlays it onto the input video with ffmpeg. Video is
re-encoded (libx264, CRF 20). Audio is stream-copied to preserve quality.
"""
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FFMPEG_EXE = os.path.join(SCRIPT_DIR, "ffmpeg.exe")
FFPROBE_EXE = os.path.join(SCRIPT_DIR, "ffprobe.exe")

VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".mkv", ".avi", ".webm"}

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def is_video_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in VIDEO_EXTS


def _resolve_ffmpeg():
    ffmpeg = FFMPEG_EXE if os.path.isfile(FFMPEG_EXE) else shutil.which("ffmpeg")
    ffprobe = FFPROBE_EXE if os.path.isfile(FFPROBE_EXE) else shutil.which("ffprobe")
    if not ffmpeg or not ffprobe:
        raise FileNotFoundError(
            "ffmpeg.exe / ffprobe.exe not found. Place them next to this script "
            "or add them to PATH."
        )
    return ffmpeg, ffprobe


def _probe_dimensions(ffprobe: str, video_path: str):
    out = subprocess.check_output(
        [
            ffprobe, "-v", "error",
            "-select_streams", "v:0",
            "-show_entries", "stream=width,height",
            "-of", "csv=p=0:s=x",
            video_path,
        ],
        creationflags=_NO_WINDOW,
    ).decode("utf-8", errors="replace").strip()
    w_str, _, h_str = out.partition("x")
    return int(w_str), int(h_str)


def _has_audio(ffprobe: str, video_path: str) -> bool:
    out = subprocess.check_output(
        [
            ffprobe, "-v", "error",
            "-select_streams", "a",
            "-show_entries", "stream=index",
            "-of", "csv=p=0",
            video_path,
        ],
        creationflags=_NO_WINDOW,
    ).decode("utf-8", errors="replace").strip()
    return bool(out)


def _find_font(font_size: int) -> ImageFont.ImageFont:
    candidates = [
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
    ]
    for path in candidates:
        if os.path.isfile(path):
            try:
                return ImageFont.truetype(path, font_size)
            except OSError:
                continue
    return ImageFont.load_default()


def _make_watermark_png(
    text: str,
    video_w: int,
    video_h: int,
    color_rgb: int,
    transparency: float,
    font_size: int,
    out_png: str,
) -> None:
    # Oversize the canvas so rotation fully covers the video frame.
    diag = int((video_w ** 2 + video_h ** 2) ** 0.5)
    canvas = diag + max(video_w, video_h)

    img = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    font = _find_font(font_size)

    r = (color_rgb >> 16) & 0xFF
    g = (color_rgb >> 8) & 0xFF
    b = color_rgb & 0xFF
    alpha = int(round(255 * (1.0 - max(0.0, min(1.0, transparency)))))
    fill = (r, g, b, alpha)

    bbox = draw.textbbox((0, 0), text, font=font)
    tw = max(1, bbox[2] - bbox[0])
    th = max(1, bbox[3] - bbox[1])

    gap_x = max(font_size, tw // 3)
    gap_y = th * 3
    step_x = tw + gap_x
    step_y = th + gap_y

    row = 0
    for y in range(0, canvas, step_y):
        offset = (step_x // 2) if (row % 2) else 0
        for x in range(-step_x, canvas, step_x):
            draw.text((x + offset, y), text, font=font, fill=fill)
        row += 1

    rotated = img.rotate(30, resample=Image.BICUBIC, expand=False)

    cx = (canvas - video_w) // 2
    cy = (canvas - video_h) // 2
    rotated.crop((cx, cy, cx + video_w, cy + video_h)).save(out_png, "PNG")


def add_video_watermark(
    video_path: str,
    watermark_text: str,
    color_rgb: int = 0xA6A6A6,
    transparency: float = 0.70,
    font_size: int | None = None,
    progress_cb=None,
) -> str:
    """Watermark a video. Returns the output path.

    progress_cb: optional callable(seconds_done, total_seconds) for UI updates.
    """
    video_path = os.path.abspath(video_path)
    if not os.path.isfile(video_path):
        raise FileNotFoundError(video_path)

    ffmpeg, ffprobe = _resolve_ffmpeg()

    base, ext = os.path.splitext(video_path)
    if not ext:
        ext = ".mp4"
    output_path = f"{base}_watermarked{ext}"

    width, height = _probe_dimensions(ffprobe, video_path)
    has_audio = _has_audio(ffprobe, video_path)

    if font_size is None:
        # Auto-scale: roughly 1/30 of frame height, clamped.
        font_size = max(18, min(72, height // 30))

    tmp_dir = tempfile.mkdtemp(prefix="vwm_")
    wm_png = os.path.join(tmp_dir, "watermark.png")
    try:
        _make_watermark_png(
            watermark_text, width, height, color_rgb, transparency, font_size, wm_png
        )

        cmd = [
            ffmpeg, "-y",
            "-hide_banner", "-loglevel", "error", "-stats",
            "-i", video_path,
            "-i", wm_png,
            "-filter_complex", "[0:v][1:v]overlay=0:0:format=auto",
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "20",
            "-pix_fmt", "yuv420p",
        ]
        if has_audio:
            cmd += ["-c:a", "copy"]
        if ext.lower() in {".mp4", ".m4v", ".mov"}:
            cmd += ["-movflags", "+faststart"]
        cmd += [output_path]

        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=_NO_WINDOW,
            universal_newlines=True,
            bufsize=1,
        )

        last_err = []
        for line in proc.stdout:
            last_err.append(line)
            if progress_cb and "time=" in line:
                # Parse "time=HH:MM:SS.xx"
                try:
                    t = line.split("time=", 1)[1].split(" ", 1)[0]
                    hh, mm, ss = t.split(":")
                    seconds = int(hh) * 3600 + int(mm) * 60 + float(ss)
                    progress_cb(seconds, None)
                except Exception:
                    pass

        proc.wait()
        if proc.returncode != 0:
            tail = "".join(last_err[-20:]).strip() or "(no output)"
            raise RuntimeError(f"ffmpeg failed (exit {proc.returncode}):\n{tail}")

        return output_path
    finally:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Usage: python Video_Watermark.py input.mp4 "WATERMARK TEXT"')
        sys.exit(1)
    out = add_video_watermark(sys.argv[1], sys.argv[2])
    print(f"Saved: {out}")
