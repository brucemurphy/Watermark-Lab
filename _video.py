"""Video watermarking via ffmpeg.

Generates a tiled, diagonal, semi-transparent text watermark as a PNG
using Pillow, then overlays it onto the input video with ffmpeg. Video is
re-encoded (libx264, CRF 20). Audio is stream-copied to preserve quality.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FFMPEG_EXE = os.path.join(SCRIPT_DIR, "ffmpeg.exe")

VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".mkv", ".avi", ".webm"}

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def is_video_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in VIDEO_EXTS


def _next_available_video(base, suffix, ext):
    """Return base+suffix+ext, appending (1), (2)... if the file exists."""
    candidate = f"{base}{suffix}{ext}"
    n = 1
    while os.path.exists(candidate):
        candidate = f"{base}{suffix}({n}){ext}"
        n += 1
    return candidate


def _resolve_ffmpeg() -> str:
    """Return the path to a usable ffmpeg binary.

    Resolution order:
    1. imageio-ffmpeg bundled binary (preferred — minimal ~30 MB build).
    2. ffmpeg.exe sitting next to this script (user-supplied legacy location).
    3. ffmpeg on PATH.
    """
    # 1. imageio-ffmpeg
    try:
        import imageio_ffmpeg
        exe = imageio_ffmpeg.get_ffmpeg_exe()
        if exe and os.path.isfile(exe):
            return exe
    except Exception:
        pass

    # 2. Script-directory fallback
    if os.path.isfile(FFMPEG_EXE):
        return FFMPEG_EXE

    # 3. PATH
    on_path = shutil.which("ffmpeg")
    if on_path:
        return on_path

    raise FileNotFoundError(
        "ffmpeg not found. Install imageio-ffmpeg (pip install imageio-ffmpeg) "
        "or place ffmpeg.exe next to this script, or add ffmpeg to PATH."
    )


def _probe_video(ffmpeg: str, video_path: str):
    """Return (width, height, has_audio) by parsing 'ffmpeg -i' stderr.

    ffmpeg -i with no output file always exits non-zero; that is expected.
    Stream information is written to stderr regardless of the exit code.
    """
    result = subprocess.run(
        [ffmpeg, "-hide_banner", "-i", video_path],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        creationflags=_NO_WINDOW,
    )
    info = result.stderr.decode("utf-8", errors="replace")

    # Video dimensions: "Video: ..., 1920x1080 [SAR ...]"
    width, height = 1280, 720  # safe fallback
    m = re.search(r"Video:.*?,\s*(\d+)x(\d+)", info)
    if m:
        width, height = int(m.group(1)), int(m.group(2))

    has_audio = "Audio:" in info
    return width, height, has_audio


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

    ffmpeg = _resolve_ffmpeg()

    base, ext = os.path.splitext(video_path)
    if not ext:
        ext = ".mp4"
    output_path = _next_available_video(base, "_watermarked", ext)

    width, height, has_audio = _probe_video(ffmpeg, video_path)

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
