"""Video watermarking via ffmpeg.

Generates a tiled, diagonal, semi-transparent text watermark as a PNG
using Pillow, then overlays it onto the input video with ffmpeg. Video is
re-encoded (libx264, CRF 20). Audio is stream-copied to preserve quality.
"""
import os
import re
import subprocess
import sys
import tempfile
import shutil

from PIL import Image, ImageDraw, ImageFont

from _ffmpeg import get_ffmpeg_exe, FfmpegNotReadyError  # noqa: F401 (re-exported)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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
    """Return the path to the cached ffmpeg binary via _ffmpeg module."""
    return get_ffmpeg_exe()


def _probe_video(ffmpeg: str, video_path: str):
    """Return (width, height, has_audio) by parsing 'ffmpeg -i' stderr.

    ffmpeg -i with no output file always exits non-zero; that is expected.
    Stream information is written to stderr regardless of the exit code.
    """
    result = subprocess.run(
        [ffmpeg, "-hide_banner", "-i", video_path],
        stdin=subprocess.DEVNULL,
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


# CENC / DRM box atoms in an encrypted MP4's metadata (moov/moof). 'ffmpeg -i'
# HIDES these (it reports the underlying avc1/mp4a via the sinf/frma box), so we
# parse the container metadata directly instead of scanning ffmpeg's output.
_ENC_ATOMS = (b"encv", b"enca", b"tenc", b"sinf", b"schm", b"pssh")
# 'ffmpeg -i' stderr markers (fallback for non-MP4 containers we don't parse).
_ENC_MARKERS = ("(encv", "(enca", "pssh")


def _iter_top_boxes(fh, size):
    """Yield (type, content_start, box_end) for each top-level ISO-BMFF box."""
    pos = 0
    while pos + 8 <= size:
        fh.seek(pos)
        hdr = fh.read(8)
        if len(hdr) < 8:
            break
        box_size = int.from_bytes(hdr[0:4], "big")
        btype = hdr[4:8]
        content_start = pos + 8
        if box_size == 1:                     # 64-bit extended size
            ext = fh.read(8)
            if len(ext) < 8:
                break
            box_size = int.from_bytes(ext, "big")
            content_start = pos + 16
        elif box_size == 0:                   # runs to EOF
            box_size = size - pos
        if box_size < 8:
            break
        yield btype, content_start, pos + box_size
        pos += box_size


def _mp4_is_encrypted(path) -> bool:
    """Detect CENC/DRM by scanning only the MP4 metadata boxes (moov/moof).

    Scanning the whole file for atom names would risk false positives from mdat
    media bytes, so we walk the top-level boxes and only inspect the metadata.
    """
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as fh:
            for btype, cstart, cend in _iter_top_boxes(fh, size):
                if btype == b"pssh":
                    return True
                if btype in (b"moov", b"moof"):
                    fh.seek(cstart)
                    content = fh.read(min(cend - cstart, 8_000_000))
                    if any(a in content for a in _ENC_ATOMS):
                        return True
    except Exception:
        return False
    return False


def _looks_encrypted(ffmpeg_info: str) -> bool:
    info = (ffmpeg_info or "").lower()
    return any(m in info for m in _ENC_MARKERS)


def is_protected_video(path: str) -> bool:
    """Return True if the video is DRM/CENC-encrypted (ffmpeg can't decode it).

    Parses the MP4/MOV metadata for encryption boxes first (reliable, and works
    even before ffmpeg is downloaded); falls back to an 'ffmpeg -i' marker scan
    for other containers.
    """
    if os.path.splitext(path)[1].lower() in (".mp4", ".m4v", ".mov"):
        if _mp4_is_encrypted(path):
            return True
    try:
        ffmpeg = _resolve_ffmpeg()
    except Exception:
        return False
    try:
        result = subprocess.run(
            [ffmpeg, "-hide_banner", "-i", path],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            creationflags=_NO_WINDOW,
            timeout=20,
        )
    except Exception:
        return False
    return _looks_encrypted(result.stderr.decode("utf-8", errors="replace"))


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

    # DRM/CENC-encrypted videos can't be decoded, so refuse rather than fail
    # mid-encode with a cryptic ffmpeg error.
    if is_protected_video(video_path):
        raise RuntimeError(
            "This video is DRM/encryption-protected and can't be watermarked.")

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
            stdin=subprocess.DEVNULL,
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
