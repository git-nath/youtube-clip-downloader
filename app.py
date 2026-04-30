import json
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

try:
    import yt_dlp
except ImportError:
    yt_dlp = None


INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
DEFAULT_WINDOW_SIZE = "1040x780"
HISTORY_LIMIT = 6

APP_BG = "#f5f1e8"
CARD_BG = "#fffaf0"
HERO_BG = "#113f3c"
INK = "#1d2b2a"
MUTED = "#667572"
ACCENT = "#0f766e"
ACCENT_DARK = "#0b5f59"
BORDER = "#ded6c8"
SOFT_ACCENT = "#e3f4ef"


def format_has_video(fmt: dict | None) -> bool:
    return bool(fmt and fmt.get("vcodec") not in {None, "none"})


def format_has_audio(fmt: dict | None) -> bool:
    return bool(fmt and fmt.get("acodec") not in {None, "none"})


def same_media_resource(video_format: dict, audio_format: dict | None) -> bool:
    if not audio_format:
        return False
    if video_format.get("format_id") and video_format.get("format_id") == audio_format.get("format_id"):
        return True
    return bool(video_format.get("url") and video_format.get("url") == audio_format.get("url"))


def source_host(url: str) -> str:
    try:
        return urlparse(url).netloc.lower()
    except ValueError:
        return ""


def prefers_local_download(url: str) -> bool:
    host = source_host(url)
    return "tiktok.com" in host


def supports_timestamp_url(url: str) -> bool:
    host = source_host(url)
    return "youtube.com" in host or "youtu.be" in host


def with_timestamp(url: str, seconds: float) -> str:
    if not supports_timestamp_url(url):
        return url

    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["t"] = f"{max(0, int(seconds))}s"
    return urlunparse(parsed._replace(query=urlencode(query)))


def get_settings_path() -> Path:
    base_dir = Path(os.getenv("APPDATA", Path.home() / ".portion_downloader"))
    return base_dir / "PortionDownloader" / "settings.json"


def default_videos_folder() -> Path:
    return Path.home() / "Videos"


def sanitize_filename(name: str) -> str:
    cleaned = INVALID_FILENAME_CHARS.sub("_", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.rstrip(". ")
    return cleaned or "clip"


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def parse_timestamp(value: str) -> float:
    text = value.strip()
    if not text:
        raise ValueError("Timestamp is required.")

    if ":" not in text:
        seconds = float(text)
        if seconds < 0:
            raise ValueError("Timestamp cannot be negative.")
        return seconds

    parts = text.split(":")
    if len(parts) > 3:
        raise ValueError("Use SS, MM:SS, or HH:MM:SS format.")

    try:
        if len(parts) == 2:
            minutes = int(parts[0])
            seconds = float(parts[1])
            hours = 0
        else:
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = float(parts[2])
    except ValueError as exc:
        raise ValueError("Timestamps must be numeric.") from exc

    if hours < 0 or minutes < 0 or seconds < 0:
        raise ValueError("Timestamp cannot be negative.")
    if len(parts) == 3 and minutes >= 60:
        raise ValueError("Minutes must be below 60.")
    if seconds >= 60:
        raise ValueError("Seconds must be below 60.")

    return hours * 3600 + minutes * 60 + seconds


def format_seconds_for_display(seconds: float) -> str:
    total_milliseconds = int(round(seconds * 1000))
    hours, remainder = divmod(total_milliseconds, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    whole_seconds, milliseconds = divmod(remainder, 1000)
    if milliseconds:
        return f"{hours:02d}:{minutes:02d}:{whole_seconds:02d}.{milliseconds:03d}"
    return f"{hours:02d}:{minutes:02d}:{whole_seconds:02d}"


def format_seconds_for_filename(seconds: float) -> str:
    total_milliseconds = int(round(seconds * 1000))
    hours, remainder = divmod(total_milliseconds, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    whole_seconds, milliseconds = divmod(remainder, 1000)
    if milliseconds:
        return f"{hours:02d}-{minutes:02d}-{whole_seconds:02d}_{milliseconds:03d}"
    return f"{hours:02d}-{minutes:02d}-{whole_seconds:02d}"


def build_default_filename(title: str, start: float, end: float) -> str:
    return sanitize_filename(
        f"{title} [{format_seconds_for_filename(start)} to {format_seconds_for_filename(end)}]"
    )


def codec_label(codec: str | None) -> str:
    if not codec or codec == "none":
        return "none"
    return codec.split(".")[0]


def best_video_sort_key(fmt: dict) -> tuple:
    return (
        int(fmt.get("height") or 0),
        int(fmt.get("fps") or 0),
        float(fmt.get("tbr") or 0),
        1 if fmt.get("ext") == "mp4" else 0,
        1 if str(fmt.get("vcodec") or "").startswith("avc1") else 0,
    )


def best_audio_sort_key(fmt: dict) -> tuple:
    codec = str(fmt.get("acodec") or "")
    return (
        float(fmt.get("abr") or 0),
        float(fmt.get("asr") or 0),
        1 if fmt.get("ext") in {"m4a", "mp4"} else 0,
        1 if codec.startswith("mp4a") or codec.startswith("aac") else 0,
    )


def collect_video_formats(formats: list[dict]) -> list[dict]:
    all_video = [
        fmt
        for fmt in formats
        if fmt.get("vcodec") not in {None, "none"}
        and (fmt.get("url") or fmt.get("manifest_url"))
    ]
    video_only = [fmt for fmt in all_video if fmt.get("acodec") == "none"]
    candidates = video_only or all_video
    return sorted(candidates, key=best_video_sort_key, reverse=True)


def collect_audio_formats(formats: list[dict]) -> list[dict]:
    all_audio = [
        fmt
        for fmt in formats
        if fmt.get("acodec") not in {None, "none"}
        and (fmt.get("url") or fmt.get("manifest_url"))
    ]
    audio_only = [fmt for fmt in all_audio if fmt.get("vcodec") in {None, "none"}]
    candidates = audio_only or all_audio
    return sorted(candidates, key=best_audio_sort_key, reverse=True)


def build_video_label(fmt: dict) -> str:
    resolution = fmt.get("resolution")
    if not resolution:
        height = fmt.get("height")
        resolution = f"{height}p" if height else "video"

    fps = fmt.get("fps")
    fps_text = f"{int(fps)}fps" if fps else "?"
    note = fmt.get("format_note") or ""
    size = fmt.get("filesize") or fmt.get("filesize_approx")
    size_text = f"{size / (1024 * 1024):.1f} MB" if size else "size?"
    pieces = [
        resolution,
        str(fmt.get("ext") or "?").upper(),
        codec_label(fmt.get("vcodec")),
        fps_text,
        size_text,
        f"id {fmt.get('format_id')}",
    ]
    if note:
        pieces.insert(1, note)
    return " | ".join(pieces)


def build_audio_label(fmt: dict) -> str:
    abr = fmt.get("abr")
    abr_text = f"{abr:.0f} kbps" if abr else "audio"
    size = fmt.get("filesize") or fmt.get("filesize_approx")
    size_text = f"{size / (1024 * 1024):.1f} MB" if size else "size?"
    return " | ".join(
        [
            abr_text,
            str(fmt.get("ext") or "?").upper(),
            codec_label(fmt.get("acodec")),
            size_text,
            f"id {fmt.get('format_id')}",
        ]
    )


def build_headers_blob(headers: dict | None) -> str:
    if not headers:
        return ""
    lines = []
    for key, value in headers.items():
        if not value:
            continue
        if key.lower() in {"user-agent", "accept-encoding"}:
            continue
        lines.append(f"{key}: {value}")
    return "\r\n".join(lines) + ("\r\n" if lines else "")


def ffmpeg_time(seconds: float) -> str:
    return format_seconds_for_display(seconds)


def ydl_options() -> dict:
    return {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "noplaylist": True,
    }


def build_ydl_format_selector(video_format: dict, audio_format: dict | None) -> str:
    video_id = str(video_format.get("format_id") or "best")
    if not audio_format:
        return video_id

    audio_id = str(audio_format.get("format_id") or "")
    if not audio_id or same_media_resource(video_format, audio_format) or format_has_video(audio_format):
        return video_id

    return f"{video_id}+{audio_id}/{video_id}/best"


def ydl_download_options(temp_path: Path, video_format: dict, audio_format: dict | None, progress_hook) -> dict:
    return {
        "quiet": True,
        "no_warnings": True,
        "noplaylist": True,
        "format": build_ydl_format_selector(video_format, audio_format),
        "outtmpl": str(temp_path / "source.%(ext)s"),
        "merge_output_format": "mp4",
        "progress_hooks": [progress_hook],
    }


class PortionDownloaderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("YouTube Portion Downloader")
        self.geometry(DEFAULT_WINDOW_SIZE)
        self.minsize(960, 720)

        self.settings_path = get_settings_path()
        self.settings = self.load_settings()
        self.ui_queue: queue.Queue[tuple] = queue.Queue()
        self.download_history = self.load_history()

        self.video_info: dict | None = None
        self.video_formats: list[dict] = []
        self.audio_formats: list[dict] = []
        self.video_options: list[dict] = []
        self.audio_options: list[dict] = []
        self.output_path: Path | None = None

        self.auto_filename = True
        self._updating_filename = False
        self._busy = False

        self.url_var = tk.StringVar()
        self.title_var = tk.StringVar(value="Load a YouTube URL to see the video title and formats.")
        self.duration_var = tk.StringVar(value="Duration: --")
        self.quality_hint_var = tk.StringVar(value="Fetch formats to unlock quality choices.")
        self.video_var = tk.StringVar()
        self.audio_var = tk.StringVar()
        self.start_var = tk.StringVar(value="00:00:00")
        self.end_var = tk.StringVar()
        self.clip_length_var = tk.StringVar(value="Clip length: set an end time")
        self.preview_slider_var = tk.DoubleVar(value=0.0)
        self.preview_time_var = tk.StringVar(value="Cursor: 00:00:00")
        self.filename_var = tk.StringVar()
        self.folder_var = tk.StringVar(value=str(self.initial_output_folder()))
        self.remember_folder_var = tk.BooleanVar(value=bool(self.settings.get("remember_folder")))
        self.status_var = tk.StringVar(value="Ready.")
        self.final_path_var = tk.StringVar(value="")

        self.configure_styles()
        self.build_ui()
        self.filename_var.trace_add("write", self.on_filename_changed)
        self.start_var.trace_add("write", self.on_time_changed)
        self.end_var.trace_add("write", self.on_time_changed)
        self.folder_var.trace_add("write", self.on_folder_changed)
        self.remember_folder_var.trace_add("write", self.on_folder_preference_changed)

        self.update_clip_length()
        self.set_busy(False)
        self.set_dependency_status()
        self.after(150, self.process_queue)

    def load_settings(self) -> dict:
        try:
            return json.loads(self.settings_path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return {}
        except json.JSONDecodeError:
            return {}

    def load_history(self) -> list[dict]:
        raw_history = self.settings.get("download_history", [])
        if not isinstance(raw_history, list):
            return []

        history = []
        for entry in raw_history[:HISTORY_LIMIT]:
            if not isinstance(entry, dict):
                continue
            path = str(entry.get("path") or "").strip()
            if not path:
                continue
            history.append(
                {
                    "path": path,
                    "name": str(entry.get("name") or Path(path).name),
                    "range": str(entry.get("range") or ""),
                    "created": str(entry.get("created") or ""),
                }
            )
        return history

    def save_settings(self) -> None:
        history = getattr(self, "download_history", [])
        payload = {
            "remember_folder": bool(self.remember_folder_var.get()),
            "output_folder": self.folder_var.get().strip() if self.remember_folder_var.get() else "",
            "download_history": history[:HISTORY_LIMIT],
        }
        self.settings_path.parent.mkdir(parents=True, exist_ok=True)
        self.settings_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def initial_output_folder(self) -> Path:
        saved = self.settings.get("output_folder")
        if self.settings.get("remember_folder") and saved:
            saved_path = Path(saved)
            if saved_path.exists():
                return saved_path
        return default_videos_folder()

    def configure_styles(self) -> None:
        self.configure(background=APP_BG)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        body_font = ("Segoe UI", 10)
        small_font = ("Segoe UI", 9)
        title_font = ("Segoe UI Semibold", 19)
        section_font = ("Segoe UI Semibold", 11)

        style.configure(".", font=body_font, background=APP_BG, foreground=INK)
        style.configure("App.TFrame", background=APP_BG)
        style.configure("Card.TFrame", background=CARD_BG, relief="flat")
        style.configure("Hero.TFrame", background=HERO_BG)

        style.configure("TLabel", background=APP_BG, foreground=INK)
        style.configure("Card.TLabel", background=CARD_BG, foreground=INK)
        style.configure("Muted.TLabel", background=CARD_BG, foreground=MUTED, font=small_font)
        style.configure("Section.TLabel", background=CARD_BG, foreground=INK, font=section_font)
        style.configure("HeroTitle.TLabel", background=HERO_BG, foreground="#fff7e8", font=title_font)
        style.configure("HeroSubtitle.TLabel", background=HERO_BG, foreground="#d6ebe7", font=("Segoe UI", 10))
        style.configure("HeroEyebrow.TLabel", background=HERO_BG, foreground="#f6d18b", font=("Segoe UI Semibold", 9))
        style.configure("HeroBadge.TLabel", background="#1b5550", foreground="#f8ead0", font=small_font, padding=(10, 5))
        style.configure("Pill.TLabel", background=SOFT_ACCENT, foreground=ACCENT_DARK, font=small_font, padding=(10, 5))
        style.configure("Status.TLabel", background=CARD_BG, foreground=INK)

        style.configure("TEntry", fieldbackground="#fffcf6", foreground=INK, bordercolor=BORDER, lightcolor=BORDER)
        style.configure("TCombobox", fieldbackground="#fffcf6", foreground=INK, bordercolor=BORDER, arrowcolor=ACCENT)
        style.configure("TCheckbutton", background=CARD_BG, foreground=INK)

        style.configure("TButton", padding=(12, 8), background="#efe6d8", foreground=INK, bordercolor=BORDER)
        style.map("TButton", background=[("active", "#e6dccb")])
        style.configure("Accent.TButton", background=ACCENT, foreground="white", bordercolor=ACCENT)
        style.map("Accent.TButton", background=[("active", ACCENT_DARK), ("disabled", "#9eb8b2")])
        style.configure("Ghost.TButton", background=CARD_BG, foreground=ACCENT_DARK, bordercolor=BORDER, padding=(10, 6))
        style.map("Ghost.TButton", background=[("active", SOFT_ACCENT)])

        style.configure(
            "Modern.Horizontal.TProgressbar",
            background=ACCENT,
            troughcolor="#e6dccb",
            bordercolor="#e6dccb",
            lightcolor=ACCENT,
            darkcolor=ACCENT,
        )

    def build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, style="App.TFrame", padding=18)
        root.grid(sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        hero = ttk.Frame(root, style="Hero.TFrame", padding=(20, 14))
        hero.grid(row=0, column=0, sticky="ew")
        hero.columnconfigure(0, weight=1)

        ttk.Label(hero, text="PORTION STUDIO", style="HeroEyebrow.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(hero, text="Clip the exact moment, not the whole video.", style="HeroTitle.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )
        ttk.Label(
            hero,
            text="Fetch quality options, trim by timestamp, export MP4, and keep recent clips close.",
            style="HeroSubtitle.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

        content = ttk.Frame(root, style="App.TFrame")
        content.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        content.columnconfigure(0, weight=3, uniform="content")
        content.columnconfigure(1, weight=2, uniform="content")
        content.rowconfigure(0, weight=1)

        left_column = ttk.Frame(content, style="App.TFrame")
        left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        left_column.columnconfigure(0, weight=1)

        right_column = ttk.Frame(content, style="App.TFrame")
        right_column.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        right_column.columnconfigure(0, weight=1)
        right_column.rowconfigure(2, weight=1)

        source_card = ttk.Frame(left_column, style="Card.TFrame", padding=18)
        source_card.grid(row=0, column=0, sticky="ew")
        source_card.columnconfigure(0, weight=1)

        ttk.Label(source_card, text="1. Add Source", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            source_card,
            text="Paste a YouTube, Instagram, or TikTok link, then fetch streams before trimming.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))

        url_row = ttk.Frame(source_card, style="Card.TFrame")
        url_row.grid(row=2, column=0, sticky="ew")
        url_row.columnconfigure(0, weight=1)
        url_entry = ttk.Entry(url_row, textvariable=self.url_var)
        url_entry.grid(row=0, column=0, sticky="ew")
        url_entry.bind("<Return>", lambda _event: self.fetch_formats())
        ttk.Button(url_row, text="Paste", style="Ghost.TButton", command=self.paste_url).grid(
            row=0, column=1, padx=(10, 0)
        )
        self.fetch_button = ttk.Button(
            url_row,
            text="Fetch Formats",
            style="Accent.TButton",
            command=self.fetch_formats,
        )
        self.fetch_button.grid(row=0, column=2, padx=(10, 0))

        title_row = ttk.Frame(source_card, style="Card.TFrame")
        title_row.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        title_row.columnconfigure(0, weight=1)
        ttk.Label(title_row, textvariable=self.title_var, wraplength=570, justify="left", style="Card.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(title_row, textvariable=self.duration_var, style="Pill.TLabel").grid(
            row=0, column=1, sticky="e", padx=(10, 0)
        )

        formats_card = ttk.Frame(left_column, style="Card.TFrame", padding=18)
        formats_card.grid(row=1, column=0, sticky="ew", pady=(18, 0))
        formats_card.columnconfigure(1, weight=1)

        ttk.Label(formats_card, text="3. Choose Quality", style="Section.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(formats_card, textvariable=self.quality_hint_var, style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 12)
        )

        ttk.Label(formats_card, text="Video", style="Card.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 12))
        self.video_combo = ttk.Combobox(
            formats_card,
            textvariable=self.video_var,
            state="readonly",
            values=["Fetch formats first"],
        )
        self.video_combo.grid(row=2, column=1, sticky="ew")
        self.video_combo.set("Fetch formats first")

        ttk.Label(formats_card, text="Audio", style="Card.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 12), pady=(12, 0)
        )
        self.audio_combo = ttk.Combobox(
            formats_card,
            textvariable=self.audio_var,
            state="readonly",
            values=["Fetch formats first"],
        )
        self.audio_combo.grid(row=3, column=1, sticky="ew", pady=(12, 0))
        self.audio_combo.set("Fetch formats first")

        range_card = ttk.Frame(right_column, style="Card.TFrame", padding=18)
        range_card.grid(row=0, column=0, sticky="ew")
        range_card.columnconfigure(0, weight=1)

        range_header = ttk.Frame(range_card, style="Card.TFrame")
        range_header.grid(row=0, column=0, sticky="ew")
        range_header.columnconfigure(0, weight=1)
        ttk.Label(range_header, text="2. Set Clip Range", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(range_header, textvariable=self.clip_length_var, style="Pill.TLabel").grid(
            row=0, column=1, sticky="e"
        )
        ttk.Label(
            range_card,
            text="Use the fields, quick buttons, or preview cursor to choose the exact portion.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))

        time_grid = ttk.Frame(range_card, style="Card.TFrame")
        time_grid.grid(row=2, column=0, sticky="ew")
        time_grid.columnconfigure(1, weight=1)
        time_grid.columnconfigure(3, weight=1)
        ttk.Label(time_grid, text="Start", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(time_grid, textvariable=self.start_var).grid(row=0, column=1, sticky="ew")
        ttk.Label(time_grid, text="End", style="Card.TLabel").grid(row=0, column=2, sticky="w", padx=(14, 10))
        ttk.Entry(time_grid, textvariable=self.end_var).grid(row=0, column=3, sticky="ew")

        self.preview_scale = ttk.Scale(
            range_card,
            from_=0,
            to=0,
            orient="horizontal",
            variable=self.preview_slider_var,
            command=self.on_preview_slider_changed,
        )
        self.preview_scale.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        self.preview_scale.configure(state="disabled")

        preview_row = ttk.Frame(range_card, style="Card.TFrame")
        preview_row.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        preview_row.columnconfigure(0, weight=1)
        ttk.Label(preview_row, textvariable=self.preview_time_var, style="Muted.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.set_start_button = ttk.Button(
            preview_row,
            text="Set Start",
            style="Ghost.TButton",
            command=self.set_start_from_preview,
            state="disabled",
        )
        self.set_start_button.grid(row=0, column=1, sticky="e", padx=(8, 0))
        self.set_end_button = ttk.Button(
            preview_row,
            text="Set End",
            style="Ghost.TButton",
            command=self.set_end_from_preview,
            state="disabled",
        )
        self.set_end_button.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.open_preview_button = ttk.Button(
            preview_row,
            text="Open At Cursor",
            style="Ghost.TButton",
            command=self.open_source_at_preview,
            state="disabled",
        )
        self.open_preview_button.grid(row=0, column=3, sticky="e", padx=(8, 0))

        quick_row = ttk.Frame(range_card, style="Card.TFrame")
        quick_row.grid(row=5, column=0, sticky="w", pady=(14, 0))
        ttk.Button(quick_row, text="+15 sec", style="Ghost.TButton", command=lambda: self.apply_duration_preset(15)).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(quick_row, text="+30 sec", style="Ghost.TButton", command=lambda: self.apply_duration_preset(30)).grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )
        ttk.Button(quick_row, text="+60 sec", style="Ghost.TButton", command=lambda: self.apply_duration_preset(60)).grid(
            row=0, column=2, sticky="w", padx=(8, 0)
        )
        ttk.Button(quick_row, text="Use Full Video", style="Ghost.TButton", command=self.use_full_duration).grid(
            row=0, column=3, sticky="w", padx=(8, 0)
        )

        output_card = ttk.Frame(right_column, style="Card.TFrame", padding=18)
        output_card.grid(row=1, column=0, sticky="ew", pady=(18, 0))
        output_card.columnconfigure(0, weight=1)

        ttk.Label(output_card, text="4. Save Clip", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(output_card, text="Pick a file name and destination. The app keeps names Windows-safe.", style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 12)
        )

        ttk.Label(output_card, text="Filename", style="Card.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(output_card, textvariable=self.filename_var).grid(row=3, column=0, sticky="ew", pady=(6, 12))

        ttk.Label(output_card, text="Folder", style="Card.TLabel").grid(row=4, column=0, sticky="w")
        folder_row = ttk.Frame(output_card, style="Card.TFrame")
        folder_row.grid(row=5, column=0, sticky="ew", pady=(6, 0))
        folder_row.columnconfigure(0, weight=1)
        ttk.Entry(folder_row, textvariable=self.folder_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(folder_row, text="Browse", style="Ghost.TButton", command=self.choose_folder).grid(
            row=0, column=1, padx=(10, 0)
        )

        ttk.Checkbutton(output_card, text="Remember this folder", variable=self.remember_folder_var).grid(
            row=6, column=0, sticky="w", pady=(12, 0)
        )
        self.download_button = ttk.Button(
            output_card,
            text="Download Section",
            style="Accent.TButton",
            command=self.start_download,
        )
        self.download_button.grid(row=7, column=0, sticky="ew", pady=(18, 0))

        history_card = ttk.Frame(right_column, style="Card.TFrame", padding=18)
        history_card.grid(row=2, column=0, sticky="nsew", pady=(18, 0))
        history_card.columnconfigure(0, weight=1)
        history_card.rowconfigure(2, weight=1)

        ttk.Label(history_card, text="Recent Downloads", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(history_card, text="Reopen or copy paths from the last few finished clips.", style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 12)
        )

        history_body = ttk.Frame(history_card, style="Card.TFrame")
        history_body.grid(row=2, column=0, sticky="nsew")
        history_body.columnconfigure(0, weight=1)
        history_body.rowconfigure(0, weight=1)

        self.history_listbox = tk.Listbox(
            history_body,
            height=7,
            bg="#fffcf6",
            fg=INK,
            selectbackground=ACCENT,
            selectforeground="white",
            highlightthickness=1,
            highlightbackground=BORDER,
            relief="flat",
            activestyle="none",
            font=("Segoe UI", 9),
        )
        self.history_listbox.grid(row=0, column=0, sticky="nsew")
        self.history_listbox.bind("<<ListboxSelect>>", lambda _event: self.update_history_buttons())
        self.history_listbox.bind("<Double-Button-1>", lambda _event: self.open_history_file())

        history_scroll = ttk.Scrollbar(history_body, orient="vertical", command=self.history_listbox.yview)
        history_scroll.grid(row=0, column=1, sticky="ns")
        self.history_listbox.configure(yscrollcommand=history_scroll.set)

        history_buttons = ttk.Frame(history_card, style="Card.TFrame")
        history_buttons.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self.open_history_button = ttk.Button(
            history_buttons,
            text="Open",
            style="Ghost.TButton",
            command=self.open_history_file,
            state="disabled",
        )
        self.open_history_button.grid(row=0, column=0, sticky="w")
        self.open_history_folder_button = ttk.Button(
            history_buttons,
            text="Folder",
            style="Ghost.TButton",
            command=self.open_history_folder,
            state="disabled",
        )
        self.open_history_folder_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.copy_history_button = ttk.Button(
            history_buttons,
            text="Copy Path",
            style="Ghost.TButton",
            command=self.copy_history_path,
            state="disabled",
        )
        self.copy_history_button.grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.clear_history_button = ttk.Button(
            history_buttons,
            text="Clear",
            style="Ghost.TButton",
            command=self.clear_history,
        )
        self.clear_history_button.grid(row=0, column=3, sticky="w", padx=(8, 0))

        status_card = ttk.Frame(root, style="Card.TFrame", padding=18)
        status_card.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        status_card.columnconfigure(1, weight=1)

        ttk.Label(status_card, text="Status", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.progress_bar = ttk.Progressbar(
            status_card,
            mode="determinate",
            maximum=100,
            style="Modern.Horizontal.TProgressbar",
        )
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=(16, 0))
        ttk.Label(status_card, textvariable=self.status_var, wraplength=860, justify="left", style="Status.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(12, 0)
        )

        final_row = ttk.Frame(status_card, style="Card.TFrame")
        final_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        final_row.columnconfigure(1, weight=1)
        ttk.Label(final_row, text="Final file", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(final_row, textvariable=self.final_path_var, state="readonly").grid(row=0, column=1, sticky="ew")

        self.open_file_button = ttk.Button(final_row, text="Open File", style="Ghost.TButton", command=self.open_file, state="disabled")
        self.open_file_button.grid(row=0, column=2, sticky="e", padx=(10, 0))

        self.open_folder_button = ttk.Button(
            final_row,
            text="Open Folder",
            style="Ghost.TButton",
            command=self.open_folder,
            state="disabled",
        )
        self.open_folder_button.grid(row=0, column=3, sticky="e", padx=(8, 0))

        self.copy_path_button = ttk.Button(
            final_row,
            text="Copy Path",
            style="Ghost.TButton",
            command=self.copy_final_path,
            state="disabled",
        )
        self.copy_path_button.grid(row=0, column=4, sticky="e", padx=(8, 0))

        self.refresh_history()

    def set_dependency_status(self) -> None:
        warnings = []
        if yt_dlp is None:
            warnings.append("yt-dlp is not installed yet, so format loading will fail until you run pip install -r requirements.txt.")
        if not shutil.which("ffmpeg"):
            warnings.append("ffmpeg was not found in PATH yet, so downloads will fail until it is installed.")

        if warnings:
            self.status_var.set("Ready. " + " ".join(warnings))
        else:
            self.status_var.set("Ready.")

    def on_filename_changed(self, *_args) -> None:
        if self._updating_filename:
            return
        self.auto_filename = False

    def on_time_changed(self, *_args) -> None:
        self.update_clip_length()
        if self.auto_filename and self.video_info:
            self.set_auto_filename()

    def on_folder_preference_changed(self, *_args) -> None:
        self.save_settings()

    def on_folder_changed(self, *_args) -> None:
        if self.remember_folder_var.get():
            self.save_settings()

    def update_clip_length(self) -> None:
        if not self.end_var.get().strip():
            self.clip_length_var.set("Clip length: set an end time")
            return

        try:
            start = parse_timestamp(self.start_var.get())
            end = parse_timestamp(self.end_var.get())
        except ValueError:
            self.clip_length_var.set("Clip length: waiting for a valid range")
            return

        if end <= start:
            self.clip_length_var.set("Clip length: end must be after start")
            return

        self.clip_length_var.set(f"Clip length: {format_seconds_for_display(end - start)}")

    def paste_url(self) -> None:
        try:
            text = self.clipboard_get().strip()
        except tk.TclError:
            self.status_var.set("Clipboard is empty.")
            return

        if not text:
            self.status_var.set("Clipboard is empty.")
            return

        self.url_var.set(text)
        self.status_var.set("URL pasted. Fetch formats when ready.")

    def apply_duration_preset(self, seconds: int) -> None:
        try:
            start = parse_timestamp(self.start_var.get())
        except ValueError:
            start = 0.0
            self.start_var.set(format_seconds_for_display(start))

        duration = self.video_info.get("duration") if self.video_info else None
        if duration and start >= float(duration):
            self.status_var.set("Start time is beyond the loaded video duration.")
            return

        end = start + seconds
        if duration:
            end = min(end, float(duration))

        self.end_var.set(format_seconds_for_display(end))
        self.status_var.set(f"Set a {seconds}-second clip from the current start time.")

    def use_full_duration(self) -> None:
        duration = self.video_info.get("duration") if self.video_info else None
        if not duration:
            self.status_var.set("Fetch formats first so the app knows the full duration.")
            return

        self.start_var.set("00:00:00")
        self.end_var.set(format_seconds_for_display(float(duration)))
        self.status_var.set("Range set to the full loaded video.")

    def on_preview_slider_changed(self, value: str) -> None:
        try:
            seconds = float(value)
        except ValueError:
            seconds = 0.0
        self.preview_time_var.set(f"Cursor: {format_seconds_for_display(seconds)}")

    def set_preview_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.preview_scale.configure(state=state)
        self.set_start_button.configure(state=state)
        self.set_end_button.configure(state=state)
        self.open_preview_button.configure(state=state)

    def reset_preview_selector(self) -> None:
        self.preview_scale.configure(to=0)
        self.preview_slider_var.set(0.0)
        self.preview_time_var.set("Cursor: 00:00:00")
        self.set_preview_controls_enabled(False)

    def configure_preview_selector(self, duration: float | None) -> None:
        if not duration or duration <= 0:
            self.reset_preview_selector()
            return

        self.preview_scale.configure(to=float(duration))
        self.preview_slider_var.set(0.0)
        self.preview_time_var.set("Cursor: 00:00:00")
        self.set_preview_controls_enabled(True)

    def preview_seconds(self) -> float:
        return max(0.0, float(self.preview_slider_var.get()))

    def set_start_from_preview(self) -> None:
        self.start_var.set(format_seconds_for_display(self.preview_seconds()))
        self.status_var.set("Start time set from the preview cursor.")

    def set_end_from_preview(self) -> None:
        self.end_var.set(format_seconds_for_display(self.preview_seconds()))
        self.status_var.set("End time set from the preview cursor.")

    def source_url(self) -> str:
        return (self.video_info.get("webpage_url") if self.video_info else "") or self.url_var.get().strip()

    def open_source_at_preview(self) -> None:
        url = self.source_url()
        if not url:
            self.status_var.set("Load a source first.")
            return

        webbrowser.open(with_timestamp(url, self.preview_seconds()))
        if supports_timestamp_url(url):
            self.status_var.set("Opened the source near the preview cursor.")
        else:
            self.status_var.set("Opened the source page. This site does not support timestamp links reliably.")

    def copy_text_to_clipboard(self, text: str, status: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set(status)

    def copy_final_path(self) -> None:
        path = self.final_path_var.get().strip()
        if path:
            self.copy_text_to_clipboard(path, "Final file path copied.")

    def history_label(self, entry: dict) -> str:
        created = entry.get("created") or "recent"
        range_text = entry.get("range") or "clip"
        name = entry.get("name") or Path(entry.get("path") or "").name
        return f"{created}  |  {range_text}  |  {name}"

    def refresh_history(self) -> None:
        self.history_listbox.delete(0, tk.END)
        if not self.download_history:
            self.history_listbox.insert(tk.END, "No downloads yet")
        else:
            for entry in self.download_history:
                self.history_listbox.insert(tk.END, self.history_label(entry))
        self.update_history_buttons()

    def update_history_buttons(self) -> None:
        has_selection = self.selected_history_entry() is not None
        state = "normal" if has_selection else "disabled"
        self.open_history_button.configure(state=state)
        self.open_history_folder_button.configure(state=state)
        self.copy_history_button.configure(state=state)
        self.clear_history_button.configure(state="normal" if self.download_history else "disabled")

    def selected_history_entry(self) -> dict | None:
        selection = self.history_listbox.curselection()
        if not selection:
            return None
        index = selection[0]
        if index >= len(self.download_history):
            return None
        return self.download_history[index]

    def add_to_history(self, path: Path) -> None:
        try:
            start = parse_timestamp(self.start_var.get())
            end = parse_timestamp(self.end_var.get())
            range_text = f"{format_seconds_for_display(start)} to {format_seconds_for_display(end)}"
        except ValueError:
            range_text = "clip"

        path_text = str(path)
        entry = {
            "path": path_text,
            "name": path.name,
            "range": range_text,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        self.download_history = [item for item in self.download_history if item.get("path") != path_text]
        self.download_history.insert(0, entry)
        self.download_history = self.download_history[:HISTORY_LIMIT]
        self.save_settings()
        self.refresh_history()

    def open_history_file(self) -> None:
        entry = self.selected_history_entry()
        if not entry:
            return

        path = Path(entry["path"])
        if not path.exists():
            messagebox.showwarning("Missing File", "That recent file no longer exists at its saved path.")
            return
        os.startfile(str(path))  # type: ignore[attr-defined]

    def open_history_folder(self) -> None:
        entry = self.selected_history_entry()
        if not entry:
            return

        path = Path(entry["path"])
        folder = path.parent
        if not folder.exists():
            messagebox.showwarning("Missing Folder", "That recent file's folder no longer exists.")
            return
        os.startfile(str(folder))  # type: ignore[attr-defined]

    def copy_history_path(self) -> None:
        entry = self.selected_history_entry()
        if entry:
            self.copy_text_to_clipboard(entry["path"], "Recent file path copied.")

    def clear_history(self) -> None:
        if not self.download_history:
            return
        if not messagebox.askyesno("Clear Recent Downloads", "Clear the recent downloads list?"):
            return
        self.download_history = []
        self.save_settings()
        self.refresh_history()
        self.status_var.set("Recent downloads cleared.")

    def choose_folder(self) -> None:
        initial_dir = self.folder_var.get().strip() or str(default_videos_folder())
        chosen = filedialog.askdirectory(initialdir=initial_dir)
        if not chosen:
            return
        self.folder_var.set(chosen)
        if self.remember_folder_var.get():
            self.save_settings()

    def fetch_formats(self) -> None:
        if self._busy:
            return
        if yt_dlp is None:
            messagebox.showerror(
                "Missing yt-dlp",
                "yt-dlp is not installed. Run: python -m pip install -r requirements.txt",
            )
            return

        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("Missing URL", "Enter a YouTube URL first.")
            return

        self.output_path = None
        self.final_path_var.set("")
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.copy_path_button.configure(state="disabled")
        self.reset_preview_selector()
        self.set_busy(True)
        self.start_indeterminate_progress()
        self.status_var.set("Fetching formats...")

        worker = threading.Thread(target=self._fetch_formats_worker, args=(url,), daemon=True)
        worker.start()

    def _fetch_formats_worker(self, url: str) -> None:
        try:
            with yt_dlp.YoutubeDL(ydl_options()) as ydl:
                info = ydl.extract_info(url, download=False)

            if not info:
                raise RuntimeError("No video information was returned.")

            formats = info.get("formats") or []
            video_formats = collect_video_formats(formats)
            audio_formats = collect_audio_formats(formats)

            if not video_formats:
                raise RuntimeError("No downloadable video formats were found.")
            if not audio_formats:
                raise RuntimeError("No downloadable audio formats were found.")

            self.ui_queue.put(("formats_loaded", info, video_formats, audio_formats))
        except Exception as exc:
            self.ui_queue.put(("error", f"Could not fetch formats: {exc}"))

    def process_queue(self) -> None:
        try:
            while True:
                item = self.ui_queue.get_nowait()
                kind = item[0]

                if kind == "formats_loaded":
                    _, info, video_formats, audio_formats = item
                    self.on_formats_loaded(info, video_formats, audio_formats)
                elif kind == "status":
                    _, text = item
                    self.status_var.set(text)
                elif kind == "progress":
                    _, value = item
                    self.progress_bar.configure(mode="determinate")
                    self.progress_bar.stop()
                    self.progress_bar["value"] = max(0, min(100, value))
                elif kind == "error":
                    _, message = item
                    self.finish_busy()
                    self.status_var.set(message)
                    messagebox.showerror("Error", message)
                elif kind == "done":
                    _, path = item
                    self.on_download_finished(Path(path))
        except queue.Empty:
            pass

        self.after(150, self.process_queue)

    def on_formats_loaded(self, info: dict, video_formats: list[dict], audio_formats: list[dict]) -> None:
        self.video_info = info
        self.video_formats = video_formats
        self.audio_formats = audio_formats

        title = info.get("title") or "Untitled video"
        duration = info.get("duration")
        if duration and not self.end_var.get().strip():
            self.end_var.set(format_seconds_for_display(float(duration)))

        duration_text = f" ({format_seconds_for_display(float(duration))})" if duration else ""
        self.duration_var.set(f"Duration: {format_seconds_for_display(float(duration))}" if duration else "Duration: --")
        self.title_var.set(f"{title}{duration_text}")
        self.configure_preview_selector(float(duration) if duration else None)
        self.quality_hint_var.set(
            f"Loaded {len(video_formats)} video streams and {len(audio_formats)} audio streams. Best picks are selected."
        )

        self.video_options = [{"kind": "best", "label": "Best video"}]
        self.video_options.extend(
            {"kind": "format", "format_id": fmt.get("format_id"), "label": build_video_label(fmt)}
            for fmt in video_formats
        )

        self.audio_options = [
            {"kind": "best", "label": "Best audio"},
            {"kind": "none", "label": "No audio"},
        ]
        self.audio_options.extend(
            {"kind": "format", "format_id": fmt.get("format_id"), "label": build_audio_label(fmt)}
            for fmt in audio_formats
        )

        self.video_combo["values"] = [option["label"] for option in self.video_options]
        self.audio_combo["values"] = [option["label"] for option in self.audio_options]
        self.video_combo.current(0)
        self.audio_combo.current(0)

        self.auto_filename = True
        self.set_auto_filename()

        self.finish_busy()
        self.status_var.set("Formats loaded. Pick the resolution, audio, and time range.")

    def set_busy(self, busy: bool) -> None:
        self._busy = busy
        fetch_state = "disabled" if busy else "normal"
        download_state = "disabled" if busy or not self.video_info else "normal"
        self.fetch_button.configure(state=fetch_state)
        self.download_button.configure(state=download_state)
        self.video_combo.configure(state="disabled" if busy or not self.video_options else "readonly")
        self.audio_combo.configure(state="disabled" if busy or not self.audio_options else "readonly")

    def finish_busy(self) -> None:
        self.stop_indeterminate_progress()
        self.set_busy(False)

    def start_indeterminate_progress(self) -> None:
        self.progress_bar.stop()
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start(10)

    def stop_indeterminate_progress(self) -> None:
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")

    def set_auto_filename(self) -> None:
        if not self.video_info:
            return

        title = self.video_info.get("title") or "clip"
        try:
            start = parse_timestamp(self.start_var.get())
        except ValueError:
            start = 0.0

        try:
            end = parse_timestamp(self.end_var.get())
        except ValueError:
            duration = self.video_info.get("duration")
            end = float(duration) if duration else start

        if end <= start:
            duration = self.video_info.get("duration")
            end = float(duration) if duration and float(duration) > start else start

        suggested = build_default_filename(title, start, end)
        self._updating_filename = True
        self.filename_var.set(suggested)
        self._updating_filename = False

    def selected_option(self, combo: ttk.Combobox, options: list[dict]) -> dict:
        index = combo.current()
        if index < 0 and options:
            return options[0]
        return options[index]

    def choose_video_format(self, option: dict, video_formats: list[dict]) -> dict:
        if option["kind"] == "best":
            return video_formats[0]
        for fmt in video_formats:
            if fmt.get("format_id") == option.get("format_id"):
                return fmt
        raise RuntimeError("Selected video format is no longer available.")

    def choose_audio_format(self, option: dict, audio_formats: list[dict]) -> dict | None:
        if option["kind"] == "none":
            return None
        if option["kind"] == "best":
            return audio_formats[0]
        for fmt in audio_formats:
            if fmt.get("format_id") == option.get("format_id"):
                return fmt
        raise RuntimeError("Selected audio format is no longer available.")

    def start_download(self) -> None:
        if self._busy:
            return
        if not self.video_info:
            messagebox.showerror("No Video Loaded", "Fetch formats before downloading.")
            return

        try:
            start = parse_timestamp(self.start_var.get())
            end = parse_timestamp(self.end_var.get())
        except ValueError as exc:
            messagebox.showerror("Invalid Timestamp", str(exc))
            return

        if end <= start:
            messagebox.showerror("Invalid Range", "End time must be greater than start time.")
            return

        folder_text = self.folder_var.get().strip()
        output_folder = Path(folder_text) if folder_text else default_videos_folder()
        output_folder.mkdir(parents=True, exist_ok=True)

        raw_name = self.filename_var.get().strip()
        if not raw_name:
            raw_name = build_default_filename(self.video_info.get("title") or "clip", start, end)
            self._updating_filename = True
            self.filename_var.set(raw_name)
            self._updating_filename = False

        final_name = sanitize_filename(raw_name) + ".mp4"
        output_path = ensure_unique_path(output_folder / final_name)
        video_option = self.selected_option(self.video_combo, self.video_options)
        audio_option = self.selected_option(self.audio_combo, self.audio_options)
        source_url = self.source_url()

        if self.remember_folder_var.get():
            self.save_settings()

        self.output_path = None
        self.final_path_var.set("")
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.copy_path_button.configure(state="disabled")
        self.set_busy(True)
        self.progress_bar["value"] = 0
        self.status_var.set("Preparing selected section...")

        worker = threading.Thread(
            target=self._download_worker,
            args=(start, end, video_option, audio_option, output_path, source_url),
            daemon=True,
        )
        worker.start()

    def _download_worker(
        self,
        start: float,
        end: float,
        video_option: dict,
        audio_option: dict,
        output_path: Path,
        source_url: str,
    ) -> None:
        ffmpeg_path = shutil.which("ffmpeg")
        if not ffmpeg_path:
            self.ui_queue.put(("error", "ffmpeg was not found in PATH. Install ffmpeg and try again."))
            return

        try:
            duration = end - start
            if yt_dlp is None:
                raise RuntimeError("yt-dlp is not installed. Run: python -m pip install -r requirements.txt")

            self.ui_queue.put(("status", "Refreshing stream URLs..."))
            with yt_dlp.YoutubeDL(ydl_options()) as ydl:
                info = ydl.extract_info(source_url, download=False)

            video_formats = collect_video_formats(info.get("formats") or [])
            audio_formats = collect_audio_formats(info.get("formats") or [])
            if not video_formats:
                raise RuntimeError("No downloadable video formats were found.")
            if audio_option["kind"] != "none" and not audio_formats:
                raise RuntimeError("No downloadable audio formats were found.")

            video_format = self.choose_video_format(video_option, video_formats)
            audio_format = self.choose_audio_format(audio_option, audio_formats)

            if prefers_local_download(source_url):
                self.download_source_then_clip(
                    source_url=source_url,
                    video_format=video_format,
                    audio_format=audio_format,
                    start=start,
                    duration=duration,
                    output_path=output_path,
                    ffmpeg_path=ffmpeg_path,
                )
            else:
                self.ui_queue.put(("status", "Starting direct stream clip..."))
                command = self.build_ffmpeg_command(
                    ffmpeg_path=ffmpeg_path,
                    video_format=video_format,
                    audio_format=audio_format,
                    start=start,
                    duration=duration,
                    output_path=output_path,
                )
                try:
                    self.run_ffmpeg_with_progress(command, duration, "Downloading and converting")
                except RuntimeError as exc:
                    if not self.should_retry_with_local_source(exc, source_url):
                        raise
                    self.ui_queue.put(("status", "Direct stream was blocked. Downloading source first..."))
                    self.download_source_then_clip(
                        source_url=source_url,
                        video_format=video_format,
                        audio_format=audio_format,
                        start=start,
                        duration=duration,
                        output_path=output_path,
                        ffmpeg_path=ffmpeg_path,
                    )

            self.ui_queue.put(("done", str(output_path)))
        except Exception as exc:
            self.ui_queue.put(("error", f"Download failed: {exc}"))

    def should_retry_with_local_source(self, exc: RuntimeError, source_url: str) -> bool:
        message = str(exc).lower()
        blocked = "403" in message or "forbidden" in message or "access denied" in message
        return blocked or prefers_local_download(source_url)

    def run_ffmpeg_with_progress(
        self,
        command: list[str],
        duration: float,
        status_label: str,
        progress_start: float = 0.0,
        progress_span: float = 100.0,
    ) -> None:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )

        if not process.stdout:
            raise RuntimeError("Could not read ffmpeg progress output.")

        for line in process.stdout:
            text = line.strip()
            if not text or "=" not in text:
                continue
            key, value = text.split("=", 1)
            if key == "out_time":
                try:
                    current = parse_timestamp(value)
                except ValueError:
                    continue
                percent = (current / duration) * 100 if duration > 0 else 0
                percent = min(percent, 100)
                overall = progress_start + (percent * progress_span / 100)
                self.ui_queue.put(("progress", overall))
                self.ui_queue.put(("status", f"{status_label}... {percent:.1f}%"))
            elif key == "progress" and value == "end":
                self.ui_queue.put(("progress", progress_start + progress_span))

        return_code = process.wait()
        stderr_text = process.stderr.read().strip() if process.stderr else ""
        if return_code != 0:
            detail = stderr_text or "ffmpeg exited with a non-zero status."
            raise RuntimeError(detail)

    def download_source_then_clip(
        self,
        source_url: str,
        video_format: dict,
        audio_format: dict | None,
        start: float,
        duration: float,
        output_path: Path,
        ffmpeg_path: str,
    ) -> None:
        with tempfile.TemporaryDirectory(prefix="portion_downloader_") as temp_dir:
            temp_path = Path(temp_dir)
            self.ui_queue.put(("status", "Downloading protected source with yt-dlp..."))
            download_opts = ydl_download_options(
                temp_path=temp_path,
                video_format=video_format,
                audio_format=audio_format,
                progress_hook=self.ytdlp_progress_hook,
            )
            with yt_dlp.YoutubeDL(download_opts) as ydl:
                downloaded_info = ydl.extract_info(source_url, download=True)

            source_file = self.find_downloaded_media(temp_path, downloaded_info)
            self.ui_queue.put(("status", "Trimming downloaded source..."))
            command = self.build_local_clip_command(
                ffmpeg_path=ffmpeg_path,
                source_file=source_file,
                include_audio=audio_format is not None,
                start=start,
                duration=duration,
                output_path=output_path,
            )
            self.run_ffmpeg_with_progress(command, duration, "Trimming local source", 65, 35)

    def ytdlp_progress_hook(self, data: dict) -> None:
        status = data.get("status")
        if status == "downloading":
            total = data.get("total_bytes") or data.get("total_bytes_estimate")
            downloaded = data.get("downloaded_bytes") or 0
            if total:
                percent = min((downloaded / total) * 100, 100)
                self.ui_queue.put(("progress", percent * 0.65))
                self.ui_queue.put(("status", f"Downloading protected source... {percent:.1f}%"))
            else:
                self.ui_queue.put(("status", "Downloading protected source..."))
        elif status == "finished":
            self.ui_queue.put(("progress", 65))
            self.ui_queue.put(("status", "Source downloaded. Preparing trim..."))

    def find_downloaded_media(self, temp_path: Path, downloaded_info: dict) -> Path:
        for download in downloaded_info.get("requested_downloads") or []:
            filepath = download.get("filepath")
            if filepath and Path(filepath).exists():
                return Path(filepath)

        filename = downloaded_info.get("_filename")
        if filename and Path(filename).exists():
            return Path(filename)

        media_files = [
            path
            for path in temp_path.iterdir()
            if path.is_file() and path.suffix.lower() not in {".part", ".ytdl", ".json"}
        ]
        if not media_files:
            raise RuntimeError("yt-dlp finished but no downloaded media file was found.")
        return max(media_files, key=lambda path: path.stat().st_size)

    def build_ffmpeg_command(
        self,
        ffmpeg_path: str,
        video_format: dict,
        audio_format: dict | None,
        start: float,
        duration: float,
        output_path: Path,
    ) -> list[str]:
        command = [ffmpeg_path, "-y", "-hide_banner", "-loglevel", "error"]
        command.extend(self.ffmpeg_input_args(video_format, start, duration))

        single_input_audio = bool(
            audio_format and (same_media_resource(video_format, audio_format) or format_has_video(audio_format))
        )
        if audio_format and not single_input_audio:
            command.extend(self.ffmpeg_input_args(audio_format, start, duration))

        command.extend(["-map", "0:v:0"])
        if audio_format:
            command.extend(["-map", "0:a:0?" if single_input_audio else "1:a:0"])
        else:
            command.append("-an")

        can_copy_video = str(video_format.get("vcodec") or "").startswith("avc1")
        can_copy_audio = False
        if audio_format:
            acodec = str(audio_format.get("acodec") or "")
            can_copy_audio = acodec.startswith("mp4a") or acodec.startswith("aac")

        if can_copy_video and (can_copy_audio or not audio_format):
            command.extend(["-c", "copy"])
        else:
            command.extend(["-c:v", "libx264", "-preset", "veryfast", "-crf", "20"])
            if audio_format:
                command.extend(["-c:a", "aac", "-b:a", "192k"])

        command.extend(["-movflags", "+faststart", "-progress", "pipe:1", "-nostats", str(output_path)])
        return command

    def build_local_clip_command(
        self,
        ffmpeg_path: str,
        source_file: Path,
        include_audio: bool,
        start: float,
        duration: float,
        output_path: Path,
    ) -> list[str]:
        command = [
            ffmpeg_path,
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-ss",
            ffmpeg_time(start),
            "-t",
            ffmpeg_time(duration),
            "-i",
            str(source_file),
            "-map",
            "0:v:0",
        ]
        if include_audio:
            command.extend(["-map", "0:a:0?"])
        else:
            command.append("-an")

        command.extend(["-c:v", "libx264", "-preset", "veryfast", "-crf", "20"])
        if include_audio:
            command.extend(["-c:a", "aac", "-b:a", "192k"])

        command.extend(["-movflags", "+faststart", "-progress", "pipe:1", "-nostats", str(output_path)])
        return command

    def ffmpeg_input_args(self, fmt: dict, start: float, duration: float) -> list[str]:
        url = fmt.get("url") or fmt.get("manifest_url")
        if not url:
            raise RuntimeError(f"Format {fmt.get('format_id')} does not have a usable media URL.")

        args = ["-ss", ffmpeg_time(start), "-t", ffmpeg_time(duration)]
        headers = fmt.get("http_headers") or (self.video_info.get("http_headers") if self.video_info else {})
        user_agent = ""
        if headers:
            user_agent = next((value for key, value in headers.items() if key.lower() == "user-agent"), "")
        headers_blob = build_headers_blob(headers)

        if user_agent:
            args.extend(["-user_agent", user_agent])
        if headers_blob:
            args.extend(["-headers", headers_blob])

        args.extend(["-i", url])
        return args

    def on_download_finished(self, path: Path) -> None:
        self.finish_busy()
        self.output_path = path
        self.final_path_var.set(str(path))
        self.open_file_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        self.copy_path_button.configure(state="normal")
        self.progress_bar["value"] = 100
        self.add_to_history(path)
        self.status_var.set("Done. Saved clip and added it to Recent Downloads.")

    def open_file(self) -> None:
        if self.output_path and self.output_path.exists():
            os.startfile(str(self.output_path))  # type: ignore[attr-defined]

    def open_folder(self) -> None:
        if self.output_path:
            folder = self.output_path.parent if self.output_path.exists() else Path(self.final_path_var.get()).parent
            os.startfile(str(folder))  # type: ignore[attr-defined]


if __name__ == "__main__":
    app = PortionDownloaderApp()
    app.mainloop()
