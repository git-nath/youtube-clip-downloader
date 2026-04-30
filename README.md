# Portion YouTube Video Downloader

Simple Windows desktop app for downloading only a selected portion of a YouTube video with `yt-dlp` and `ffmpeg`.

## What it does

- accepts a YouTube URL
- provides a modern card-based desktop UI with a quick paste button
- fetches available video and audio formats
- lets you choose video resolution and audio format, including `Best` options
- clips only the selected start/end section
- shows live clip length feedback while you edit timestamps
- includes quick range buttons for 15, 30, 60 seconds, or the full loaded video
- outputs an `.mp4` file and converts when a direct MP4 copy is not possible
- lets you choose and remember an output folder
- suggests a Windows-safe filename based on the video title and selected time range
- shows status, progress, final file path, and buttons to open or copy the file path
- keeps a small recent-downloads list so you can reopen or copy previous clip paths

## Requirements

- Windows
- Python 3.11+ recommended
- `ffmpeg` installed and available on `PATH`

## Setup

```powershell
py -3 -m pip install -r requirements.txt
```

Make sure `ffmpeg` works in a terminal:

```powershell
ffmpeg -version
```

## Run

```powershell
py -3 app.py
```

If `py -3` says no installed Python was found, install Python for Windows first and then rerun the commands above.

## Notes

- The app stores the remembered output folder and recent-downloads list in `%APPDATA%\PortionDownloader\settings.json`.
- If no folder is remembered, it defaults to your `Videos` folder.
- If the selected formats are not already MP4-friendly, the clip is converted to H.264/AAC in an `.mp4` file.
