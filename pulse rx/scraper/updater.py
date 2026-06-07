"""
Self-update for the Pulse Rx Scraper.

On launch the app asks GitHub for the latest published release. If its version
is newer than the running build, the new packaged ``.exe`` is downloaded next to
the current one and swapped in by a tiny relauncher batch script (a running
Windows ``.exe`` cannot overwrite itself directly, so the helper waits for this
process to exit, replaces the file, then starts it again).

Only the packaged ``--onefile`` build can self-replace. Run from source it will
still report whether a newer release exists, but applying it is a no-op (there
is no single exe to swap), so the app just points the user at the releases page.

The releases are served from GITHUB_REPO below — change it if you fork/rename.
"""

from __future__ import annotations

import os
import sys
import json
import tempfile
import subprocess
from dataclasses import dataclass
from typing import Callable, Optional
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

from version import __version__

# owner/repo that hosts the GitHub Releases this app updates from.
GITHUB_REPO = "user8088/scrapper"
API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases/latest"
USER_AGENT = "PulseRxScraper-Updater"

# Name of the packaged binary (must match the PyInstaller --name / CI build).
EXE_NAME = "PulseRxScraper.exe"


@dataclass
class ReleaseInfo:
    version: str        # e.g. "1.1.0"
    tag: str            # e.g. "v1.1.0"
    download_url: str   # browser_download_url of the .exe asset
    notes: str          # release body / changelog


def is_frozen() -> bool:
    """True when running as a PyInstaller-built executable."""
    return bool(getattr(sys, "frozen", False))


def _parse_version(v: str) -> tuple[int, int, int]:
    """'v1.2.3' / '1.2' -> (1, 2, 3). Non-numeric suffixes are ignored."""
    parts: list[int] = []
    for chunk in v.strip().lstrip("vV").split("."):
        num = ""
        for ch in chunk:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])  # type: ignore[return-value]


def is_newer(remote: str, local: str) -> bool:
    return _parse_version(remote) > _parse_version(local)


def check_for_update(timeout: float = 6.0) -> Optional[ReleaseInfo]:
    """Return ReleaseInfo if a newer release with an .exe asset exists, else None.

    Network/parse failures return None (offline = silently stay on current build).
    """
    req = Request(API_LATEST, headers={
        "User-Agent": USER_AGENT,
        "Accept": "application/vnd.github+json",
    })
    try:
        with urlopen(req, timeout=timeout) as resp:
            data = json.load(resp)
    except (URLError, HTTPError, TimeoutError, ValueError, OSError):
        return None

    tag = data.get("tag_name") or ""
    if not tag or not is_newer(tag, __version__):
        return None

    download_url = None
    for asset in data.get("assets", []):
        if (asset.get("name") or "").lower().endswith(".exe"):
            download_url = asset.get("browser_download_url")
            break
    if not download_url:
        return None

    return ReleaseInfo(
        version=tag.lstrip("vV"),
        tag=tag,
        download_url=download_url,
        notes=(data.get("body") or "").strip(),
    )


def download_update(info: ReleaseInfo,
                    on_progress: Optional[Callable[[int, int], None]] = None) -> str:
    """Download the new exe and return its path on disk.

    When frozen it is placed alongside the running exe so the in-place swap is a
    fast same-volume move; otherwise it goes to a temp dir.
    """
    target_dir = os.path.dirname(sys.executable) if is_frozen() else tempfile.gettempdir()
    new_path = os.path.join(target_dir, "PulseRxScraper.new.exe")

    req = Request(info.download_url, headers={"User-Agent": USER_AGENT})
    with urlopen(req, timeout=60) as resp, open(new_path, "wb") as fh:
        total = int(resp.headers.get("Content-Length") or 0)
        read = 0
        while True:
            chunk = resp.read(262144)
            if not chunk:
                break
            fh.write(chunk)
            read += len(chunk)
            if on_progress:
                on_progress(read, total)
    return new_path


def apply_update_and_restart(new_path: str) -> None:
    """Swap the downloaded exe over the running one and relaunch it.

    Spawns a detached batch script that retries the move until this process has
    released the file lock (i.e. after the app exits), then starts the new exe
    and deletes itself. The caller must terminate the app right after calling.
    """
    if not is_frozen():
        raise RuntimeError("Self-update only works in the packaged .exe build.")

    exe = sys.executable
    bat = os.path.join(tempfile.gettempdir(), "pulse_rx_update.bat")
    script = (
        "@echo off\r\n"
        ":retry\r\n"
        f'move /y "{new_path}" "{exe}" >nul 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        "  ping 127.0.0.1 -n 2 >nul\r\n"
        "  goto retry\r\n"
        ")\r\n"
        f'start "" "{exe}"\r\n'
        'del "%~f0"\r\n'
    )
    with open(bat, "w", encoding="ascii") as fh:
        fh.write(script)

    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    subprocess.Popen(
        ["cmd", "/c", bat],
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )
