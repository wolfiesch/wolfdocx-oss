"""The WolfDocx user cache directory shared by every on-disk cache."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def default_cache_root() -> Path:
    """Return ``WOLFDOCX_CACHE_DIR`` when set, else the platform user cache directory."""
    configured = os.environ.get("WOLFDOCX_CACHE_DIR")
    if configured:
        return Path(configured).expanduser()
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Caches" / "wolfdocx"
    if os.name == "nt":
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            return Path(local_app_data) / "wolfdocx" / "Cache"
    xdg_cache = os.environ.get("XDG_CACHE_HOME")
    if xdg_cache:
        return Path(xdg_cache).expanduser() / "wolfdocx"
    return Path.home() / ".cache" / "wolfdocx"
