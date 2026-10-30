"""Shared test configuration."""

from __future__ import annotations

import os

# The persistent inspection store would let one test's packages change how a later test
# inspects, and would write into the user's cache. The suite runs without it.
os.environ.setdefault("WOLFDOCX_INSPECTION_CACHE", "off")
