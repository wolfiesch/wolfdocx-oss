"""Persistent per-part inspection facts.

Structural inspection parses every XML part of a package. Whether one part parses
under the bounded parser, and how many elements it holds, depends only on the part's
bytes, its parse mode, the effective :class:`~wolfdocx.inspect.PackageLimits`, and the
inspection code. This store remembers that fact for the parts of packages that
inspected valid, so a later process can defer parsing a part whose bytes it has seen
until something reads its tree.

The descriptive stage (part texts, active-content facts, the document fingerprint) is a
pure function of parsed trees, so its results are stored too, as JSON values keyed by
the digests of every part they read.

The store is a SQLite database below the WolfDocx cache root
(``WOLFDOCX_CACHE_DIR`` or the platform cache directory). Set
``WOLFDOCX_INSPECTION_CACHE=off`` to disable it. Gates run with it suspended unless
``WOLFDOCX_GATE_INSPECTION_CACHE=on`` (see :func:`gate_scope`). Every failure to open,
read, or write it disables it for the rest of the process, so inspection falls back to
parsing every part.
"""

from __future__ import annotations

import contextlib
import contextvars
import hashlib
import json
import os
import stat
import sys
import threading
import zlib
from collections.abc import Iterator, Mapping
from pathlib import Path
from typing import TYPE_CHECKING, Any

from ._cache_root import default_cache_root

if TYPE_CHECKING:
    import sqlite3

_FORMAT = b"wolfdocx-inspection-part-facts-v2"
_DATABASE_NAME = "inspection-part-facts-v1.sqlite3"
_DISABLED_VALUES = frozenset({"0", "off", "false", "no"})
_ENABLED_VALUES = frozenset({"1", "on", "true", "yes"})
# Rows are about 60 bytes, so this bounds the database near 12 MB; the oldest rows go
# first once it is reached.
MAX_ROWS = 200_000
# Descriptive values are compressed JSON. A value above the per-value bound is not
# stored; once the total exceeds its bound the oldest values go first.
MAX_VALUE_BYTES = 4 * 1024 * 1024
MAX_VALUE_TOTAL_BYTES = 128 * 1024 * 1024
# Stored facts depend on the bounded parser, descriptive stage, hashing and model.
_IDENTITY_MODULES = ("inspect.py", "canonical.py", "model.py")

_lock = threading.Lock()
# One store (or ``None`` once it failed) per process and database path.
_stores: dict[tuple[int, Path], PartFactStore | None] = {}
_code_identity: bytes | None = None
# True inside a scope that must not trust the store (see ``gate_scope``).
_suspended: contextvars.ContextVar[bool] = contextvars.ContextVar(
    "wolfdocx_inspection_store_suspended", default=False
)


def enabled() -> bool:
    """False when ``WOLFDOCX_INSPECTION_CACHE`` disables the store or a scope suspends it."""
    if _suspended.get():
        return False
    value = os.environ.get("WOLFDOCX_INSPECTION_CACHE")
    return value is None or value.strip().lower() not in _DISABLED_VALUES


@contextlib.contextmanager
def gate_scope() -> Iterator[None]:
    """Suspend the store for this thread's context unless gates opt in.

    Inside the scope nothing is read from or written to the store, and in-process
    inspections built with it are neither returned from the memo nor lent as reuse
    donors, so every verdict and report derives from the package bytes alone. Set
    ``WOLFDOCX_GATE_INSPECTION_CACHE=on`` to let gates use the store.
    """
    value = os.environ.get("WOLFDOCX_GATE_INSPECTION_CACHE", "")
    if value.strip().lower() in _ENABLED_VALUES:
        yield
        return
    token = _suspended.set(True)
    try:
        yield
    finally:
        _suspended.reset(token)


def database_path() -> Path:
    return default_cache_root() / "inspection" / _DATABASE_NAME


def open_store() -> PartFactStore | None:
    """Return this process's store, or ``None`` when it is disabled or unusable."""
    if not enabled():
        return None
    path = database_path()
    key = (os.getpid(), path)
    with _lock:
        if key in _stores:
            return _stores[key]
        store: PartFactStore | None = None
        try:
            import sqlite3
        except ImportError:
            pass
        else:
            try:
                store = PartFactStore(path, _identity())
            except sqlite3.DatabaseError as exc:
                # A database that is not one (a torn or foreign file) is rebuilt once;
                # busy or locked databases raise subclasses and are left alone.
                if type(exc) is sqlite3.DatabaseError:
                    store = _rebuild(path)
            except (sqlite3.Error, OSError):
                pass
        _stores[key] = store
        return store


def _rebuild(path: Path) -> PartFactStore | None:
    import sqlite3

    try:
        for suffix in ("", "-wal", "-shm"):
            path.with_name(path.name + suffix).unlink(missing_ok=True)
        return PartFactStore(path, _identity())
    except (sqlite3.Error, OSError):
        return None


def _identity() -> bytes:
    """Digest of everything besides the limits and the part that decides its facts."""
    global _code_identity
    if _code_identity is None:
        import pyexpat

        digest = hashlib.sha256(_FORMAT)
        digest.update(f"\0{sys.version_info[0]}.{sys.version_info[1]}".encode())
        digest.update(f"\0{pyexpat.EXPAT_VERSION}".encode())
        package = Path(__file__).parent
        for name in _IDENTITY_MODULES:
            digest.update(b"\0" + name.encode() + b"\0")
            digest.update((package / name).read_bytes())
        _code_identity = digest.digest()
    return _code_identity


def _private(mode: int, uid: int) -> bool:
    """True when the path belongs to this user and no one else may write to it."""
    getuid = getattr(os, "getuid", None)
    if getuid is None:
        return True
    return uid == getuid() and not mode & (stat.S_IWGRP | stat.S_IWOTH)


class PartFactStore:
    """Element counts of XML parts that parsed cleanly inside valid packages, and the
    descriptive values computed from them."""

    def __init__(self, path: Path, identity: bytes) -> None:
        import sqlite3

        self._identity = identity
        self._errors = (sqlite3.Error, OSError)
        self._suppress = contextlib.suppress(sqlite3.Error)
        directory = path.parent
        directory.mkdir(mode=0o700, parents=True, exist_ok=True)
        info = directory.stat()
        if not _private(info.st_mode, info.st_uid):
            raise PermissionError(f"{directory} is not private to this user")
        created = not path.exists()
        connection = sqlite3.connect(
            path, timeout=0.25, isolation_level=None, check_same_thread=False
        )
        try:
            if created:
                os.chmod(path, 0o600)
            info = path.stat()
            if not _private(info.st_mode, info.st_uid):
                raise PermissionError(f"{path} is not private to this user")
            if connection.execute("PRAGMA journal_mode").fetchone()[0] != "wal":
                connection.execute("PRAGMA journal_mode=WAL")
            connection.execute("PRAGMA synchronous=NORMAL")
            connection.execute(
                "CREATE TABLE IF NOT EXISTS part_facts "
                "(key BLOB NOT NULL UNIQUE, elements INTEGER NOT NULL)"
            )
            connection.execute(
                "CREATE TABLE IF NOT EXISTS descriptive_facts "
                "(key BLOB NOT NULL UNIQUE, value BLOB NOT NULL)"
            )
        except BaseException:
            connection.close()
            raise
        self._connection: sqlite3.Connection | None = connection
        self._lock = threading.Lock()

    def facts(self, limits: Mapping[str, Any], mode: str) -> PartFacts:
        """Facts for parts parsed in ``mode`` under the given effective limits."""
        prefix = hashlib.sha256(self._identity)
        prefix.update(json.dumps(dict(limits), sort_keys=True).encode())
        prefix.update(b"\0" + mode.encode() + b"\0")
        return PartFacts(self, prefix)

    def _disable(self) -> None:
        connection, self._connection = self._connection, None
        if connection is not None:
            with self._suppress:
                connection.close()

    def lookup(self, key: bytes) -> int | None:
        with self._lock:
            if self._connection is None:
                return None
            try:
                row = self._connection.execute(
                    "SELECT elements FROM part_facts WHERE key = ?", (key,)
                ).fetchone()
            except self._errors:
                self._disable()
                return None
        if row is None:
            return None
        elements = row[0]
        if type(elements) is not int or elements < 0:
            self.forget(key)
            return None
        return elements

    def record(self, facts: Mapping[bytes, int]) -> None:
        if not facts:
            return
        with self._lock:
            connection = self._connection
            if connection is None:
                return
            try:
                connection.execute("BEGIN IMMEDIATE")
                try:
                    connection.executemany(
                        "INSERT OR IGNORE INTO part_facts (key, elements) VALUES (?, ?)",
                        facts.items(),
                    )
                    newest = connection.execute("SELECT max(rowid) FROM part_facts").fetchone()[0]
                    if newest is not None and newest > MAX_ROWS:
                        connection.execute(
                            "DELETE FROM part_facts WHERE rowid <= ?", (newest - MAX_ROWS,)
                        )
                    connection.execute("COMMIT")
                except BaseException:
                    connection.execute("ROLLBACK")
                    raise
            except self._errors:
                self._disable()

    def forget(self, key: bytes) -> None:
        with self._lock:
            if self._connection is None:
                return
            try:
                self._connection.execute("DELETE FROM part_facts WHERE key = ?", (key,))
            except self._errors:
                self._disable()

    def lookup_value(self, key: bytes) -> Any:
        """Return the stored JSON value, or ``None``; an undecodable value is evicted."""
        with self._lock:
            if self._connection is None:
                return None
            try:
                row = self._connection.execute(
                    "SELECT value FROM descriptive_facts WHERE key = ?", (key,)
                ).fetchone()
            except self._errors:
                self._disable()
                return None
        if row is None:
            return None
        try:
            return json.loads(zlib.decompress(row[0]))
        except (TypeError, ValueError, zlib.error):
            self.forget_value(key)
            return None

    def record_values(self, values: Mapping[bytes, Any]) -> None:
        encoded = []
        for key, value in values.items():
            blob = zlib.compress(
                json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode(), 1
            )
            if len(blob) <= MAX_VALUE_BYTES:
                encoded.append((key, blob))
        if not encoded:
            return
        with self._lock:
            connection = self._connection
            if connection is None:
                return
            try:
                connection.execute("BEGIN IMMEDIATE")
                try:
                    connection.executemany(
                        "INSERT OR IGNORE INTO descriptive_facts (key, value) VALUES (?, ?)",
                        encoded,
                    )
                    total = connection.execute(
                        "SELECT coalesce(sum(length(value)), 0) FROM descriptive_facts"
                    ).fetchone()[0]
                    if total > MAX_VALUE_TOTAL_BYTES:
                        # Drop the oldest values until the rest fit in half the bound.
                        kept = 0
                        cutoff = None
                        for rowid, size in connection.execute(
                            "SELECT rowid, length(value) FROM descriptive_facts ORDER BY rowid DESC"
                        ):
                            kept += size
                            if kept > MAX_VALUE_TOTAL_BYTES // 2:
                                cutoff = rowid
                                break
                        if cutoff is not None:
                            connection.execute(
                                "DELETE FROM descriptive_facts WHERE rowid <= ?", (cutoff,)
                            )
                    connection.execute("COMMIT")
                except BaseException:
                    connection.execute("ROLLBACK")
                    raise
            except self._errors:
                self._disable()

    def forget_value(self, key: bytes) -> None:
        with self._lock:
            if self._connection is None:
                return
            try:
                self._connection.execute("DELETE FROM descriptive_facts WHERE key = ?", (key,))
            except self._errors:
                self._disable()


class PartFacts:
    """Lookups for one parse mode and limit set, keyed by part digest."""

    __slots__ = ("_prefix", "store")

    def __init__(self, store: PartFactStore, prefix: Any) -> None:
        self.store = store
        self._prefix = prefix

    def key(self, part_sha256: str) -> bytes:
        digest = self._prefix.copy()
        digest.update(part_sha256.encode())
        return digest.digest()

    def lookup_value(self, key: bytes) -> Any:
        return self.store.lookup_value(key)

    def lookup(self, key: bytes) -> int | None:
        return self.store.lookup(key)


def _reset_for_tests() -> None:
    """Close and forget every store this process opened."""
    with _lock:
        for store in _stores.values():
            if store is not None:
                store._disable()
        _stores.clear()
