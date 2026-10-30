"""Canonical, private-safe serialization and identity primitives.

This module separates persisted presentation JSON from compact semantic identity
JSON.  Callers that need to omit volatile data must construct an explicit
model-specific projection before calling :func:`semantic_sha256`.
"""

from __future__ import annotations

import hashlib
import json
import re
import stat
import zipfile
from collections.abc import Callable, Collection
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import NoReturn, TypeAlias

JsonScalar: TypeAlias = None | bool | int | float | str
JsonValue: TypeAlias = JsonScalar | list["JsonValue"] | dict[str, "JsonValue"]
Clock: TypeAlias = Callable[[], datetime]

_DIAGNOSTIC_CODE_PATTERN = re.compile(r"[a-z][a-z0-9_]{0,127}")
_STATE_PATTERN = re.compile(r"[a-z][a-z0-9_]{0,63}")
_SHA256_PATTERN = re.compile(r"[0-9a-f]{64}")
_TIMESTAMP_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?Z")
_DETERMINISTIC_ZIP_TIME = (2020, 1, 1, 0, 0, 0)


class StrictJsonError(ValueError):
    """Raised when JSON is structurally unsafe or contains a non-finite value."""


def pretty_json_text(value: object) -> str:
    """Return canonical persisted JSON text with a final newline."""

    return json.dumps(value, sort_keys=True, indent=2, ensure_ascii=True, allow_nan=False) + "\n"


def pretty_json_bytes(value: object) -> bytes:
    """Return canonical persisted JSON encoded as UTF-8 bytes."""

    return pretty_json_text(value).encode("utf-8")


def semantic_json_bytes(value: object) -> bytes:
    """Return compact UTF-8 JSON bytes for an explicit semantic projection."""

    return json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
        allow_nan=False,
    ).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    """Return the exact-byte SHA-256 identity of ``value``."""

    return hashlib.sha256(value).hexdigest()


def semantic_sha256(value: object) -> str:
    """Return the SHA-256 identity of compact semantic JSON bytes."""

    return sha256_bytes(semantic_json_bytes(value))


def sha256_file(path: str | Path, *, chunk_size: int = 64 * 1024) -> str:
    """Return exact-byte SHA-256, using a bounded 64 KiB read buffer by default."""

    if chunk_size <= 0:
        raise ValueError("chunk_size must be positive")
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def strict_json_loads(value: str | bytes | bytearray) -> JsonValue:
    """Load JSON while rejecting duplicate object keys and non-finite constants."""

    decoded: object = json.loads(
        value,
        object_pairs_hook=_reject_duplicate_pairs,
        parse_constant=_reject_nonfinite_constant,
    )
    return _json_value(decoded)


def portable_display_path(
    value: str | Path,
    *,
    relative_to: str | Path | None = None,
) -> str:
    """Return a portable relative path or a redacted basename, never an absolute path."""

    raw_value = str(value)
    if "\\" in raw_value or (len(raw_value) > 1 and raw_value[1] == ":"):
        return _redacted_basename(raw_value)
    path = Path(value)
    root = Path.cwd() if relative_to is None else Path(relative_to)
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except (OSError, ValueError):
        return _redacted_basename(raw_value)


def is_portable_display_path(value: object) -> bool:
    """Return whether ``value`` is a bounded, relative POSIX display path."""

    if (
        type(value) is not str
        or not value
        or len(value) > 1024
        or value.startswith("/")
        or "\\" in value
        or ":" in value
        or "\r" in value
        or "\n" in value
    ):
        return False
    return all(part not in {"", ".", ".."} for part in value.split("/"))


def utc_now(*, clock: Clock | None = None) -> datetime:
    """Return a UTC instant, accepting a deterministic clock for callers and tests."""

    value = datetime.now(UTC) if clock is None else clock()
    if value.tzinfo is None or value.utcoffset() != timedelta(0):
        raise ValueError("clock must return a UTC datetime")
    return value.astimezone(UTC)


def serialize_utc_timestamp(value: datetime) -> str:
    """Serialize an aware UTC timestamp using the canonical ``Z`` spelling."""

    if value.tzinfo is None or value.utcoffset() != timedelta(0):
        raise ValueError("timestamp must be UTC")
    return value.isoformat(timespec="microseconds").replace("+00:00", "Z")


def parse_utc_timestamp(value: object) -> datetime:
    """Parse only canonical UTC timestamps that use the ``Z`` suffix."""

    if type(value) is not str or _TIMESTAMP_PATTERN.fullmatch(value) is None:
        raise ValueError("timestamp must use canonical UTC format")
    try:
        return datetime.fromisoformat(f"{value[:-1]}+00:00")
    except ValueError as exc:
        raise ValueError("timestamp must use canonical UTC format") from exc


def deterministic_zip_info(name: str, *, mode: int = 0o600) -> zipfile.ZipInfo:
    """Build deterministic regular-file ZIP metadata with private permissions."""

    if type(name) is not str or not name:
        raise ValueError("zip entry name must be non-empty")
    if type(mode) is not int or not 0 <= mode <= 0o777:
        raise ValueError("zip entry mode must be a permission mask")
    metadata = zipfile.ZipInfo(name, date_time=_DETERMINISTIC_ZIP_TIME)
    metadata.create_system = 3
    metadata.compress_type = zipfile.ZIP_DEFLATED
    metadata.external_attr = (stat.S_IFREG | mode) << 16
    return metadata


def is_diagnostic_code(value: object) -> bool:
    """Return whether ``value`` is a bounded structural diagnostic code."""

    return type(value) is str and _DIAGNOSTIC_CODE_PATTERN.fullmatch(value) is not None


def safe_diagnostic_code(value: object, *, fallback: str) -> str:
    """Keep only an allow-listed diagnostic shape; otherwise return ``fallback``."""

    if not is_diagnostic_code(fallback):
        raise ValueError("diagnostic fallback must be a code")
    if type(value) is str and _DIAGNOSTIC_CODE_PATTERN.fullmatch(value) is not None:
        return value
    return fallback


def is_state(value: object) -> bool:
    """Return whether ``value`` is a bounded structural state code."""

    return type(value) is str and _STATE_PATTERN.fullmatch(value) is not None


def safe_state(value: object, *, fallback: str) -> str:
    """Keep only a structural state code; otherwise return ``fallback``."""

    if not is_state(fallback):
        raise ValueError("state fallback must be a state code")
    if type(value) is str and _STATE_PATTERN.fullmatch(value) is not None:
        return value
    return fallback


def allow_state(value: object, *, allowed: Collection[str], fallback: str) -> str:
    """Return an explicitly allowed state or a caller-declared safe fallback."""

    if fallback not in allowed:
        raise ValueError("state fallback must be allowed")
    if type(value) is str and value in allowed:
        return value
    return fallback


def is_sha256(value: object) -> bool:
    """Return whether ``value`` is a lowercase SHA-256 hex digest."""

    return type(value) is str and _SHA256_PATTERN.fullmatch(value) is not None


def safe_sha256(value: object) -> str | None:
    """Return a structurally valid digest, otherwise redact it to ``None``."""

    if type(value) is str and _SHA256_PATTERN.fullmatch(value) is not None:
        return value
    return None


def _reject_duplicate_pairs(pairs: list[tuple[str, object]]) -> dict[str, JsonValue]:
    result: dict[str, JsonValue] = {}
    for key, item in pairs:
        if key in result:
            raise StrictJsonError("duplicate JSON object key")
        result[key] = _json_value(item)
    return result


def _reject_nonfinite_constant(value: str) -> NoReturn:
    raise StrictJsonError(f"non-finite JSON constant: {value}")


def _json_value(value: object) -> JsonValue:
    if value is None:
        return None
    if type(value) is bool:
        return value
    if type(value) is int:
        return value
    if type(value) is float:
        return value
    if type(value) is str:
        return value
    if type(value) is list:
        return [_json_value(item) for item in value]
    if type(value) is dict:
        result: dict[str, JsonValue] = {}
        for key, item in value.items():
            if type(key) is not str:
                raise StrictJsonError("JSON object key must be a string")
            result[key] = _json_value(item)
        return result
    raise StrictJsonError("JSON value has an unsupported type")


def _redacted_basename(value: str) -> str:
    candidate = value.replace("\\", "/").rsplit("/", maxsplit=1)[-1]
    if candidate in {"", ".", ".."} or ":" in candidate:
        return "artifact"
    return candidate
