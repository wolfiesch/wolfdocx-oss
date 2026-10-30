"""Strict, bounded resource profiles for DOCX package execution."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path

from .canonical import strict_json_loads

PACKAGE_PROFILE_SCHEMA_VERSION = 1
MIN_MAX_MATERIALIZED_BYTES = 1024
MAX_MAX_MATERIALIZED_BYTES = 256 * 1024 * 1024
DEFAULT_MAX_MATERIALIZED_BYTES = MAX_MAX_MATERIALIZED_BYTES


class PackageProfileError(ValueError):
    """Raised when a package resource profile is not an exact bounded configuration."""

    def __init__(self, code: str) -> None:
        self.code = code
        super().__init__(code)


@dataclass(frozen=True)
class PackageMemoryBounds:
    """Explicit upper bound on uncompressed DOCX bytes retained during inspection."""

    max_materialized_bytes: int = DEFAULT_MAX_MATERIALIZED_BYTES

    def __post_init__(self) -> None:
        if type(self.max_materialized_bytes) is not int:
            raise PackageProfileError("profile_memory_max_materialized_bytes_invalid")
        if (
            not MIN_MAX_MATERIALIZED_BYTES
            <= self.max_materialized_bytes
            <= MAX_MAX_MATERIALIZED_BYTES
        ):
            raise PackageProfileError("profile_memory_max_materialized_bytes_out_of_range")

    def to_dict(self) -> dict[str, int]:
        return {"max_materialized_bytes": self.max_materialized_bytes}


@dataclass(frozen=True)
class PackageProfile:
    """Versioned package-execution profile with no unbounded resource defaults."""

    schema_version: int = PACKAGE_PROFILE_SCHEMA_VERSION
    memory: PackageMemoryBounds = field(default_factory=PackageMemoryBounds)

    def __post_init__(self) -> None:
        if (
            type(self.schema_version) is not int
            or self.schema_version != PACKAGE_PROFILE_SCHEMA_VERSION
        ):
            raise PackageProfileError("profile_schema_version_invalid")
        if type(self.memory) is not PackageMemoryBounds:
            raise PackageProfileError("profile_memory_invalid")

    def to_dict(self) -> dict[str, object]:
        return {
            "schema_version": self.schema_version,
            "memory": self.memory.to_dict(),
        }


def _require_exact_keys(
    value: Mapping[object, object],
    expected: frozenset[str],
    *,
    error_code: str,
) -> None:
    if set(value) != expected:
        raise PackageProfileError(error_code)


def _require_int(value: object, *, error_code: str) -> int:
    if type(value) is not int:
        raise PackageProfileError(error_code)
    return value


def parse_package_profile(value: object) -> PackageProfile:
    """Parse one exact versioned profile mapping without coercion or implicit fields."""

    if not isinstance(value, Mapping):
        raise PackageProfileError("profile_not_object")
    _require_exact_keys(
        value,
        frozenset({"schema_version", "memory"}),
        error_code="profile_unknown_field",
    )

    schema_version = _require_int(
        value["schema_version"],
        error_code="profile_schema_version_invalid",
    )
    if schema_version != PACKAGE_PROFILE_SCHEMA_VERSION:
        raise PackageProfileError("profile_schema_version_invalid")

    memory = value["memory"]
    if not isinstance(memory, Mapping):
        raise PackageProfileError("profile_memory_not_object")
    _require_exact_keys(
        memory,
        frozenset({"max_materialized_bytes"}),
        error_code="profile_memory_unknown_field",
    )
    max_materialized_bytes = _require_int(
        memory["max_materialized_bytes"],
        error_code="profile_memory_max_materialized_bytes_invalid",
    )
    return PackageProfile(memory=PackageMemoryBounds(max_materialized_bytes))


def load_package_profile(path: str | Path) -> PackageProfile:
    """Load a profile JSON file through the strict JSON and profile boundaries."""

    try:
        encoded = Path(path).read_bytes()
    except OSError as exc:
        raise PackageProfileError("profile_read_failed") from exc
    try:
        value = strict_json_loads(encoded)
    except (TypeError, UnicodeDecodeError, ValueError) as exc:
        raise PackageProfileError("profile_invalid_json") from exc
    return parse_package_profile(value)
