"""Immutable diagnostics for document operations and rendering.

Import ``RenderWarning``, ``FontFallbackWarning``, ``WarningCategory``,
``WarningSource`` and ``WarningCollection`` from ``wolfdocx.render_warnings``.
Document warning collections retain cumulative operation history until ``clear()``;
clearing a collection does not change document content or operation reports.
``RenderReport.warnings`` and the report adapters return immutable tuple snapshots.
Locators and diagnostic ``details`` are detached, deeply read-only copies of
JSON-like data: mappings remain inspectable mappings and sequences become tuples.
Raw reports are never modified. Unknown diagnostic kinds and keys are retained.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from enum import Enum
from types import MappingProxyType
from typing import Any, overload


class WarningCategory(str, Enum):
    """The effect or origin of a diagnostic, independent of its severity."""

    OMISSION = "omission"
    APPROXIMATION = "approximation"
    RECOVERY = "recovery"
    PUBLICATION = "publication"
    FONT_FALLBACK = "font_fallback"
    FONT_DISCOVERY = "font_discovery"
    OTHER = "other"


class WarningSource(str, Enum):
    """The operation that emitted a diagnostic."""

    LOAD = "load"
    SAVE = "save"
    RENDER = "render"
    FONT = "font"


def _freeze(value: Any) -> Any:
    """Detach JSON-like containers without reducing structured data to text."""
    if value is None or isinstance(value, (str, bool, int, float, bytes, Enum)):
        return value
    if isinstance(value, Mapping):
        return MappingProxyType({_freeze(key): _freeze(item) for key, item in value.items()})
    if isinstance(value, Sequence):
        return tuple(_freeze(item) for item in value)
    if isinstance(value, (set, frozenset)):
        return frozenset(_freeze(item) for item in value)
    raise TypeError("warning locators and details must contain immutable scalars or containers")


@dataclass(frozen=True, slots=True)
class RenderWarning:
    """One diagnostic, retaining its original severity and structured location.

    ``details`` preserves the complete diagnostic record, including keys that the
    typed fields do not recognize. Neither it nor ``locator`` shares mutable
    containers with its input.
    """

    category: WarningCategory
    source: WarningSource
    description: str
    kind: str | None = None
    locator: Any = None
    reason: str | None = None
    severity: str = "warning"
    details: Mapping[str, Any] = field(default_factory=dict, kw_only=True)

    def __post_init__(self) -> None:
        object.__setattr__(self, "locator", _freeze(self.locator))
        object.__setattr__(self, "details", _freeze(self.details))


@dataclass(frozen=True, slots=True, kw_only=True)
class FontFallbackWarning(RenderWarning):
    """A substitution with both the requested and resolved family recorded."""

    requested_family: str
    resolved_family: str


class WarningCollection(Sequence[RenderWarning]):
    """Read-only warning sequence with explicit history clearing.

    Document operations add records through the internal ``_extend`` hook; callers
    can inspect or clear history, but cannot publicly append diagnostics. Slices
    and iterators retain their snapshot when the collection later changes.
    """

    __slots__ = ("_items",)

    def __init__(self, warnings: Iterable[RenderWarning] = ()) -> None:
        self._items = tuple(warnings)

    def __len__(self) -> int:
        return len(self._items)

    def __iter__(self) -> Iterator[RenderWarning]:
        return iter(self._items)

    @overload
    def __getitem__(self, index: int) -> RenderWarning: ...

    @overload
    def __getitem__(self, index: slice) -> tuple[RenderWarning, ...]: ...

    def __getitem__(self, index: int | slice) -> RenderWarning | tuple[RenderWarning, ...]:
        return self._items[index]

    def clear(self) -> None:
        """Discard collected history without changing any report or document."""
        self._items = ()

    def _extend(self, warnings: Iterable[RenderWarning]) -> None:
        self._items += tuple(warnings)


def _record_warning(
    record: Mapping[str, Any],
    category: WarningCategory,
    source: WarningSource,
    *,
    default_description: str = "",
    default_kind: str | None = None,
) -> RenderWarning:
    return RenderWarning(
        category=category,
        source=source,
        description=record.get(
            "description", record.get("message", record.get("reason", default_description))
        ),
        kind=record.get("kind", default_kind),
        locator=record.get("locator"),
        reason=record.get("reason", record.get("code")),
        severity=record.get("severity", "warning"),
        details=record,
    )


def _font_substitution_warning(record: Mapping[str, Any]) -> RenderWarning:
    requested = record.get("requested_family", record.get("requested"))
    resolved = record.get("resolved_family", record.get("used"))
    reason = record.get("reason")
    description = record.get("description", record.get("message"))
    if description is None:
        description = f"{requested!r} resolves to {resolved!r}"
        if reason is not None:
            description += f" ({reason})"
    fields = {
        "category": WarningCategory.FONT_FALLBACK,
        "source": WarningSource.FONT,
        "description": description,
        "kind": record.get("kind", "font_substitution"),
        "locator": record.get("locator"),
        "reason": reason,
        "severity": record.get("severity", "warning"),
        "details": record,
    }
    if requested is not None and resolved is not None:
        return FontFallbackWarning(**fields, requested_family=requested, resolved_family=resolved)
    return RenderWarning(**fields)


def warnings_from_render_report(raw: Mapping[str, Any]) -> tuple[RenderWarning, ...]:
    """Adapt every output's diagnostics, preserving order and original records.

    Every omission and approximation is retained, including unfamiliar kinds,
    locators and severities. Font substitutions and generic output warnings are
    retained as well; summary counts are not additional diagnostics.
    """
    warnings: list[RenderWarning] = []
    for output in raw.get("outputs", ()):
        for key, category in (
            ("omissions", WarningCategory.OMISSION),
            ("approximations", WarningCategory.APPROXIMATION),
        ):
            warnings.extend(
                _record_warning(record, category, WarningSource.RENDER)
                for record in output.get(key, ())
            )
        warnings.extend(
            _font_substitution_warning(record) for record in output.get("font_substitutions", ())
        )
        warnings.extend(
            _record_warning(
                record,
                WarningCategory.OTHER,
                WarningSource.RENDER,
                default_kind=record.get("code"),
            )
            for record in output.get("warnings", ())
        )
    return tuple(warnings)


def warnings_from_native_report(report: Any) -> tuple[RenderWarning, ...]:
    """Adapt all native boundaries, substitutions and any refusal code.

    Same-family substitutions are retained: embedded faces and style changes are
    diagnostics too. A diagnostic-free published report produces an empty tuple.
    """
    warnings: list[RenderWarning] = [
        RenderWarning(
            WarningCategory.APPROXIMATION,
            WarningSource.RENDER,
            boundary,
            kind="native_render_boundary",
            reason=boundary,
        )
        for boundary in report.boundaries
    ]
    warnings.extend(_font_substitution_warning(record) for record in report.substitutions or ())
    if report.code is not None:
        warnings.append(
            RenderWarning(
                WarningCategory.PUBLICATION,
                WarningSource.RENDER,
                report.code,
                kind=report.code,
                reason=report.code,
                severity="error",
            )
        )
    return tuple(warnings)


def font_discovery_warning(record: Mapping[str, Any]) -> RenderWarning:
    """Adapt a skipped font discovery callback, never inventing a substitution.

    Native callbacks contain ``code``, ``file`` and ``source_kind``. Discovery
    records do not prove which family resolved, so they stay ordinary warnings.
    """
    code = record.get("code")
    locator = record.get("locator", record.get("file"))
    description = record.get("description", record.get("message"))
    if description is None:
        description = f"{locator}: {code}" if locator is not None else str(code or "")
    return RenderWarning(
        category=WarningCategory.FONT_DISCOVERY,
        source=WarningSource.FONT,
        description=description,
        kind=record.get("kind", code),
        locator=locator,
        reason=record.get("reason", code),
        severity=record.get("severity", "warning"),
        details=record,
    )
