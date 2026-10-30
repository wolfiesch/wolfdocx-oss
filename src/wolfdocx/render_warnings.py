"""Immutable diagnostics collected by document save operations."""

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


