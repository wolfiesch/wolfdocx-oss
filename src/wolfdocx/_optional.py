"""Resolve optional features through installed distribution entry points.

A source checkout used without installing it (``PYTHONPATH=src``) has no
distribution metadata, so the provider shipped beside this module stands in for
the entry point. The MIT core ships no provider and stays core-only.
"""

from __future__ import annotations

from collections.abc import Callable
from functools import cache
from importlib import import_module
from importlib.metadata import entry_points
from types import ModuleType
from typing import Literal, overload

#: Provider module of the commercial engine, resolved only when no entry point is installed.
_SOURCE_PROVIDER = "_commercial_features"


@cache
def _provider() -> Callable[[str], ModuleType] | None:
    providers = entry_points(group="wolfdocx.features", name="commercial")
    if len(providers) > 1:
        raise RuntimeError("commercial_feature_provider_ambiguous")
    if providers:
        return next(iter(providers)).load()
    name = f"{__package__}.{_SOURCE_PROVIDER}"
    try:
        module = import_module(name)
    except ModuleNotFoundError as error:
        if error.name != name:
            raise
        return None
    return module.load_module


@overload
def feature_module(name: str, *, required: Literal[True] = True) -> ModuleType: ...


@overload
def feature_module(name: str, *, required: Literal[False]) -> ModuleType | None: ...


def feature_module(name: str, *, required: bool = True) -> ModuleType | None:
    """Resolve a feature; optional discovery is reserved for provider extensions."""
    provider = _provider()
    if provider is None:
        if not required:
            return None
        raise ModuleNotFoundError("commercial_feature_unavailable")
    return provider(name)
