"""Reusable decorators for the migration tool."""

from __future__ import annotations

from collections.abc import Callable
from functools import cache, wraps
from typing import Generic, Protocol, Self, TypeVar


class Copyable(Protocol):
    """A value that can return an independent copy of itself (e.g. ``dict``,
    ``pandas.DataFrame``)."""

    def copy(self) -> Self: ...


_C = TypeVar("_C", bound=Copyable)


class cached_copy(Generic[_C]):
    """Decorator: load a constant value once, then hand out a fresh copy per call.

    Wraps a zero-argument *loader* that produces an expensive but
    input-independent value -- e.g. a config ``DataFrame`` read from a pickle, or
    a lookup ``dict`` built from a workbook. The loader runs at most once per
    process (lazily, on first call); every call thereafter returns a fresh
    ``.copy()`` of that cached value.

    Returning a copy is the point: each caller gets an isolated object, so
    accidentally mutating it cannot corrupt the shared master or leak into other
    callers. This keeps the "fresh object per call" guarantee you would get by
    reloading the source every time, while paying the load cost only once.

    The loader must return a :class:`Copyable` (anything exposing ``.copy()``);
    the precise return type is preserved for callers. Caching is per process, so
    under pytest-xdist each worker warms its own copy -- still a net win across
    that worker's many migrations.

    Example::

        @cached_copy
        def _table_of_contents() -> dict[str, int]:
            ...  # expensive load, run once

        toc = _table_of_contents()  # a private copy, safe to mutate
    """

    def __init__(self, loader: Callable[[], _C]) -> None:
        self._load = cache(loader)
        wraps(loader)(self)

    def __call__(self) -> _C:
        return self._load().copy()
