"""Unit tests for the ``cached_copy`` decorator in migration_tool.decorators."""

from __future__ import annotations

import pandas as pd

from migration_tool.decorators import cached_copy


def test_loads_lazily_and_only_once():
    calls = []

    @cached_copy
    def loader() -> dict:
        calls.append(1)
        return {"a": 1}

    # Decorating must not run the loader.
    assert calls == []

    loader()
    loader()
    loader()
    # The underlying loader runs at most once, on first call.
    assert calls == [1]


def test_returns_equal_but_independent_dicts():
    @cached_copy
    def loader() -> dict:
        return {"a": 1, "b": 2}

    first = loader()
    second = loader()

    assert first == second
    assert first is not second  # a fresh copy each call

    # Mutating a returned copy must not leak into later calls.
    first["a"] = 999
    assert loader() == {"a": 1, "b": 2}


def test_returns_independent_dataframes():
    @cached_copy
    def loader() -> pd.DataFrame:
        return pd.DataFrame({"old": [1, 2], "new": [3, 4]})

    first = loader()
    assert first is not loader()

    # Mutating the returned frame must not corrupt the cached master.
    first["extra"] = [9, 9]
    assert "extra" not in loader().columns


def test_preserves_wrapped_metadata():
    @cached_copy
    def my_loader() -> dict:
        """Doc for the loader."""
        return {}

    assert my_loader.__name__ == "my_loader"
    assert my_loader.__doc__ == "Doc for the loader."


def test_independent_caches_per_decorated_function():
    @cached_copy
    def a() -> dict:
        return {"which": "a"}

    @cached_copy
    def b() -> dict:
        return {"which": "b"}

    assert a() == {"which": "a"}
    assert b() == {"which": "b"}
