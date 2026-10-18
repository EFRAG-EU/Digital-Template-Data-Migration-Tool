"""Packaged data is found wherever the caller's working directory is."""

from __future__ import annotations

import pytest

from migration_tool.tool import _missing_NR, _NR_changes


@pytest.mark.parametrize("loader", [_missing_NR, _NR_changes])
def test_loads_from_any_working_directory(loader, tmp_path, monkeypatch):
    loader._load.cache_clear()  # otherwise an earlier call hides the problem
    monkeypatch.chdir(tmp_path)
    try:
        assert loader()
    finally:
        loader._load.cache_clear()
