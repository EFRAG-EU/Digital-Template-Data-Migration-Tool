"""Unit tests for migration_tool.outils."""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import absolute_coordinate, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
import pytest

from migration_tool.outils import (
    get_DFcell,
    set_DFcell,
    try_fetch_NR_refs,
    get_NR_singlevalue,
)
from migration_tool.tool import IssuesCollector


def _create_NR(wb: Workbook, name: str, *sheetRngPairs: tuple[str, str]) -> None:
    refs: list[str] = []
    for sheet, rng in sheetRngPairs:
        ref = f"{sheet}!{absolute_coordinate(rng)}"
        refs.append(ref)
    all = ",".join(refs)
    defn = DefinedName(name=name, attr_text=all)
    wb.defined_names.add(defn)


def _create_NR_unsafe(wb: Workbook, name: str, *sheetRngPairs: tuple[str, str]) -> None:
    refs: list[str] = []
    for sheet, rng in sheetRngPairs:
        ref = f"{sheet}!{rng}"
        refs.append(ref)
    all = ",".join(refs)
    defn = DefinedName(name=name, attr_text=all)
    wb.defined_names.add(defn)


@pytest.fixture
def data_frame() -> pd.DataFrame:
    """A small table representative of the DataFrames used by the tool."""
    return pd.DataFrame(
        {
            "name_ranges": ["FirstRange", "SecondRange", "ThirdRange"],
            "sheets": ["Sheet 1", "Sheet 2", "Sheet 2"],
            "cell_values": ["old value", "another value", "something"],
        }
    )


@pytest.fixture
def wb() -> Workbook:
    """A 1-sheet workbook with a 3x3 block of values and two NRs (single- and multi-cells)"""
    wb = Workbook()
    ws = wb.active
    for row in range(1, 4):
        for col in range(1, 4):
            ws.cell(row=row, column=col).value = row * 10 + col

    ws_title = quote_sheetname(ws.title)
    _create_NR(wb, "singleCell", (ws_title, "A1"))
    _create_NR(wb, "multiCell", (ws_title, "C1:C3"))
    _create_NR_unsafe(wb, "unreadableRef", (ws_title, "C1:C3,C5"))
    _create_NR_unsafe(wb, "invalidRange", (ws_title, "ciao"))

    return wb


@pytest.fixture
def collector() -> IssuesCollector:
    c = IssuesCollector()
    return c


class TestDataFrames:
    def test_get_DFcell_returns_value_matching_filters(self, data_frame):
        assert (
            get_DFcell(
                data_frame,
                "cell_values",
                ("name_ranges", "SecondRange"),
                ("sheets", "Sheet 2"),
            )
            == "another value"
        )

    def test_set_DFcell_updates_value_matching_filters(self, data_frame):
        set_DFcell(
            data_frame,
            "cell_values",
            "updated value",
            ("name_ranges", "FirstRange"),
            ("sheets", "Sheet 1"),
        )

        assert data_frame.loc[0, "cell_values"] == "updated value"


class TestNameRanges:
    def test_fetching_NR_refs(self, wb, collector):
        defNames = wb.defined_names

        sheet, rng = try_fetch_NR_refs("singleCell", defNames, collector)
        assert sheet == "Sheet"
        assert rng == absolute_coordinate("A1")

        sheet, rng = try_fetch_NR_refs("multiCell", defNames, collector)
        assert rng == absolute_coordinate("C1:C3")

        # inexistent name range
        assert ("", "") == try_fetch_NR_refs(".", defNames, collector)
        assert "not valid" in collector.issues[-1]

        # unreadable destination
        assert ("", "") == try_fetch_NR_refs("unreadableRef", defNames, collector)
        assert "unreadable" in collector.issues[-1]

        # invalid range
        _, rng = try_fetch_NR_refs("invalidRange", defNames, collector)
        assert ("", "") == try_fetch_NR_refs("invalidRange", defNames, collector)
        assert "invalid" in collector.issues[-1]

        # multiple destinations
        _create_NR(wb, "multiCell", ("Sheet", "C1:C2"), ("Sheet", "C3:C4"))
        assert ("", "") == try_fetch_NR_refs("multiCell", defNames, collector)
        assert "expected 1" in collector.issues[-1]

    def test_getting_single_NR(self, wb, collector, capsys):
        assert get_NR_singlevalue(wb, ".", collector) is None
        assert get_NR_singlevalue(wb, "unreadableRef", collector) is None
        assert get_NR_singlevalue(wb, "invalidRange", collector) is None

        assert 11 == get_NR_singlevalue(wb, "singleCell", collector)

        with pytest.raises(Exception) as exc_info:
            get_NR_singlevalue(wb, "multiCell", collector)
        assert "Do not use get_NR_singlevalue for a range" in str(exc_info.value)


### to continue with other unit tests for functions in outils.py
