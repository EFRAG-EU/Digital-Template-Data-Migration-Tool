"""Unit tests for migration_tool.outils."""

from __future__ import annotations

import pandas as pd
import pytest
from openpyxl import Workbook
from openpyxl.utils import absolute_coordinate, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName

from migration_tool.classes import make_value
from migration_tool.outils import (
    access_NR_table,
    adjust_countriesOfOperation,
    assess_energyConsumption_validation,
    clean_NR_with_no_data,
    get_DFcell,
    get_NR_singlevalue,
    markerNames,
    set_DFcell,
    try_fetch_NR_refs,
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


class TestMarkerNames:
    """A name whose cells are all covered by disjoint, more specific names only
    marks an area (e.g. a table); the data is in the names inside it."""

    @pytest.fixture
    def grid(self) -> Workbook:
        """GridTable spans A1:C2 and A5:C5. Inside it: single-cell names, an
        alias pair on A2, a merged B2:C2 named by its top-left, formulas in C1
        and in the merged A5:C5, and RowTable, itself a marker of A1:B1."""
        wb = Workbook()
        ws = wb.active
        ws["A1"], ws["B1"], ws["C1"] = 1, 2, "=A1+B1"
        ws["A2"], ws["B2"] = 3, 4
        ws.merge_cells("B2:C2")
        ws["A5"] = "=SUM(A1:B2)"
        ws.merge_cells("A5:C5")
        for name, rng in [
            ("First", "A1"),
            ("Second", "B1"),
            ("AliasOne", "A2"),
            ("AliasTwo", "A2"),
            ("Merged", "B2"),
        ]:
            _create_NR(wb, name, ("Sheet", rng))
        _create_NR(wb, "GridTable", ("Sheet", "A1:C2"), ("Sheet", "A5:C5"))
        _create_NR(wb, "RowTable", ("Sheet", "A1:B1"))
        return wb

    def test_covered_names_are_markers(self, grid):
        assert {"GridTable", "RowTable"} <= markerNames(grid)

    def test_names_without_inner_names_are_not_markers(self, grid):
        _create_NR(grid, "PlainBlock", ("Sheet", "D1:D3"))
        assert markerNames(grid).isdisjoint({"First", "AliasOne", "PlainBlock"})

    def test_an_uncovered_input_cell_is_data(self, grid):
        grid.active["C1"] = 7
        assert "GridTable" not in markerNames(grid)

    def test_overlapping_inner_names_are_not_a_cover(self, grid):
        del grid.defined_names["First"], grid.defined_names["Second"]
        _create_NR(grid, "Left", ("Sheet", "A1:B1"))
        _create_NR(grid, "Right", ("Sheet", "B1:C1"))
        assert "GridTable" not in markerNames(grid)

    def test_markers_are_skipped_silently(self, grid, collector):
        df = access_NR_table(grid, collector)
        assert "GridTable" not in set(df["name_ranges"])
        assert collector.issues == []

    def test_other_multi_destination_names_are_still_reported(self, grid, collector):
        _create_NR(grid, "BrokenPair", ("Sheet", "D1"), ("Sheet", "D3"))
        access_NR_table(grid, collector)
        assert any("'BrokenPair' has 2 destinations" in i for i in collector.issues)

    def test_cleaning_keeps_data_names_whatever_they_are_called(self):
        df = pd.DataFrame(
            {
                "name_ranges": [
                    "IdentifierOfSiteTypedAxis_SitesInBiodiversitySensitiveAreasTable",
                    "template_label_x",
                    "enum_ListYear",
                    "BreakdownOfEnergyConsumptionAxis",
                ]
            }
        )
        assert list(clean_NR_with_no_data(df)["name_ranges"]) == [
            "IdentifierOfSiteTypedAxis_SitesInBiodiversitySensitiveAreasTable"
        ]


class TestTargetCellsByLabel:
    """Target-template helper cells are looked up by label, not by address, so
    they are found wherever a template release puts them."""

    @pytest.fixture
    def missing(self) -> pd.DataFrame:
        return pd.DataFrame(
            {
                "labels": [
                    "OperationInMoreThanOneCountry",
                    "FuelTotCalculatedResult",
                    "FuelAmount3",
                ],
                "sheets": ["Social Disclosures", "Fuel Converter", "Fuel Converter"],
                "cell_ranges": ["$E$99", "$D$99", "$C$99"],
                "cell_values": [None, None, make_value([[5]])],
            }
        )

    def test_third_fuel_is_flagged(self, missing, collector):
        assess_energyConsumption_validation(missing, collector)
        assert any("Energy Consumption" in issue for issue in collector.issues)

    def test_countries_and_fuel_transfer_are_set(self, missing):
        countries = pd.DataFrame(
            {
                "name_ranges": ["CountryOfEmploymentContractAxis"],
                "cell_values": [make_value([["FR"], ["DE"], [None]])],
            }
        )
        old_wb = Workbook()
        old_wb.create_sheet("Fuel Converter")["B10"] = "Diesel"

        adjust_countriesOfOperation(countries, missing, old_wb)

        for label in ("OperationInMoreThanOneCountry", "FuelTotCalculatedResult"):
            value = get_DFcell(missing, "cell_values", ("labels", label))
            assert value.topleft() is True, label


### to continue with other unit tests for functions in outils.py
