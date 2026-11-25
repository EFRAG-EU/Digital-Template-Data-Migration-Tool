"""B2's 'Own workforce' and 'Workers in the value chain' tick boxes are one in VS 2.0.1.
The old columns are found by their headers: either ticked ticks the merged box, and the
issues after them stay in their own columns."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
from conftest import B2_MERGED_NOTE, SAMPLE_PARAMS, migrated
from openpyxl import Workbook

from migration_tool.classes import Value, make_shape, make_value
from migration_tool.outils import SUSTAINABILITY_ISSUES_NR, mergeSustainabilityIssues
from migration_tool.tool import IssuesCollector, load_workbook_quietly

OLD_HEADERS = [
    "=template_label_climate_change",
    "=template_label_pollution",
    "=template_label_water_and_marine_resources",
    "=template_label_biodiversity_and_ecosystems",
    "=template_label_circular_economy",
    "=template_label_own_workforce",
    "=template_label_workers_in_the_value_chain",
    "=template_label_affected_communities",
    "=template_label_consumers_and_endusers_validation",
    "=template_label_business_conduct",
]
UNMIGRATED = "B2) could not be migrated"


def merge(row: list, headers: list = OLD_HEADERS) -> tuple[Value, list[str]]:
    """Run the merge on an old B2 row under headers, into the 9-box 2.0.1 range."""
    wb = Workbook()
    ws = wb.create_sheet("General Information")
    ws.append(headers)
    ws.append(row)
    oldRange = f"A2:{chr(ord('A') + len(row) - 1)}2"
    dfOld = pd.DataFrame(
        {
            "name_ranges": [SUSTAINABILITY_ISSUES_NR],
            "sheets": [ws.title],
            "cell_shapes": [make_shape(oldRange)],
        }
    )
    dfOldToMerge = pd.DataFrame(
        {"name_ranges": [SUSTAINABILITY_ISSUES_NR], "cell_values": [make_value([row])]}
    )
    dfNew = pd.DataFrame(
        {
            "name_ranges": [SUSTAINABILITY_ISSUES_NR],
            "cell_shapes": [make_shape("D508:L508")],
        }
    )
    c = IssuesCollector()
    mergeSustainabilityIssues(dfOld, dfOldToMerge, dfNew, wb, c)
    return dfOldToMerge["cell_values"][0], c.issues


def ticks(own: bool, workers: bool) -> list:
    return [True, False, True, False, False, own, workers, True, False, True]


@pytest.mark.parametrize(
    ("own", "workers", "merged"),
    [
        (False, False, False),
        (True, False, True),
        (False, True, True),
        (True, True, True),
    ],
)
def test_either_tick_ticks_the_merged_box(own: bool, workers: bool, merged: bool):
    value, issues = merge(ticks(own, workers))

    assert value.values() == [
        [True, False, True, False, False, merged, True, False, True]
    ]
    assert issues == ([B2_MERGED_NOTE] if workers else [])


def test_plain_text_headers_of_1_0():
    headers = ["Climate change", "Pollution", "Water and marine resources"]
    headers += ["Biodiversity and ecosystems", "Circular economy", "Own workforce"]
    headers += ["Workers in the value chain", "Affected communities"]
    headers += ["Consumers and end-users", "Business conduct"]

    value, _ = merge(ticks(False, True), headers)

    assert value.values() == [
        [True, False, True, False, False, True, True, False, True]
    ]


def test_unknown_headers_leave_b2_empty_and_say_so():
    headers = [*OLD_HEADERS[:5], "Own staff", *OLD_HEADERS[6:]]

    value, issues = merge(ticks(True, False), headers)

    assert not value
    assert len(issues) == 1 and UNMIGRATED in issues[0]


def test_a_row_already_the_new_width_is_left_alone():
    row = [True, False, True, False, False, True, True, False, True]
    headers = [h for h in OLD_HEADERS if "workers" not in h]

    value, issues = merge(row, headers)

    assert value.values() == [row]
    assert issues == []


def b2(wb: Workbook) -> tuple[list[object], list[object]]:
    """B2's column headers and ticks in wb."""
    [(sheet, rng)] = wb.defined_names[SUSTAINABILITY_ISSUES_NR].destinations
    cells = wb[sheet][rng.replace("$", "")][0]
    headers: list[object] = [
        wb[sheet].cell(cells[0].row - 1, c.column).value for c in cells
    ]
    return headers, [c.value for c in cells]


@pytest.mark.parametrize("sample_path", SAMPLE_PARAMS)
def test_samples_keep_each_issue_in_its_own_column(sample_path: Path):
    oldHeaders, oldTicks = b2(load_workbook_quietly(sample_path))
    newHeaders, newTicks = b2(migrated(sample_path).workbook)
    old = dict(zip(oldHeaders, oldTicks, strict=True))

    def was(*headers: str) -> object:
        [tick] = [old[h] for h in headers if h in old]
        return tick

    own = was("=template_label_own_workforce", "Own workforce")
    workers = was(
        "=template_label_workers_in_the_value_chain", "Workers in the value chain"
    )
    expected = {
        "=template_label_own_workforce_and_workers_in_the_value_chain": (
            True if own is True or workers is True else own
        ),
        "=template_label_climate_change": was(
            "=template_label_climate_change", "Climate change"
        ),
        "=template_label_water": was(
            "=template_label_water_and_marine_resources", "Water and marine resources"
        ),
        "=template_label_affected_communities": was(
            "=template_label_affected_communities", "Affected communities"
        ),
        "=template_label_business_conduct": was(
            "=template_label_business_conduct", "Business conduct"
        ),
    }
    new = dict(zip(newHeaders, newTicks, strict=True))
    assert {h: new[h] for h in expected} == expected
