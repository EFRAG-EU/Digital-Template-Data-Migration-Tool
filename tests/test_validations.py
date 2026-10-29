"""missing_NR.json entries are objects: {"sheet": ..., "range": ...}, or
{"sheet": ..., "anchor": {...}} for cells found by their row's template label."""

from __future__ import annotations

import pytest

from migration_tool.data.validations import VersionCollection

DATA = {
    "1.0.0": {
        "Present": {"sheet": "Social Disclosures", "range": "$D$152:$D$153"},
        "Absent": {"sheet": "Social Disclosures", "range": None},
    }
}


def test_loads_sheets_and_ranges():
    collection = VersionCollection.from_dict(DATA)
    assert collection.get_labels("1.0.0") == ["Present", "Absent"]
    assert collection.get_sheets("1.0.0") == ["Social Disclosures"] * 2
    assert collection.get_ranges("1.0.0") == ["$D$152:$D$153", None]


def test_round_trips():
    assert VersionCollection.from_dict(DATA).dump_simple() == DATA


ANCHORED = {
    "2.0.0": {
        "TrainingHours": {
            "sheet": "Social Disclosures",
            "anchor": {"label": "template_label_training", "cols": "D", "rows": 2},
        },
        "GPSCoordinates": {
            "sheet": "General Information",
            "anchor": {"label": "template_label_gps", "cols": "H", "rowOffset": 3},
        },
    }
}


def test_round_trips_anchors():
    assert VersionCollection.from_dict(ANCHORED).dump_simple() == ANCHORED


def test_anchors_have_no_range_until_resolved():
    collection = VersionCollection.from_dict(ANCHORED)
    with pytest.raises(ValueError, match="TrainingHours"):
        collection.get_ranges("2.0.0")


@pytest.mark.parametrize(
    "ref",
    [
        {"sheet": "Social Disclosures"},
        {"range": "$D$7"},
        {"sheet": "Social Disclosures", "range": "$D$7", "colour": "red"},
        {"sheet": "S", "range": "$D$7", "anchor": {"label": "x", "cols": "D"}},
        {"sheet": "S", "anchor": {"cols": "D"}},
        {"sheet": "S", "anchor": {"label": "x"}},
        {"sheet": "S", "anchor": {"label": "x", "cols": "D", "colour": "red"}},
    ],
)
def test_rejects_malformed_entries(ref):
    with pytest.raises(ValueError, match="Broken"):
        VersionCollection.from_dict({"1.0.0": {"Broken": ref}})
