"""missing_NR.json entries are objects: {"sheet": ..., "range": ...}."""

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


@pytest.mark.parametrize(
    "ref",
    [
        {"sheet": "Social Disclosures"},
        {"range": "$D$7"},
        {"sheet": "Social Disclosures", "range": "$D$7", "colour": "red"},
    ],
)
def test_rejects_malformed_entries(ref):
    with pytest.raises(ValueError, match="Broken"):
        VersionCollection.from_dict({"1.0.0": {"Broken": ref}})
