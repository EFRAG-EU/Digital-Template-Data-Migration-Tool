"""VS 2.0.1's B7 waste names cover 'Technical Sheet' formulas mirroring the 'Environmental
Disclosures' waste table, so their data is pasted into that table, and no template formula is
ever overwritten."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
from conftest import SAMPLE_PARAMS, SAMPLES_DIR, migrated

from migration_tool.classes import check_formula, make_shape, make_value
from migration_tool.data.validations import VersionCollection
from migration_tool.outils import mergeSharedInputs, redirectToInputs, resolveAnchors
from migration_tool.tool import (
    NEW_TEMPLATE_NAME,
    IssuesCollector,
    _input_NR,
    load_workbook_quietly,
)

TEMPLATE = (
    Path(__file__).parents[1] / "src" / "migration_tool" / "data" / NEW_TEMPLATE_NAME
)
SAMPLE_2_0_0 = SAMPLES_DIR / "VS-Digital-Template-Sample-2.0.0.xlsx"


def formulas(path: Path) -> dict[tuple[str, str], object]:
    wb = load_workbook_quietly(path)
    return {
        (ws.title, cell.coordinate): getattr(cell.value, "text", cell.value)
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for cell in row
        if check_formula(cell)
    }


@pytest.mark.parametrize("sample_path", SAMPLE_PARAMS)
def test_samples_leave_template_formulas_alone(sample_path: Path):
    template = formulas(TEMPLATE)
    output = formulas(migrated(sample_path).output_path)
    assert {k: v for k, v in template.items() if output.get(k) != v} == {}


def test_2_0_0_waste_table_is_entered_on_environmental_disclosures():
    old = load_workbook_quietly(SAMPLE_2_0_0)
    new = migrated(SAMPLE_2_0_0).workbook
    inputs = VersionCollection.from_dict(_input_NR())
    resolveAnchors(inputs, "2.0.1", load_workbook_quietly(TEMPLATE))
    for ref in inputs.getRefs("2.0.1"):
        [(sheet, rng)] = old.defined_names[ref.label].destinations
        oldValues = [[c.value for c in row] for row in old[sheet][rng.replace("$", "")]]
        newValues = [[c.value for c in row] for row in new[ref.sheet][ref.range]]
        assert ref.sheet == "Environmental Disclosures"
        assert any(v is not None for row in oldValues for v in row), ref.label
        assert newValues == oldValues, ref.label


def table(*rows) -> pd.DataFrame:
    return pd.DataFrame(
        rows, columns=["name_ranges", "sheets", "cell_ranges", "cell_values"]
    )


def test_shared_inputs_are_pasted_once_filling_each_others_gaps():
    c = IssuesCollector()
    df = mergeSharedInputs(
        table(
            ("Mass", "ED", "J1:J3", make_value([[1], [None], [None]])),
            ("Volume", "ED", "J1:J3", make_value([[None], [2], [None]])),
            ("Other", "ED", "K1:K3", make_value([[3], [4], [5]])),
        ),
        c,
    )
    assert list(df["name_ranges"]) == ["Mass", "Other"]
    assert df["cell_values"][0].values() == [[1], [2], [None]]
    assert c.issues == []


def test_differing_shared_inputs_keep_the_first_and_say_so():
    c = IssuesCollector()
    df = mergeSharedInputs(
        table(
            ("Mass", "ED", "J1:J2", make_value([[1], [None]])),
            ("Volume", "ED", "J1:J2", make_value([[7], [2]])),
        ),
        c,
    )
    assert df["cell_values"][0].values() == [[1], [2]]
    assert len(c.issues) == 1 and "'Mass' and 'Volume'" in c.issues[0]


def test_an_input_of_another_size_is_refused():
    inputs = VersionCollection.from_dict(
        {"9.9.9": {"Mass": {"sheet": "ED", "range": "$J$1:$J$2"}}}
    )
    df = pd.DataFrame(
        {
            "name_ranges": ["Mass"],
            "sheets": ["TS"],
            "cell_ranges": ["$H$1:$H$3"],
            "cell_shapes": [make_shape("$H$1:$H$3")],
        }
    )
    with pytest.raises(ValueError, match="differs in size"):
        redirectToInputs(df, inputs, "9.9.9")
