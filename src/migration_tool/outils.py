from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any, cast

if TYPE_CHECKING:
    from .tool import IssuesCollector

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment
from openpyxl.utils import (
    absolute_coordinate,
    quote_sheetname,
    range_boundaries,
)
from openpyxl.workbook.defined_name import DefinedName, DefinedNameDict
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from .classes import Shape, Value, check_formula, make_shape, make_value
from .data.validations import VersionCollection

VERSION_NR = "template_reporting_template_version"
# The 2.0 name: paste_values() sees the new template's name ranges, after
# apply_changes_NR() has renamed the old ListOfOmittedDisclosuresDeemedToBe...
OMITTEDDISCL_NR = "ListOfOmittedDisclosures"


class MigrationError(ValueError):
    """The workbook cannot be migrated; the message is suitable for users."""


def get_DFcell(df: pd.DataFrame, return_col: str, *colfilters_pairs: tuple[str, Any]):

    if not colfilters_pairs:
        raise TypeError("At least one pair (column, filter) is required.")

    mask = pd.Series(True, index=df.index)

    for col, filter in colfilters_pairs:
        # & operator progressively builds the mask
        mask &= df[col] == filter

    result = df.loc[mask, return_col]

    if len(result) != 1:
        raise ValueError(f"Expected exactly one result, got {len(result)}")

    return result.iloc[0]


def set_DFcell(
    df: pd.DataFrame, paste_col: str, value, *colfilters_pairs: tuple[str, Any]
):

    if not colfilters_pairs:
        raise TypeError("At least one pair (column, filter) is required.")

    mask = pd.Series(True, index=df.index)

    for col, filter in colfilters_pairs:
        mask &= df[col] == filter

    if len(cell := df.loc[mask, paste_col]) != 1:
        raise ValueError(f"Expected exactly one result, got {len(cell)}")

    df.loc[mask, paste_col] = value


def try_fetch_NR_refs(
    NR: str, definedNames: DefinedNameDict, c: IssuesCollector
) -> tuple:
    try:
        defn: DefinedName = definedNames[NR]
    except KeyError:
        c.issues.append(
            f"Name range '{NR}' is not valid and has been ignored in the migration."
        )
        return "", ""

    try:
        dest = list(defn.destinations)
    except AttributeError:
        c.issues.append(
            f"Name range '{NR}' has an unreadable destination (value: {defn.attr_text!r}). This name range has been ignored in the migration."
        )
        return "", ""

    if len(dest) != 1:
        c.issues.append(
            f"Name range '{NR}' has {len(dest)} destinations (expected 1; value: {defn.attr_text!r}). This name range has been ignored in the migration."
        )
        return "", ""

    sheet, range = dest[0]

    # range could silently be wrong
    if None in range_boundaries(range) and range:
        c.issues.append(
            f"Name range '{NR}' has an invalid range (rng: {range!r}), and has been ignored in the migration."
        )
        return "", ""

    return sheet, range


def get_NR_singlevalue(wb: Workbook, name: str, c: IssuesCollector) -> Any:
    sheet, range = try_fetch_NR_refs(name, wb.defined_names, c)
    if not sheet and not range:
        return None

    shape = make_shape(range)
    if shape._left == shape._right and shape._top == shape._bottom:
        return wb[sheet].cell(shape._top, shape._left).value
    else:
        raise ValueError(
            f"Do not use get_NR_singlevalue for a range. Name range '{name}' is a range."
        )


def get_version(wb: Workbook, c: IssuesCollector) -> str:
    version = get_NR_singlevalue(wb, VERSION_NR, c)
    if version is None or not str(version).strip():
        raise MigrationError(
            f"Cannot determine the template version: named range '{VERSION_NR}' "
            "is missing or its cell is empty."
        )
    return version


def check_status_incomplete(wb: Workbook, status: str, c: IssuesCollector) -> bool:
    """
    Check if the workbook is filled out or not based on the value of template_overall_validation_status.
    Check is manual as openpyxl displays only the resulting label (in the selected language).
    """
    status_cell = get_NR_singlevalue(wb, status, c)
    return status_cell in [
        "INCOMPLETE",
        "UFÆRDIG",
        "ONVOLLEDIG",
        "INCOMPLET",
        "UNVOLLSTÄNDIG",
        "NEAMHIOMLÁN",
        "INCOMPLETO",
        "NEBAIGTA",
        "NIEKOMPLETNY",
        "INCOMPLETO",
        "NEPOPOLNO",
        "INCOMPLETO",
    ]


def create_table_of_contents(wb_values) -> dict[str, int]:
    """Create a dict with keys as names of ToC sections and values as the corresponding row numbers in 'Table of Contents & Validation' sheet.
    If there is any change in where the ToC is located in the sheet, or changes in the range of cells utilised for it, this function should be updated."""
    _keys = [
        cell[0].value for cell in wb_values["Table of Contents & Validation"]["B9:B65"]
    ]
    _values = list(range(9, 66))  # row numbers
    return dict(zip(_keys, _values))


Cell = tuple[str, int, int]  # sheet, column, row


def _isFormula(value: Any) -> bool:
    return isinstance(value, ArrayFormula) or (
        isinstance(value, str) and value.startswith("=")
    )


def markerNames(wb: Workbook) -> set[str]:
    """Names that only mark out an area (e.g. a table): every cell is covered by
    disjoint, more specific names, which hold the data. Names with identical
    areas are aliases and count once, markers nested inside don't count;
    formula cells, and merged cells whose top-left is covered or a formula,
    need no name."""
    areas: dict[str, frozenset[Cell]] = {}
    for name, defn in wb.defined_names.items():
        try:
            destinations = list(defn.destinations)
        except AttributeError:
            continue
        cells: set[Cell] = set()
        for sheet, ref in destinations:
            left, top, right, bottom = range_boundaries(ref)
            if sheet not in wb.sheetnames or None in (left, top, right, bottom):
                break
            cells.update(
                (sheet, col, row)
                for col in range(cast(int, left), cast(int, right) + 1)
                for row in range(cast(int, top), cast(int, bottom) + 1)
            )
        else:
            if cells:
                areas[name] = frozenset(cells)

    topLefts: dict[Cell, Cell] = {
        (ws.title, col, row): (ws.title, merged.min_col, merged.min_row)
        for ws in wb.worksheets
        for merged in ws.merged_cells.ranges
        for col in range(merged.min_col, merged.max_col + 1)
        for row in range(merged.min_row, merged.max_row + 1)
    }

    def needsNoName(cell: Cell, covered: set[Cell]) -> bool:
        sheet, col, row = cell
        return cell in covered or _isFormula(wb[sheet].cell(row, col).value)

    markers: set[str] = set()
    # smallest first, so a table nested in another is known to be a marker
    # (and not one of the outer table's data names) by the time we get there
    for name, area in sorted(areas.items(), key=lambda item: len(item[1])):
        inner = {
            a
            for n, a in areas.items()
            if n not in markers and len(a) < len(area) and a <= area
        }
        covered: set[Cell] = set()
        for innerArea in inner:
            if covered & innerArea:
                break
            covered |= innerArea
        else:
            if inner and all(
                needsNoName(cell, covered)
                or ((topLeft := topLefts.get(cell)) and needsNoName(topLeft, covered))
                for cell in area
            ):
                markers.add(name)
    return markers


def access_NR_table(
    wb: Workbook,
    c: IssuesCollector,
    old_sheet_names: list[str] | None = None,
) -> pd.DataFrame:
    """Access names, cell references and coordinates of each name range of the workbook.
    Returns a (pandas) DataFrame.
    There were some VSME samples that, after processing in openpyxl, returned wrong sheet references for their name ranges (e.g. "[1]General Information" instead of "General Information");
    To handle this, a list of issues (str explaining where the issue is) is returned alongside the DataFrame."""

    rows: list[tuple[str, str | None, str | None, Shape]] = []
    markers = markerNames(wb)

    for nr in wb.defined_names:
        if nr in markers:
            continue
        sheet, rng = try_fetch_NR_refs(nr, wb.defined_names, c)

        if old_sheet_names is not None and sheet not in old_sheet_names:
            c.issues.append(
                f"Name range '{nr}' refers to sheet '{sheet}' which is not present in the old workbook. This name range has been ignored in the migration."
            )
            sheet = None

        # `rng or None` keeps the cell_ranges column NA-detectable (empty -> None);
        # make_shape independently treats an empty range as a NullShape.
        rows.append((nr, sheet, rng or None, make_shape(rng)))

    df_populated = pd.DataFrame(
        rows, columns=["name_ranges", "sheets", "cell_ranges", "cell_shapes"]
    )
    return df_populated


def access_missingNR_table(missingNR: VersionCollection, version: str) -> pd.DataFrame:
    """
    Access sheets, cell references and coordinates of each missing name range.
    The json data contained in missingNR is listed by hand because certain name ranges were not present in old VSME versions.
    The json is saved in src/migration_tool/data (see validations.py for data structure); the target version's
    anchors must be resolved first (see resolveAnchors).
    """

    list_of_labels = missingNR.get_labels(version)
    list_of_sheets = missingNR.get_sheets(version)
    list_of_ranges = missingNR.get_ranges(version)
    list_of_shapes = [make_shape(rng) for rng in list_of_ranges]

    return pd.DataFrame(
        {
            "labels": list_of_labels,
            "sheets": list_of_sheets,
            "cell_ranges": list_of_ranges,
            "cell_shapes": list_of_shapes,
        }
    )


TEMPLATE_LABEL = re.compile(r"=\s*(template_label_\w+)\s*")


def _labelRows(ws: Worksheet) -> dict[str, set[int]]:
    """The rows of each `=template_label_…` cell (whole formula, not a mention)
    and of each plain-text cell, for sheets that have no template labels."""
    rows: dict[str, set[int]] = {}
    for rowNumber, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for value in row:
            text = value.text if isinstance(value, ArrayFormula) else value
            if not isinstance(text, str):
                continue
            if match := TEMPLATE_LABEL.fullmatch(text):
                rows.setdefault(match.group(1), set()).add(rowNumber)
            elif not text.startswith("="):
                rows.setdefault(text, set()).add(rowNumber)
    return rows


def resolveAnchors(missingNR: VersionCollection, version: str, wb: Workbook) -> None:
    """Fill in the range of each of the version's anchored refs from wb, the
    template the anchors describe."""
    labelRows: dict[str, dict[str, set[int]]] = {}
    for ref in missingNR.getRefs(version):
        if not (anchor := ref.anchor):
            continue
        if ref.sheet not in labelRows:
            labelRows[ref.sheet] = _labelRows(wb[ref.sheet])
        rows = labelRows[ref.sheet].get(anchor.label, set())
        if len(rows) != 1:
            raise MigrationError(
                f"The target template has {len(rows)} rows labelled '{anchor.label}' in '{ref.sheet}' (for '{ref.label}'); expected exactly one."
            )
        (labelRow,) = rows
        top = labelRow + anchor.rowOffset
        bottom = top + anchor.rows - 1
        first, _, last = anchor.cols.partition(":")
        last = last or first
        if (first, top) == (last, bottom):
            ref.range = absolute_coordinate(f"{first}{top}")
        else:
            ref.range = absolute_coordinate(f"{first}{top}:{last}{bottom}")


def apply_changes_NR(
    df: pd.DataFrame,
    version_cell: str,
    version_cell_new: str,
    NR_changes: dict[str, str],
) -> pd.DataFrame:
    """Change name ranges based on the version of the old workbook, and return a df with the updated name ranges.
    This is necessary because some name ranges have changed from version to version.
    Data comes from NR_changes.json, to be updated each time a NR name changes."""

    df_of_changes_NR = pd.DataFrame(
        {
            "name_ranges": NR_changes[version_cell],
            "name_ranges_new": NR_changes[version_cell_new],
        }
    )

    df = df.merge(df_of_changes_NR, on="name_ranges", how="left")

    for i in range(len(df)):
        if pd.notna(df.loc[i, "name_ranges_new"]):
            df.loc[i, "name_ranges"] = df.loc[i, "name_ranges_new"]

    df = df.drop(columns=["name_ranges_new"])

    return df


def change_wastes(df: pd.DataFrame, mapping, c: IssuesCollector):
    """Change waste categories based on mapping_wastes.pkl provided in src/migration_tool/data/mapping.
    Mapping is based on the changes in waste categories in the new EU Regulation (see more in wastes.xlsx in base dir).
    Issues appended are old waste categories not present in the new Regulation)."""

    WASTECATEGORIES_NR = "TypeOfWasteAxis"

    waste_block = cast(
        Value, get_DFcell(df, "cell_values", ("name_ranges", WASTECATEGORIES_NR))
    )
    old_wastes = waste_block.first_element_row(double_list=False)

    new_wastes: list[list[str]] = []
    print_return = "not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955"
    for i in old_wastes:
        if isinstance(i, str):
            if pd.isna(new_cat := get_DFcell(mapping, "new", ("old", i))):
                new_wastes.append([f"Waste category '{i.strip()}' {print_return}"])
                c.issues.append(f"Waste category '{i.strip()}' {print_return}")
            else:
                new_wastes.append([new_cat])

    set_DFcell(
        df, "cell_values", make_value(new_wastes), ("name_ranges", WASTECATEGORIES_NR)
    )


def adjust_wasteValues(
    old_df: pd.DataFrame, new_df: pd.DataFrame, new_wb: Workbook, c: IssuesCollector
) -> pd.DataFrame:
    REUSE_RECYCLE_NR = "WasteDivertedToRecycleOrReuseMass"
    DIVERTED_NR = "WasteDirectedToDisposalMass"
    TOTMASS_NR = "WasteGeneratedMass"

    # getting shape of new name range
    sheet, rng = try_fetch_NR_refs(TOTMASS_NR, new_wb.defined_names, c)
    shape = make_shape(rng)

    # calculation from previous mass of reused/recycled and diverted waste
    reuse_recycle = cast(
        Value, get_DFcell(old_df, "cell_values", ("name_ranges", REUSE_RECYCLE_NR))
    )
    diverted = cast(
        Value, get_DFcell(old_df, "cell_values", ("name_ranges", DIVERTED_NR))
    )
    tot_waste = reuse_recycle.sum_two_ranges(diverted)

    # add helper comments to each populated cell in the range
    comment = "Old 'Waste diverted to recycle or reuse' cell value: "
    block = reuse_recycle.values()
    for row in range(shape.top(), shape.top() + len(block)):
        if block and (reuse_recycle_sum := block[row - shape.top()][0]) is not None:
            cell = new_wb[sheet].cell(row=row, column=shape.left())
            cell.comment = Comment(f"{comment}{reuse_recycle_sum}", "Migration tool")

    # adding new row with calculated values to dataframe
    new_df_row = {
        "name_ranges": TOTMASS_NR,
        "sheets": sheet,
        "cell_ranges": rng,
        "cell_shapes": shape,
        "cell_values": tot_waste,
    }
    new_df.loc[len(new_df)] = new_df_row
    return new_df


def convert_months(df: pd.DataFrame, version: str, dateLabels: tuple) -> None:
    if version == "1.0.0":
        mask = df["labels"].isin(dateLabels)
        for _, row in df[mask].iterrows():
            row["cell_values"].convert_month_to_numbers()


def clean_NR_with_no_data(df):
    """Function to filter name ranges to process.
    Some (most starting with template_) were created for Digital Converter and not for storing data.
    Markers (e.g. tables) never get here: access_NR_table skips them (see markerNames).
    Unfortunately, there were some exceptions and edge cases (see list_of_NRs_toNOTremove)."""

    list_of_keywords = [
        "template_",
        "enum_",
    ]
    # the energy table's row labels (template_label_ formulas), not data
    list_of_NRs_toremove = ["BreakdownOfEnergyConsumptionAxis"]

    list_of_NRs_toNOTremove = [
        "template_reporting_entity_name",
        "template_reporting_entity_identifier_scheme",
        "template_reporting_entity_identifier",
        "template_currency",
    ]
    df_toattach = pd.concat(
        [
            df[df["name_ranges"].isin(list_of_NRs_toNOTremove)],
            df[df["name_ranges"].str.contains("template_checkbox", na=False)],
        ]
    )

    for kw in list_of_keywords:
        df = df[~df["name_ranges"].str.contains(kw, na=False)]
    df = df[~df["name_ranges"].isin(list_of_NRs_toremove)]

    return pd.concat([df.reset_index(drop=True), df_toattach]).reset_index(drop=True)


def get_indexes_of_NAs(df, column) -> list[int]:
    """Get the indexes of the rows where the specified column has NA values"""
    return df.loc[df[column].isna()].index.tolist()


def copy_values(pyxl: Workbook, df: pd.DataFrame, *, NRflag: bool) -> pd.DataFrame:
    """
    Copy values from openpyxl workbook based on cell shapes, returns DataFrame with col of cell values.
    If NRflag is True, DataFrame is returned with the name range col;
    if False, with the labels col (see missing_NR.json).
    """

    cell_values = []
    # Python None always becomes NaN after pd.DataFrame creation
    index_sheets = get_indexes_of_NAs(df, "sheets")
    index_ranges = get_indexes_of_NAs(df, "cell_ranges")

    for i in range(len(df)):
        if i not in index_sheets:
            sheet = pyxl[df["sheets"][i]]
            shape = cast(Shape, df["cell_shapes"][i])

            # handling the #REF ranges
            if i not in index_ranges:
                cell_values.append(make_value(shape.build_values(sheet)))
            else:
                cell_values.append(make_value(None))
        # handling issue with sheet names (ex "[1]General Information") in old workbooks
        else:
            cell_values.append(make_value(None))

    df["cell_values"] = cell_values

    if NRflag:
        return df[["name_ranges", "cell_values"]]
    else:
        return df[["labels", "cell_values"]]


def paste_values(pyxl: Workbook, df, version, NR=None, table_of_contents=None):
    """Paste values from a DataFrame column into an openpyxl workbook.

    Treatment for merged cells (merged only across columns in VSMEs):
    - thanks to .merged_cells.ranges method, coordinates of merged cells are extracted and saved in a DataFrame (merged_loc);
    - before pasting, check if top-left cell of the shape is in the merged_loc df, and if so,
      we keep only the first element of each row of values (with .first_element_row method) to fit merged shape.

    Treatment for enlarged ranges (throughout VSME versions) are handled with .enlarged_range_correction method,
    which pastes values only in the original shape of the name range and None(s) for the rest of the enlarged range.

    The NR=True flag argument is used for specifically handling the checkboxes of the biodiversity sites, as if True is pasted the rest of the cols must have False for uniformity (see Template).
    """

    sheet_names = pyxl.sheetnames
    merged_loc = pd.DataFrame()

    for sheet in sheet_names:
        merged = list(pyxl[sheet].merged_cells.ranges)
        merged_list = []
        for i in range(len(merged)):
            merged_list.append(range_boundaries(str(merged[i]))[:2])
        merged_loc = pd.concat([merged_loc, pd.DataFrame({sheet: merged_list})], axis=1)

    SITE_IN_BIOAREA_NR = "SiteLocatedInABiodiversitySensitiveArea"
    SITE_NEAR_BIOAREA_NR = "SiteLocatedNearABiodiversitySensitiveArea"

    add_checkbox = (SITE_IN_BIOAREA_NR, SITE_NEAR_BIOAREA_NR)

    # the following missingNR ranges have shrinked in VS, so they need specific handling just for v 1.2.0
    # (v 1.2.0 introduced the checkboxes in the ToC but not their related name ranges)
    PROBLEMATIC_TOC_REFERENCES = [
        "B3_B4_ToC",
        "B5_ToC",
        "B8_B9_ToC",
        "C8_C9_ToC",
    ]

    for i in range(len(df)):
        sheet = pyxl[df["sheets"][i]]
        rng = df["cell_ranges"][i]
        shape = cast(Shape, df["cell_shapes"][i])
        value = cast(Value, df["cell_values"][i])

        # specific handling for list of classified information
        if table_of_contents is not None:
            if df["name_ranges"][i] == OMITTEDDISCL_NR:
                classified_info_handling(value, table_of_contents, pyxl, version)
                continue
        else:
            if (
                version == "1.2.0"
                and (label := df["labels"][i]) in PROBLEMATIC_TOC_REFERENCES
            ):
                value = classified_info_handling_120(value, label)

        if value.values() is None or pd.isna(rng):  # nothing to paste
            continue

        if shape.isonecell():
            sheet[rng] = value.topleft()  # when it's one cell, paste topleft

        else:
            tuple_to_check = (shape.left(), shape.top())

            if (
                tuple_to_check in merged_loc[df["sheets"][i]].values.tolist()
            ):  # merged cells check
                value.first_element_row()  # keeping only the first value of each row
                value.enlarged_range_correction(shape)  # enlarged ranges check
                value.paste(sheet, shape)

            else:
                value.enlarged_range_correction(shape)  # enlarged ranges check
                if NR is not None and df["name_ranges"][i] in add_checkbox:
                    value.add_checkboxes()  # add checkboxes for the specific NRs (see above)
                value.paste(sheet, shape)


def classified_info_handling(value, table_of_contents, pyxl, version) -> None:
    """Specific handling for the list of classified information (see new functionality added in version 1.2.0).
    Check where the first 2 characters (ex B1) of "value" (choices about classified info in old versions) match in the ToC,
    and add True in the corresponding row in col D of new "Table of Contents & Validation" sheet where formulas are not detected (formula cells get automatically updated upon opening new version).

    For migrations from 1.2.0 to newer versions, the cells are hard-coded in missingNR_df.
    """

    if version in ["1.0.0", "1.0.1", "1.1.0", "1.1.1"]:
        for input in value.values():
            if input[0] is not None:
                match = [
                    s for s in table_of_contents if input[0][0:2] in s
                ]  # matches based on first 2 characters

                if match:
                    row_n = [table_of_contents.get(key) for key in match]
                    for row in row_n:
                        cell = pyxl["Table of Contents & Validation"].cell
                        if not check_formula(cell(row=row, column=4)):
                            cell(row=row, column=4).value = True


def classified_info_handling_120(value: Value, label: str) -> Value:
    """Remove based on changes in the 1.2.0 and 2.0.0 ToCs. See the Templates for more."""
    match label:
        case "B3_B4_ToC":
            value.remove_row(3)  # 3 means the 4th item
        case "B5_ToC":
            value.remove_row(1)
        case "B8_B9_ToC":
            value.remove_row(4)
        case "C8_C9_ToC":
            value.remove_row(1)

    return value


def adjust_countriesOfOperation(
    df_new_wv: pd.DataFrame, missingNR_df_new_wv: pd.DataFrame, old_wb: Workbook
) -> None:

    def add_TrueOrFalse_to_df(df, label: str, bool) -> None:
        val = make_value([[bool]])
        set_DFcell(df, "cell_values", val, ("labels", label))

    # resolving whether undertaking operates in more than one country
    COUNTRIES_EMPLOYMENT_NR = "CountryOfEmploymentContractAxis"
    length = get_DFcell(
        df_new_wv, "cell_values", ("name_ranges", COUNTRIES_EMPLOYMENT_NR)
    ).count_uniques()

    # > 2 because one value is None
    add_TrueOrFalse_to_df(
        missingNR_df_new_wv, "OperationInMoreThanOneCountry", length > 2
    )

    # resolving Fuel Converter transfer cell
    first_fuel = old_wb["Fuel Converter"]["B10"].value
    add_TrueOrFalse_to_df(
        missingNR_df_new_wv, "FuelTotCalculatedResult", bool(first_fuel)
    )


def add_helperComments_social(
    old_wb_values: Workbook,
    new_wb: Workbook,
    missingNR_df: pd.DataFrame,
    version: str,
    c: IssuesCollector,
) -> None:
    """In VS (2.0.0), new cell data without name ranges have been added in Social Disclosures.
    This function adds helper comments in the new workbook, taking relevant refs from the missing NR dataframe.
    In the case of training hours, the average time per employee"""

    # CELL: "Select whether the rate is per 100 or 500 full-time workers over a yearly time frame"
    shape = cast(
        Shape,
        get_DFcell(
            missingNR_df, "cell_shapes", ("labels", "WhetherRateIsPer50OrPer100")
        ),
    )
    cell = new_wb["Social Disclosures"].cell(row=shape.top(), column=shape.left())
    cell.comment = Comment("New field included with the VS release.", "Migration tool")

    # CELL: "Is the undertaking is already required by EU law or other national regulations to report the percentage gap in pay between its female and male employees?"
    shape = cast(
        Shape,
        get_DFcell(
            missingNR_df, "cell_shapes", ("labels", "GenderPercentageGapRequiredByLaw")
        ),
    )
    cell = new_wb["Social Disclosures"].cell(row=shape.top(), column=shape.left())
    cell.comment = Comment(
        "New field included with the VS release. Modify to TRUE if applicable.",
        "Migration tool",
    )

    # TRAINING HOURS (old references will never change)
    refs_averageTrainingHours = {
        "1.0.0": "C71",
        "1.0.1": "C71",
        "1.1.0": "C72",
        "1.1.1": "C162",
        "1.2.0": "C162",
        "1.3.0": "C162",
    }
    averageTrainingHours = old_wb_values["Social Disclosures"][
        refs_averageTrainingHours[version]
    ].value

    shape = cast(
        Shape,
        get_DFcell(missingNR_df, "cell_shapes", ("labels", "TrainingHours")),
    )
    cell = new_wb["Social Disclosures"].cell(row=shape.top(), column=shape.left())
    cell.comment = Comment(
        f"Old 'Average number of annual training hours per employee' was {averageTrainingHours}",
        "Migration tool",
    )

    c.issues.append(
        "New required information in the 'Social Disclosures' sheet (see cell-level comments for more)."
    )


def create_or_update_migration_status(pyxl: Workbook) -> None:
    """Create or update the name range template_migration_status in the Introduction sheet, cell D1.
    The cell returns TRUE if the migration process has been completed, but will be processed as None in openpyxl (data_only=True) mode
    if the workbook has not been opened after migration.
    The name range is used as starting check in the Converter, in the case of workbooks of the newest version."""

    if pyxl.defined_names.get("template_migration_status") is None:
        ws = pyxl["Introduction"]
        ref = f"{quote_sheetname(ws.title)}!{absolute_coordinate('D2')}"
        defn = DefinedName(name="template_migration_status", attr_text=ref)

        pyxl.defined_names.add(defn)
        ws["D1"].value = "Migration status"
        ws["D1"].alignment = Alignment(horizontal="center")
        ws["D2"].value = "=AND(TRUE,OR(FALSE,TRUE))"


def assess_energyConsumption_validation(df: pd.DataFrame, c: IssuesCollector):
    """Fuel Converter for first 3 versions does not sum from the third added fuel.
    This function checks whether a third fuel (in C12) has been added,
    and appends an issue to issuesCollector list since
    the newest version should display a validation error next to the energy cons. cells."""

    if get_DFcell(df, "cell_values", ("labels", "FuelAmount3")).topleft():
        c.issues.append(
            "Issues in Energy Consumption sums. Please check the Environmental Disclosures and the Fuel Converter sheets."
        )
