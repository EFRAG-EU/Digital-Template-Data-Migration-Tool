import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.utils import absolute_coordinate, quote_sheetname, range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from typing import Dict

from .classes import check_formula, shapes, values


def check_status_incomplete(openpyxl_obj) -> bool:
    """Check if the workbook is filled out or not based on the value of the 'Status' cell in the 'Table of Contents & Validation' sheet"""

    status_cell = openpyxl_obj["Table of Contents & Validation"]["C3"].value
    if status_cell in [
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
        "INCOMPLETO",
    ]:
        return True
    else:
        return False


def create_table_of_contents(wb_values) -> Dict[str, int]:
    """Create a dict with keys as names of ToC sections and values as the corresponding row numbers in 'Table of Contents & Validation' sheet.
    If there is any change in where the ToC is located in the sheet, or changes in the range of cells utilised for it, this function should be updated."""
    _keys = [
        cell[0].value for cell in wb_values["Table of Contents & Validation"]["B9:B68"]
    ]
    _values = list(range(9, 69))
    return dict(zip(_keys, _values))


def access_NR_table(pyxl_NR, sheets=None):
    """Access names, cell references and coordinates of each name range from the python object containing name ranges.
    Returns a (pandas) DataFrame.
    There were some VSME samples that, after processing in openpyxl, returned wrong sheet references for their name ranges (e.g. "[1]General Information" instead of "General Information");
    To handle this, a list of issues (str explaining where the issue is) is returned alongside the DataFrame."""

    list_NR = list(pyxl_NR.keys())
    list_sheets = []
    list_ranges = []
    list_shapes = []

    list_NR_issues = []

    for NR in list_NR:
        list_temp = pyxl_NR[NR].attr_text.split("!")[0].replace("'", "")
        if sheets is not None and list_temp not in sheets:
            list_NR_issues.append(
                f"Name range '{NR}' refers to sheet '{list_temp}' which is not present in the old workbook. This name range has been ignored in the migration."
            )
            list_temp = None
        list_sheets.append(list_temp)

        rng_temp = pyxl_NR[NR].attr_text.split("!")[1]
        if rng_temp == "#REF":  # handling #REF name ranges (in Translations sheet)
            rng_temp = None
        list_ranges.append(rng_temp)

    for rng in list_ranges:
        if rng is None:
            list_shapes.append(shapes(None))
        else:
            list_shapes.append(shapes(range_boundaries(rng)))

    df_populated = pd.DataFrame(
        {
            "name_ranges": list_NR,
            "sheets": list_sheets,
            "cell_ranges": list_ranges,
            "cell_shapes": list_shapes,
        }
    )
    if sheets is not None:
        return df_populated, list_NR_issues
    else:
        return df_populated


def access_missingNR_table(df_missingNR, version_cell):
    """Access sheets, cell references and coordinates of each missing name range.
    the func argument df_missingNR is an hard-coded DataFrame created because certain name ranges were not present in old VSME versions.
    The DataFrame (missingNR_df.pkl) is saved in src/migration_tool/data, and must be updated for every new VSME version,
    with a new column for each version (and its corresponding references for every missing name range)."""

    list_of_sheets = []
    list_of_ranges = []

    for i in range(len(df_missingNR)):
        list_of_sheets.append(df_missingNR.loc[:, version_cell].values[i].split("!")[0])
        list_of_ranges.append(df_missingNR.loc[:, version_cell].values[i].split("!")[1])

    for rang in list_of_ranges:
        if rang == "None":
            list_of_ranges[list_of_ranges.index(rang)] = None

    list_of_shapes = [
        shapes(range_boundaries(rng)) if rng is not None else shapes(None)
        for rng in list_of_ranges
    ]

    return pd.DataFrame(
        {
            "sheets": list_of_sheets,
            "cell_ranges": list_of_ranges,
            "cell_shapes": list_of_shapes,
        }
    )


def apply_changes_NR(df, version_cell, version_cell_new):
    """Change name ranges based on the version of the old workbook, and return a df with the updated name ranges.
    This is necessary because some name ranges have changed from version to version.
    This dictionary could be transformed into a pd DataFrame and saved it as a pickle like missingNR_df if
    many name changes will occur in the future, but for now it is hard-coded here."""

    dict_of_changes_NR = {
        "1.0.0": [
            "NumberOfPermanentContactEmployees",
            "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherinterestedPartiesOrCommunitiesInGovernance",
            "MostSeniorLevelAccountableForImplementationOfPracticesPoliciesAndOrFutureInitiatives",
        ],
        "1.0.1": [
            "NumberOfPermanentContactEmployees",
            "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherinterestedPartiesOrCommunitiesInGovernance",
            "MostSeniorLevelAccountableForImplementationOfPracticesPoliciesAndOrFutureInitiatives",
        ],
        "1.1.0": [
            "NumberOfPermanentContractEmployees",
            "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherInterestedPartiesOrCommunitiesInGovernance",
            "MostSeniorLevelAccountableForImplementationOfPolicies",
        ],
        "1.1.1": [
            "NumberOfPermanentContractEmployees",
            "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherInterestedPartiesOrCommunitiesInGovernance",
            "MostSeniorLevelAccountableForImplementationOfPolicies",
        ],
        "1.2.0": [
            "NumberOfPermanentContractEmployees",
            "DescriptionOfTheEffectiveParticipationOfWorkersUsersOrOtherInterestedPartiesOrCommunitiesInGovernance",
            "MostSeniorLevelAccountableForImplementationOfPolicies",
        ],
    }

    df_of_changes_NR = pd.DataFrame(
        {
            "name_ranges": dict_of_changes_NR[version_cell],
            "name_ranges_new": dict_of_changes_NR[version_cell_new],
        }
    )

    df = df.merge(df_of_changes_NR, on="name_ranges", how="left")

    for i in range(len(df)):
        if pd.notna(df.loc[i, "name_ranges_new"]):
            df.loc[i, "name_ranges"] = df.loc[i, "name_ranges_new"]

    df = df.drop(columns=["name_ranges_new"])

    return df


def change_wastes(df, mapping) -> list[str]:
    """Change waste categories based on mapping_wastes.pkl provided in src/migration_tool/data/mapping.
    Mapping is based on the changes in waste categories in the new EU Regulation (see more in wastes.xlsx in base dir).
    Returns a list of issues encountered during the change (e.g. old waste category not present in the new Regulation)."""

    old_wastes = (
        df.loc[df["name_ranges"] == "TypeOfWasteAxis", "cell_values"]
        .values[0]
        .first_element_row(double_list=False)
    )
    new_wastes = []
    list_wasteissues: list[str] = []
    for i in old_wastes:
        if i is not None:
            if pd.isna(mapping.loc[mapping["old"] == i, "new"].values[0]):
                new_wastes.append(
                    [
                        f"Waste category {i} not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955"
                    ]
                )
                list_wasteissues.append(
                    f"Waste category --{i}-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955"
                )
            else:
                new_wastes.append([mapping.loc[mapping["old"] == i, "new"].values[0]])
    df.loc[df["name_ranges"] == "TypeOfWasteAxis", "cell_values"] = values(new_wastes)

    return list_wasteissues


def clean_NR_with_no_data(df):
    """Function to filter name ranges to process.
    Some (most starting with template_) were created for Digital Converter and not for storing data.
    Unfortunately, there were some exceptions and edge cases (see list_of_NRs_toNOTremove)."""

    list_of_keywords = [
        "template_",
        "enum_",
        "Table",
        "BreakdownOfEnergyConsumptionAxis",
        "Hypercube",
    ]

    list_of_NRs_toNOTremove = [
        "template_reporting_entity_name",
        "template_reporting_entity_identifier_scheme",
        "template_reporting_entity_identifier",
        "template_currency",
    ]
    df_toattach = df[df["name_ranges"].isin(list_of_NRs_toNOTremove)]

    for kw in list_of_keywords:
        df = df[~df["name_ranges"].str.contains(kw, na=False)]

    return pd.concat([df.reset_index(drop=True), df_toattach]).reset_index(drop=True)


def get_indexes_of_NAs(df, column) -> list[int]:
    """Get the indexes of the rows where the specified column has NA values"""
    return df.loc[df[column].isna()].index.tolist()


def copy_values(pyxl, df, key=None):
    """Copy values from openpyxl workbook based on cell shapes, returns DataFrame with col of cell values.
    2 cases: 1) if key is provided, a DataFrame with 2 cols (name ranges and values) is returned;
             2) if key is not provided, just the cell values are returned as a pandas Series."""

    cell_values = []
    index_sheets = get_indexes_of_NAs(df, "sheets")
    index_ranges = get_indexes_of_NAs(df, "cell_ranges")

    for i in range(len(df)):
        if i not in index_sheets:
            sheet = pyxl[df["sheets"][i]]
            shape = df["cell_shapes"][i]

            # handling the #REF ranges
            if i not in index_ranges:
                cell_values.append(values(shape.build_values(sheet)))
            else:
                cell_values.append(values(None))
        # to handle issue with sheet names (ex "[1]General Information") in old workbooks
        else:
            cell_values.append(values(None))

    df["cell_values"] = cell_values

    if key is not None:
        return df[[key, "cell_values"]]
    else:
        return df["cell_values"]


def paste_values(pyxl, df, NR=None, table_of_contents=None):
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

    add_checkbox = [
        "SiteLocatedInABiodiversitySensitiveArea",
        "SiteLocatedNearABiodiversitySensitiveArea",
    ]

    for i in range(len(df)):
        sheet = pyxl[df["sheets"][i]]
        rng = sheet[df["cell_ranges"][i]]
        shape = df["cell_shapes"][i]
        value = df["cell_values"][i]

        # specific handling for list of classified information
        if table_of_contents is not None:
            if (
                df["name_ranges"][i]
                == "ListOfOmittedDisclosuresDeemedToBeClassifiedOrSensitiveInformation"
            ):
                classified_info_handling(value, table_of_contents, pyxl)
                continue

        if shape.isonecell():
            rng.value = value.topleft()  # when it's one cell, paste topleft

        else:
            tuple_to_check = (shape.left(), shape.top())

            if value.values() is not None:  # avoiding formulas
                if (
                    tuple_to_check in merged_loc[df["sheets"][i]].values.tolist()
                ):  # merged cells check
                    value.first_element_row()  # keeping only the first value of each row
                    value.enlarged_range_correction(shape)  # enlarged ranges check
                    value.paste(sheet, shape)

                else:
                    value.enlarged_range_correction(shape)  # enlarged ranges check
                    if NR is not None:
                        if df["name_ranges"][i] in add_checkbox:
                            value.add_checkboxes()  # add checkboxes for the specific NRs (see above)
                    value.paste(sheet, shape)


def classified_info_handling(value, table_of_contents, pyxl):
    """Specific handling for the list of classified information (see new functionality added in version 1.2.0).
    Check where the first 2 characters (ex B1) of "value" (choices about classified info in old versions) match in the ToC,
    and add True in the corresponding row in col D of new "Table of Contents & Validation" sheet where formulas are not detected (formula cells get automatically updated upon opening new version)."""

    for input in value.values():
        if input[0] is not None:
            match = [
                s for s in table_of_contents.keys() if input[0][0:2] in s
            ]  # matches based on first 2 characters
            if match:
                row_n = [table_of_contents.get(key) for key in match]
                for row in row_n:
                    cell = pyxl["Table of Contents & Validation"].cell
                    if not check_formula(cell(row=row, column=4)):
                        cell(row=row, column=4).value = True


def create_or_update_migration_status(pyxl) -> None:
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
