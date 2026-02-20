from typing import Dict
import pandas as pd
from openpyxl.utils import range_boundaries

from .classes import shapes, values, check_formula


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
    _keys = [
        cell[0].value for cell in wb_values["Table of Contents & Validation"]["B9:B68"]
    ]
    _values = list(range(9, 69))
    return dict(zip(_keys, _values))


def access_NR_table(pyxl_NR, sheets=None):
    """Access names, cell references and coordinates of each name range from the python object containing name ranges"""

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
    """Access sheets, cell references and coordinates of each missing name range"""

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
    """Copy values from old workbook to new workbook based on name ranges and cell shapes, and return a df with key and cell values"""

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
    """Paste values from a DataFrame column into an openpyxl workbook"""

    sheet_names = pyxl.sheetnames
    merged_loc = pd.DataFrame()

    for sheet in sheet_names:
        merged = list(pyxl[sheet].merged_cells.ranges)
        merged_list = []
        for i in range(len(merged)):
            str(merged[i])
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
