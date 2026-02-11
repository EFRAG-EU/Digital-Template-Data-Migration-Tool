import time
import warnings
from importlib.resources import as_file, files

import pandas as pd
from openpyxl import Workbook, _ZipFileFileProtocol, load_workbook

from .classes import values
from .outils import (
    access_missingNR_table,
    access_NR_table,
    apply_changes_NR,
    change_wastes,
    clean_NR_with_no_data,
    copy_values,
    paste_values,
)


def migrate_workbook(
    old_wb: Workbook | _ZipFileFileProtocol,
) -> tuple[Workbook, float, list[str]]:
    start_time = time.time()

    mapping_wastes = pd.read_pickle(
        files("migration_tool.data").joinpath("Mapping_wastes.pkl")
    )
    missingNR_df = pd.read_pickle(
        files("migration_tool.data").joinpath("missingNR_df.pkl")
    )

    # load old filled-out Template
    if isinstance(old_wb, _ZipFileFileProtocol):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            old_wb_obj = load_workbook(old_wb, data_only=False)
    elif isinstance(old_wb, Workbook):
        # already a workbook-like object
        old_wb_obj = old_wb
    else:
        raise ValueError(
            f"old_wb [{type(old_wb)}] must be a file path or an openpyxl Workbook"
        )

    # load new empty Template
    with as_file(
        files("migration_tool.data").joinpath("VSME-Digital-Template-1.1.1.xlsx")
    ) as path:
        new_wb_empty = load_workbook(path, data_only=False)

    list_migrationissues: list[str] = []

    version_cell = old_wb_obj["Introduction"].cell(row=1, column=3).value
    version_cell_new = new_wb_empty["Introduction"].cell(row=1, column=3).value

    df_old = access_NR_table(old_wb_obj.defined_names)

    df_new = access_NR_table(new_wb_empty.defined_names)

    missingNR_df_old = access_missingNR_table(missingNR_df, version_cell)
    missingNR_df_new = access_missingNR_table(missingNR_df, version_cell_new)

    df_old_wv = copy_values(old_wb_obj, df_old, key="name_ranges")
    missingNR_df_old_values = copy_values(old_wb_obj, missingNR_df_old, key=None)

    if version_cell == "1.0.0":
        for position in [0, 1, 6]:
            missingNR_df_old_values[position].convert_month_to_numbers()

    df_old_tomerge = apply_changes_NR(df_old_wv, version_cell, version_cell_new)
    df_old_tomerge = clean_NR_with_no_data(df_old_tomerge)
    if version_cell in ["1.0.0", "1.0.1"]:
        list_migrationissues.extend(change_wastes(df_old_tomerge, mapping_wastes))

    df_new_wv = df_new.merge(df_old_tomerge)
    missingNR_df_new_wv = pd.concat([missingNR_df_new, missingNR_df_old_values], axis=1)

    if version_cell in ["1.0.0", "1.0.1"]:
        length = (
            df_new_wv.loc[
                df_new_wv["name_ranges"] == "CountryOfEmploymentContractAxis",
                "cell_values",
            ]
            .values[0]
            .count_uniques()
        )
        if (
            length > 2
        ):  # ignore PyLance warning for pandas dataframes, no issues at runtime
            missingNR_df_new_wv.loc[
                (missingNR_df_new_wv["sheets"] == "Social Disclosures")
                & (missingNR_df_new_wv["cell_ranges"] == "$E$27"),
                "cell_values",
            ] = values([[True]])
        else:
            missingNR_df_new_wv.loc[
                (missingNR_df_new_wv["sheets"] == "Social Disclosures")
                & (missingNR_df_new_wv["cell_ranges"] == "$E$27"),
                "cell_values",
            ] = values([[False]])

    paste_values(new_wb_empty, missingNR_df_new_wv)
    paste_values(new_wb_empty, df_new_wv, NR=True)

    elapsed = time.time() - start_time
    return new_wb_empty, elapsed, list_migrationissues
