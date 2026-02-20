from openpyxl import load_workbook, Workbook
import pandas as pd
import warnings
import time
import os
from importlib.resources import files, as_file

from .outils import (
    create_table_of_contents,
    access_NR_table,
    access_missingNR_table,
    apply_changes_NR,
    change_wastes,
    clean_NR_with_no_data,
    copy_values,
    paste_values,
    flatten_sublists_lc,
)
from .classes import values


def tool(old_wb: Workbook | str | os.PathLike):
    start_time = time.time()

    mapping_wastes = pd.read_pickle(
        files("migration_tool.data").joinpath("Mapping_wastes.pkl")
    )
    missingNR_df = pd.read_pickle(
        files("migration_tool.data").joinpath("missingNR_df.pkl")
    )

    # load old filled-out Template
    if isinstance(old_wb, (str, os.PathLike)):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            old_wb_obj = load_workbook(old_wb, data_only=False)
    elif old_wb is None:
        raise ValueError("old_wb must be a file path or an openpyxl Workbook")
    else:
        # already a workbook-like object
        old_wb_obj = old_wb

    # load new empty Template
    template_dir = files("migration_tool.data")
    with as_file(template_dir) as template_path:
        template_files = list(template_path.glob("VSME-Digital-Template-*.xlsx"))
        if not template_files:
            raise FileNotFoundError(
                "No template file found matching pattern VSME-Digital-Template-*.xlsx"
            )
        new_wb_empty = load_workbook(template_files[0], data_only=False)
        new_wb_empty_values = load_workbook(template_files[0], data_only=True)

    list_migrationissues = []

    table_of_contents = create_table_of_contents(new_wb_empty_values)

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
        list_migrationissues.append(change_wastes(df_old_tomerge, mapping_wastes))

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
    paste_values(new_wb_empty, df_new_wv, NR=True, table_of_contents=table_of_contents)

    elapsed = time.time() - start_time
    return new_wb_empty, elapsed, flatten_sublists_lc(list_migrationissues)
