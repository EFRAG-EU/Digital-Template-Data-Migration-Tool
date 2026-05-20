import io
import time
import warnings
from importlib.resources import as_file, files
from io import BytesIO
from pathlib import Path
from typing import TypeAlias

import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook as openpyxl_load_workbook

from .outils import (
    access_missingNR_table,
    access_NR_table,
    adjust_classified_info,
    adjust_data_missing_first2versions,
    apply_changes_NR,
    change_wastes,
    check_status_incomplete,
    clean_NR_with_no_data,
    copy_values,
    create_table_of_contents,
    paste_values,
)

FilePathOrBinaryBlob: TypeAlias = str | Path | io.BufferedIOBase


def load_workbook_quietly(
    file: FilePathOrBinaryBlob, data_only: bool = False
) -> Workbook:
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message=".*? extension is not supported and will be removed",
            category=UserWarning,
            module=r"openpyxl\.worksheet\._reader",
        )
        return openpyxl_load_workbook(file, data_only=data_only)


def migrate_workbook_as_bytes(
    old_wb: FilePathOrBinaryBlob,
) -> tuple[bytes, float, list[str]]:
    new_wb, elapsed, issues = migrate_workbook(old_wb)
    new_wb_bytes = BytesIO()
    new_wb.save(new_wb_bytes)
    return new_wb_bytes.getvalue(), elapsed, issues


def migrate_workbook(
    old_wb: Workbook | FilePathOrBinaryBlob,
) -> tuple[Workbook, float, list[str]]:
    start_time = time.time()

    mapping_wastes = pd.read_pickle(
        files("migration_tool.data").joinpath("Mapping_wastes.pkl").open("rb")
    )
    missingNR_df = pd.read_pickle(
        files("migration_tool.data").joinpath("missingNR_df.pkl").open("rb")
    )

    # load old filled-out Template
    if isinstance(old_wb, FilePathOrBinaryBlob):
        old_wb_obj = load_workbook_quietly(old_wb, data_only=False)
        old_wb_obj_values = load_workbook_quietly(old_wb, data_only=True)
    else:
        raise TypeError(
            f"old_wb [{type(old_wb)}] must be a file path or an openpyxl Workbook"
        )

    # list of migration issues to return (and to be displayed in webpage after migration), remember to update if any new potential issues arise.
    list_migrationissues: list[str] = []
    if check_status_incomplete(old_wb_obj_values):
        list_migrationissues.append(
            "The old workbook is incomplete. Migration happened only for the filled-out cells, but some data might be missing."
        )

    # load new empty Template
    with as_file(
        files("migration_tool.data").joinpath("VSME-Digital-Template-1.2.1.xlsx")
    ) as path:
        new_wb_empty = load_workbook_quietly(path, data_only=False)
        new_wb_empty_values = load_workbook_quietly(path, data_only=True)

    table_of_contents = create_table_of_contents(new_wb_empty_values)

    version_cell = old_wb_obj["Introduction"].cell(row=1, column=3).value
    version_cell_new = new_wb_empty["Introduction"].cell(row=1, column=3).value

    old_wb_sheets = [sheet.title for sheet in old_wb_obj.worksheets]

    df_old, sheets_issues = access_NR_table(old_wb_obj.defined_names, old_wb_sheets)
    if isinstance(sheets_issues, list) and sheets_issues:
        list_migrationissues.extend(sheets_issues)
    df_new, _ = access_NR_table(new_wb_empty.defined_names)

    missingNR_df_old = access_missingNR_table(missingNR_df, version_cell)
    missingNR_df_new = access_missingNR_table(missingNR_df, version_cell_new)

    df_old_wv = copy_values(old_wb_obj, df_old, key="name_ranges")
    missingNR_df_old_values = copy_values(old_wb_obj, missingNR_df_old, key=None)

    if version_cell == "1.0.0":
        for position in [0, 1, 6]:  # careful if rows in missingNR_df change
            missingNR_df_old_values[position].convert_month_to_numbers()

    if version_cell not in ["1.0.0", "1.0.1", "1.1.0", "1.1.1"]:
        adjust_classified_info(df_old, df_old_wv, old_wb_obj_values)

    df_old_tomerge = apply_changes_NR(df_old_wv, version_cell, version_cell_new)
    df_old_tomerge = clean_NR_with_no_data(df_old_tomerge)
    if version_cell in ["1.0.0", "1.0.1"]:
        list_migrationissues.extend(change_wastes(df_old_tomerge, mapping_wastes))

    df_new_wv = df_new.merge(df_old_tomerge)
    missingNR_df_new_wv = pd.concat([missingNR_df_new, missingNR_df_old_values], axis=1)

    if version_cell in ["1.0.0", "1.0.1"]:
        adjust_data_missing_first2versions(df_new_wv, missingNR_df_new_wv, old_wb_obj)

    paste_values(new_wb_empty, missingNR_df_new_wv)
    paste_values(
        new_wb_empty,
        df_new_wv,
        NR=True,
        table_of_contents=table_of_contents,
        version=version_cell,
    )

    elapsed = time.time() - start_time
    return new_wb_empty, elapsed, list_migrationissues
