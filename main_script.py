from openpyxl import load_workbook
import numpy as np
import pandas as pd
from tkinter import Tk, filedialog
import warnings
import time
import sys

from outils.check_template_fill import check_template_fill
from outils.apply_changes_NR import apply_changes_NR, change_wastes
from outils.clean_NR_with_no_data import clean_NR_with_no_data
from outils.paste_values import paste_values
from outils.access_NR_tables import access_NR_table, access_missingNR_table
from outils.copy_values import copy_values
from outils.classes import values
from pickles.missingNR_df import missingNR_df
mapping_wastes = pd.read_pickle("pickles/Mapping_wastes.pkl")

root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Select Your Old Template Version (with data, .xlsx format)", filetypes=[("Excel Files", "*.xlsx")])
if file_path:
    print(f"File selected: {file_path}")
else:
    print("No file selected.")
    sys.exit()

with warnings.catch_warnings():
    warnings.simplefilter("ignore", category=UserWarning)
    old_wb = load_workbook(file_path)
    old_wb_values = load_workbook(file_path, data_only=True)

start_time = time.time()

if not check_template_fill(old_wb_values):
    print("!Warning!: The selected file is not a FILLED vsme digital template.")
    sys.exit()

with warnings.catch_warnings():
    warnings.simplefilter("ignore", category=UserWarning)
    new_wb_empty = load_workbook("Template/VSME-Digital-Template-1.1.1.xlsx")

version_cell = old_wb["Introduction"].cell(row=1, column=3).value
version_cell_new = new_wb_empty["Introduction"].cell(row=1, column=3).value

df_old = access_NR_table(old_wb.defined_names)
df_new = access_NR_table(new_wb_empty.defined_names)

missingNR_df_old = access_missingNR_table(missingNR_df, version_cell)
missingNR_df_new = access_missingNR_table(missingNR_df, version_cell_new)

df_old_wv = copy_values(old_wb, df_old, key="name_ranges")
# df_old_wv.to_pickle("pickles/df_old_wv.pkl") 
missingNR_df_old_values = copy_values(old_wb, missingNR_df_old, key=None)

if(version_cell == "1.0.0"): # (only in version 1.0.0)
    for position in [0,1,6]: # (see missingNR_df)
        missingNR_df_old_values[position].convert_month_to_numbers()
# missingNR_df_old_values.to_pickle("pickles/missingNR_df_old_values.pkl")

df_old_tomerge = apply_changes_NR(df_old_wv, version_cell, version_cell_new) # apply migration NR changes
df_old_tomerge = clean_NR_with_no_data(df_old_tomerge) # clean NRs which have no data to transfer
if version_cell in ["1.0.0", "1.0.1", "1.1.0"]: # (only in the 3 oldest versions)
    change_wastes(df_old_tomerge, mapping_wastes) # waste categories update

df_new_wv = df_new.merge(df_old_tomerge)
# df_new_wv.to_pickle("pickles/df_new_wv.pkl")
missingNR_df_new_wv = pd.concat([missingNR_df_new, missingNR_df_old_values], axis=1)
# missingNR_df_new_wv.to_pickle("pickles/missingNR_df_new_wv.pkl")

if version_cell in ["1.0.0", "1.0.1"]: # (only in the 2 oldest versions, add the checkbox value based on adjacent NR)
    len = df_new_wv.loc[df_new_wv["name_ranges"] == "CountryOfEmploymentContractAxis", "cell_values"].values[0].count_uniques()
    if len > 2:                 # '> 2' because None is one unique
        missingNR_df_new_wv.loc[(missingNR_df_new_wv["sheets"]== "Social Disclosures") & (missingNR_df_new_wv["cell_ranges"]== "$E$27"), "cell_values"] = values([[True]])
    else:
        missingNR_df_new_wv.loc[(missingNR_df_new_wv["sheets"]== "Social Disclosures") & (missingNR_df_new_wv["cell_ranges"]== "$E$27"), "cell_values"] = values([[False]])

paste_values(new_wb_empty, missingNR_df_new_wv)
paste_values(new_wb_empty, df_new_wv, NR=True)

new_wb_empty.save(f"Template/result_from_{version_cell}.xlsx")

for i in range(1000000):
    pass
end_time = time.time()
print(f"Execution time: {end_time - start_time:.4f} seconds")