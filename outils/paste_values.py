import pandas as pd
from openpyxl.utils.cell import range_boundaries


def paste_values(pyxl, df, NR=None):
    """ Paste values from a DataFrame column into an openpyxl workbook"""
    
    sheet_names = pyxl.sheetnames
    merged_loc = pd.DataFrame()

    for sheet in sheet_names:
        merged = list(pyxl[sheet].merged_cells.ranges)
        merged_list = []
        for i in range(len(merged)):
            str(merged[i])
            merged_list.append(range_boundaries(str(merged[i]))[:2])
        merged_loc = pd.concat([merged_loc, pd.DataFrame({sheet: merged_list})], axis=1)

    add_checkbox = ["SiteLocatedInABiodiversitySensitiveArea","SiteLocatedNearABiodiversitySensitiveArea"]
    # yellow_checkbox = ["CountryOfEmploymentContractAxis"].values
    # if NR is "CountryOfEmploymentContractAxis":
    for i in range(len(df)):

        sheet = pyxl[df["sheets"][i]]
        rng = sheet[df["cell_ranges"][i]]
        shape = df["cell_shapes"][i]
        value = df["cell_values"][i]


        if shape.isonecell():
            rng.value = value.topleft() # when it's one cell, paste topleft 
            
        else:

            tuple_to_check = (shape.left(), shape.top()) 

            if value.values() is not None: # avoiding formulas

                if tuple_to_check in merged_loc[df["sheets"][i]].values.tolist(): # merged cells check

                    value.first_element_row() # keeping only the first value of each row
                    value.enlarged_range_correction(shape) # enlarged ranges check
                    value.paste(sheet, shape)
                                   
                else:

                    value.enlarged_range_correction(shape) # enlarged ranges check
                    if NR is not None:
                        if df["name_ranges"][i] in add_checkbox:
                            value.add_checkboxes() # add checkboxes for the specific NRs (see above)
                    value.paste(sheet, shape)