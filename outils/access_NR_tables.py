import pandas as pd
from openpyxl.utils import range_boundaries

from outils.classes import shapes

def access_NR_table(pyxl_NR):
    """ Access names, cell references and coordinates of each name range from the python object containing name ranges"""

    list_NR = list(pyxl_NR.keys())
    list_sheets = []
    list_ranges = []
    list_shapes = []

    for NR in list_NR:
        list_sheets.append(pyxl_NR[NR].attr_text.split("!")[0].replace("'", ""))
        temp = pyxl_NR[NR].attr_text.split("!")[1]
        if temp == "#REF": # handling #REF name ranges (in Translations sheet)
            temp = None
        list_ranges.append(temp)

    for rng in list_ranges:
        if rng is None:
            list_shapes.append(shapes(None))
        else:
            list_shapes.append(shapes(range_boundaries(rng)))

    return pd.DataFrame({
        "name_ranges": list_NR,
        "sheets": list_sheets,
        "cell_ranges": list_ranges,
        "cell_shapes": list_shapes
    })

def access_missingNR_table(df_missingNR, version_cell):
    """ Access sheets, cell references and coordinates of each missing name range"""

    list_of_sheets = []
    list_of_ranges = []

    for i in range(len(df_missingNR)):
        list_of_sheets.append(df_missingNR.loc[:,version_cell].values[i].split("!")[0])
        list_of_ranges.append(df_missingNR.loc[:,version_cell].values[i].split("!")[1])

    for rang in list_of_ranges:
        if rang == "None":
            list_of_ranges[list_of_ranges.index(rang)] = None

    list_of_shapes = [shapes(range_boundaries(rng)) if rng is not None else shapes(None) for rng in list_of_ranges]

    return pd.DataFrame({
        "sheets": list_of_sheets,
        "cell_ranges": list_of_ranges,
        "cell_shapes": list_of_shapes
    }) 

