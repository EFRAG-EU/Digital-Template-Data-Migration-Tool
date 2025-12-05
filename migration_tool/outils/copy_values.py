from outils.classes import values

def copy_values(pyxl, df, key=None):
    """ Copy values from old workbook to new workbook based on name ranges and cell shapes, and return a df with key and cell values"""

    cell_values = []

    for i in range(len(df)):

        sheet = pyxl[df["sheets"][i]]
        shape = df["cell_shapes"][i]
        rng = df["cell_ranges"][i]

        if rng is not None:
            cell_values.append(values(shape.build_values(sheet)))
        else:
            cell_values.append(values(None))

    df["cell_values"] = cell_values

    if key is not None:
        return df[[key, "cell_values"]]
    else:
        return df["cell_values"]