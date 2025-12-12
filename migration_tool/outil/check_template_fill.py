def check_template_fill(old_wb_values):

    list_of_incompletes = ["INCOMPLETE", "INCOMPLETE", "NEBAIGTA", "NIEKOMPLETNY", "INCOMPLETO", "INCOMPLETO"]

    value = old_wb_values["Table of Contents & Validation"].cell(row=3, column=3).value

    if value in list_of_incompletes:
        return False
    return True
