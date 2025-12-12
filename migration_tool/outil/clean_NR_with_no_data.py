import pandas as pd

def clean_NR_with_no_data(df):

    list_of_keywords = ["template_", "enum_", "Table", "BreakdownOfEnergyConsumptionAxis", "Hypercube"]
    
    list_of_NRs_toNOTremove = ["template_reporting_entity_name","template_reporting_entity_identifier_scheme","template_reporting_entity_identifier","template_currency"]
    df_toattach = df[df["name_ranges"].isin(list_of_NRs_toNOTremove)]

    for kw in list_of_keywords:
        df = df[df["name_ranges"].str.contains(kw) == False]

    return pd.concat([df.reset_index(drop=True), df_toattach]).reset_index(drop=True)

