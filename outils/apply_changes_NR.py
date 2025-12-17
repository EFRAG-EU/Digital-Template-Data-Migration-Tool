import pandas as pd

from outils.classes import values


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


def change_wastes(df, mapping):
    old_wastes = (
        df.loc[df["name_ranges"] == "TypeOfWasteAxis", "cell_values"]
        .values[0]
        .first_element_row(double_list=False)
    )
    new_wastes = []
    list_wasteissues = []
    for i in old_wastes:
        if i is not None:
            if pd.isna(mapping.loc[mapping["old"] == i, "new"].values[0]):
                new_wastes.append(
                    [
                        f"Waste category {i} not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955"
                    ]
                )
                list_wasteissues.append(
                    [
                        f"Waste category --{i}-- not present in new Regulation. Please, see https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32014D0955"
                    ]
                )
            else:
                new_wastes.append([mapping.loc[mapping["old"] == i, "new"].values[0]])
    df.loc[df["name_ranges"] == "TypeOfWasteAxis", "cell_values"] = values(new_wastes)

    return list_wasteissues
