
import pandas as pd
from utils import find_cut_off_year


def payorXincomeXtypeXrevXhalf(df: pd.DataFrame):
    third_parties = df.groupby(['Third_Party_9LC'])
    outputs = []
    cut_off_year = find_cut_off_year(df)

    for third_party, third_party_df in third_parties:
        third_party_output = []

        income_types = third_party_df.groupby(['Normalized_Income_Type_9LC'])
        for income_type, income_df in income_types:
            income_output = {}

            halfs = income_df.groupby(['Year_Statement_9LC', 'Half_Statement_9LC'])
            for (year, half), half_df in halfs:
                if int(year) < cut_off_year:
                    continue
                half_royalty = half_df['Royalty_Payable_SB'].sum()
                income_output[year + ' ' + half] = [round(half_royalty, 2)]

            if not income_output.items():
                # income_output is an empty dictionary.
                continue

            income_output['Total'] = round(sum(i[0] for i in income_output.values()), 2)
            income_output['Third Party'] = third_party
            income_output['Income Type'] = income_type
            income_output = pd.DataFrame.from_dict(income_output)

            third_party_output.append(income_output)

        third_party_output = pd.concat(third_party_output, ignore_index=True).sort_values(by='Total', ascending=False).reset_index(drop=True)
        third_party_output['% Of Revenue'] = (100 * (third_party_output['Total'] / third_party_output['Total'].sum()))
        third_party_output['Cumulative %'] = third_party_output['% Of Revenue'].cumsum().iloc[::-1]

        third_party_output['% Of Revenue'] = third_party_output['% Of Revenue'].round(2).astype(str) + '%'
        third_party_output['Cumulative %'] = third_party_output['Cumulative %'].round(2).astype(str) + '%'

        cols = third_party_output.columns.values.tolist()
        cols_wo_thirdparty_n_source = [i for i in cols if i not in ['Third Party', 'Income Type']]
        third_party_output = third_party_output[['Third Party', 'Income Type'] + cols_wo_thirdparty_n_source]

        cols = third_party_output.columns.values.tolist()
        year_cols = [i for i in cols if i not in ['% Of Revenue', 'Cumulative %', 'Total', 'Third Party', 'Income Type']]
        sorted_year_cols = sorted(year_cols, key=lambda x: int(x.split()[0]) + int(x.split()[1][-1]) * 0.1)

        if year_cols != sorted_year_cols:
            third_party_output = third_party_output[['Third Party', 'Income Type'] +
                                                    sorted_year_cols +
                                                    ['% Of Revenue', 'Cumulative %']]

        outputs.append(third_party_output)

    # i = 0
    # for output in outputs:
    #     output.to_csv(f'{i}.csv')
    #     i += 1

    return outputs


if __name__ == "__main__":
    master = pd.read_parquet('master.parquet.gzip')
    print(master.columns.values.tolist())
    payorXincomeXtypeXrevXhalf(master)