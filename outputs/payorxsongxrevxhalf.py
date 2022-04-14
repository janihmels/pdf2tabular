
import pandas as pd
from utils import find_cut_off_year


def payorxsongxrevxhalf(df: pd.DataFrame):

    third_parties = df.groupby(['Third_Party_9LC'])
    outputs = []
    cut_off_year = find_cut_off_year(df)

    for third_party, third_party_df in third_parties:
        third_party_output = []

        songs = third_party_df.groupby(['Song_Name_9LC'])
        for song_title, song_title_df in songs:
            song_output = {}

            halfs = song_title_df.groupby(['Year_Statement_9LC', 'Half_Statement_9LC'])
            for (year, half), half_df in halfs:
                if int(year) < cut_off_year:
                    continue

                half_royalty = half_df['Royalty_Payable_SB'].sum()
                song_output[year + ' ' + half] = [round(half_royalty, 2)]

            if not song_output.items():
                # song_output is an empty dictionary.
                continue

            song_output['Total'] = round(sum(i[0] for i in song_output.values()), 2)
            song_output['Third Party'] = third_party
            song_output['Song Title'] = song_title
            song_output = pd.DataFrame.from_dict(song_output)

            third_party_output.append(song_output)

        third_party_output = pd.concat(third_party_output, ignore_index=True).sort_values(by='Total', ascending=False).reset_index(drop=True)
        third_party_output['% Of Revenue'] = (100 * (third_party_output['Total'] / third_party_output['Total'].sum()))
        third_party_output['Cumulative %'] = third_party_output['% Of Revenue'].cumsum().iloc[::-1]

        third_party_output['% Of Revenue'] = third_party_output['% Of Revenue'].round(2).astype(str) + '%'
        third_party_output['Cumulative %'] = third_party_output['Cumulative %'].round(2).astype(str) + '%'

        cols = third_party_output.columns.values.tolist()
        cols_wo_thirdparty_n_source = [i for i in cols if i not in ['Third Party', 'Song Title']]
        third_party_output = third_party_output[['Third Party', 'Song Title'] + cols_wo_thirdparty_n_source]

        cols = third_party_output.columns.values.tolist()
        year_cols = [i for i in cols if i not in ['% Of Revenue', 'Cumulative %', 'Total', 'Third Party', 'Song Title']]
        sorted_year_cols = sorted(year_cols, key=lambda x: int(x.split()[0]) + int(x.split()[1][-1]) * 0.1)

        if year_cols != sorted_year_cols:
            third_party_output = third_party_output[['Third Party', 'Song Title'] +
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
    payorxsongxrevxhalf(master)

