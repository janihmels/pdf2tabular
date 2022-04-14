
import pandas as pd
from utils import find_cut_off_year


def artistxrevxhalf(df: pd.DataFrame):

    artists = df.groupby(['Release_Artist_9LC'])
    output = []
    cut_off_year = find_cut_off_year(df)

    for release_artist, artist_df in artists:
        artist_output = {}

        halfs = artist_df.groupby(['Year_Statement_9LC', 'Half_Statement_9LC'])
        for (year, half), half_df in halfs:
            if int(year) < cut_off_year:
                continue

            half_royalty = half_df['Royalty_Payable_SB'].sum()
            artist_output[year + ' ' + half] = [round(half_royalty, 2)]

        if not artist_output.items():
            # artist_output is an empty dictionary.
            continue

        artist_output['Total'] = round(sum(i[0] for i in artist_output.values()), 2)
        artist_output['Release Artist'] = release_artist if release_artist.strip() != '' else 'Unknown Artists'
        artist_output = pd.DataFrame.from_dict(artist_output)

        output.append(artist_output)

    output = pd.concat(output, ignore_index=True).sort_values(by='Total', ascending=False).reset_index(drop=True)
    output['% Of Revenue'] = (100 * (output['Total'] / output['Total'].sum()))
    output['Cumulative %'] = output['% Of Revenue'].cumsum().iloc[::-1]

    output['% Of Revenue'] = output['% Of Revenue'].round(2).astype(str) + '%'
    output['Cumulative %'] = output['Cumulative %'].round(2).astype(str) + '%'

    cols = output.columns.values.tolist()
    cols_wo_release_artist = [i for i in cols if i not in ['Release Artist']]
    output = output[['Release Artist'] + cols_wo_release_artist]

    cols = output.columns.values.tolist()
    year_cols = [i for i in cols if i not in ['% Of Revenue', 'Cumulative %', 'Total', 'Release Artist']]
    sorted_year_cols = sorted(year_cols, key=lambda x: int(x.split()[0]) + int(x.split()[1][-1]) * 0.1)

    if year_cols != sorted_year_cols:
        output = output[['Third Party', 'Income Type'] +
                        sorted_year_cols +
                        ['% Of Revenue', 'Cumulative %']]

    # i = 0
    # output.to_csv(f'{i}.csv')

    return output


if __name__ == "__main__":
    master = pd.read_parquet('master.parquet.gzip')
    print(master.columns.values.tolist())
    artistxrevxhalf(master)

