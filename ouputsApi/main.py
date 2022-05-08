import pandas as pd

from ouputsApi.utils import find_cut_off_year

# 4
def SongxIncomexRevxHalf(parquet_file):
    # Years = ["2016 H1","2016 H2","2017 H1","2017 H2","2018 H1","2018 H2","2019 H1","2019 H2","2020 H1","2020 H2","2021 H1","2021 H2"]
    # 'make Dataframe and take specific columns'
    newDataFrame = parquet_file[
        ["Statement_Period_Half_9LC", "Song_Name_9LC", "Normalized_Income_Type_9LC", "Royalty_Payable_SB"]]

    # 'groupby Song Name and make sum of Royalties round the sum to ^.^^ and reset index'
    newDataFrame = newDataFrame.groupby(
        ["Song_Name_9LC", "Normalized_Income_Type_9LC", "Statement_Period_Half_9LC"]).sum().round(2)
    newDataFrame = newDataFrame.reset_index()

    # 'pivot table by index,column and value'
    newDataFrame = \
    newDataFrame.pivot(index=["Song_Name_9LC", "Normalized_Income_Type_9LC"], columns=["Statement_Period_Half_9LC"],
                       values=["Royalty_Payable_SB"])["Royalty_Payable_SB"]

    # 'column list and rest index and column name'
    Years = list(newDataFrame.keys())[8:]
    newDataFrame = newDataFrame[Years].reset_index()
    newDataFrame.columns.name = None

    newDataFrame['Total'] = newDataFrame.sum(axis=1, numeric_only=True)

    newDataFrame = newDataFrame.groupby(["Song_Name_9LC"])
    newDataFrameAry = [newDataFrame.get_group(x) for x in newDataFrame.groups]
    newDataFrameAryfinished = []
    for newDataFrame in newDataFrameAry:

        # 'add finel Line(sum of all column)'
        df = {}
        for year in Years:
            df[year] = newDataFrame[year].sum().round(2)

        ################################################################################
        # 'add Total title and total value(sum of all row)'##############################
        Total = sum(df.values()).round(2)
        # 'add % Of Revenue sum/total *100 and round to 2'
        newDataFrame['% Of Revenue'] = round((newDataFrame["Total"] / Total) * 100, 2)
        # 'add Cumulative % sum/total *100 and round to 2'
        newDataFrame['Cumulative %'] = newDataFrame['% Of Revenue'].cumsum().round(2)
        df["Total"] = Total
        df["Normalized_Income_Type_9LC"] = "Total"
        df = pd.DataFrame(df, index=[0])
        newDataFrame = pd.concat([newDataFrame, df], ignore_index=True, axis=0)
        # newDataFrame = pd.concat([pd.DataFrame({"Song_Name_9LC" : ""},index=[0]), newDataFrame], ignore_index=True, axis=0)
        newDataFrameAryfinished.append(newDataFrame)
        ################################################################################
    return newDataFrameAryfinished


# 1,2,3
def SimpleExtract(TheColumn, parquet_file):
    # Years = ["2016 H1","2016 H2","2017 H1","2017 H2","2018 H1","2018 H2","2019 H1","2019 H2","2020 H1","2020 H2","2021 H1","2021 H2"]

    # 'make Dataframe and take specific columns'
    newDataFrame = parquet_file[["Statement_Period_Half_9LC", TheColumn, "Royalty_Payable_SB"]]

    # 'groupby Song Name and make sum of Royalties round the sum to ^.^^ and reset index'
    newDataFrame = newDataFrame.groupby([TheColumn, "Statement_Period_Half_9LC"]).sum().round(2)
    newDataFrame = newDataFrame.reset_index()

    # 'pivot table by index,column and value'
    newDataFrame = \
    newDataFrame.pivot(index=[TheColumn], columns=["Statement_Period_Half_9LC"], values=["Royalty_Payable_SB"])[
        "Royalty_Payable_SB"]

    # 'column list and rest index and column name'
    Years = list(newDataFrame.keys())[8:]
    newDataFrame = newDataFrame[Years].reset_index()
    newDataFrame.columns.name = None

    # 'Total column is the sum of all columns except song title'
    newDataFrame['Total'] = newDataFrame.sum(axis=1, numeric_only=True).round(2)

    # 'add finel Line(sum of all column)'
    df = {}
    for year in Years:
        df[year] = newDataFrame[year].sum().round(2)

    ################################################################################
    # 'add Total title and total value(sum of all row)'##############################
    Total = sum(df.values()).round(2)
    # 'add % Of Revenue sum/total *100 and round to 2'
    newDataFrame['% Of Revenue'] = round((newDataFrame["Total"] / Total) * 100, 2)
    # 'add Cumulative % sum/total *100 and round to 2'
    newDataFrame['Cumulative %'] = newDataFrame['% Of Revenue'].cumsum().round(2)
    df["Total"] = Total
    df[TheColumn] = "Total"
    df = pd.DataFrame(df, index=[0])
    newDataFrame = pd.concat([newDataFrame, df], ignore_index=True, axis=0)
    df = pd.DataFrame({TheColumn: "(" + ",".join(
        parquet_file[["Third_Party_9LC"]].drop_duplicates()["Third_Party_9LC"].to_list()) + ")"}, index=[0])
    newDataFrame = pd.concat([df, newDataFrame], ignore_index=True, axis=0)
    ################################################################################
    return newDataFrame


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

    output = pd.concat(output, ignore_index=True).sort_values(by='Total', ascending=False)
    output.reset_index(drop=True, inplace=True)
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

    totals_row = output.sum(axis=0, numeric_only=True).round(2)
    totals_row['Release Artist'] = 'Total'
    output.loc[len(output)] = totals_row

    return output


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

        third_party_output = pd.concat(third_party_output, ignore_index=True).sort_values(by='Total', ascending=False)
        third_party_output.reset_index(drop=True, inplace=True)
        third_party_output['% Of Revenue'] = (100 * (third_party_output['Total'] / third_party_output['Total'].sum()))
        third_party_output['Cumulative %'] = third_party_output['% Of Revenue'].cumsum().iloc[::-1]

        third_party_output['% Of Revenue'] = third_party_output['% Of Revenue'].round(2).astype(str) + '%'
        third_party_output['Cumulative %'] = third_party_output['Cumulative %'].round(2).astype(str) + '%'

        cols = third_party_output.columns.values.tolist()
        cols_wo_thirdparty_n_source = [i for i in cols if i not in ['Third Party', 'Income Type']]
        third_party_output = third_party_output[['Third Party', 'Income Type'] + cols_wo_thirdparty_n_source]

        cols = third_party_output.columns.values.tolist()
        year_cols = [i for i in cols if
                     i not in ['% Of Revenue', 'Cumulative %', 'Total', 'Third Party', 'Income Type']]
        sorted_year_cols = sorted(year_cols, key=lambda x: int(x.split()[0]) + int(x.split()[1][-1]) * 0.1)

        if year_cols != sorted_year_cols:
            third_party_output = third_party_output[['Third Party', 'Income Type'] +
                                                    sorted_year_cols +
                                                    ['% Of Revenue', 'Cumulative %']]

        totals_row = third_party_output.sum(axis=0, numeric_only=True).round(2)
        totals_row['Income Type'] = 'Total'
        third_party_output.loc[len(third_party_output)] = totals_row

        outputs.append(third_party_output)

    # i = 0
    # for output in outputs:
    #     output.to_csv(f'{i}.csv')
    #     i += 1

    return outputs


def payorXsongXrevXhalf(df: pd.DataFrame):
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

        third_party_output = pd.concat(third_party_output, ignore_index=True).sort_values(by='Total', ascending=False)
        third_party_output.reset_index(drop=True, inplace=True)
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

        totals_row = third_party_output.sum(axis=0, numeric_only=True).round(2)
        totals_row['Song Title'] = 'Total'
        third_party_output.loc[len(third_party_output)] = totals_row

        outputs.append(third_party_output)

    # i = 0
    # for output in outputs:
    #     output.to_csv(f'{i}.csv')
    #     i += 1

    return outputs


def payorXsourceXrevXhalf(df: pd.DataFrame):
    third_parties = df.groupby(['Third_Party_9LC'])
    outputs = []
    cut_off_year = find_cut_off_year(df)

    for third_party, third_party_df in third_parties:
        third_party_output = []

        sources = third_party_df.groupby(['Source_SB'])
        for source, source_title_df in sources:
            source_output = {}

            halfs = source_title_df.groupby(['Year_Statement_9LC', 'Half_Statement_9LC'])
            halfs = sorted(halfs, key=lambda x: x[0][0])

            for (year, half), half_df in halfs:
                if int(year) < cut_off_year:
                    continue

                half_royalty = half_df['Royalty_Payable_SB'].sum()
                source_output[year + ' ' + half] = [round(half_royalty, 2)]

            if not source_output.items():
                # source_output is an emtpy dictionary.
                continue

            source_output['Total'] = round(sum(i[0] for i in source_output.values()), 2)
            source_output['Third Party'] = third_party
            source_output['Source'] = source
            source_output = pd.DataFrame.from_dict(source_output)

            # print(source_output.columns) .

            third_party_output.append(source_output)
        if len(third_party_output) > 0:
            third_party_output = pd.concat(third_party_output, ignore_index=True).sort_values(by='Total', ascending=False)
            third_party_output.reset_index(drop=True, inplace=True)
            third_party_output['% Of Revenue'] = (100 * (third_party_output['Total'] / third_party_output['Total'].sum()))
            third_party_output['Cumulative %'] = third_party_output['% Of Revenue'].cumsum().iloc[::-1]
    
            third_party_output['% Of Revenue'] = third_party_output['% Of Revenue'].round(2).astype(str) + '%'
            third_party_output['Cumulative %'] = third_party_output['Cumulative %'].round(2).astype(str) + '%'
            cols = third_party_output.columns.values.tolist()
            cols_wo_thirdparty_n_source = [i for i in cols if i not in ['Third Party', 'Source']]
            third_party_output = third_party_output[['Third Party', 'Source'] + cols_wo_thirdparty_n_source]
    
            cols = third_party_output.columns.values.tolist()
            year_cols = [i for i in cols if i not in ['% Of Revenue', 'Cumulative %', 'Total', 'Third Party', 'Source']]
            sorted_year_cols = sorted(year_cols, key=lambda x: int(x.split()[0]) + int(x.split()[1][-1]) * 0.1)
    
            if year_cols != sorted_year_cols:
                third_party_output = third_party_output[
                    ['Third Party', 'Source'] + sorted_year_cols + ['% Of Revenue', 'Cumulative %']]
    
            totals_row = third_party_output.sum(axis=0, numeric_only=True).round(2)
            totals_row['Source'] = 'Total'
            third_party_output.loc[len(third_party_output)] = totals_row
    
            outputs.append(third_party_output)

    return outputs


def defualtDetails(parquet_file):
    df = parquet_file[
        ["Payout_Currency_9LC", "Third_Party_9LC", "Contract_ID_9LC", "Royalty_Payable_SB","Adjusted_Royalty_SB", "Adjusted_Royalty_USD_SB"]]
    df = df.groupby(['Payout_Currency_9LC', 'Third_Party_9LC', "Contract_ID_9LC"])
    lineCount = df.size()
    df = df.sum()
    df["Line Count"] = lineCount.tolist()
    df = df.reset_index()
    currency = df["Payout_Currency_9LC"].drop_duplicates()[0]
    df.rename(columns={"Payout_Currency_9LC": "Currency", "Third_Party_9LC": "Party","Contract_ID_9LC" : "Contract","Adjusted_Royalty_SB" : "Adjusted("+currency+")","Adjusted_Royalty_USD_SB" : "Adjusted($)","Royalty_Payable_SB" : "Nominal"}, inplace=True)
    #print(df)
    df["Party/Contract"] = df["Party"]+"/"+df["Contract"]
    df = df[["Party/Contract","Line Count","Currency","Nominal","Adjusted($)","Adjusted("+currency+")"]]
    return df


    #parquet_file = pd.read_parquet("databases/61f04127fead509ee33d2280/master.parquet.gzip", engine='pyarrow')
#print(defualtDetails(parquet_file))
