import pandas as pd

#4
def SongxIncomexRevxHalf(parquet_file):
    # Years = ["2016 H1","2016 H2","2017 H1","2017 H2","2018 H1","2018 H2","2019 H1","2019 H2","2020 H1","2020 H2","2021 H1","2021 H2"]
    # 'make Dataframe and take specific columns'
    newDataFrame = parquet_file[["Statement_Period_Half_9LC","Song_Name_9LC","Normalized_Income_Type_9LC","Royalty_Payable_SB"]]

    # 'groupby Song Name and make sum of Royalties round the sum to ^.^^ and reset index'
    newDataFrame = newDataFrame.groupby(["Song_Name_9LC","Normalized_Income_Type_9LC","Statement_Period_Half_9LC"]).sum().round(2)
    newDataFrame = newDataFrame.reset_index()

    # 'pivot table by index,column and value'
    newDataFrame = newDataFrame.pivot(index=["Song_Name_9LC","Normalized_Income_Type_9LC"],columns=["Statement_Period_Half_9LC"],values=["Royalty_Payable_SB"])["Royalty_Payable_SB"]

    # 'column list and rest index and column name'
    Years = list(newDataFrame.keys())[8:]
    newDataFrame = newDataFrame[Years].reset_index()
    newDataFrame.columns.name = None

    newDataFrame['Total'] = newDataFrame.sum(axis=1,numeric_only=True)

    newDataFrame =  newDataFrame.groupby(["Song_Name_9LC"])
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
        newDataFrame['% Of Revenue'] = round((newDataFrame["Total"] / Total) * 100,2)
        # 'add Cumulative % sum/total *100 and round to 2'
        newDataFrame['Cumulative %'] = newDataFrame['% Of Revenue'].cumsum().round(2)
        df["Total"] = Total
        df["Normalized_Income_Type_9LC"] = "Total"
        df = pd.DataFrame(df,index=[0])
        newDataFrame = pd.concat([newDataFrame, df], ignore_index = True, axis = 0)
        newDataFrame = pd.concat([pd.DataFrame({"Song_Name_9LC" : ""},index=[0]), newDataFrame], ignore_index=True, axis=0)
        newDataFrameAryfinished.append(newDataFrame)
        ################################################################################
    return newDataFrameAryfinished

# 1,2,3
def SimpleExtract(TheColumn,parquet_file):
    # Years = ["2016 H1","2016 H2","2017 H1","2017 H2","2018 H1","2018 H2","2019 H1","2019 H2","2020 H1","2020 H2","2021 H1","2021 H2"]


    # 'make Dataframe and take specific columns'
    newDataFrame = parquet_file[["Statement_Period_Half_9LC",TheColumn,"Royalty_Payable_SB"]]

    # 'groupby Song Name and make sum of Royalties round the sum to ^.^^ and reset index'
    newDataFrame = newDataFrame.groupby([TheColumn,"Statement_Period_Half_9LC"]).sum().round(2)
    newDataFrame = newDataFrame.reset_index()

    # 'pivot table by index,column and value'
    newDataFrame = newDataFrame.pivot(index=[TheColumn],columns=["Statement_Period_Half_9LC"],values=["Royalty_Payable_SB"])["Royalty_Payable_SB"]

    # 'column list and rest index and column name'
    Years = list(newDataFrame.keys())[8:]
    newDataFrame = newDataFrame[Years].reset_index()
    newDataFrame.columns.name = None

    # 'Total column is the sum of all columns except song title'
    newDataFrame['Total'] = newDataFrame.sum(axis=1,numeric_only=True).round(2)

    # 'add finel Line(sum of all column)'
    df = {}
    for year in Years:
        df[year] = newDataFrame[year].sum().round(2)

    ################################################################################
    # 'add Total title and total value(sum of all row)'##############################
    Total = sum(df.values()).round(2)
    # 'add % Of Revenue sum/total *100 and round to 2'
    newDataFrame['% Of Revenue'] = round((newDataFrame["Total"] / Total) * 100,2)
    # 'add Cumulative % sum/total *100 and round to 2'
    newDataFrame['Cumulative %'] = newDataFrame['% Of Revenue'].cumsum().round(2)
    df["Total"] = Total
    df[TheColumn] = "Total"
    df = pd.DataFrame(df,index=[0])
    newDataFrame = pd.concat([newDataFrame, df], ignore_index = True, axis = 0)
    df = pd.DataFrame({TheColumn : "("+",".join(parquet_file[["Third_Party_9LC"]].drop_duplicates()["Third_Party_9LC"].to_list())+")" },index=[0])
    newDataFrame = pd.concat([df,newDataFrame ], ignore_index = True, axis = 0)
    ################################################################################
    return newDataFrame


'''
while True:
    number = int(input("1 - SongxRevxHal f\n"
                       "2 - Income x Rev x Half \n"
                       "3 - Source x Rev x Half \n"
                       "Enter: "))
    if number == 1:
        newDataFrame = SimpleExtract("Song_Name_9LC")
        print(newDataFrame)
    elif number == 2:
        newDataFrame = SimpleExtract("Normalized_Income_Type_9LC")
        print(newDataFrame)
    elif number == 3:
        newDataFrame = SimpleExtract("Normalized_Source_9LC")
        print(newDataFrame)
    elif number == 4:
    ######################################################################
    ######################################################################
    ######################################################################
    #                        Total need fix                              #
    ######################################################################
    ######################################################################
    ######################################################################
        newDataFrame = SongxIncomexRevxHalf()
        print(newDataFrame)
        
    #MakeOut(newDataFrame)
'''