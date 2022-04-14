import pandas as pd
import MySQLdb
import random

from Levenshtein import distance as levenshtein_distance

DATA_FILE = './data-cleaned-v3.csv'


def extract_data(dbname, query):

    def cmd2df(cursor, query):
        """
        :param cursor: A cursor object from the MySQLdb package, which gives us access to a database.
        :param query: A string that contains a query to perform on the database.
        :return: The table returned from the query, as a pandas.DataFrame.
        """

        def get_col_names():
            """
            :return: A list of the columns names of the table extracted by the given query.
            """

            return [i[0] for i in cursor.description]

        cursor.execute(query)
        data = cursor.fetchall()  # fetching the data
        df = pd.DataFrame.from_records(data=data, columns=get_col_names())

        return df

    db_connection = MySQLdb.connect('34.65.111.142',
                                    'external',
                                    'musicpass',
                                    dbname)

    # -- connected --
    print(f"connected to the db '{dbname}'")

    cursor = db_connection.cursor()

    return cmd2df(cursor, query)


def find_similar_normalized_sources(data):
    normalized_sources = data['Normalized_Source'].drop_duplicates()

    cnt = 0
    for i in range(len(normalized_sources)):
        for j in range(i+1, len(normalized_sources)):
            a = normalized_sources.iloc[i]
            b = normalized_sources.iloc[j]
            if a != b and distance(a, b) <= 0.22:
                print(a, "||", b)
                cnt += 1

    print(f"Found {cnt} similar pairs")


def extract_data_from_db():
    dbnames = ['Tempo - Lukas Graham (New)_6212643631bc177582886e3a',
               'Bob Morrison (New 2022)_62103829659e5fd3d84d20f8',
               'Songvest - Ketih Thomas_62103671659e5fd3d84d20f7',
               'Songvest - Zairyus Jackson_620e5d0ad6d0bedf062aa551',
               'Songvest - Christoffer - Final_620ab81050477b98465c78d4',
               'Songvest - Gianni (New)_6203ef7fd5c5799999e9a38d',
               'Shawty Redd (New)_61ff3cd2c98a37867c06137b',
               'Amanda Ghost (New)_61fdbe07a8b2e77c67c486df',
               'John Borger (New)_61fb1dbfa8b2e77c67c486dc']

    queries = ['SELECT DISTINCT Source_SB AS Source, Normalized_Source_9LC AS Normalized_Source FROM Master'] * \
              len(dbnames)
    data = pd.concat([extract_data(dbname=dbname, query=query) for dbname, query in zip(dbnames, queries)], axis=0)

    return data


def get_data():
    """
    :return: Two lists:
            * A list of tuples representing features of song titles.
            * A list of tuple representing features of phrases that aren't song titles.
    """

    data = pd.read_csv(DATA_FILE)

    # def english_rows(df):
    #     res = []
    #     j = 0
    #
    #     for i, row in df.iterrows():
    #         source = row['Source']
    #
    #         if pd.isna(source):
    #             res.append(0)
    #         else:
    #             if not all([ch <= '~' for ch in source]):
    #                 j += 1
    #             res.append(all([ch <= '~' for ch in source]))
    #
    #     return pd.Series(res)

    # -- cleaning nans --
    data = data.loc[pd.notna(data['Source']) & pd.notna(data['Normalized_Source'])]

    # -- removing rows that are duplications of other ones --
    data = data.drop_duplicates()

    return data


def organize_data(data):
    """
    :param data: A pandas.DataFrame containing two columns: 'Source' and 'Normalized Source'.
    :return: * A list of tuples of the form (<list of normalized source names>, their integer label).
             * A dictionary mapping between an integer (label) to the name it represents.
    """

    dfs = [x for _, x in data.groupby('Normalized_Source')]
    data = []
    int_to_name = {}

    label = 0
    for df in dfs:
        data.append((list(df['Source']), label))
        int_to_name[label] = df['Normalized_Source'].iloc[0]

        label += 1

    return data, int_to_name


def augment_name(name):
    def get_random_character():
        return str(random.randint(40, 126))

    policy = random.randint(0, 1)

    if policy == 0:
        # Add sum random character at a random index
        rand_idx = random.randint(0, len(name))
        name = name[:rand_idx] + get_random_character() + name[rand_idx:]

    if policy == 1:
        # Remove some character at a random index
        rand_idx = random.randrange(0, len(name))
        name = name[:rand_idx] + name[rand_idx+1:]

    return name


MAX_LENGTH = 100
def prepro(tokenizer, name):
    """
    :param tokenizer: Some bert model tokenizer.
    :param name: A string.
    :return: A dictionary containing the tokenized data for the string.
    """

    name = tokenizer.encode_plus(name)

    input_ids = name['input_ids']
    attention_masks = name['attention_mask']
    token_type_ids = name['token_type_ids']

    if len(input_ids) < MAX_LENGTH:
        input_ids += [0] * (MAX_LENGTH - len(input_ids))
        attention_masks += [0] * (MAX_LENGTH - len(attention_masks))
        token_type_ids += [0] * (MAX_LENGTH - len(token_type_ids))

    return {'input_ids': input_ids,
            'attention_mask': attention_masks,
            'token_type_ids': token_type_ids}


def distance(str1, str2):
    # return sum([(c in str2) for c in str1]) / max(len(str1), len(str2))
    return levenshtein_distance(str1, str2) / max(len(str1), len(str2))


def test_history_classifier():
    data = get_data()
    train_size = int(len(data) * (9 / 10))
    validation_size = len(data) - train_size

    train, validation = data[:train_size], data[train_size:]

    mapping = {row['Source']: row['Normalized_Source'] for i, row in train.iterrows()}

    accuracy = 0

    for i, row in validation.iterrows():
        source = row['Source']
        normalized_source = row['Normalized_Source']

        closest_source = min([(k, distance(source, k)) for k in mapping.keys()], key=lambda x: x[1])[0]
        # print(source, '|||', closest_source)
        prediction = mapping[closest_source]

        accuracy += (prediction == normalized_source)
        loss = distance(prediction, normalized_source)

    accuracy /= validation_size
    loss /= validation_size

    print(f"Validation Loss: {loss}")
    print(f"Validation Accuracy: {accuracy}")


if __name__ == "__main__":
    # test_history_classifier()
    data = get_data()
    final_data = organize_data(data)
