import pandas as pd
import MySQLdb
import os
# import enchant
import matplotlib.pyplot as plt
import itertools

DATA_FILE = './data.csv'
#
# previous_saved_names = ['netflix',
#                         'hbo',
#                         'vh1',
#                         'ADD. DISTRIBUTION',
#                         'ADD DISTRIBUTION',
#                         'ADDITIONAL DISTRIBUTION',
#                         'ADDITIONAL DIST.',
#                         'Settlement',
#                         'REEMISSION',
#                         'Proxy'
#                         'ADD. DISTR.',
#                          'facebook',
#                         'e!',
#                         'youtube',
#                         'dmx',
#                         'ADD.DISTRIBUTION',
#                         'PERFORMANCE']

saved_names = ['Distribution',
               'Dist',
               'Sacem',
               'Socan',
               'Performance',
               'Normalzuschlag',
               'Settlement',
               'Supplement',
               'Youtube',
               'Facebook',
               'Netflix',
               'Residual',
               'HBO',
               'Fox',
               'MTV',
               'MTVJ',
               'MTV2',
               'NBC',
               'Univ',
               'VH1',
               'VH1S',
               'HBOL',
               'Reemission',
               'Sony',
               'unknown']


saved_names += [str(i) for i in range(2000, 2050)]

saved_names = [word.lower() for word in saved_names]


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


def split_to_classes(data_df):
    """
    :param data_df: A pandas.DataFrame containing names and 'not names'.
    :return: A tuple of (list_of_song_names, list_of_not_song_name).
    """

    return list(data_df[data_df['Revenue'] != 'X']['Name']), list(data_df[data_df['Revenue'] == 'X']['Name'])


def get_data(featurize_data=False):
    """
    :param featurize_data: A boolean telling if to featurize the data before passing it or not.
    :return: Two lists:
            * A list of tuples representing features of song titles.
            * A list of tuple representing features of phrases that aren't song titles.
    """

    def featurize(english_dict, phrase):
        """
        :param english_dict: An object from the 'enchant' package.
        :param phrase: A string that may represent a song title, or not a song title.
        :return: A feature vector for the phrase, i.e. a tuple of (# words in phrase that appear at English dictionary,
                                                                   # words in phrase from the list bellow).
        """

        original_words = phrase.split()
        original_words = itertools.chain.from_iterable([word.split('.') for word in original_words])
        original_words = [word for word in original_words if word != '']
        words = [word.lower() for word in original_words if word != '']

        att1 = sum(english_dict.check(word) for word in words if not (word in saved_names)) / len(words)
        # ^ : rate of words in English dictionary but not at the 'saved words' list

        att2 = sum([(word in saved_names) for word in words]) > 0
        # ^ : bool telling if a word from the 'saved words' list appears

        att3 = len(words)
        att4 = sum([(word.isupper() and word.isalpha()) for word in original_words]) / len(original_words)  # rate of words in CAPS
        att5 = sum([1 for i in range(len(phrase)) if phrase[i] == '-'])  # number of times '-' appears at the phrase

        att6 = "'" in phrase  # phrase contains '

        att7 = sum([1 for i in range(len(phrase)) if phrase[i] == '('])  # number of times '(' appears at the phrase
        att8 = sum([1 for i in range(len(phrase)) if phrase[i] == ')'])  # number of times ')' appears at the phrase

        return att1, att2, att3, att4, att5, att6, att7, att8

    if os.path.isfile(DATA_FILE):
        # file exists
        data = pd.read_csv(DATA_FILE)
    else:
        # extract the data from the database
        dbnames = ['Tempo - Lukas Graham (New)_6212643631bc177582886e3a',
                   'Bob Morrison (New 2022)_62103829659e5fd3d84d20f8',
                   'Songvest - Ketih Thomas_62103671659e5fd3d84d20f7',
                   'Songvest - Zairyus Jackson_620e5d0ad6d0bedf062aa551',
                   'Songvest - Christoffer - Final_620ab81050477b98465c78d4',
                   'Songvest - Gianni (New)_6203ef7fd5c5799999e9a38d',
                   'Shawty Redd (New)_61ff3cd2c98a37867c06137b',
                   'Amanda Ghost (New)_61fdbe07a8b2e77c67c486df',
                   'John Borger (New)_61fb1dbfa8b2e77c67c486dc']

        queries = ['SELECT DISTINCT Original_Song_Title_SB AS Name, Pool_Revenue_9LC AS Revenue FROM Master'] * \
                  len(dbnames)
        data = pd.concat([extract_data(dbname=dbname, query=query) for dbname, query in zip(dbnames, queries)], axis=0)
        data.to_csv(DATA_FILE)

    names, not_names = split_to_classes(data_df=data)

    def clean_nans(lst):
        return [i for i in lst if pd.notna(i)]

    names = clean_nans(names)
    not_names = clean_nans(not_names)

    # en_dict = enchant.Dict("en_US")

    if featurize_data:
        en_dict = None
        return [featurize(english_dict=en_dict, phrase=str(name)) for name in names], \
               [featurize(english_dict=en_dict, phrase=str(not_name)) for not_name in not_names]
    else:
        return names, not_names


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


def plot_data(featurized_names, featurized_not_names):

    x_yes = [i[0] for i in featurized_names]
    y_yes = [i[1] for i in featurized_names]

    x_no = [i[0] for i in featurized_not_names]
    y_no = [i[1] for i in featurized_not_names]

    plt.scatter(x=x_yes, y=y_yes, c='blue', marker='^')
    plt.scatter(x=x_no, y=y_no, c='red', marker='o')

    plt.xlabel("% words that appear at the English dictionary")
    plt.ylabel("% words that appear at the 'saved words' list")

    plt.show()


if __name__ == "__main__":

    names, not_names = get_data()

    not_names_features = [i[1] for i in not_names]

    i = 0

    for name, feature in names:
        for not_name, feature1 in not_names:
            if feature == feature1:
                i += 1
                print(name, ":", not_name, ":", feature)

    print(i)
    #
    print(len(set(names).intersection(set(not_names))))
    #
    # print(set(names))
    # print(set(not_names))
    #
    print(len(set(names + not_names)))

    print(f"{len([i for i in names if i in not_names])} collisions, out of total {len(names) + len(not_names)} samples")

    # plot_data(featurized_names=names,
    #           featurized_not_names=not_names



