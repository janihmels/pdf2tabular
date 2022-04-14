#import MySQLdb  # IMPORTANT : pip3 install mysqlclient
import pandas as pd


def sql2xlsx(dbname, queries, output_filename):

    """
    :param dbname: The name of the database from which we want to extract information.
    :param queries: A string that contains a single query, or a dictionary of {query_name : query} where query is a str.
    :param output_filename: The name of .xlsx file at which we'll store the extracted data.
    :return: 1 if succeeded. -1 otherwise.
    """

    try:
        db_connection = MySQLdb.connect('34.65.111.142',
                                        'external',
                                        'musicpass',
                                        dbname)
    except:
        print("Error: Couldn't connect to the specified database: '{0}'.".format(dbname))
        return -1

    # -- connected --

    cursor = db_connection.cursor()

    if type(queries) == str:
        queries = {'sheet 1': queries}

    sheets = []

    for sheet_name, query in queries.items():
        try:
            sheet_df = cmd2xlsx(cursor, query)
        except:
            print("Error: Couldn't execute the query: '{0}'.".format(query))
            return -1

        # else:
        sheets += [(sheet_name, sheet_df)]

    db_connection.close()

    try:
        with pd.ExcelWriter(output_filename) as writer:
            for sheet_name, sheet_df in sheets:
                sheet_df.to_excel(writer, sheet_name=sheet_name)
    except:
        print("Error: Couldn't write the output to the specified file: '{0}'.".format(output_filename))
        return -1

    return 1


def cmd2xlsx(cursor, query):

    """
    :param cursor: A cursor object from the MySQLdb package, which gives us access to a database.
    :param query: A string that contains a query to perform on the database.
    :return: The table returned from the query, as a pandas.DataFrame.
    """

    def get_col_names():
        """
        :return: A list of the columns names of the table extracted by the given query.
        """

        field_names = [i[0] for i in cursor.description]

        return field_names

    cursor.execute(query)
    data = cursor.fetchall()  # fetching the data
    df = pd.DataFrame.from_records(data=data, columns=get_col_names())

    return df


if __name__ == "__main__":

    example_dbname = 'Adman Khan_61b775f7c94b68a289900e81'

    example_queries = {'sheet 1': 'SELECT Territory_SB FROM Master',
                       'sheet 2': 'SELECT * FROM Master'}

    # example_queries = 'SELECT Territory_SB FROM Master'

    example_output_filename = 'output.xlsx'

    sql2xlsx(dbname=example_dbname,
             queries=example_queries,
             output_filename=example_output_filename)