import MySQLdb  # IMPORTANT : pip3 install mysqlclient
import pandas as pd
import pickle


def sql2pandas(dbname):

    """
    :param dbname: The name of the database from which we want to extract information.
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
    df = cmd2csv(cursor, 'SELECT * FROM Master')
    db_connection.close()

    return df


def cmd2csv(cursor, query):

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

    # example_queries = 'SELECT Territory_SB FROM Master'

    example_output_filename = 'output.xlsx'

    sql2pandas(dbname=example_dbname)

