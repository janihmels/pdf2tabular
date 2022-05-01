
from pandasql import sqldf


class pandas_cursor:
    def __init__(self, df):
        self.df = df
        self.temporal_result = None
        self.apply_query_fn = lambda q: sqldf(q, {'Master': df})

    def execute(self, query):
        self.temporal_result = self.apply_query_fn(q=query).values.tolist()

    def fetchall(self):
        result = self.temporal_result
        self.temporal_result = None
        return result

