
def sql2pandas(query_string):


if __name__ == "__main__":
    '''sum( CASE WHEN Year_Statement_9LC = "{}" AND {} = "{}" AND
                          Normalized_Source_9LC = "Pool Revenue"
                          THEN Adjusted_Royalty_SB ELSE NULL END) AS `{} {}'''