import pandas as pd
from datetime import date


def find_cut_off_year(df: pd.DataFrame):

    recent_years = df['Year_Statement_9LC'].drop_duplicates()

    todays_date = date.today()
    current_year = todays_date.year

    recent_year = recent_years.iloc[-1]
    if recent_year == current_year:
        find_cut_off_year = current_year
    else:
        find_cut_off_year = current_year - 1

    # Find cut off
    statement_period_half_blank = \
        df['Statement_Period_Half_9LC'][df['Year_Statement_9LC'].astype(int) <= find_cut_off_year].\
            drop_duplicates().sort_values(ascending=False)

    def check_blank(period):
        if period == '':
            return False
        else:
            return True

    remove_blank = filter(check_blank, statement_period_half_blank)
    statement_period_minus_blank = list(remove_blank)

    if len(statement_period_minus_blank) > 10:
        standard_cut_off = statement_period_minus_blank[-10]
    else:
        standard_cut_off = statement_period_minus_blank[0]

    year_cut_off = int(standard_cut_off[0:4])

    return year_cut_off