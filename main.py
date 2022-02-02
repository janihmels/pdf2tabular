
import tabula
import pandas as pd

PDF_NAME = '../example_data/2016/EssexDavid_070172307_2016121_052_089780_26722.PDF'

ATTRIBUTES = ['Work Title',
              'Work Creator',
              'Place',
              'Publisher',
              'Work No',
              'Usage & Territory',
              'Broadcast Region',
              'Period',
              'Production',
              'Performances',
              'Royalty']


def pdf_to_pages(pdf_filepath):

    """

    :param pdf_filepath:
    :return: a list of dataframes, where each corresponds to a page in a dataframe.
    """

    return tabula.read_pdf(PDF_NAME, pages='all', area=[50, 0, 1000, 1000], guess=False)


def parse_page(total_df, page_df):
    """

    :param total_df: the data collected till now, organized in the desired format.
    :param page_df: the data of the current page. needs to be organized and added to the total data.
    :return: the new total_df.
    """

    songs_dfs = page_df_to_blocks(page_df)  # assigns a df for each song. every df is called a 'block'
    cols_names = list(songs_dfs[0].columns)  # specify the format of the extracted data frame of the page

    if cols_names[1] == 'Unnamed: 0':
        # 'Work No' and 'Usage and Territory' are -- separated --
        are_separated1 = True

        if cols_names[2] == 'IP1':
            # 'IP1' and 'IP2' are -- separated --
            are_separated2 = True
        else:
            # 'IP1' and 'IP2' are -- united --
            are_separated2 = False
    else:
        # 'Work No' and 'Usage and Territory' are -- united --
        are_separated1 = False

        if cols_names[1] == 'IP1':
            # 'IP1' and 'IP2' are -- separated --
            are_separated2 = True
        else:
            # 'IP1' and 'IP2' are -- united --
            are_separated2 = False

    for song_block in songs_dfs:
        parse_block(total_df, song_block, are_separated1, are_separated2)

    return total_df


def page_df_to_blocks(page_df):
    """

    :param page_df: the non organized dataframe of the page we're parsing.
    :return: a list of dataframes, where each is for a different song.
    """

    # there are 3 types of lines:
    #
    # those that start a block (whom 'Your Share' value is a percentage) which define a *new song* - 'metadata'
    # those in a block (whom 'Your Share' value in them is an integer) - the 'data' ones
    # those that end a block (whom 'Your Share' value in them is not a numeric one, but is the string 'Sub Total') -
    #                                                                                                         'metadata'

    curr_block = pd.DataFrame(columns=page_df.columns)
    blocks = []

    for i in range(1, curr_block.size):
        if page_df.iloc[i]['Your Share %'][-1] == '%':
            # we are starting a new block
            curr_block_start_idx = i
        if page_df.iloc[i]['Your Share %'] == 'Sub Total':
            # we are ending a block
            blocks.append(page_df.iloc[curr_block_start_idx: i])

    return blocks


def parse_block(total_df, block, are_separated1, are_separated2):
    """

    :param total_df: a pandas.df that contains the total data extracted till now from the pdf file.
    :param block: a pandas.df of a song.
    :param are_separated1: a boolean that == True <---> 'Work No' and 'Usage and Territory' have SEPARATED columns
    :param are_separated2: a boolean that == True <---> 'IP1' and 'IP2' have SEPARATED columns (at the given df)
    :return: the block at the desired format (as a pandas.df).
    """

    song_name = block['Work Title'][0]

    for i in range(1, block.size):

        if are_separated1:
            # 'Work No' and 'Usage and Territory' are -- separated --

            work_no = block['Work Title'][i]
            usage_and_territory = block['Unnamed: 0'][i]

            if are_separated2:
                # 'IP1' and 'IP2' are -- separated --
                ip1 = block['IP1'][i]
                ip2 = block['IP2'][i]
            else:
                # 'IP1' and 'IP2' are -- united --
                # TODO - extract ip1 and ip2
                pass
        else:
            # 'Work No' and 'Usage and Territory' are -- united --

            work_no, usage_and_territory = line['Work Title'].split(' ', maxsplit=1)

            if are_separated1:
                ip1 = block['IP1'][i]
                ip2 = block['IP2'][i]
            else:
                # 'IP1' and 'IP2' are -- united --
                #  TODO - extract ip1 and ip2
                pass

        if 'Unnamed: 7' in block.columns:
            royalty = block['Unnamed: 7'][i]
        else:
            royalty = block['Unnamed: 6'][i]

        # TODO - extract more desired attributes

if __name__ == "__main__":

    pages_dfs = pdf_to_pages(pdf_filepath=PDF_NAME)  # WILL BE CHANGED TO BE SPECIFIED IN THE CMD
    # ^ : almost parses all the data at each page. assigns a data frame for each page.

    total_df = pd.DataFrame(columns = ATTRIBUTES)

    for page_df in pages_dfs:
        parse_page(total_df=total_df,
                    page_df=page_df)

    total_df.to_csv("Result.csv")  # WILL BE CHANGED TO BE SPECIFIED IN THE CMD