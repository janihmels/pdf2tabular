
import tabula
import pandas as pd
import PyPDF2

class WixenParser:

    def __init__(self, pdf_filepath):

        self.pdf_filepath = pdf_filepath

        pdf_file = open(self.pdf_filepath, "rb")
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        self.num_pages = read_pdf.getNumPages()

        self.pages_to_parse = None
        self.init_pages_to_parse()

        self.parsed_pages = []

        self.result = None

    def init_pages_to_parse(self):
        """
        sacns the document and initializes the fields start_page, end_page, and pages_to_parse.
        """

        pages = tabula.read_pdf(self.pdf_filepath,
                                pages='all',
                                area=(111.6, 30, 777.72, 575.28),
                                columns=(111.6, 133.92, 218.88, 250.56, 302.4, 419.76, 465.84, 582))

        self.pages_to_parse = []
        parsing_page_column_names = ['Unnamed: 0',
                                     'Unnamed: 1',
                                     'Unnamed: 2',
                                     'Unnamed: 3',
                                     'Unnamed: 4',
                                     'Amt Rcvd/Price',
                                     'Your',
                                     'Amount']

        saw_page_1 = False

        for page in pages:
            if [name for name in page.columns] == parsing_page_column_names:
                self.pages_to_parse.append(page.iloc[1:].reset_index(drop=True))
            else:
                if 'Unnamed: 5' in page.columns and page['Unnamed: 5'][0] == 'Page: 1':

                    if not saw_page_1:
                        saw_page_1 = True
                        continue

                    # this is the first page we're going to parse at the document:
                    # performing some initial preprocessing:

                    page = page.iloc[5:].reset_index(drop=True).rename(columns={original_name: new_name for
                                                                                original_name, new_name in
                                                                        zip(page.columns, parsing_page_column_names)})

                    self.pages_to_parse.append(page)

    def save_result(self, output_filepath):

        if self.result is None:
            print("Call parse() before calling save_result()")
            raise Exception

        # else:

        self.result.to_csv(output_filepath)

    def parse(self):

        """
        :return: nothing. saves the result at self.result.
        """

        curr_song_name, curr_artist, curr_territory = None, None, None

        for page_df in self.pages_to_parse:
            curr_song_name, curr_artist, curr_territory = self.parse_page(page_df,
                                                                          curr_song_name,
                                                                          curr_artist,
                                                                          curr_territory)

        self.result = pd.concat(self.parsed_pages, ignore_index=True)

    def parse_page(self, page_df, curr_song_name, curr_artist, curr_territory):
        """
        adds the parsed data frame to the list self.parsed_pages.

        :param curr_song_name: the song name of the last block in the previous page we were parsing.
        :param curr_artist: the artist name "".
        :param curr_territory: the territory "".
        :param page_df: tabula's extracted df of the page we currently want to parse at the document.
        :return: curr_song_name, curr_artist, curr_territory.
        """

        blocks, curr_song_name, curr_artist, curr_territory = WixenParser.page_df_to_blocks(page_df,
                                                                                            curr_song_name,
                                                                                            curr_artist,
                                                                                            curr_territory)

        parsed_blocks = []

        for block in blocks:
            parsed_blocks += [self.parse_block(block=block)]

        if len(parsed_blocks) > 0:
            self.parsed_pages += [pd.concat(parsed_blocks, ignore_index=True)]
        else:
            # we are at the end of the document and identified this page as a data page,
            # while it's not.
            pass

        return curr_song_name, curr_artist, curr_territory

    @staticmethod
    def page_df_to_blocks(page_df, curr_song_name, curr_artist, curr_territory):
        """
        :param curr_song_name: the song name of the last block in the previous page we were parsing.
        :param curr_artist: the artist name "".
        :param curr_territory: the territory "".
        :param page_df: the non organized dataframe of the page we're parsing.
        :return: a list of dataframes, where each is for a different song.
        """

        # ----- cleaning the page's dataframe ----- :

        i = 0
        rows_to_remove = []

        for i in range(0, len(page_df)):
            if pd.isna(page_df['Unnamed: 0'][i]):
                rows_to_remove += [i]

        page_df = page_df.drop(rows_to_remove, axis=0).reset_index(drop=True)

        # ----- dividing the page into blocks ----- :

        def is_header_line(line):
            """

            :param line: a pd.DataFrame corresponding to a line in the block we're working on.
            :return: True <--> the current line defines some header (song name, artist name, or usage description).
            """

            if pd.isna(line['Amount']) or pd.isna(line['Unnamed: 4']):
                return True
            # else:

            return False

        blocks = []

        header_cnt = 0  # counts the amount of lines that contains headers (song_name, artist, or territory) that
        # we've seen before reaching the current line

        curr_block_is_open = False

        for i in range(0, len(page_df)):

            # reminder : header_cnt is the amount of rows before this one whom 'Amt Rcvd/Price Your' column was empty

            if is_header_line(page_df.iloc[i]):
                # curr line is a header:

                if header_cnt == 0 and i != 0:
                    # we've just finished to move over a new block
                    blocks.append(Block(block_df=page_df.iloc[curr_block_start_idx: i].reset_index(drop=True),
                                        song_name=curr_song_name,
                                        artist=curr_artist,
                                        territory=curr_territory))

                    curr_block_is_open = False

                # updating the header counter
                header_cnt += 1
            else:
                # curr line is a data line and not a header one:
                if header_cnt > 0:
                    # we've reached a new block:

                    # updating metadata
                    if header_cnt >= 1:
                        curr_territory = page_df['Unnamed: 0'].fillna('')[i-1] + \
                                         page_df['Unnamed: 1'].fillna('')[i-1] + \
                                         page_df['Unnamed: 2'].fillna('')[i-1]
                    if header_cnt >= 2:
                        curr_artist = page_df['Unnamed: 0'].fillna('')[i-2] + \
                                      page_df['Unnamed: 1'].fillna('')[i-2] + \
                                      page_df['Unnamed: 2'].fillna('')[i-2]
                    if header_cnt >= 3:
                        curr_song_name = page_df['Unnamed: 0'].fillna('')[i-3] + \
                                         page_df['Unnamed: 1'].fillna('')[i-3] + \
                                         page_df['Unnamed: 2'].fillna('')[i-3]

                    if header_cnt == 4:
                        curr_song_name = page_df['Unnamed: 0'].fillna('')[i-4] + \
                                         page_df['Unnamed: 1'].fillna('')[i-4] + \
                                         page_df['Unnamed: 2'].fillna('')[i-4]

                        curr_artist = page_df['Unnamed: 0'].fillna('')[i-3] + \
                                      page_df['Unnamed: 1'].fillna('')[i-3] + \
                                      page_df['Unnamed: 2'].fillna('')[i-3]

                        curr_territory = page_df['Unnamed: 0'].fillna('')[i-1] + \
                                         page_df['Unnamed: 1'].fillna('')[i-1] + \
                                         page_df['Unnamed: 2'].fillna('')[i-1]

                    if header_cnt == 5:
                        curr_song_name = page_df['Unnamed: 0'].fillna('')[i-5] + \
                                         page_df['Unnamed: 1'].fillna('')[i-5] + \
                                         page_df['Unnamed: 2'].fillna('')[i-5]

                        curr_artist = page_df['Unnamed: 0'].fillna('')[i-2] + \
                                      page_df['Unnamed: 1'].fillna('')[i-2] + \
                                      page_df['Unnamed: 2'].fillna('')[i-2]

                        curr_territory = page_df['Unnamed: 0'].fillna('')[i-1] + \
                                         page_df['Unnamed: 1'].fillna('')[i-1] + \
                                         page_df['Unnamed: 2'].fillna('')[i-1]

                    if header_cnt >= 6:
                        curr_song_name = page_df['Unnamed: 0'].fillna('')[i-3] + \
                                         page_df['Unnamed: 1'].fillna('')[i-3] + \
                                         page_df['Unnamed: 2'].fillna('')[i-3]

                        curr_artist = page_df['Unnamed: 0'].fillna('')[i-2] + \
                                      page_df['Unnamed: 1'].fillna('')[i-2] + \
                                      page_df['Unnamed: 2'].fillna('')[i-2]

                        curr_territory = page_df['Unnamed: 0'].fillna('')[i-1] + \
                                         page_df['Unnamed: 1'].fillna('')[i-1] + \
                                         page_df['Unnamed: 2'].fillna('')[i-1]

                    # declaring a new starting of a block
                    curr_block_start_idx = i
                    curr_block_is_open = True

                    # resetting the header counter
                    header_cnt = 0

                else:
                    # we're just inside some block
                    pass

        if curr_block_is_open:
            blocks.append(Block(block_df=page_df.iloc[curr_block_start_idx: len(page_df)].reset_index(drop=True),
                                song_name=curr_song_name,
                                artist=curr_artist,
                                territory=curr_territory))

        return blocks, curr_song_name, curr_artist, curr_territory

    def parse_block(self, block):

        """
        :param block: an instance of the 'Block' class.
        :return: The information at the block at the desired format, ad a pandas.DataFrame.
        """

        result = pd.DataFrame(columns=['Song Name',
                                       'Artist',
                                       'Territory',
                                       'Usage',
                                       'A',
                                       'B',
                                       'Units',
                                       'Period - Start',
                                       'Period - End',
                                       'Price',
                                       'Share',
                                       'Amount'])

        block.df.to_csv('last_block.csv')

        curr_usage = None

        for i in range(0, len(block.df)):
            line = block.df.iloc[i]

            if ~pd.isna(line['Unnamed: 0']):
                curr_usage = line['Unnamed: 0']

            A = line['Unnamed: 1']
            B = line['Unnamed: 2']

            units = line['Unnamed: 3']
            period = line['Unnamed: 4']

            price = line['Amt Rcvd/Price']
            share = line['Your']
            amount = line['Amount']

            line_df = pd.DataFrame({'Song Name': [block.song_name],
                                    'Artist': [block.artist],
                                    'Territory': [block.territory],
                                    'Usage': [curr_usage],
                                    'A': [A],
                                    'B': [B],
                                    'Units': [units],
                                    'Period - Start': [period.split('-')[0]],
                                    'Period - End': [period.split('-')[1]],
                                    'Price': [price],
                                    'Share': [share],
                                    'Amount': [amount]
                                    })

            result = pd.concat([result, line_df], ignore_index=True, axis=0)

        return result


class Block:

    def __init__(self, block_df, song_name, artist, territory):

        self.df = block_df

        self.song_name = song_name
        self.artist = artist
        self.territory = territory