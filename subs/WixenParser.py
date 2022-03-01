import tabula
import pandas as pd
import PyPDF2
from collections import defaultdict


class WixenParser:

    def __init__(self, pdf_filepath):

        self.pdf_filepath = pdf_filepath

        pdf_file = open(self.pdf_filepath, "rb")
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        self.num_pages = read_pdf.getNumPages()

        self.contract_to_pages = defaultdict(lambda: [])
        self.pages_to_parse = []
        self.init_pages_to_parse()

        self.parsed_pages = []
        self.result = {}

    def init_pages_to_parse(self):
        """
        scans the document and initializes the fields start_page, end_page, and pages_to_parse.
        """

        def scan(pages, contracts_ids):
            parsing_page_column_names = ['Unnamed: 0',
                                         'Unnamed: 1',
                                         'Unnamed: 2',
                                         'Unnamed: 3',
                                         'Unnamed: 4',
                                         'Amt Rcvd/Price',
                                         'Your',
                                         'Amount']

            parsing_page_column_names_1 = parsing_page_column_names.copy()
            parsing_page_column_names_1[5] = 'A mt Rcvd/Price'

            saw_page_1 = False
            i = 0

            for page, contract_id in zip(pages, contracts_ids):

                if [name for name in page.columns] == parsing_page_column_names:
                    final_page = page.iloc[1:].reset_index(drop=True)
                    self.pages_to_parse.append((i, final_page))
                    self.contract_to_pages[contract_id] += [i]

                elif [name for name in page.columns] == (parsing_page_column_names + ['Song']):
                    final_page = page.iloc[1:, :-1].reset_index(drop=True)
                    self.pages_to_parse.append((i, final_page))
                    self.contract_to_pages[contract_id] += [i]

                elif [name for name in page.columns] == (parsing_page_column_names_1 + ['Song']):
                    final_page = page.iloc[1:, :-1].reset_index(drop=True).rename(columns={'A mt Rcvd/Price': 'Amt Rcvd/Price'})
                    self.pages_to_parse.append((i, final_page))
                    self.contract_to_pages[contract_id] += [i]

                else:
                    if 'Page: 1' in [s.strip() for s in page.iloc[:, -1] if type(s) is str]:
                        if not saw_page_1:
                            saw_page_1 = True
                            continue

                        # this is the first page we're going to parse at the document:
                        # performing some initial preprocessing:

                        final_page = page.iloc[5:].reset_index(drop=True).rename(columns={original_name: new_name for
                                                                                    original_name, new_name in
                                                                                    zip(page.columns,
                                                                                        parsing_page_column_names)})

                        self.pages_to_parse.append((i, final_page))
                        self.contract_to_pages[contract_id].append(i)
                i += 1

        contracts_ids_dfs = tabula.read_pdf(self.pdf_filepath,
                                            pages='all',
                                            area=(74.88, 28.08, 195.84, 557.28),
                                            columns=(43.92, 108.72, 214.56, 249.12, 299.52, 418.32, 458.64, 578.16))

        contracts_ids = [[item.strip() for item in df.columns[2].split()] for df in contracts_ids_dfs]
        contracts_ids = [splitted_string[min([i for i in range(len(splitted_string)) if '(' in splitted_string[i]] + [len(splitted_string)-1])][1:-1] for splitted_string in contracts_ids]

        pages = tabula.read_pdf(self.pdf_filepath,
                                pages='all',
                                area=(111.6, 30, 777.72, 575.28),
                                columns=(111.6, 133.92, 218.88, 250.56, 302.4, 419.76, 465.84, 582))

        self.pages_to_parse = []
        scan(pages, contracts_ids)

        if len(self.pages_to_parse) / len(pages) < 0.25:
            # this is the format of years >= 2021

            self.pages_to_parse = []
            pages = tabula.read_pdf(self.pdf_filepath,
                                    pages='all',
                                    area=(111.6, 30, 777.72, 575.28),
                                    columns=(97.92, 118.08, 174.96, 220.32, 259.92, 366.48, 401.76, 504.72))
            scan(pages, contracts_ids)

    def save_result(self, output_filepath):

        if self.result is None:
            print("Call parse() before calling save_result()")
            raise Exception

        # else:

        for contract_id, df in self.result.items():
            df.to_csv('.'.join(output_filepath.split('.')[:-1]) + f'_{contract_id}.' + output_filepath.split('.')[-1],
                      index=False)

    def parse(self):

        """
        :return: nothing. saves the result at self.result.
        """

        curr_song_name, curr_artist, curr_territory = None, None, None

        for i, page_df in self.pages_to_parse:
            curr_song_name, curr_artist, curr_territory = self.parse_page(page_df=page_df,
                                                                          page_idx=i,
                                                                          curr_song_name=curr_song_name,
                                                                          curr_artist=curr_artist,
                                                                          curr_territory=curr_territory)

        # -- extra care for the last page --

        hypens = ['-' * j for j in range(2, 20)]

        if self.parsed_pages[-1][1].iloc[-1]['Share'] in hypens:
            self.parsed_pages[-1][1].iloc[-1]['Share'] = self.parsed_pages[-1][1][:-1]

        # -- concatenating the df's for each contract id --

        for contract_id, pages_idxs in self.contract_to_pages.items():
            self.result[contract_id] = pd.concat([page_df for i, page_df in self.parsed_pages if i in pages_idxs],
                                                 ignore_index=True)

            # finalize
            curr_df = self.result[contract_id]

            # changing the columns names
            curr_df = curr_df.rename(columns={'Song Name': 'song',
                                              'C': 'credit',
                                              'Territory': 'channel',
                                              'Usage': 'source',
                                              'A': 'country',
                                              'B': 'specification',
                                              'Units': 'units',
                                              'Price': 'price',
                                              'Share': 'share',
                                              'Amount': 'amount'})

            # finalizing the country and the period columns
            curr_df['country'] = [''.join(item.split()) if pd.notna(item) else '' for item in curr_df['country']]
            curr_df['period'] = curr_df['Period - Start'] + '-' + curr_df['Period - End']

            # dropping those 2 columns
            curr_df.drop('Period - Start', inplace=True, axis=1)
            curr_df.drop('Period - End', inplace=True, axis=1)

            # moving the new 'period' column to be before 'price', 'share' and 'amount' columns
            cols = list(curr_df.columns)
            curr_df = curr_df[cols[:-4] + [cols[-1]] + cols[-4: -1]]

            # adding the new contract id column
            curr_df['contract'] = [contract_id for i in range(len(curr_df))]

            self.result[contract_id] = curr_df

    def parse_page(self, page_df, page_idx, curr_song_name, curr_artist, curr_territory):
        """
        adds the parsed data frame to the list self.parsed_pages.

        :param page_df: tabula's extracted df of the page we currently want to parse at the document.
        :param page_idx: The identifier of the page (it's name at the document).
        :param curr_song_name: The song name of the last block in the previous page we were parsing.
        :param curr_artist: The artist name " " ".
        :param curr_territory: The territory " " ".

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
            self.parsed_pages += [(page_idx, pd.concat(parsed_blocks, ignore_index=True))]
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

        rows_to_remove = []

        if not pd.isna(page_df.iloc[0]['Amount']):
            rows_to_remove.append(0)

        for i in range(0, len(page_df)):
            if type(page_df['Amount'][i]) is str and \
                    page_df['Amount'][i].strip() in (['-' * j for j in range(0, 25)] + ['Due']):

                # it's a row that come before/after some total some row, or an unnecessary row
                rows_to_remove += [i]

            if 0 < i < len(page_df) - 1 and \
                    (type(page_df['Amount'][i-1]) is str and \
                     page_df['Amount'][i-1].strip() in (['-' * j for j in range(25)] + ['=' * j for j in range(25)])) and \
                    (type(page_df['Amount'][i+1]) is str and \
                     page_df['Amount'][i+1].strip() in (['-' * j for j in range(25)] + ['=' * j for j in range(25)])):

                # it's a total sum row
                rows_to_remove += [i]

            if sum([pd.isna(l) for l in page_df.iloc[i]]) >= 7 and (not pd.isna(page_df.iloc[i]['Amount'])):
                rows_to_remove += [i]

            if [name for name in page_df.columns if (not pd.isna(page_df.iloc[i][name]))] == ['Amt Rcvd/Price', 'Amount']:
                # it's a sub total row that wasn't removed
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

        try:
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
                            curr_territory = page_df['Unnamed: 0'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 1]
                        if header_cnt >= 2:
                            curr_artist = page_df['Unnamed: 0'].fillna('')[i - 2] + \
                                          page_df['Unnamed: 1'].fillna('')[i - 2] + \
                                          page_df['Unnamed: 2'].fillna('')[i - 2]
                        if header_cnt >= 3:
                            curr_song_name = page_df['Unnamed: 0'].fillna('')[i - 3] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 3] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 3]

                        if header_cnt == 4:
                            curr_song_name = page_df['Unnamed: 0'].fillna('')[i - 4] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 4] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 4]

                            curr_artist = page_df['Unnamed: 0'].fillna('')[i - 3] + \
                                          page_df['Unnamed: 1'].fillna('')[i - 3] + \
                                          page_df['Unnamed: 2'].fillna('')[i - 3]

                            curr_territory = page_df['Unnamed: 0'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 1]

                        if header_cnt == 5:
                            curr_song_name = page_df['Unnamed: 0'].fillna('')[i - 5] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 5] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 5]

                            curr_artist = page_df['Unnamed: 0'].fillna('')[i - 2] + \
                                          page_df['Unnamed: 1'].fillna('')[i - 2] + \
                                          page_df['Unnamed: 2'].fillna('')[i - 2]

                            curr_territory = page_df['Unnamed: 0'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 1]

                        if header_cnt >= 6:
                            curr_song_name = page_df['Unnamed: 0'].fillna('')[i - 3] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 3] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 3]

                            curr_artist = page_df['Unnamed: 0'].fillna('')[i - 2] + \
                                          page_df['Unnamed: 1'].fillna('')[i - 2] + \
                                          page_df['Unnamed: 2'].fillna('')[i - 2]

                            curr_territory = page_df['Unnamed: 0'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 1'].fillna('')[i - 1] + \
                                             page_df['Unnamed: 2'].fillna('')[i - 1]

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
        except UnboundLocalError:
            # it's some of the last pages, which we don't need to parse
            pass

        return blocks, curr_song_name, curr_artist, curr_territory

    def parse_block(self, block):

        """
        :param block: an instance of the 'Block' class.
        :return: The information at the block at the desired format, ad a pandas.DataFrame.
        """

        result = pd.DataFrame(columns=['Song Name',
                                       'C',
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

        curr_usage = None

        for i in range(0, len(block.df)):
            line = block.df.iloc[i]

            if not pd.isna(line['Unnamed: 0']):
                curr_usage = line['Unnamed: 0']

            A = line['Unnamed: 1']
            B = line['Unnamed: 2']

            units = line['Unnamed: 3']
            period = line['Unnamed: 4']

            price = line['Amt Rcvd/Price']
            share = line['Your']
            amount = line['Amount']

            if len(period.split('-')) == 1:
                # we are in some of the last blocks at the file and this line is not necessary, we stop parsing here
                break

            line_df = pd.DataFrame({'Song Name': [block.song_name],
                                    'C': [block.artist],
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
