
import tabula
import pandas as pd

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


class PRSParser:

    def __init__(self, pdf_filepath):

        self.pdf_filepath = pdf_filepath
        self.extracted_names = set()  # Here we'll save all the names we've extracted from the file

        self.result = None

        self.pages_to_parse = [(i, df) for i, df in enumerate(self.pdf_to_pages())]
        # ^ : Here we'll store the pages that we've not parsed yet due to some error occurred at the parsing
        #     - as tuples of (page_number, original_df)

        self.parsed_pages = []
        # ^ : Here we'll store the pages that we could extract from the file - as tuples of (page_number, formatted_df)

    def pdf_to_pages(self):

        """
        :return: a list of tuples of (page_number, page's dataframe).
        """

        return tabula.read_pdf(self.pdf_filepath, pages='all', area=[50, 0, 1000, 1000], guess=False)

    def parse(self, start_page, end_page):
        """

        :param start_page: the page to start to parse the file from.
        :param end_page: the last page to parse at the file.
        :return: nothing. saves the results at self.result. call save_result() if you wan't the parser to save it.
        """

        amount_of_pages_to_parse = end_page - start_page + 1

        while len(self.pages_to_parse) > amount_of_pages_to_parse:

            for item in self.pages_to_parse[start_page: end_page + 1]:
                page_number, page_df = item

                try:
                    self.parse_page(page_number=page_number, page_df=page_df)
                    self.pages_to_parse.remove(item)

                except ParsingError:
                    # we couldn't extract some data from the block at the moment. keeping the page for parsing
                    # at a later moment when we'll have more data that will help us extract the information from
                    # the file (namely - the ip's)
                    # print(self.extracted_names)
                    pass

        # Here we have all of the pages parsed

        self.parsed_pages.sort(key=lambda item:item[0])
        self.result = pd.concat([item[1] for item in self.parsed_pages], ignore_index=True)

    def save_result(self, output_filepath):

        if self.result is None:
            print("call parse() before you call save_result()")
            exit(1)

        # else:

        self.result.to_csv(output_filepath)

    def parse_page(self, page_df, page_number):
        """
        Extracts and adds the information from the given page to the list of dataframes called 'parsed_pages'.

        :param total_df: the data collected till now, organized in the desired format.
        :param page_df: the data of the current page. needs to be organized and added to the total data.
        :param page_number: the number of the page at the given file.

        """

        songs_dfs = self.page_df_to_blocks(page_df)  # assigns a df for each song. every df is called a 'block'
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

        formatted_blocks = []

        for song_block in songs_dfs:
            formatted_blocks += [self.parse_block(song_block, are_separated1, are_separated2, page_number)]

        self.parsed_pages += [(page_number, pd.concat(formatted_blocks))]

    @staticmethod
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

        blocks = []

        curr_block_start_idx = 1

        for i in range(1, len(page_df)):

            if (type(page_df['Your Share %'][i]) == str) and (page_df['Your Share %'][i][-1] == '%'):
                # we are starting a new block
                curr_block_start_idx = i

            if page_df['Your Share %'][i] == 'Sub Total':
                # we are ending a block
                blocks.append(page_df.iloc[curr_block_start_idx: i])

        if len(blocks) == 0:
            blocks.append(page_df.iloc[curr_block_start_idx: len(page_df)-2])

        return blocks

    def parse_block(self, block, are_separated1, are_separated2, page_number):
        """

        :param page_number: the page number of the block we're parsing. used mostly for debugging.
        :param total_df: a pandas.df that contains the total data extracted till now from the pdf file.
        :param block: a pandas.df of a song.
        :param are_separated1: a boolean that == True <---> 'Work No' and 'Usage and Territory' have SEPARATED columns
        :param are_separated2: a boolean that == True <---> 'IP1' and 'IP2' have SEPARATED columns (at the given df)
        :return: the block at the desired format (as a pandas.df).
        """

        result = pd.DataFrame(columns=['Song Name', 'Work No', 'Usage and Territory', 'IP1', 'IP2', 'Royalty'])

        block = block.reset_index()
        song_name = block['Work Title'][0]

        if are_separated1:
            # 'Work No' and 'Usage and Territory' are -- separated --

            if are_separated2:
                # 'IP1' and 'IP2' are -- separated --
                ip1 = block['IP1'][0]
                ip2 = block['IP2'][0]
                self.extracted_names.add(ip1)
                self.extracted_names.add(ip2)
            else:
                # 'IP1' and 'IP2' are -- united --
                ip1, ip2 = self.extract_ips(block['IP1 IP2'][0])
        else:
            # 'Work No' and 'Usage and Territory' are -- united --

            if are_separated2:
                ip1 = block['IP1'][0]
                ip2 = block['IP2'][0]
                self.extracted_names.add(ip1)
                self.extracted_names.add(ip2)
            else:
                # 'IP1' and 'IP2' are -- united --
                ip1, ip2 = self.extract_ips(block['IP1 IP2'][0])

        ip3 = block['IP3'][0]
        ip4 = block['IP4'][0]

        for i in range(1, len(block)):

            if are_separated1:
                # 'Work No' and 'Usage and Territory' are -- separated --

                work_no = block['Work Title'][i]
                usage_and_territory = block['Unnamed: 0'][i]

            else:
                # 'Work No' and 'Usage and Territory' are -- united --
                work_no, usage_and_territory = block['Work Title'][i].split(' ', maxsplit=1)

            if 'Unnamed: 7' in block.columns:
                royalty = block['Unnamed: 7'][i]
            else:
                if 'Unnamed: 6' in block.columns:
                    royalty = block['Unnamed: 6'][i]
                else:
                    royalty = block['Unnamed: 5'][i]

            period = block['IP3'][i]

            line = pd.DataFrame({'Song Name': [song_name],
                                 'Work No': [work_no],
                                 'Usage and Territory': [usage_and_territory],
                                 'IP1': ip1,
                                 'IP2': ip2,
                                 'IP3': ip3,
                                 'IP4': ip4,
                                 'Period': period,
                                 'Royalty': royalty})

            # TODO - extract more desired attributes

            result = pd.concat([result, line], ignore_index=True, axis=0)

        return result

    def extract_ips(self, ips):
        """

        :param ips: a string that contains the two ip's.
        :return: a tuple containing the two ip's, or -1 if couldn't be extracted. if succeed to extract
                    saves the names at the 'extracted_names' set for a later use.

                 Raises a ParsingError if couldn't extract the names from the string.
        """

        for name in self.extracted_names:
            if ips.find(name) != -1:
                # the name was found at the ip's string
                splitting_idx = ips.find(name)

                if splitting_idx > 0:
                    ip1, ip2 = ips[:splitting_idx-1], ips[splitting_idx:]
                else:
                    ip1, ip2 = ips[:len(name)], ips[len(name) + 1:]

                self.extracted_names.add(ip1)
                self.extracted_names.add(ip2)

                return ip1, ip2

        raise ParsingError("Couldn't separate between ip1 and ip2")


class ParsingError(Exception):

    def __init__(self, message):
        super().__init__(self, message)


if __name__ == "__main__":

    PDF_NAME = '../example_data/2016/EssexDavid_070172307_2016121_052_089780_26722.PDF'

    parser = PRSParser(pdf_filepath=PDF_NAME)
    parser.parse(start_page=9, end_page=117)

    parser.save_result(output_filepath='result.csv')

