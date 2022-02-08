
import tabula
import pandas as pd
import sys
from subs.PRSParser_Finalizer import PRSParser_Finalizer

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

        pd.options.mode.chained_assignment = None  # default='warn'

        self.pdf_filepath = pdf_filepath
        self.extracted_names = set()  # Here we'll save all the names we've extracted from the file

        self.result = None

        self.pages_to_parse = [(i+1, df) for i, df in enumerate(self.pdf_to_pages())]
        # ^ : Here we'll store the pages that we've not parsed yet due to some error occurred at the parsing
        #     - as tuples of (page_number, original_df)

        self.parsed_pages = []
        # ^ : Here we'll store the pages that we could extract from the file - as tuples of (page_number, formatted_df)

        self.starting_page = min([i for i, df in self.pages_to_parse if len(df.columns) > 8])
        # ^ : The page to start to parse the document from.

        self.end_page = max([i for i, df in self.pages_to_parse if len(df.columns) > 8 and 'Your Share %' in df.columns])
        # ^ : The page to end the parsing end (including).

        self.pages_to_parse = [(i, df) for i, df in self.pages_to_parse if self.starting_page <= i <= self.end_page]

        self.work_details = dict()
        # ^ : a 'list' of dictionaries containing the work details for every page between
        #     [self.starting_page, self.end_page]

        self.init_work_details()

        self.amount_of_pages_to_parse = self.end_page - self.starting_page + 1

        self.finalizer = PRSParser_Finalizer(pdf_filepath=self.pdf_filepath,
                                             starting_page=self.end_page + 1)

        self.finalizer_on = False

    def pdf_to_pages(self):

        """
        :return: a list of tuples of (page_number, page's dataframe).
        """

        return tabula.read_pdf(self.pdf_filepath,
                               pages='all',
                               area=(68, 17.28, 555.84, 818.64),
                               guess=False,
                               silent=True)

    def init_work_details(self):
        """

        :return: saves a 'list' of dictionaries containing the 'Work Details' information for every page,
                    at the field self.work_details.
        """

        dfs = tabula.read_pdf(self.pdf_filepath,
                              pages=str(self.starting_page) + '-' + str(self.end_page),
                              area=(10, 647.28, 61.92, 812.16),
                              guess=False,
                              silent=True)

        page_number = self.starting_page

        for df in dfs:
            details = {'Member Name': df['Work Detail'].iloc[0].split(':')[1].strip(),
                       'CAE Number': df['Work Detail'].iloc[1].split(':')[1].strip(),
                       'Distribution Number': df['Work Detail'].iloc[2].split(':')[1].strip().split()[0]}

            self.work_details[page_number] = details
            page_number += 1

    def get_work_details(self, page_number):
        """
        returns the dictionary containing the work details for the specified page number.
        """

        return self.work_details[page_number]

    def parse(self):
        """
        parses the document and stores the result at self.result. call save_result() if you want the parser to save it.
        :return: none.
        """

        previous_amount_of_pages_to_parse = self.amount_of_pages_to_parse

        while self.amount_of_pages_to_parse > 0:

            for item in self.pages_to_parse:

                page_number, page_df = item

                try:
                    self.parse_page(page_number=page_number, page_df=page_df)
                    self.pages_to_parse.remove(item)
                    self.amount_of_pages_to_parse -= 1

                except ParsingError:
                    # we couldn't extract some data from the block at the moment. keeping the page for parsing
                    # at a later moment when we'll have more data that will help us extract the information from
                    # the file (namely - the ip's)

                    pass

            if self.amount_of_pages_to_parse == previous_amount_of_pages_to_parse:
                # we can't extract more ip's naively as we did before. forcing the extraction:
                self.force_extract_pages_names(pages_numbers=[i for i, df in self.pages_to_parse])

            previous_amount_of_pages_to_parse = self.amount_of_pages_to_parse

        # Here we have all of the pages parsed

        self.parsed_pages.sort(key=lambda item:item[0])
        self.result = pd.concat([item[1] for item in self.parsed_pages], ignore_index=True)

        if self.finalizer_on:
            self.finalizer.parse()
            finalizers_result = self.finalizer.get_result()

            self.result = pd.concat([self.result, finalizers_result], ignore_index=True)

    def save_result(self, output_filepath):

        if self.result is None:
            raise ParsingError("call parse() before you call save_result()")

        # else:

        self.result.to_csv(output_filepath)

    def parse_page(self, page_df, page_number):
        """
        Extracts and adds the information from the given page to the list of dataframes called 'parsed_pages'.

        :param total_df: the data collected till now, organized in the desired format.
        :param page_df: the data of the current page. needs to be organized and added to the total data.
        :param page_number: the number of the page at the given file.

        """

        work_details = self.get_work_details(page_number)
        self.add_name(work_details['Member Name'])
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
            formatted_blocks += [self.parse_block(song_block, are_separated1, are_separated2, page_number, work_details)]

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
        #                                                                                                     'metadata'

        # unify the pages
        if all(pd.isna(page_df['Your Share %'])):
            # The 'Your Share %' column and 'Unnamed: {mixed_unnamed_index}' columns were mixed
            mixed_unnamed_index = max([i for i in range(0, 8) if 'Unnamed: {0}'.format(i) in page_df.columns]) - 1
            for row_index in range(0, len(page_df)):
                if type(page_df.iloc[row_index]['Unnamed: {0}'.format(mixed_unnamed_index)]) == str and \
                        (page_df.iloc[row_index]['Unnamed: {0}'.format(mixed_unnamed_index)][-1] == '%'
                         or page_df.iloc[row_index]['Unnamed: {0}'.format(mixed_unnamed_index)] == "Sub Total"):

                    page_df['Your Share %'][row_index] = page_df.iloc[row_index]['Unnamed: {0}'.format(mixed_unnamed_index)]
                    page_df['Unnamed: {0}'.format(mixed_unnamed_index)][row_index] = None

        blocks = []

        curr_block_is_open = False  # tells us if we've started a new block but didn't end it

        for i in range(1, len(page_df)):

            if (type(page_df['Your Share %'][i]) == str) and (page_df['Your Share %'][i][-1] == '%'):
                # we are starting a new block
                curr_block_start_idx = i
                curr_block_is_open = True

            if page_df['Your Share %'][i] == 'Sub Total':
                # we are ending a block
                blocks.append(page_df.iloc[curr_block_start_idx: i])
                curr_block_is_open = False

        if curr_block_is_open:
            # cutting of the last line (which contains only the page number),
            # creating the block and adding it to the list

            blocks.append(page_df.iloc[curr_block_start_idx: len(page_df)-1])

        return blocks

    def parse_block(self, block, are_separated1, are_separated2, page_number, work_details):
        """

        :param total_df: a pandas.df that contains the total data extracted till now from the pdf file.
        :param block: a pandas.df of a song.
        :param are_separated1: a boolean that == True <---> 'Work No' and 'Usage and Territory' have SEPARATED columns
        :param are_separated2: a boolean that == True <---> 'IP1' and 'IP2' have SEPARATED columns (at the given df)
        :param page_number: the page number of the block we're parsing. used mostly for debugging.
        :param work_details: a dictionary that contains the 'Member Name', 'CAE Number' and 'Distribution Number'
                             of the current page.
        :return: the block at the desired format (as a pandas.df).
        """

        result = pd.DataFrame(columns=['Work Title',
                                       'ISWC',
                                       'Usage Narrative',
                                       'IP1',
                                       'IP2',
                                       'IP3',
                                       'IP4',
                                       'Perf Start Date',
                                       'Perf End Date',
                                       'Production',
                                       'Share',
                                       'Number of Perfs',
                                       'Amount (performance revenue)',
                                       'Member Name',
                                       'CAE Number',
                                       'Distribution (posted)'])

        block = block.reset_index()
        work_title = block['Work Title'][0]
        your_share_percent = block['Your Share %'][0]

        if are_separated1:
            # 'Work No' and 'Usage and Territory' are -- separated --

            if are_separated2:
                # 'IP1' and 'IP2' are -- separated --
                ip1 = block['IP1'][0]
                ip2 = block['IP2'][0]
                self.add_name(ip1)
                self.add_name(ip2)
            else:
                # 'IP1' and 'IP2' are -- united --
                ip1, ip2 = self.extract_ips(block['IP1 IP2'][0])
        else:
            # 'Work No' and 'Usage and Territory' are -- united --

            if are_separated2:
                ip1 = block['IP1'][0]
                ip2 = block['IP2'][0]
                self.add_name(ip1)
                self.add_name(ip2)
            else:
                # 'IP1' and 'IP2' are -- united --
                ip1, ip2 = self.extract_ips(block['IP1 IP2'][0])

        ip3 = block['IP3'][0]
        ip4 = block['IP4'][0]

        self.add_name(ip3)
        self.add_name(ip4)

        for i in range(1, len(block)):

            if are_separated1:
                # 'Work No' and 'Usage and Territory' are -- separated --

                work_no = block['Work Title'][i]
                usage_and_territory = block['Unnamed: 0'][i]

                # they may have different column names, but all the data is at work_no.
                # we're checking that option and if it's the case, we're separating them

                if len(work_no.split()) == 2:
                    work_no, usage_and_territory = work_no.split()
            else:
                # 'Work No' and 'Usage and Territory' are -- united --
                work_no, usage_and_territory = block['Work Title'][i].split(' ', maxsplit=1)

            if 'Unnamed: 7' in block.columns:
                royalty = block['Unnamed: 7'][i]
                production = block['Unnamed: 6'][i]
            else:
                if 'Unnamed: 6' in block.columns:
                    royalty = block['Unnamed: 6'][i]
                    production = block['Unnamed: 5'][i]
                else:
                    royalty = block['Unnamed: 5'][i]
                    production = block['Unnamed: 4'][i]

            performances = block['Your Share %'][i]

            period = block['IP3'][i]

            line = pd.DataFrame({'Work Title': [work_title],
                                 'ISWC': [work_no],
                                 'Usage Narrative': [usage_and_territory],
                                 'IP1': [ip1],
                                 'IP2': [ip2],
                                 'IP3': [ip3],
                                 'Perf Start Date': [None if pd.isna(period) else period.split('-')[0]],
                                 'Perf End Date': [None if pd.isna(period) else period.split('-')[1]],
                                 'IP4': [ip4],
                                 'Production': [production],
                                 'Share': [your_share_percent],
                                 'Number of Perfs': [performances],
                                 'Amount (performance revenue)': [royalty],
                                 'Member Name': [work_details['Member Name']],
                                 'CAE Number': [work_details['CAE Number']],
                                 'Distribution (posted)': [work_details['Distribution Number']]})

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

        ips = ips.strip()

        maximum_min_overlap = -1
        ip1, ip2 = None, None

        for name in self.extracted_names:
            if ips.find(name) != -1:
                # the name was found at the ip's string

                splitting_idx = ips.find(name)

                if splitting_idx > 0:
                    extracted_ip1, extracted_ip2 = ips[:splitting_idx-1], ips[splitting_idx:]
                else:
                    extracted_ip1, extracted_ip2 = ips[:len(name)], ips[len(name) + 1:]

                if min(len(extracted_ip1), len(extracted_ip2)) > maximum_min_overlap:
                    ip1, ip2 = extracted_ip1, extracted_ip2
                    maximum_min_overlap = min(len(ip1), len(ip2))

        if not (ip1 is None):
            self.add_name(ip1)
            self.add_name(ip2)

            return ip1, ip2

        # else: we didn't found an overlapping and couldn't extract the ip's from the united string

        raise ParsingError("Couldn't separate between ip1 and ip2")

    def add_name(self, name):

        """

        :param name: a name to append to our list of names we've extracted from the file so far.
        :return: nothing. adds the name with some permutations (as they may appear at the file) to our list
                 of known names.
        """

        if pd.isna(name):
            return

        if not type(name) == str or name == '':
            return

        self.extracted_names.add(name)

        if len(name.split()) == 2:
            # it may be a < first name, family name >
            self.extracted_names.add(' '.join(list(reversed(name.split()))))

        if len(name.split()) == 3:
            first, second, third = name.split()
            self.extracted_names.add(third + ' ' + first + ' ' + second)

    def force_extract_pages_names(self, pages_numbers):

        """
        :param pages_numbers: a list of page numbers that we didn't suucceed in extracting from them the ip's.
        :return: none. adds the extracted names for the 'extracted_names' set.
        """

        dfs = tabula.read_pdf(self.pdf_filepath,
                              pages=pages_numbers,
                              area=(68, 17.28, 555.84, 818.64),
                              guess=False,
                              columns=(15.84, 136.8, 260.64, 380.16, 504, 681.12),
                              silent=True)

        for page_df in dfs:
            for row_index in range(0, len(page_df) - 1):
                row = page_df.iloc[row_index]
                if row['Work Title'] == 'Work No Usage & Territory' or pd.isna(row['Work Title']):
                    self.add_name(page_df.iloc[row_index+1]['IP1'])
                    self.add_name(page_df.iloc[row_index+1]['IP2'])
                    self.add_name(page_df.iloc[row_index+1]['IP3'])
                    self.add_name(page_df.iloc[row_index+1]['IP4'])


class ParsingError(Exception):

    def __init__(self, message):
        super().__init__(self, message)


if __name__ == "__main__":

    filepath = sys.argv[1]
    parser = PRSParser(pdf_filepath=filepath)
    parser.parse()
    parser.save_result(filepath + '_extracted.csv')

