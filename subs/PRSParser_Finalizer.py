import PyPDF2
import pandas as pd
import tabula
import re

class PRSParser_Finalizer:

    cnt = 0

    def __init__(self, pdf_filepath, starting_page):

        self.pdf_filepath = pdf_filepath
        self.starting_page = starting_page

        file = open(self.pdf_filepath, 'rb')
        self.end_page = PyPDF2.PdfFileReader(file).numPages - 1

        self.parsed_pages = []

        self.names = set()
        self.init_names()  # extracts all the names from the relevant pages (between starting_page and end_page)

    def init_names(self):
        """
        initializes the field self.names of the instance, by extracting all of the names that appear at the file.
        """

        pass

    def parse(self):

        for page_number in range(self.starting_page, self.end_page):
            self.parse_page(page_number)

    def get_result(self):

        if len(self.parsed_pages) == 0:
            print("call parse() before get_result()")
            exit(1)

        return pd.concat(self.parsed_pages, ignore_index=True)

    def parse_page(self, page_number):

        page_df = tabula.read_pdf(self.pdf_filepath,
                                  area=(77, 14, 570, 815),
                                  pages=[page_number],
                                  guess=False,
                                  columns=(14, 56.16, 260, 408, 495, 535, 628, 707, 764, 817))[0]

        blocks = PRSParser_Finalizer.page_to_blocks(page_df)

        parsed_blocks = []

        for block in blocks:
            parsed_blocks.append(self.parse_block(block))

        page_parsed_df = pd.concat(parsed_blocks)

        self.parsed_pages.append(page_parsed_df)

    def parse_block(self, block):

        result = pd.DataFrame(columns=['Work Title',
                                       'ISWC',
                                       'Usage Narrative',
                                       'IP1',
                                       'IP2',
                                       'IP3',
                                       'Perf Start Date',
                                       'Perf End Date',
                                       'IP4',
                                       'Production',
                                       'Old Share',
                                       'New Share',
                                       'Number of Perfs',
                                       'Amount (performance revenue)',
                                       'Member Name',
                                       'CAE Number',
                                       'Distribution (posted)'])

        work_title, ip1 = self.extract_names(block['Work Title'][0] + block['IP1'][0])
        ip2, ip3 = self.extract_names(block['IP2 IP3'][0] + block['Unnamed: 1'][0])

        for i in range(1, len(block) // 2 + 1):
            row1 = block.iloc[2*i - 1]
            row2 = block.iloc[2*i]

            old_share = row1['IP1'].split()[0]
            new_share = row1['IP1'].split()[4]

            work_no = row2['Work Title']
            usage_n_territory = row2['IP1']
            reason = row1['IP2 IP3']

            performances = row2['Unnamed: 4']
            royalty = row2['Unnamed: 5']

            # block.to_csv('sample_block.csv')

            # if PRSParser_Finalizer.cnt == 12:
            #     print("Here")
            #     exit(1)
            #
            # PRSParser_Finalizer.cnt += 1

            if not pd.isna(row2['IP2 IP3']) and (len((row2['IP2 IP3'] + row2['Unnamed: 1']).split('-')) == 2):
                period_start, period_end = [d.strip() for d in (row2['IP2 IP3'] + row2['Unnamed: 1']).split('-')]
            elif len(row2['Unnamed: 1'].split('-')) == 2:
                period_start, period_end = [d.strip() for d in row2['Unnamed: 1'].split('-')]
            else:
                period_start, period_end = None, None

            distribution_number = row1['Unnamed: 3'].split('-')[1].strip()

            line = pd.DataFrame({'Work Title': [work_title],
                                 'ISWC': [work_no],
                                 'Usage Narrative': [usage_n_territory],
                                 'IP1': [ip1],
                                 'IP2': [ip2],
                                 'IP3': [ip3],
                                 'Perf Start Date': period_start,
                                 'Perf End Date': period_end,
                                 'IP4': None,
                                 'Production': None,
                                 'Old Share': [old_share],
                                 'New Share': [new_share],
                                 'Number of Perfs': [performances],
                                 'Amount (performance revenue)': [royalty],
                                 'Member Name': None,
                                 'CAE Number': None,
                                 'Distribution (posted)': distribution_number})

            result = pd.concat([result, line], ignore_index=True, axis=0)

        result.to_csv('sample_block_extraction.csv')

        return result

    def extract_names(self, names_string):

        return names_string, names_string

    @staticmethod
    def page_to_blocks(page_df):

        """
        :param page_df: a data frame of a page in the 'old share, new share' format.
        :return: a list of df, each corresponds to a block of a song.
        """

        def token_type(token):

            if token is None or pd.isna(token):
                return 'NONE'

            if ((not (re.match('[T][0-9]+', token) is None)) and re.match('[T][0-9]+', token).span() == (0, len(token))) or \
                    ((not (re.match('[0-9]+', token) is None)) and re.match('[0-9]+', token).span() == (0, len(token))):

                return 'WORK NO'

            # else:

            if token == 'Old Share -':
                return 'OLD SHARE'

            # else:

            return 'SONG NAME'

        # --------------- clean page's dataframe ---------------

        page_df = page_df.drop([0], axis=0)
        page_df = page_df.drop(['Unnamed: 0'], axis=1)

        rows_to_drop = []

        for i in range(1, len(page_df)):

            if page_df['Unnamed: 4'][i] == 'Sub Total':
                rows_to_drop += [i]

            if type(page_df['Unnamed: 3'][i]) == str and type(page_df['Unnamed: 4'][i]) == str and \
                (page_df['Unnamed: 3'][i] + page_df['Unnamed: 4'][i] == 'Sub Total'):
                rows_to_drop += [i]

        page_df = page_df.drop(rows_to_drop, axis=0)
        page_df = page_df.reset_index(drop=True)
        page_df = page_df.drop([len(page_df) - 1], axis=0)

        page_df.to_csv('sample_extraction_n_cleaning.csv')

        # --------------- divide it into blocks ---------------

        blocks = []

        for i in range(0, len(page_df)):
            curr_row_header = page_df['Work Title'][i]

            if token_type(curr_row_header) == 'SONG NAME':
                if i != 0:
                    blocks += [page_df.iloc[curr_block_start_idx: i].reset_index(drop=True)]
                curr_block_start_idx = i

        blocks += [page_df.iloc[curr_block_start_idx: len(page_df)].reset_index(drop=True)]

        return blocks


