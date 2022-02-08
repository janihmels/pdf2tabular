
import tabula
import pandas as pd
import PyPDF2
import sys


class CMGParser:

    def __init__(self, pdf_filepath):

        self.pdf_filepath = pdf_filepath

        pdf_file = open(self.pdf_filepath, "rb")
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        self.num_pages = read_pdf.getNumPages()

        self.parsed_pages = []

        self.result = None

    def parse(self, start_page=2, end_page=-1):

        """
        :param start_page: the page to start o parse from.
        :param end_page: the page to end the parsing at (we include also the 'end_page' at the parsing).

        :return: nothing. saves the result as a list of pandas.DataFrame's,
                each corresponding to a page in the document whose path is self.pdf_filepath.
        """

        if end_page == -1:
            end_page = self.num_pages

        for page_number in range(start_page, end_page + 1):
            self.parsed_pages.append(self.parse_page(page_number))
            if self.parsed_pages[-1].empty:
                # we are at the last page which contains no data, so we delete the last line of the
                # real last page (whose data is 'total earnings this statement')
                if len(self.parsed_pages) >= 2:
                    self.parsed_pages[-2] = self.parsed_pages[-2][:-1]

        self.result = pd.concat(self.parsed_pages)

    def save_result(self, output_filepath):

        if self.result is None:
            raise ParsingError("call parse() before you call save_result()")

        # else:

        self.result.to_csv(output_filepath)

    def parse_page(self, page_number):

        """

        :param page_number: the number of the page in self.pdf_filepath to parse.
        :return: the parsed pandas.DataFrame for the page.
        """

        if page_number == 1:
            df = tabula.read_pdf(self.pdf_filepath, pages='1', area=(270.08, 18, 555.84, 778))[0]
            df = df.iloc[:-1, :]  # delete the last line of the total earnings

            return df

        elif page_number == 2:
            return tabula.read_pdf(self.pdf_filepath, pages='2', area=(42.48, 18, 555.12, 778))[0]

        # else:

        dfs = tabula.read_pdf(self.pdf_filepath, pages=[page_number], area=(22, 18, 556.12, 778), guess=False)

        if len(dfs) == 0:
            return pd.DataFrame()

        df = dfs[0]

        if 'Unnamed: 0' in df.columns:
            # Then 'Song Number' and 'Song Title' columns have been merged
            numbers_and_titles = [u.split(' ', maxsplit=1) for u in df['Song Number Song Title']]

            numbers = [a[0] for a in numbers_and_titles]
            titles = [a[1] for a in numbers_and_titles]

            df = pd.DataFrame({'Song Number': numbers,
                               'Song Title': titles,
                               'Share': df['Share'],
                               'Net Earnings': df['Net Earnings']})

        if page_number == self.num_pages:
            # this is the last page in the pdf file, and we need to delete the last row from it
            df = df.iloc[:-1, :]

        return df


class ParsingError(Exception):

    def __init__(self, message):
        super().__init__(self, message)


if __name__ == "__main__":

    PDF_NAME = sys.argv[1]

    parser = CMGParser(pdf_filepath=PDF_NAME)
    parser.parse()

    parser.save_result(output_filepath=PDF_NAME + '_extracted.csv')






