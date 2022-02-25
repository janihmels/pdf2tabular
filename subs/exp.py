
from WixenParser import WixenParser
import os

if __name__ == "__main__":

    pdfs_folder = '../Jeremy R Wixen PDFs'
    csvs_folder = '../csvs'

    for filename in os.listdir(pdfs_folder):

        if filename.endswith('.pdf'):

            input_file = f'{pdfs_folder}/{filename}'
            output_file = f"{csvs_folder}/{filename.split('.')[0] + '.csv'}"

            print(f"Parsing {input_file}")

            parser = WixenParser(pdf_filepath=input_file)
            parser.parse()
            parser.save_result(output_filepath=output_file)

            print(f"Parsed {input_file} to {output_file}")