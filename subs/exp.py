
from WixenParser import WixenParser
import os

if __name__ == "__main__":

    pdfs_folder = '/mnt/disks/birddata2/birddata//621d093388b8063d867e3c30/Jeremy R Wixen PDFs-20220223T222936Z-001/Jeremy R Wixen PDFs'
    csvs_folder = '../csvs'

    filename = '2018 Q1 WIXEN.pdf'

    input_file = f'{pdfs_folder}/{filename}'
    output_file = f"{csvs_folder}/{filename.split('.')[0] + '.csv'}"

    print(f"Parsing {input_file}")

    parser = WixenParser(pdf_filepath=input_file)
    parser.parse()
    parser.save_result(output_filepath=output_file)

    print(f"Parsed {input_file} to {output_file}")

    for filename in os.listdir(pdfs_folder):

        if filename.endswith('.pdf'):

            input_file = f'{pdfs_folder}/{filename}'
            output_file = f"{csvs_folder}/{filename.split('.')[0] + '.csv'}"

            print(f"Parsing {input_file}")

            parser = WixenParser(pdf_filepath=input_file)
            parser.parse()
            parser.save_result(output_filepath=output_file)

            print(f"Parsed {input_file} to {output_file}")