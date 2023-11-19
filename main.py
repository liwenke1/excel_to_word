import os
import argparse
from loader import ExcelLoader
from parse import WordParser


def parse_options():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i',
                        '--input',
                        action='store',
                        default='',
                        type=str,
                        help='excel file path')
    parser.add_argument('-o',
                        '--output_dir',
                        action='store',
                        default='',
                        type=str,
                        help='word file output path dir')
    parser.add_argument('-n',
                        '--name_index',
                        action='store',
                        default=0,
                        type=int,
                        help='extract word file name from column index number')
    args = parser.parse_args()
    if len(args.input) == 0 or len(args.output_dir) == 0:
        raise ValueError("args `input` or `output_dir` is not defined")

    return args


def main():
    args = parse_options()
    excel_loader = ExcelLoader(excel_path=args.input)
    for row in excel_loader.data:
        word_parser = WordParser(header=excel_loader.header, source=row)

        file_name = row[excel_loader.header[args.name_index]]
        word_parser.dump(os.path.join(args.output_dir, f"{file_name}.docx"))


if __name__ == "__main__":
    main()
