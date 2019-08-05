#!/usr/bin/env python

import csv
import os
import argparse
# import unidecode

from openpyxl import Workbook, load_workbook

def convert_csv_to_xlsx(file_in):
    """Convert csv to xlsx"""
    if not file_in.endswith('.csv'):
        print('File in must be CSV')
        exit()

    file_out = file_in[:-4] + '.xlsx'

    with open(file_in, 'r') as file_in:
        csv_in = csv.reader(file_in)

        wb = Workbook()
        ws = wb.active

        for line in csv_in:
            # print(type(line))
            # print(line)
            lineout = []
            for col in line:
                if isinstance(col, str):
                    lineout.append(
                        col.encode('unicode_escape').decode('utf-8')
                    )
                else:
                    lineout.append(col)
            ws.append(lineout)

        wb.save(file_out)


# def format_string(s):
#     """Convert unicode string to ascii"""
#     if isinstance(s, unicode):
#         try:
#             s.encode('ascii')
#             return s
#         except:
#             return unidecode(s)
#     else:
#         return s


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert CSV to XLSX')
    parser.add_argument('in_files', metavar='I', type=str, nargs='+')
    args = parser.parse_args()

    for file in args.in_files:
        convert_csv_to_xlsx(file)
