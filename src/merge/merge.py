from Diff.diff import make_diff
import sys
import os
import shutil
import xlwings as xw
from xlwings.utils import rgb_to_int

backgroundRGB = (255, 217, 179)
red_RGB = (255, 0, 0)
green_RGB = (0, 255, 0)


def make_merge(workbook_a, workbook_b):

    diffs = make_diff(workbook_a, workbook_b)
    if not diffs:
        print("not diffs")
        return
    book_a_path = os.path.abspath(
        workbook_a) if workbook_a != 'nul' and workbook_a != '/dev/null' else None
    book_b_path = os.path.abspath(
        workbook_b) if workbook_b != 'nul' and workbook_b != '/dev/null' else None
    book_a = xw.Book(book_a_path) if book_a_path else None
    book_b = xw.Book(book_b_path) if book_b_path else None
    print("merge start")
    sheets = []

    for sht in book_a.sheets:
        sheets.append(sht.name)
    for sht_name in sheets:
        print("merge sheet: {}".format(sht_name))
        if not book_b.sheets[sht.name]:
            # 添加sheet
            book_b.sheets.add(sht.name)
        sheet_a = book_a.sheets[sht.name]
        sheet_b = book_b.sheets[sht.name]

        for diff in diffs[sht.name]:
            sheet_a.range(diff['address']).color = red_RGB
            sheet_a.range(diff['address']).value = '<<<<<<<\n{}\n=======\n{}\n>>>>>>>'.format(
                diff['diff'][0], diff['diff'][1])
            
    book_a.save()
    keys = xw.apps.keys()
    for key in keys:
        xw.apps[key].kill()


if __name__ == '__main__':
    # print(sys.path)
    print('start merge')
    print(sys.argv)
    file_o, file_a, file_b = sys.argv[1:]
    shutil.copyfile(file_o, 'temp_o.xlsx')
    shutil.copyfile(file_a, 'temp_a.xlsx')
    shutil.copyfile(file_b, 'temp_b.xlsx')
    
    # make_merge(sys.argv[1], sys.argv[2])
    print("Conflict resolved!")

