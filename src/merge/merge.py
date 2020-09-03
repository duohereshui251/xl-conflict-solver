from Diff.diff import make_diff
import sys
import os
import shutil
import xlwings as xw
from xlwings.utils import rgb_to_int

backgroundRGB = (255, 217, 179)
red_RGB = (255, 0, 0)
green_RGB = (0, 255, 0)

isConflict = False


def make_merge(workbook_o, workbook_a, workbook_b):

    diffs_oa = make_diff(workbook_a, workbook_o, False)
    diffs_ob = make_diff(workbook_b, workbook_o, False)

    book_a_path = os.path.abspath(
        workbook_a) if workbook_a != 'nul' and workbook_a != '/dev/null' else None
    book_b_path = os.path.abspath(
        workbook_b) if workbook_b != 'nul' and workbook_b != '/dev/null' else None
    book_a = xw.Book(book_a_path) if book_a_path else None
    book_b = xw.Book(book_b_path) if book_b_path else None

    sheets = []

    for sht in book_a.sheets:
        sheets.append(sht.name)

    for sht_name in sheets:
        print("merge sheet: {}".format(sht_name))
        if not book_b.sheets[sht_name]:
            # 添加sheet
            book_b.sheets.add(sht_name)

        sheet_a = book_a.sheets[sht_name]
        sheet_b = book_b.sheets[sht_name]

        for addr, _ in diffs_ob[sht_name].items():
            # 同一单元格，本地和他人都改了
            if addr in diffs_oa[sht_name]:
                # 标记有冲突
                isConflict = True
                sheet_a.range(addr).color = red_RGB
                sheet_a.range(addr).value = '<<<<<<< our change\n{}\n=======\n{}\n>>>>>>> their change'.format(
                    sheet_a.range(addr).value, sheet_b.range(addr).value)
            else:
                # 同一单元格，本地没改别人改了, 那么就采取别人的改动
                sheet_a.range(addr).value = sheet_b.range(addr).value

    book_a.save()
    keys = xw.apps.keys()
    for key in keys:
        xw.apps[key].kill()


if __name__ == '__main__':
    # print(sys.path)
    print('start merging {}'.format(sys.argv[4]))
    print(sys.argv)
    file_o, file_a, file_b, filename = sys.argv[1:5]
    copy_o = 'temp_o.xlsx'
    copy_a = 'temp_a.xlsx'
    copy_b = 'temp_b.xlsx'
    shutil.copy(file_o, copy_o)
    shutil.copy(file_a, copy_a)
    shutil.copy(file_b, copy_b)
    make_merge(file_o, file_a, file_b)
    os.system('cat {} > {}'.format(file_a, file_a))
    os.system('rm {} {} {}'.format(file_a, file_a, file_b))
    if not isConflict:
        print("Conflict resolved!")
        exit(0)
    else:
        exit(1)
