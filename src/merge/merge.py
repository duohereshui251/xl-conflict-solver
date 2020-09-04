from Diff.diff import make_diff, DiffType
import sys
import os
import shutil
import xlwings as xw
from xlwings.utils import rgb_to_int

backgroundRGB = (255, 217, 179)
red_RGB = (236, 173, 158)
green_RGB = (160, 238, 225)

isConflict = False


def make_merge(workbook_o, workbook_a, workbook_b):

    diffs_oa = make_diff(workbook_a, workbook_o, True)
    diffs_ob = make_diff(workbook_b, workbook_o, True)

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

        for addr, diff in diffs_ob[sht_name].items():
            # 同一单元格，本地和他人都改了
            if addr in diffs_oa[sht_name]:
                # 标记有冲突
                print('{} has conflict in cell {}'.format(sht_name, addr))
                global isConflict
                isConflict = True
                # 值冲突为红色， 函数冲突为绿色
                print(diff['type'])
                print(diffs_oa[sht_name][addr]['type'])
                if diff['type'] == DiffType.formula or diffs_oa[sht_name][addr]['type'] == DiffType.formula:
                    print('formula conflict')
                    sheet_a.range(addr).formula = None
                    sheet_a.range(addr).color = green_RGB
                    sheet_a.range(addr).value = '<<<<<<< 表格函数冲突,请重新设置 our change\n{}\n=======\n{}\n>>>>>>> their change'.format(
                    diffs_oa[sht_name]['diff'][0], diffs_ob[sht_name]['diff'][0])
                else:
                    print('value conflict')
                    sheet_a.range(addr).color = red_RGB
                    sheet_a.range(addr).value = '<<<<<<< our change\n{}\n=======\n{}\n>>>>>>> their change'.format(
                        sheet_a.range(addr).value, sheet_b.range(addr).value)

            else:
                # 同一单元格，本地没改别人改了, 那么就采取别人的改动
                if diff['type'] == DiffType.value:
                    sheet_a.range(addr).value = sheet_b.range(addr).value
                else:
                    sheet_a.range(addr).formula = sheet_b.range(addr).formula

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
    make_merge(copy_o, copy_a, copy_b)
    os.system('cat {} > {}'.format(copy_a, file_a))
    os.system('rm {} {} {}'.format(copy_o, copy_a, copy_b))

    if isConflict:
        print("{} has conflict".format(filename))
        exit(-1)
    else:
        print('merge file {} success'.format(filename))
        exit(0)
