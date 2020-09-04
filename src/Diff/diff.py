import sys
import os
# from difflib import unified_diff

# from oletools.olevba3 import VBA_Parser

# excel处理工具
import xlwings as xw 
# 命令行输出显示颜色
from colorama import Fore, Back, Style, init, deinit 
from enum import Enum

class DiffType(Enum):
    value = 1
    formula = 2

# 主要针对windows powershell颜色无法显示的问题


def print_diff(diffs):
    init(wrap=True, autoreset=True)
    for k, v in diffs.items():
        print('in sheet ' + k)
        for _,diff in v.items():
            if diff['a'] :  print('{}+++ a/{}/{}'.format(Fore.WHITE,diff['a'],diff['address']))
            if diff['b']:   print('{}--- b/{}/{}'.format(Fore.WHITE,diff['b'],diff['address']))
            if diff['diff'][0]: print('{}+{}'.format(Fore.GREEN,diff['diff'][0]))
            if diff['diff'][1]: print('{}-{}'.format(Fore.RED,diff['diff'][1]))
    deinit()

def make_diff(workbook_a =None, workbook_b= None, printOn = True):
    diffs = {}
    # print(sys.argv)
    if not workbook_a or not workbook_b: 
        args = sys.argv
        if not 8 <= len(args) <= 9:
            print('Unexpected number of arguments: {0}'.format(len(args)))
            sys.exit(0)
        # 参数为8个
        if len(args) == 8:
            workbook_b, workbook_a = args[2], args[5]
            # _, workbook_name, workbook_b, _, _, workbook_a, _ , _ = args
        # 参数为9个
        if len(args) == 9:
            workbook_b, workbook_a = args[3], args[6]
            # _, _, workbook_name, workbook_b, _, _, workbook_a, _, _ = args

    book_a_path = os.path.abspath(workbook_a) if workbook_a != 'nul' and workbook_a != '/dev/null' else None
    book_b_path = os.path.abspath(workbook_b) if workbook_b != 'nul' and workbook_b != '/dev/null' else None

    try:
        book_a = xw.Book(book_a_path) if book_a_path else None
    except BaseException:
        print('a 打开失败')

    try:
        book_b = xw.Book(book_b_path) if book_b_path else None
    except BaseException:
        print("b 打开失败")
        return
        
    sheets = []
    for sht in book_a.sheets:
        sheets.append(sht.name)
    for sht_name in sheets:
        if not book_b.sheets[sht_name]:
            # todo 记录表的不同
            continue
        diffs[sht_name] = {}
        sheet_a = book_a.sheets[sht_name]
        sheet_b = book_b.sheets[sht_name]
        rows = max(len(sheet_a.used_range.rows), len(sheet_b.used_range.rows))
        columns = max(len(sheet_a.used_range.columns),
                      len(sheet_b.used_range.columns))
        # print('[Debug] row:{0}, col:{1}'.format(rows, columns) )

        for row in range(1, rows+1):
            # 整行相同跳过
            if sheet_a.range((row, 1), (row, columns)).value == sheet_b.range((row, 1), (row, columns)).value:
                continue

            for col in range(1, columns+1):
                # print('[Debug] a/{0}'.format(sheet_a.range((row,col)).value))
                # print('[Debug] b/{0}'.format(sheet_b.range((row,col)).value))

                if sheet_a.range((row, col)).value == sheet_b.range((row, col)).value:
                    continue
                # a 是当前文件
                # b 是要对比的文件
                address = sheet_b.range((row,col)).address.replace('$', '')

                # 检测函数的差异
                if sheet_a.range((row, col)).formula != sheet_b.range((row, col)).formula:
                    print('formula different:\n{}\n{}'.format(sheet_a.range((row, col)).formula, sheet_b.range((row, col)).formula))
                    diffs[sht_name][address] = {
                        'type': DiffType.formula,
                        'address':address,
                        'a': book_a.name if sheet_a.range((row, col)).formula else '',
                        'b': book_b.name if sheet_b.range((row, col)).formula else '',
                        'diff': [sheet_a.range((row, col)).formula, sheet_b.range((row, col)).formula]
                    }
                    continue
                
                # 检测值的差异
                if sheet_a.range((row,col)).value != sheet_b.range((row,col)).value:
                    # a在b的基础上删除了
                    diffs[sht_name][address] = {
                        'type': DiffType.value,
                        'address':address,
                        'a':book_a.name if sheet_a.range((row,col)).value else '',
                        'b':book_b.name if sheet_b.range((row,col)).value else '',
                        'diff': [sheet_a.range((row,col)).value,sheet_b.range((row,col)).value]
                    }
                #     #  a在b的基础上增加了
                # elif sheet_a.range((row,col)).value and not sheet_b.range((row,col)).value:
                #     diffs[sht_name][address] = {
                #         'type': DiffType.value,
                #         'address': address,
                #         'a':book_a.name,
                #         'b':'',
                #         'diff': [sheet_a.range((row,col)).value,'']
                #     }
                # elif sheet_a.range((row,col)).value != sheet_b.range((row,col)).value:
                #     diffs[sht_name][address] = {
                #         'type': DiffType.value,
                #         'address': address,
                #         'a':book_a.name,
                #         'b':book_b.name,
                #         'diff': [sheet_a.range((row,col)).value,sheet_b.range((row,col)).value]
                #     }
    book_a.close()
    book_b.close()
    if printOn:
        print_diff(diffs)
    
    keys = xw.apps.keys()
    for key in keys:
        xw.apps[key].kill()
    return diffs

if __name__ == '__main__':
    make_diff()

                    
