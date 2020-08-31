import sys
import os
# from difflib import unified_diff

# from oletools.olevba3 import VBA_Parser

# excel处理工具
import xlwings as xw 
# 命令行输出显示颜色
from colorama import Fore, Back, Style, init, deinit 

# 主要针对windows powershell颜色无法显示的问题
# init(wrap=True, autoreset=True)

diffs = {}

if __name__ == '__main__':
    if not 8 <= len(sys.argv) <= 9:
        print('Unexpected number of arguments: {0}'.format(len(sys.argv)))
        print(sys.argv)
        sys.exit(0)
    # 参数为8个
    if len(sys.argv) == 8:
        _, workbook_name, workbook_b, _, _, workbook_a, _ , _ = sys.argv
        numlines = 3

    # 参数为9个
    if len(sys.argv) == 9:
        _, numlines, workbook_name, workbook_b, _, _, workbook_a, _, _ = sys.argv
        numlines = int(numlines)
    

    book_a_path = os.path.abspath(workbook_a) if workbook_a != 'nul' and workbook_a != '/dev/null' else None
    book_b_path = os.path.abspath(workbook_b) if workbook_b != 'nul' and workbook_b != '/dev/null' else None
    book_a = xw.Book(book_a_path) if book_a_path else None
    book_b = xw.Book(book_b_path) if book_b_path else None
    sheets = []
    for sht in book_a.sheets:
        sheets.append(sht.name)
    for sht_name in sheets:
        if not book_b.sheets[sht_name]:
            # todo 记录表的不同
            continue
        diffs[sht_name] = []
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
                if not sheet_a.range((row,col)).value and sheet_b.range((row,col)).value:
                    # a在b的基础上删除了
                    diffs[sht_name].append({
                        'a':'',
                        'b':'{}--- b/{}/{}'.format(Fore.WHITE,book_b.name,sheet_b.range((row,col)).address),
                        'diff': '{}-{}'.format(Fore.RED,sheet_b.range((row,col)).value)
                    })
                    #  a 在 b的基础上增加了
                elif sheet_a.range((row,col)).value and not sheet_b.range((row,col)).value:
                    diffs[sht_name].append({
                        'a':'{}+++ a/{}/{}'.format(Fore.WHITE,book_a.name,sheet_a.range((row,col)).address ),
                        'b':'',
                        'diff': '{}+{}'.format(Fore.GREEN,sheet_a.range((row,col)).value)
                    })
                elif sheet_a.range((row,col)).value != sheet_b.range((row,col)).value:
                    diffs[sht_name].append({
                        'a':'{}+++ a/{}/{}'.format(Fore.WHITE,book_a.name,sheet_a.range((row,col)).address ),
                        'b':'{}--- b/{}/{}'.format(Fore.WHITE,book_b.name,sheet_b.range((row,col)).address),
                        'diff': '{}+{}\n{}-{}'.format(Fore.GREEN , sheet_a.range((row,col)).value,  Fore.RED , sheet_b.range((row,col)).value)
                    })
    book_a.close()
    book_b.close()

    for k, v in diffs.items():
        print('in sheet ' + k)
        for diff in v:
            if diff['a'] :  print(diff['a'])
            if diff['b']:   print(diff['b'])
            print(diff['diff'])
            # print('\n')
    # deinit()

                    
