import sys
import os
# from difflib import unified_diff

# import colorama
# from oletools.olevba3 import VBA_Parser
import xlwings as xw
from colorama import Fore, Back, Style, init, deinit

init(wrap=True, autoreset=True)

diffs = {}

if __name__ == '__main__':
    if not 3 == len(sys.argv):
        print('Unexpected number of arguments: ' + len(sys.argv))
        sys.exit(0)
    book_a_path, book_b_path = sys.argv[1], sys.argv[2]
    book_a = xw.Book(book_a_path)
    book_b = xw.Book(book_b_path)
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
            print(diff['a'])
            print(diff['b'])
            print(diff['diff'])
            # print('\n')
    deinit()

                    
