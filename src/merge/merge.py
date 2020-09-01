from Diff.diff import make_diff
import sys, os
import xlwings as xw

def make_merge(args = None):
    if not args: args = sys.argv

    if not 8 <= len(args) <= 9:
        print('Unexpected number of arguments: {0}'.format(len(args)))
        sys.exit(0)
    # 参数为8个
    if len(args) == 8:
        _, workbook_name, workbook_b, _, _, workbook_a, _ , _ = args
        numlines = 3

    # 参数为9个
    if len(args) == 9:
        _, numlines, workbook_name, workbook_b, _, _, workbook_a, _, _ = args
        numlines = int(numlines)

    diffs = make_diff()

    book_a_path = os.path.abspath(workbook_a) if workbook_a != 'nul' and workbook_a != '/dev/null' else None
    book_b_path = os.path.abspath(workbook_b) if workbook_b != 'nul' and workbook_b != '/dev/null' else None
    book_a = xw.Book(book_a_path) if book_a_path else None
    book_b = xw.Book(book_b_path) if book_b_path else None

    sheets = []
    for sht in book_a.sheets:
        sheets.append(sht.name)
    pass

if __name__ == '__main__':
    print(sys.argv)
    print("Conflict resolved!")
    # make_merge()



