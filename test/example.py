import xlwings as xw
import sys

s = sys.argv[1]

print('write {} to a.xlsx in cell (1,2)'.format(s))

book_a = xw.Book('test/a.xlsx')
book_a.sheets[0].range((1,2)).value = s
book_a.save()
pid = xw.apps.keys()[0]
xw.apps[pid].kill()