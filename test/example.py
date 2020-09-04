import xlwings as xw
import sys

s = sys.argv[1]
if len(sys.argv) > 3:
    change_f = sys.argv[2]

book_a = xw.Book('test/a.xlsx')
book_a.sheets[0].range((1, 2)).value = s

if len(sys.argv) > 3 and change_f.split():
    book_a.sheets[0].range(1,3).formula = '=SUM(A1:B1)+{}'.format(s)
    
book_a.save()
keys = xw.apps.keys()
for key in keys:
    xw.apps[key].kill()
