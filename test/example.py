import xlwings as xw
import sys

s = sys.argv[1]
col = 2
if len(sys.argv) > 2:
    option = sys.argv[2]

if option == '-f':
    print('change formula')
    book_a.sheets[0].range(1,3).formula = '=SUM(A1:B1)+{}'.format(s)
elif option == '--col':
    print('change col {}'.format(col))
    col = int(sys.argv[3])

book_a = xw.Book('test/a.xlsx')
book_a.sheets[0].range((1, col)).value = s 
book_a.save()

keys = xw.apps.keys()
for key in keys:
    xw.apps[key].kill()
