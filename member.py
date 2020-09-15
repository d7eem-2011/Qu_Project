import xlwings as xw


try:
    wb = xw.Book('test2.xlsx')

except Exception:
    wb = xw.Book()
    wb.save('test2.xlsx')

ws1 = wb.sheets[0]

ws1.range('A1').value = [["col1", 'col2'], ['name', 'number']]

