import pandas as pd
import openpyxl
from pandas import Series
from openpyxl import Workbook, load_workbook

from openpyxl.utils.dataframe import dataframe_to_rows

try:
    wb = load_workbook('test2.xlsx')

except Exception:
    wb = Workbook()
    wb.save('test2.xlsx')



df = pd.DataFrame({'Data': ['name']},{'name': ['name2']},{'number': ['name3']})


# Activate worksheet to write dataframe
active = wb.worksheets[0]



# Write dataframe to active worksheet
for x in dataframe_to_rows(df, index=False, header=False):
    active.append(x)
    print(x)
# Save workbook to write
wb.save('test2.xlsx')


