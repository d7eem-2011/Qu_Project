#load libraries
import os
from datetime import date
#import pyarabic.number
from num2words import num2words
from docxtpl import DocxTemplate
import pandas as pd

from taskAssign.ketat import ketat_def

from taskAssign.tashkel import tashkel_def


def task_do(xlxfil,xlxForTsh,docfil,doctash):

    context1 = {}
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))
    
    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))

    doc = DocxTemplate(docfil)


    
    doc.render(context1)
    today = str(date.today())
    OUTPUT = 'output اعداد محضر'
    if not os.path.exists(OUTPUT):
        os.makedirs(OUTPUT)

    save_name = today + 'template output v3 .docx'
    doc.save(OUTPUT + '/' + save_name)



    context2 = {}
    context2.update(tashkel_def(xlxForTsh))



    tashkelFil = DocxTemplate(doctash)
    tashkelFil.render(context2)
    today = str(date.today())
    OUTPUT2 = 'output اعداد التشكيل'
    if not os.path.exists(OUTPUT2):
        os.makedirs(OUTPUT2)

    save_name2 = today + 'template output v3 .docx'
    tashkelFil.save(OUTPUT2 + '/' + save_name2)










# docfil = 'C:/Users/D7eem/Documents/Team_Project/word document/template v4.docx'
#
# xlxfil = 'C:/Users/D7eem/Documents/Team_Project/taskAssign/22 1111.xlsx'
#
# xlsxFile = 'C:/Users/D7eem/Documents/Team_Project/Excel Document/تشكيل الجلسة الثانية عشر.xlsx'
#
# doctash = "C:/Users/D7eem/Documents/Team_Project/word document/تشكيل - مذكرات عرض للمجلس.docx"
#
# task_do(xlxfil, xlsxFile, docfil, doctash)