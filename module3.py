#load libraries
import os
from datetime import date
#import pyarabic.number
from docxtpl import DocxTemplate

from ketat import ketat_def

from tashkel import tashkel_def


def task_do(xlxfil,xlxForTsh,docfil,doctash,saction,datein):

    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))
    
    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))

    doc = DocxTemplate(docfil)



    if not os.path.exists(saction):
        os.makedirs(saction)

    
    doc.render(context1)
    today = str(date.today())
    OUTPUT = ' اعداد محضر'
    if not os.path.exists(saction+'/'+OUTPUT):
        os.makedirs(saction+'/'+OUTPUT)

    save_name = 'اعداد محضر {}-{} .docx'.format(today, saction)
    doc.save(saction+'/'+OUTPUT + '/' + save_name)

    context2 = {}
    context2.update(tashkel_def(xlxForTsh))



    tashkelFil = DocxTemplate(doctash)
    tashkelFil.render(context2)
    today = str(date.today())
    OUTPUT2 = 'اعداد التشكيل'
    if not os.path.exists(saction+'/'+OUTPUT2):
        os.makedirs(saction+'/'+OUTPUT2)

    save_name2 = 'اعداد التشكيل {}-{} .docx'.format(today, saction)
    tashkelFil.save(saction+'/'+OUTPUT2 + '/' + save_name2)


def task_do2(xlxfil, xlxForTsh, docfil, saction, datein):
    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))

    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))

    doc = DocxTemplate(docfil)

    if not os.path.exists(saction):
        os.makedirs(saction)

    doc.render(context1)
    today = str(date.today())
    OUTPUT = 'اعداد اعلان الجلسة'
    if not os.path.exists(saction + '/' + OUTPUT):
        os.makedirs(saction + '/' + OUTPUT)

    save_name = 'اعداد اعلان الجلسة {}-{} .docx'.format(today, saction)
    doc.save(saction + '/' + OUTPUT + '/' + save_name)

def task_do3(xlxfil, xlxForTsh, docfil, saction, datein):
    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))

    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))

    doc = DocxTemplate(docfil)

    if not os.path.exists(saction):
        os.makedirs(saction)

    doc.render(context1)
    today = str(date.today())
    OUTPUT = ' اعداد مذكرة خطة'
    if not os.path.exists(saction + '/' + OUTPUT):
        os.makedirs(saction + '/' + OUTPUT)

    save_name = 'اعداد مذكرة خطة {} {}.docx'.format(today, saction)
    doc.save(saction + '/' + OUTPUT + '/' + save_name)

def task_do4(xlxfil, xlxForTsh, docfil, saction, datein):
    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))

    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))

    doc = DocxTemplate(docfil)

    if not os.path.exists(saction):
        os.makedirs(saction)

    doc.render(context1)
    today = str(date.today())
    OUTPUT = 'ملخص الخطط والتشكيل'
    if not os.path.exists(saction + '/' + OUTPUT):
        os.makedirs(saction + '/' + OUTPUT)

    save_name = ' ملخص الخطط والتشكيل {} {} .docx'.format(today, saction)
    doc.save(saction + '/' + OUTPUT + '/' + save_name)

def task_do5(xlxfil, xlxForTsh, docfil, saction, datein):
    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))

    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))

    doc = DocxTemplate(docfil)

    if not os.path.exists(saction):
        os.makedirs(saction)

    doc.render(context1)
    today = str(date.today())
    OUTPUT = 'output اعداد اعلان الجلسة'
    if not os.path.exists(saction + '/' + OUTPUT):
        os.makedirs(saction + '/' + OUTPUT)

    save_name = '{} {} template output v3 .docx'.format(today, saction)
    doc.save(saction + '/' + OUTPUT + '/' + save_name)



# docfil = 'C:/Users/D7eem/Documents/Team_Project/word document/template v4.docx'
#
# xlxfil = 'C:/Users/D7eem/Documents/Team_Project/taskAssign/22 1111.xlsx'
#
# xlsxFile = 'C:/Users/D7eem/Documents/Team_Project/Excel Document/تشكيل الجلسة الثانية عشر.xlsx'
#
# doctash = "C:/Users/D7eem/Documents/Team_Project/word document/تشكيل - مذكرات عرض للمجلس.docx"
#
# task_do(xlxfil, xlsxFile, docfil, doctash)