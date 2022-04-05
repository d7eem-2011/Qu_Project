#load libraries
import os
from datetime import date
import pyarabic.number as nb 
from docxtpl import DocxTemplate
from ketat import ketat_def
from tashkel import tashkel_def


def corctnum(x, y=True):
    if x + 1 == 1:
        return 'الاول'
    else:
        return nb.number2ordinal(x + 1, feminin=y)

def task_do(xlxfil,xlxForTsh,docfil,doctash,saction,datein,realPath = os.getcwd() ):




    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))

    # context1.update(len(ketat_def(xlxfil))+ketat_def(xlxfil))
    # print('its work Here /n'+str(ketat_def(xlxfil))+' /n '+str(tashkel_def(xlxForTsh)))
    theTotalRec = len(context1['تشكيل']) + len(context1['خطط'])

    try:


        for i in range(theTotalRec):
            rank = corctnum(i, y=False)
            if i < len(context1['تشكيل']):
                context1['تشكيل'][i].update({'الترتيب': rank})


            if i >= len(context1['تشكيل']):

                i = i - len(context1['تشكيل'])

                context1['خطط'][i].update({'الترتيب': rank})
    except Exception:
        Exception





    doc = DocxTemplate(docfil)



    if not os.path.exists(realPath+'/'+saction):
        os.makedirs(realPath+'/'+saction)

    
    doc.render(context1)
    today = str(date.today())
    OUTPUT = ' اعداد محضر'
    if not os.path.exists(realPath+'/'+saction+'/'+OUTPUT):
        os.makedirs(realPath+'/'+saction+'/'+OUTPUT)

    save_name = 'اعداد محضر {}-{} .docx'.format(today, saction)
    doc.save(realPath+'/'+saction+'/'+OUTPUT + '/' + save_name)

    context2 = {}
    context2.update(tashkel_def(xlxForTsh))



    tashkelFil = DocxTemplate(doctash)
    tashkelFil.render(context2)
    today = str(date.today())
    OUTPUT2 = 'اعداد التشكيل'
    if not os.path.exists(realPath+'/'+saction+'/'+OUTPUT2):
        os.makedirs(realPath+'/'+saction+'/'+OUTPUT2)

    save_name2 = 'اعداد التشكيل {}-{} .docx'.format(today, saction)
    tashkelFil.save(realPath+'/'+saction+'/'+OUTPUT2 + '/' + save_name2)


def task_do2(xlxfil, xlxForTsh, docfil, saction, datein, realPath = os.getcwd()):
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
    if not os.path.exists(realPath + '/' + saction + '/' + OUTPUT):
        os.makedirs(realPath + '/' + saction + '/' + OUTPUT)

    save_name = 'اعداد اعلان الجلسة {}-{} .docx'.format(today, saction)
    doc.save(realPath + '/' + saction + '/' + OUTPUT + '/' + save_name)

def task_do3(xlxfil, xlxForTsh, docfil, saction, datein, realPath = os.getcwd()):
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
    if not os.path.exists(realPath + '/' + saction + '/' + OUTPUT):
        os.makedirs(realPath + '/' + saction + '/' + OUTPUT)

    save_name = 'اعداد مذكرة خطة {} {}.docx'.format(today, saction)
    doc.save(realPath + '/' + saction + '/' + OUTPUT + '/' + save_name)
def task_do4(xlxfil, xlxForTsh, docfil, saction, datein, realPath = os.getcwd()):
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
    if not os.path.exists(realPath + '/' + saction + '/' + OUTPUT):
        os.makedirs(realPath + '/' + saction + '/' + OUTPUT)

    save_name = ' ملخص الخطط والتشكيل {} {} .docx'.format(today, saction)
    doc.save(realPath+'/'+saction+'/'+OUTPUT + '/' + save_name)

def task_do5(xlxfil, xlxForTsh, docfil, saction, datein, realPath = os.getcwd()):
    context1 = {}
    context1.update(con=saction)
    context1.update(date=datein)
    context1.update(ketat_def(xlxfil))

    context1.update(tashkel_def(xlxForTsh))



    doc = DocxTemplate(docfil)

    if not os.path.exists(saction):
        os.makedirs(saction)

    doc.render(context1)
    today = str(date.today())
    OUTPUT = 'output اعداد اعلان الجلسة'
    if not os.path.exists(realPath + '/' + saction + '/' + OUTPUT):
        os.makedirs(realPath + '/' + saction + '/' + OUTPUT)

    save_name = '{} {} template output v3 .docx'.format(today, saction)
    doc.save(realPath+'/'+saction+'/'+OUTPUT + '/' + save_name)



